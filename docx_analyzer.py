"""
Optimized DOCX Analyzer
High-performance DOCX document analysis with efficient text extraction
"""

import io
import re
import zipfile
from typing import Dict, Any, List, Optional, Union
from pathlib import Path
import logging
from xml.etree import ElementTree as ET
from functools import lru_cache

from .base_analyzer import BaseAnalyzer
from ..config import CACHE_SIZE

logger = logging.getLogger(__name__)


class OptimizedDocxAnalyzer(BaseAnalyzer):
    """
    Optimized DOCX analyzer with improvements:
    - Efficient XML parsing using ElementTree
    - Lazy text extraction
    - Regex compilation and caching
    - Memory-efficient processing
    """
    
    # Compile regex patterns once for better performance
    FIELD_PATTERNS = {
        'business_app_id': re.compile(r'Business\s+Application\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
        'enterprise_release_id': re.compile(r'Enterprise\s+Release\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
        'project_name': re.compile(r'Project\s+Name[:\s]*([^\n\r]+)', re.IGNORECASE),
        'task_id': re.compile(r'Task\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
    }
    
    # XML namespaces for DOCX
    DOCX_NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    def __init__(self, file_input: Union[str, Path, bytes], cache_enabled: bool = True):
        """Initialize with DOCX-specific optimizations"""
        super().__init__(file_input, cache_enabled)
        self._document_xml: Optional[str] = None
        self._footer_xml: Optional[str] = None
        self._extracted_data: Optional[Dict[str, Any]] = None
    
    def _get_file_content(self) -> bytes:
        """Get file content as bytes"""
        if isinstance(self.file_input, (str, Path)):
            with open(self.file_input, 'rb') as f:
                return f.read()
        elif hasattr(self.file_input, 'getvalue'):
            return self.file_input.getvalue()
        else:
            return self.file_input
    
    @lru_cache(maxsize=1)
    def _extract_xml_content(self) -> tuple:
        """Extract XML content from DOCX with caching"""
        try:
            content = self._get_file_content()
            
            with zipfile.ZipFile(io.BytesIO(content), 'r') as docx_zip:
                # Extract main document
                document_xml = None
                if 'word/document.xml' in docx_zip.namelist():
                    document_xml = docx_zip.read('word/document.xml').decode('utf-8')
                
                # Extract footer (if exists)
                footer_xml = None
                footer_files = [f for f in docx_zip.namelist() if f.startswith('word/footer')]
                if footer_files:
                    footer_xml = docx_zip.read(footer_files[0]).decode('utf-8')
                
                return document_xml, footer_xml
                
        except Exception as e:
            logger.error(f"Error extracting XML from DOCX: {e}")
            return None, None
    
    def _extract_text_from_xml(self, xml_content: str) -> str:
        """Extract text from XML efficiently"""
        if not xml_content:
            return ""
        
        try:
            # Parse XML
            root = ET.fromstring(xml_content)
            
            # Extract text using xpath-like approach
            text_elements = []
            for elem in root.iter():
                if elem.tag.endswith('}t'):  # Text elements
                    if elem.text:
                        text_elements.append(elem.text)
            
            return ' '.join(text_elements)
            
        except Exception as e:
            logger.error(f"Error extracting text from XML: {e}")
            return ""
    
    @lru_cache(maxsize=CACHE_SIZE)
    def _extract_field_cached(self, text: str, field_name: str) -> Optional[str]:
        """Cached field extraction"""
        pattern = self.FIELD_PATTERNS.get(field_name)
        if not pattern:
            return None
        
        match = pattern.search(text)
        return match.group(1).strip() if match else None
    
    def _analyze_first_page(self, document_text: str) -> Dict[str, Any]:
        """Analyze first page with optimized extraction"""
        # Take first portion of text (approximate first page)
        first_page_text = document_text[:2000]  # Adjust based on typical page size
        
        return {
            'business_app_id': self._extract_field_cached(first_page_text, 'business_app_id'),
            'enterprise_release_id': self._extract_field_cached(first_page_text, 'enterprise_release_id'),
            'project_name': self._extract_field_cached(first_page_text, 'project_name'),
            'task_id': self._extract_field_cached(first_page_text, 'task_id'),
        }
    
    def _analyze_footer(self, footer_text: str) -> Dict[str, Any]:
        """Analyze footer content"""
        if not footer_text:
            return {'project_name_in_footer': None}
        
        # Simple project name extraction from footer
        project_match = self._extract_field_cached(footer_text, 'project_name')
        
        return {
            'project_name_in_footer': project_match
        }
    
    def _check_table_of_contents(self, document_text: str) -> Dict[str, Any]:
        """Check for table of contents"""
        toc_patterns = [
            re.compile(r'table\s+of\s+contents', re.IGNORECASE),
            re.compile(r'contents', re.IGNORECASE),
        ]
        
        has_toc = any(pattern.search(document_text) for pattern in toc_patterns)
        
        # Check for specific sections
        has_scope_section = bool(re.search(r'3\.3[:\s]*in\s+scope', document_text, re.IGNORECASE))
        
        return {
            'has_table_of_contents': has_toc,
            'has_scope_section': has_scope_section
        }
    
    def _detect_embedded_excel(self, content: bytes) -> Dict[str, Any]:
        """Detect embedded Excel files efficiently"""
        try:
            with zipfile.ZipFile(io.BytesIO(content), 'r') as docx_zip:
                # Look for embedded objects
                embedded_files = [f for f in docx_zip.namelist() 
                                if f.startswith('word/embeddings/') and f.endswith('.xlsx')]
                
                excel_data = {
                    'has_embedded_excel': len(embedded_files) > 0,
                    'embedded_excel_count': len(embedded_files),
                    'embedded_files': embedded_files
                }
                
                # If Excel files found, analyze worksheets
                if embedded_files:
                    excel_data.update(self._analyze_embedded_excel(docx_zip, embedded_files[0]))
                
                return excel_data
                
        except Exception as e:
            logger.error(f"Error detecting embedded Excel: {e}")
            return {
                'has_embedded_excel': False,
                'embedded_excel_count': 0,
                'error': str(e)
            }
    
    def _analyze_embedded_excel(self, docx_zip: zipfile.ZipFile, excel_file: str) -> Dict[str, Any]:
        """Analyze embedded Excel file"""
        try:
            import pandas as pd
            
            excel_content = docx_zip.read(excel_file)
            
            with io.BytesIO(excel_content) as excel_buffer:
                # Get sheet names efficiently
                xl_file = pd.ExcelFile(excel_buffer)
                sheet_names = xl_file.sheet_names
                
                return {
                    'worksheet_names': sheet_names,
                    'worksheet_count': len(sheet_names),
                    'has_architecture_sheet': any('architecture' in name.lower() for name in sheet_names)
                }
                
        except Exception as e:
            logger.error(f"Error analyzing embedded Excel: {e}")
            return {
                'worksheet_names': [],
                'worksheet_count': 0,
                'has_architecture_sheet': False,
                'error': str(e)
            }
    
    def _check_milestones_section(self, document_text: str) -> Dict[str, Any]:
        """Check for milestones section"""
        # Look for section 12 milestones
        milestone_pattern = re.compile(r'12[:\.\s]+milestones?', re.IGNORECASE)
        has_milestones = bool(milestone_pattern.search(document_text))
        
        # Extract implementation dates if found
        implementation_dates = []
        if has_milestones:
            date_pattern = re.compile(r'implementation[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', re.IGNORECASE)
            dates = date_pattern.findall(document_text)
            implementation_dates = dates
        
        return {
            'has_milestones_section': has_milestones,
            'implementation_dates': implementation_dates
        }
    
    def analyze_content(self) -> Dict[str, Any]:
        """Main analysis method"""
        try:
            # Get file content
            content = self._get_file_content()
            
            # Extract XML content
            document_xml, footer_xml = self._extract_xml_content()
            
            if not document_xml:
                raise ValueError("Could not extract document content")
            
            # Extract text from XML
            document_text = self._extract_text_from_xml(document_xml)
            footer_text = self._extract_text_from_xml(footer_xml) if footer_xml else ""
            
            # Perform analysis
            result = {
                'first_page_data': self._analyze_first_page(document_text),
                'footer_data': self._analyze_footer(footer_text),
                'table_of_contents': self._check_table_of_contents(document_text),
                'embedded_excel': self._detect_embedded_excel(content),
                'milestones': self._check_milestones_section(document_text),
                'document_length': len(document_text),
                'has_content': len(document_text) > 0
            }
            
            return result
            
        except Exception as e:
            logger.error(f"Error in DOCX analysis: {e}")
            raise


# Backward compatibility alias
DocxAnalyzer = OptimizedDocxAnalyzer
