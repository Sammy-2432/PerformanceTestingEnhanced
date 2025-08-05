"""
Optimized PowerPoint Analyzer
High-performance PPTX document analysis with efficient slide processing
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


class OptimizedPowerPointAnalyzer(BaseAnalyzer):
    """
    Optimized PowerPoint analyzer with improvements:
    - Efficient XML parsing for slides
    - Lazy slide processing
    - Cached pattern matching
    - Memory-efficient handling of large presentations
    """
    
    # Compiled regex patterns for better performance
    METADATA_PATTERNS = {
        'business_app_id': re.compile(r'Business\s+Application\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
        'enterprise_release_id': re.compile(r'Enterprise\s+Release\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
        'project_name': re.compile(r'Project\s+Name[:\s]*([^\n\r]+)', re.IGNORECASE),
        'task_id': re.compile(r'Task\s+ID[:\s]*([A-Z0-9]+)', re.IGNORECASE),
    }
    
    PLT_STATUS_PATTERNS = [
        re.compile(r'PLT\s+status[:\s]*([^\n\r]+)', re.IGNORECASE),
        re.compile(r'Production\s+Live\s+Testing[:\s]*([^\n\r]+)', re.IGNORECASE),
    ]
    
    # PPTX XML namespaces
    PPTX_NAMESPACES = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    def __init__(self, file_input: Union[str, Path, bytes], cache_enabled: bool = True):
        """Initialize with PPTX-specific optimizations"""
        super().__init__(file_input, cache_enabled)
        self._slide_cache: Dict[int, str] = {}
        self._slide_count: Optional[int] = None
    
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
    def _get_slide_list(self) -> List[str]:
        """Get list of slide files with caching"""
        try:
            content = self._get_file_content()
            
            with zipfile.ZipFile(io.BytesIO(content), 'r') as pptx_zip:
                slide_files = [f for f in pptx_zip.namelist() 
                             if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
                
                # Sort slides by number
                slide_files.sort(key=lambda x: int(re.search(r'slide(\d+)', x).group(1)))
                return slide_files
                
        except Exception as e:
            logger.error(f"Error getting slide list: {e}")
            return []
    
    @lru_cache(maxsize=CACHE_SIZE)
    def _extract_slide_text_cached(self, slide_file: str) -> str:
        """Extract text from a slide with caching"""
        try:
            content = self._get_file_content()
            
            with zipfile.ZipFile(io.BytesIO(content), 'r') as pptx_zip:
                slide_xml = pptx_zip.read(slide_file).decode('utf-8')
                return self._extract_text_from_xml(slide_xml)
                
        except Exception as e:
            logger.error(f"Error extracting text from {slide_file}: {e}")
            return ""
    
    def _extract_text_from_xml(self, xml_content: str) -> str:
        """Extract text from PowerPoint XML efficiently"""
        if not xml_content:
            return ""
        
        try:
            root = ET.fromstring(xml_content)
            
            # Extract text from text elements
            text_elements = []
            for elem in root.iter():
                if elem.tag.endswith('}t'):  # Text elements
                    if elem.text:
                        text_elements.append(elem.text)
            
            return ' '.join(text_elements)
            
        except Exception as e:
            logger.error(f"Error extracting text from XML: {e}")
            return ""
    
    def _get_slide_text(self, slide_number: int) -> str:
        """Get text from specific slide"""
        slide_files = self._get_slide_list()
        
        if slide_number <= 0 or slide_number > len(slide_files):
            return ""
        
        slide_file = slide_files[slide_number - 1]
        return self._extract_slide_text_cached(slide_file)
    
    @lru_cache(maxsize=CACHE_SIZE)
    def _extract_field_cached(self, text: str, field_name: str) -> Optional[str]:
        """Cached field extraction"""
        pattern = self.METADATA_PATTERNS.get(field_name)
        if not pattern:
            return None
        
        match = pattern.search(text)
        return match.group(1).strip() if match else None
    
    def _analyze_first_slide(self) -> Dict[str, Any]:
        """Analyze first slide metadata"""
        first_slide_text = self._get_slide_text(1)
        
        if not first_slide_text:
            return {
                'business_app_id': None,
                'enterprise_release_id': None,
                'project_name': None,
                'task_id': None,
                'has_content': False
            }
        
        return {
            'business_app_id': self._extract_field_cached(first_slide_text, 'business_app_id'),
            'enterprise_release_id': self._extract_field_cached(first_slide_text, 'enterprise_release_id'),
            'project_name': self._extract_field_cached(first_slide_text, 'project_name'),
            'task_id': self._extract_field_cached(first_slide_text, 'task_id'),
            'has_content': len(first_slide_text) > 0
        }
    
    def _detect_embedded_excel(self) -> Dict[str, Any]:
        """Detect embedded Excel files in PPTX"""
        try:
            content = self._get_file_content()
            
            with zipfile.ZipFile(io.BytesIO(content), 'r') as pptx_zip:
                # Look for embedded Excel files
                embedded_files = [f for f in pptx_zip.namelist() 
                                if f.startswith('ppt/embeddings/') and 
                                (f.endswith('.xlsx') or f.endswith('.xls'))]
                
                excel_data = {
                    'has_embedded_excel': len(embedded_files) > 0,
                    'embedded_excel_count': len(embedded_files),
                    'embedded_files': embedded_files
                }
                
                # Analyze first embedded Excel if exists
                if embedded_files:
                    excel_data.update(self._analyze_embedded_excel(pptx_zip, embedded_files[0]))
                
                return excel_data
                
        except Exception as e:
            logger.error(f"Error detecting embedded Excel in PPTX: {e}")
            return {
                'has_embedded_excel': False,
                'embedded_excel_count': 0,
                'error': str(e)
            }
    
    def _analyze_embedded_excel(self, pptx_zip: zipfile.ZipFile, excel_file: str) -> Dict[str, Any]:
        """Analyze embedded Excel file"""
        try:
            import pandas as pd
            
            excel_content = pptx_zip.read(excel_file)
            
            with io.BytesIO(excel_content) as excel_buffer:
                xl_file = pd.ExcelFile(excel_buffer)
                sheet_names = xl_file.sheet_names
                
                return {
                    'worksheet_names': sheet_names,
                    'worksheet_count': len(sheet_names),
                    'has_architecture_sheet': any('architecture' in name.lower() for name in sheet_names)
                }
                
        except Exception as e:
            logger.error(f"Error analyzing embedded Excel in PPTX: {e}")
            return {
                'worksheet_names': [],
                'worksheet_count': 0,
                'has_architecture_sheet': False,
                'error': str(e)
            }
    
    def _check_milestones_section(self) -> Dict[str, Any]:
        """Check for milestones across all slides"""
        slide_files = self._get_slide_list()
        
        has_milestones = False
        milestone_slides = []
        implementation_dates = []
        
        # Check each slide for milestones
        for i, slide_file in enumerate(slide_files, 1):
            slide_text = self._extract_slide_text_cached(slide_file)
            
            # Look for milestone keywords
            if re.search(r'milestone|implementation|schedule|timeline', slide_text, re.IGNORECASE):
                has_milestones = True
                milestone_slides.append(i)
                
                # Extract dates
                date_pattern = re.compile(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})')
                dates = date_pattern.findall(slide_text)
                implementation_dates.extend(dates)
        
        return {
            'has_milestones_section': has_milestones,
            'milestone_slides': milestone_slides,
            'implementation_dates': list(set(implementation_dates))  # Remove duplicates
        }
    
    def _check_plt_status(self) -> Dict[str, Any]:
        """Check for PLT (Production Live Testing) status"""
        slide_files = self._get_slide_list()
        
        plt_status = None
        plt_slides = []
        
        # Check each slide for PLT status
        for i, slide_file in enumerate(slide_files, 1):
            slide_text = self._extract_slide_text_cached(slide_file)
            
            # Check for PLT status patterns
            for pattern in self.PLT_STATUS_PATTERNS:
                match = pattern.search(slide_text)
                if match:
                    plt_status = match.group(1).strip()
                    plt_slides.append(i)
                    break
            
            if plt_status:
                break  # Found PLT status, no need to check further
        
        return {
            'has_plt_status': plt_status is not None,
            'plt_status': plt_status,
            'plt_slides': plt_slides
        }
    
    def analyze_content(self) -> Dict[str, Any]:
        """Main analysis method"""
        try:
            slide_files = self._get_slide_list()
            
            if not slide_files:
                raise ValueError("No slides found in PPTX file")
            
            # Perform analysis
            result = {
                'slide_count': len(slide_files),
                'first_slide_data': self._analyze_first_slide(),
                'embedded_excel': self._detect_embedded_excel(),
                'milestones': self._check_milestones_section(),
                'plt_status': self._check_plt_status(),
                'has_content': len(slide_files) > 0
            }
            
            return result
            
        except Exception as e:
            logger.error(f"Error in PPTX analysis: {e}")
            raise


# Backward compatibility alias
PowerPointAnalyzer = OptimizedPowerPointAnalyzer
