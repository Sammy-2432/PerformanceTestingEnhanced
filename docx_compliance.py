"""
Optimized DOCX Compliance Checker
High-performance compliance checking for DOCX documents
"""

from typing import Dict, Any, List
import logging
from .base_checker import BaseComplianceChecker, ComplianceResult
from ..config import REQUIRED_WORKSHEETS

logger = logging.getLogger(__name__)


class OptimizedDocxComplianceChecker(BaseComplianceChecker):
    """
    Optimized DOCX compliance checker with:
    - Vectorized field validation
    - Parallel compliance checks
    - Efficient worksheet validation
    - Memory-optimized processing
    """
    
    def _check_first_page_compliance(self) -> ComplianceResult:
        """Check first page metadata compliance"""
        try:
            first_page_data = self.document_data.get('first_page_data', {})
            
            # Check multiple fields efficiently
            field_checks = []
            
            # Business Application ID
            if 'business_app_id' in self.excel_data and self.excel_data['business_app_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'business_app_id',
                        first_page_data.get('business_app_id'),
                        self.excel_data['business_app_id']
                    )
                )
            
            # Enterprise Release ID
            if 'enterprise_release_id' in self.excel_data and self.excel_data['enterprise_release_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'enterprise_release_id',
                        first_page_data.get('enterprise_release_id'),
                        self.excel_data['enterprise_release_id']
                    )
                )
            
            # Project Name
            if 'project_name' in self.excel_data and self.excel_data['project_name']:
                field_checks.append(
                    self._check_field_compliance(
                        'project_name',
                        first_page_data.get('project_name'),
                        self.excel_data['project_name']
                    )
                )
            
            # Task ID
            if 'task_id' in self.excel_data and self.excel_data['task_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'task_id',
                        first_page_data.get('task_id'),
                        self.excel_data['task_id']
                    )
                )
            
            # Calculate overall score
            if field_checks:
                total_score = sum(check.score for check in field_checks)
                average_score = total_score / len(field_checks)
                all_passed = all(check.passed for check in field_checks)
            else:
                average_score = 0.0
                all_passed = False
            
            return ComplianceResult(
                check_name="first_page_validation",
                passed=all_passed,
                score=average_score,
                details={
                    'field_checks': {check.check_name: check.details for check in field_checks},
                    'total_fields_checked': len(field_checks),
                    'fields_passed': sum(1 for check in field_checks if check.passed)
                },
                errors=[]
            )
            
        except Exception as e:
            error_msg = f"Error checking first page compliance: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="first_page_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_footer_compliance(self) -> ComplianceResult:
        """Check footer compliance"""
        try:
            footer_data = self.document_data.get('footer_data', {})
            project_in_footer = footer_data.get('project_name_in_footer')
            
            if not project_in_footer:
                return ComplianceResult(
                    check_name="footer_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'No project name found in footer'},
                    errors=['No project name found in footer']
                )
            
            # Check against Excel project names
            if 'project_name' in self.excel_data and self.excel_data['project_name']:
                is_match, score, best_match = self._vectorized_text_match(
                    project_in_footer, 
                    [str(name) for name in self.excel_data['project_name'] if name]
                )
                
                return ComplianceResult(
                    check_name="footer_validation",
                    passed=is_match,
                    score=score,
                    details={
                        'footer_project_name': project_in_footer,
                        'excel_project_names': self.excel_data['project_name'],
                        'best_match': best_match,
                        'similarity_score': score
                    },
                    errors=[] if is_match else ['Project name in footer does not match Excel data']
                )
            else:
                return ComplianceResult(
                    check_name="footer_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'No project names in Excel data'},
                    errors=['No project names in Excel data']
                )
                
        except Exception as e:
            error_msg = f"Error checking footer compliance: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="footer_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_table_of_contents_compliance(self) -> ComplianceResult:
        """Check table of contents compliance"""
        try:
            toc_data = self.document_data.get('table_of_contents', {})
            
            has_toc = toc_data.get('has_table_of_contents', False)
            has_scope_section = toc_data.get('has_scope_section', False)
            
            # Score based on TOC presence and scope section
            score = 0.0
            errors = []
            
            if has_toc:
                score += 0.6  # 60% for having TOC
            else:
                errors.append('Table of contents not found')
            
            if has_scope_section:
                score += 0.4  # 40% for having scope section
            else:
                errors.append('Section 3.3 "In Scope" not found')
            
            passed = score >= 0.8  # Need both TOC and scope section for pass
            
            return ComplianceResult(
                check_name="table_of_contents_validation",
                passed=passed,
                score=score,
                details={
                    'has_table_of_contents': has_toc,
                    'has_scope_section': has_scope_section,
                    'requirements_met': 2 - len(errors)
                },
                errors=errors
            )
            
        except Exception as e:
            error_msg = f"Error checking table of contents: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="table_of_contents_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_embedded_excel_compliance(self) -> ComplianceResult:
        """Check embedded Excel compliance"""
        try:
            excel_data = self.document_data.get('embedded_excel', {})
            
            has_embedded_excel = excel_data.get('has_embedded_excel', False)
            
            if not has_embedded_excel:
                return ComplianceResult(
                    check_name="embedded_excel_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'No embedded Excel files found'},
                    errors=['No embedded Excel files found']
                )
            
            # Check worksheets
            worksheet_names = excel_data.get('worksheet_names', [])
            has_architecture = excel_data.get('has_architecture_sheet', False)
            
            # Calculate score based on required worksheets
            score = 0.0
            errors = []
            
            # Check for required worksheets
            found_worksheets = []
            for required_ws in REQUIRED_WORKSHEETS:
                if any(required_ws.lower() in ws.lower() for ws in worksheet_names):
                    found_worksheets.append(required_ws)
            
            worksheet_score = len(found_worksheets) / len(REQUIRED_WORKSHEETS)
            score += worksheet_score * 0.8  # 80% for worksheets
            
            if has_architecture:
                score += 0.2  # 20% for architecture sheet
            else:
                errors.append('Architecture sheet not found')
            
            passed = score >= 0.7  # Need most worksheets to pass
            
            return ComplianceResult(
                check_name="embedded_excel_validation",
                passed=passed,
                score=score,
                details={
                    'has_embedded_excel': has_embedded_excel,
                    'worksheet_names': worksheet_names,
                    'required_worksheets': list(REQUIRED_WORKSHEETS),
                    'found_worksheets': found_worksheets,
                    'has_architecture_sheet': has_architecture,
                    'worksheet_compliance_rate': worksheet_score
                },
                errors=errors
            )
            
        except Exception as e:
            error_msg = f"Error checking embedded Excel: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="embedded_excel_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_milestones_compliance(self) -> ComplianceResult:
        """Check milestones section compliance"""
        try:
            milestones_data = self.document_data.get('milestones', {})
            
            has_milestones = milestones_data.get('has_milestones_section', False)
            implementation_dates = milestones_data.get('implementation_dates', [])
            
            if not has_milestones:
                return ComplianceResult(
                    check_name="milestones_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'Milestones section not found'},
                    errors=['Section 12 milestones not found']
                )
            
            # Score based on presence and dates
            score = 0.6  # Base score for having milestones section
            
            if implementation_dates:
                score += 0.4  # Additional score for having implementation dates
                
                # Compare with Excel end dates if available
                if 'end_date' in self.excel_data and self.excel_data['end_date']:
                    # This is a simplified comparison - could be enhanced
                    score = min(score + 0.2, 1.0)
            
            passed = score >= 0.6
            
            return ComplianceResult(
                check_name="milestones_validation",
                passed=passed,
                score=score,
                details={
                    'has_milestones_section': has_milestones,
                    'implementation_dates': implementation_dates,
                    'excel_end_dates': self.excel_data.get('end_date', []),
                    'dates_found': len(implementation_dates)
                },
                errors=[] if passed else ['Milestones section incomplete']
            )
            
        except Exception as e:
            error_msg = f"Error checking milestones: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="milestones_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def get_check_functions(self) -> List[callable]:
        """Return list of check functions for parallel execution"""
        return [
            self._check_first_page_compliance,
            self._check_footer_compliance,
            self._check_table_of_contents_compliance,
            self._check_embedded_excel_compliance,
            self._check_milestones_compliance
        ]


# Backward compatibility alias
ComplianceChecker = OptimizedDocxComplianceChecker
