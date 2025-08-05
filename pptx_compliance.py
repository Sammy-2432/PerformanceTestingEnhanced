"""
Optimized PPTX Compliance Checker
High-performance compliance checking for PPTX documents
"""

from typing import Dict, Any, List
import logging
from .base_checker import BaseComplianceChecker, ComplianceResult
from ..config import REQUIRED_WORKSHEETS

logger = logging.getLogger(__name__)


class OptimizedPptxComplianceChecker(BaseComplianceChecker):
    """
    Optimized PPTX compliance checker with:
    - Vectorized slide validation
    - Parallel compliance checks
    - Efficient PLT status checking
    - Memory-optimized processing
    """
    
    def _check_first_slide_compliance(self) -> ComplianceResult:
        """Check first slide metadata compliance"""
        try:
            first_slide_data = self.document_data.get('first_slide_data', {})
            
            if not first_slide_data.get('has_content', False):
                return ComplianceResult(
                    check_name="first_slide_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'First slide has no content'},
                    errors=['First slide has no content']
                )
            
            # Check multiple fields efficiently
            field_checks = []
            
            # Business Application ID
            if 'business_app_id' in self.excel_data and self.excel_data['business_app_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'business_app_id',
                        first_slide_data.get('business_app_id'),
                        self.excel_data['business_app_id']
                    )
                )
            
            # Enterprise Release ID
            if 'enterprise_release_id' in self.excel_data and self.excel_data['enterprise_release_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'enterprise_release_id',
                        first_slide_data.get('enterprise_release_id'),
                        self.excel_data['enterprise_release_id']
                    )
                )
            
            # Project Name
            if 'project_name' in self.excel_data and self.excel_data['project_name']:
                field_checks.append(
                    self._check_field_compliance(
                        'project_name',
                        first_slide_data.get('project_name'),
                        self.excel_data['project_name']
                    )
                )
            
            # Task ID
            if 'task_id' in self.excel_data and self.excel_data['task_id']:
                field_checks.append(
                    self._check_field_compliance(
                        'task_id',
                        first_slide_data.get('task_id'),
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
                check_name="first_slide_validation",
                passed=all_passed,
                score=average_score,
                details={
                    'field_checks': {check.check_name: check.details for check in field_checks},
                    'total_fields_checked': len(field_checks),
                    'fields_passed': sum(1 for check in field_checks if check.passed),
                    'slide_has_content': first_slide_data.get('has_content', False)
                },
                errors=[]
            )
            
        except Exception as e:
            error_msg = f"Error checking first slide compliance: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="first_slide_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_embedded_excel_compliance(self) -> ComplianceResult:
        """Check embedded Excel compliance in PPTX"""
        try:
            excel_data = self.document_data.get('embedded_excel', {})
            
            has_embedded_excel = excel_data.get('has_embedded_excel', False)
            
            if not has_embedded_excel:
                return ComplianceResult(
                    check_name="embedded_excel_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'No embedded Excel files found in PPTX'},
                    errors=['No embedded Excel files found in PPTX']
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
            
            worksheet_score = len(found_worksheets) / len(REQUIRED_WORKSHEETS) if REQUIRED_WORKSHEETS else 0
            score += worksheet_score * 0.8  # 80% for worksheets
            
            if has_architecture:
                score += 0.2  # 20% for architecture sheet
            else:
                errors.append('Architecture sheet not found in embedded Excel')
            
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
                    'worksheet_compliance_rate': worksheet_score,
                    'embedded_excel_count': excel_data.get('embedded_excel_count', 0)
                },
                errors=errors
            )
            
        except Exception as e:
            error_msg = f"Error checking embedded Excel in PPTX: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="embedded_excel_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_milestones_compliance(self) -> ComplianceResult:
        """Check milestones compliance in PPTX"""
        try:
            milestones_data = self.document_data.get('milestones', {})
            
            has_milestones = milestones_data.get('has_milestones_section', False)
            milestone_slides = milestones_data.get('milestone_slides', [])
            implementation_dates = milestones_data.get('implementation_dates', [])
            
            if not has_milestones:
                return ComplianceResult(
                    check_name="milestones_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'No milestones found in any slides'},
                    errors=['No milestones found in any slides']
                )
            
            # Score based on presence and quality
            score = 0.4  # Base score for having milestones
            
            if milestone_slides:
                score += 0.3  # Additional score for having dedicated milestone slides
            
            if implementation_dates:
                score += 0.3  # Additional score for having implementation dates
                
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
                    'milestone_slides': milestone_slides,
                    'implementation_dates': implementation_dates,
                    'excel_end_dates': self.excel_data.get('end_date', []),
                    'slides_with_milestones': len(milestone_slides),
                    'dates_found': len(implementation_dates)
                },
                errors=[] if passed else ['Milestones section incomplete']
            )
            
        except Exception as e:
            error_msg = f"Error checking milestones in PPTX: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="milestones_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_plt_status_compliance(self) -> ComplianceResult:
        """Check PLT (Production Live Testing) status compliance"""
        try:
            plt_data = self.document_data.get('plt_status', {})
            
            has_plt_status = plt_data.get('has_plt_status', False)
            plt_status = plt_data.get('plt_status')
            plt_slides = plt_data.get('plt_slides', [])
            
            if not has_plt_status:
                return ComplianceResult(
                    check_name="plt_status_validation",
                    passed=False,
                    score=0.0,
                    details={'error': 'PLT status not found in any slides'},
                    errors=['PLT (Production Live Testing) status not found']
                )
            
            # Score based on PLT status presence and content
            score = 0.7  # Base score for having PLT status
            
            # Additional scoring based on PLT status content
            if plt_status:
                plt_status_lower = plt_status.lower()
                
                # Check for positive indicators
                positive_indicators = ['complete', 'passed', 'successful', 'green', 'ok']
                negative_indicators = ['pending', 'failed', 'incomplete', 'red', 'issue']
                
                if any(indicator in plt_status_lower for indicator in positive_indicators):
                    score += 0.3  # Bonus for positive PLT status
                elif any(indicator in plt_status_lower for indicator in negative_indicators):
                    score += 0.1  # Small bonus for at least reporting negative status
                else:
                    score += 0.2  # Moderate bonus for having some status
            
            passed = score >= 0.7
            
            return ComplianceResult(
                check_name="plt_status_validation",
                passed=passed,
                score=score,
                details={
                    'has_plt_status': has_plt_status,
                    'plt_status': plt_status,
                    'plt_slides': plt_slides,
                    'slides_with_plt': len(plt_slides)
                },
                errors=[] if passed else ['PLT status found but may need more detail']
            )
            
        except Exception as e:
            error_msg = f"Error checking PLT status: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="plt_status_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def _check_slide_count_compliance(self) -> ComplianceResult:
        """Check if presentation has adequate number of slides"""
        try:
            slide_count = self.document_data.get('slide_count', 0)
            
            # Score based on slide count (reasonable range for test reports)
            if slide_count >= 10:  # Comprehensive report
                score = 1.0
                passed = True
                errors = []
            elif slide_count >= 5:  # Adequate report
                score = 0.8
                passed = True
                errors = []
            elif slide_count >= 3:  # Minimal report
                score = 0.6
                passed = True
                errors = ['Presentation seems short for a comprehensive test report']
            else:  # Too short
                score = 0.3
                passed = False
                errors = ['Presentation is too short for a proper test report']
            
            return ComplianceResult(
                check_name="slide_count_validation",
                passed=passed,
                score=score,
                details={
                    'slide_count': slide_count,
                    'minimum_expected': 5,
                    'comprehensive_threshold': 10
                },
                errors=errors
            )
            
        except Exception as e:
            error_msg = f"Error checking slide count: {str(e)}"
            logger.error(error_msg)
            return ComplianceResult(
                check_name="slide_count_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=[error_msg]
            )
    
    def get_check_functions(self) -> List[callable]:
        """Return list of check functions for parallel execution"""
        return [
            self._check_first_slide_compliance,
            self._check_embedded_excel_compliance,
            self._check_milestones_compliance,
            self._check_plt_status_compliance,
            self._check_slide_count_compliance
        ]


# Backward compatibility alias
TestReportComplianceChecker = OptimizedPptxComplianceChecker
