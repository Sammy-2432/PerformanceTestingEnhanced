"""
Optimized Compliance Checker Base Class
High-performance compliance checking with vectorized operations and caching
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Tuple, Optional, Set
import logging
from functools import lru_cache
from dataclasses import dataclass
import numpy as np
from concurrent.futures import ThreadPoolExecutor, as_completed

from ..config import COMPLIANCE_THRESHOLD, PARTIAL_MATCH_WEIGHT, CACHE_SIZE

logger = logging.getLogger(__name__)


@dataclass
class ComplianceResult:
    """Structured compliance result for better performance"""
    check_name: str
    passed: bool
    score: float
    details: Dict[str, Any]
    errors: List[str]


class BaseComplianceChecker(ABC):
    """
    Base compliance checker with optimization features:
    - Vectorized compliance scoring
    - Parallel check execution
    - Result caching
    - Memory-efficient processing
    """
    
    def __init__(self, excel_data: Dict[str, Any], document_data: Dict[str, Any]):
        """Initialize with optimized data structures"""
        self.excel_data = excel_data
        self.document_data = document_data
        self._cache: Dict[str, ComplianceResult] = {}
        
    @lru_cache(maxsize=CACHE_SIZE)
    def _normalize_text_cached(self, text: str) -> str:
        """Cached text normalization for comparison"""
        if not text:
            return ""
        
        return str(text).strip().lower().replace(" ", "").replace("-", "").replace("_", "")
    
    def _calculate_similarity_score(self, text1: str, text2: str) -> float:
        """Calculate similarity score between two texts efficiently"""
        if not text1 or not text2:
            return 0.0
        
        norm1 = self._normalize_text_cached(text1)
        norm2 = self._normalize_text_cached(text2)
        
        if norm1 == norm2:
            return 1.0
        
        # Use Jaccard similarity for efficiency
        set1 = set(norm1)
        set2 = set(norm2)
        
        intersection = len(set1 & set2)
        union = len(set1 | set2)
        
        return intersection / union if union > 0 else 0.0
    
    def _vectorized_text_match(self, target: str, candidates: List[str], threshold: float = 0.8) -> Tuple[bool, float, Optional[str]]:
        """Vectorized text matching for multiple candidates"""
        if not target or not candidates:
            return False, 0.0, None
        
        # Calculate similarities for all candidates
        similarities = [self._calculate_similarity_score(target, candidate) for candidate in candidates]
        
        if not similarities:
            return False, 0.0, None
        
        max_similarity = max(similarities)
        best_match_idx = similarities.index(max_similarity)
        best_match = candidates[best_match_idx]
        
        return max_similarity >= threshold, max_similarity, best_match
    
    def _check_field_compliance(self, field_name: str, document_value: Any, excel_values: List[Any]) -> ComplianceResult:
        """Check compliance for a specific field"""
        errors = []
        
        try:
            if not document_value:
                return ComplianceResult(
                    check_name=f"{field_name}_validation",
                    passed=False,
                    score=0.0,
                    details={'error': f'{field_name} not found in document'},
                    errors=[f'{field_name} not found in document']
                )
            
            if not excel_values:
                return ComplianceResult(
                    check_name=f"{field_name}_validation",
                    passed=False,
                    score=0.0,
                    details={'error': f'{field_name} not found in Excel data'},
                    errors=[f'{field_name} not found in Excel data']
                )
            
            # Convert to strings for comparison
            doc_str = str(document_value)
            excel_strs = [str(val) for val in excel_values if val is not None]
            
            # Use vectorized matching
            is_match, score, best_match = self._vectorized_text_match(doc_str, excel_strs)
            
            return ComplianceResult(
                check_name=f"{field_name}_validation",
                passed=is_match,
                score=score,
                details={
                    'document_value': doc_str,
                    'excel_values': excel_strs,
                    'best_match': best_match,
                    'similarity_score': score
                },
                errors=errors
            )
            
        except Exception as e:
            error_msg = f"Error checking {field_name}: {str(e)}"
            errors.append(error_msg)
            logger.error(error_msg)
            
            return ComplianceResult(
                check_name=f"{field_name}_validation",
                passed=False,
                score=0.0,
                details={'error': error_msg},
                errors=errors
            )
    
    def _run_checks_parallel(self, check_functions: List[callable], max_workers: int = 4) -> List[ComplianceResult]:
        """Run compliance checks in parallel"""
        results = []
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all check functions
            future_to_check = {
                executor.submit(check_func): check_func.__name__ 
                for check_func in check_functions
            }
            
            # Collect results as they complete
            for future in as_completed(future_to_check):
                check_name = future_to_check[future]
                try:
                    result = future.result()
                    if isinstance(result, ComplianceResult):
                        results.append(result)
                    else:
                        # Handle legacy return format
                        results.append(ComplianceResult(
                            check_name=check_name,
                            passed=result.get('passed', False),
                            score=result.get('score', 0.0),
                            details=result,
                            errors=result.get('errors', [])
                        ))
                except Exception as e:
                    logger.error(f"Check {check_name} failed: {e}")
                    results.append(ComplianceResult(
                        check_name=check_name,
                        passed=False,
                        score=0.0,
                        details={'error': str(e)},
                        errors=[str(e)]
                    ))
        
        return results
    
    def _calculate_overall_score(self, results: List[ComplianceResult]) -> Dict[str, Any]:
        """Calculate overall compliance score"""
        if not results:
            return {
                'overall_score': 0.0,
                'passed_checks': 0,
                'total_checks': 0,
                'compliance_level': 'Non-compliant'
            }
        
        # Use numpy for efficient calculations
        scores = np.array([result.score for result in results])
        passed_count = sum(1 for result in results if result.passed)
        
        overall_score = np.mean(scores)
        pass_rate = passed_count / len(results)
        
        # Determine compliance level
        if overall_score >= 0.9:
            compliance_level = 'Fully Compliant'
        elif overall_score >= COMPLIANCE_THRESHOLD:
            compliance_level = 'Compliant'
        elif overall_score >= 0.4:
            compliance_level = 'Partially Compliant'
        else:
            compliance_level = 'Non-compliant'
        
        return {
            'overall_score': float(overall_score),
            'pass_rate': float(pass_rate),
            'passed_checks': passed_count,
            'total_checks': len(results),
            'compliance_level': compliance_level,
            'individual_scores': {result.check_name: result.score for result in results}
        }
    
    @abstractmethod
    def get_check_functions(self) -> List[callable]:
        """Return list of check functions to run"""
        pass
    
    def run_all_checks(self, parallel: bool = True, max_workers: int = 4) -> Dict[str, Any]:
        """Run all compliance checks"""
        try:
            check_functions = self.get_check_functions()
            
            if parallel and len(check_functions) > 1:
                results = self._run_checks_parallel(check_functions, max_workers)
            else:
                results = [check_func() for check_func in check_functions]
            
            # Calculate overall scores
            overall_stats = self._calculate_overall_score(results)
            
            # Convert results to legacy format for compatibility
            legacy_results = {}
            for result in results:
                legacy_results[result.check_name] = {
                    'passed': result.passed,
                    'score': result.score,
                    'details': result.details,
                    'errors': result.errors
                }
            
            return {
                'individual_checks': legacy_results,
                'overall_statistics': overall_stats,
                'metadata': {
                    'total_processing_time': 0.0,  # Could be added if needed
                    'parallel_execution': parallel,
                    'max_workers': max_workers if parallel else 1
                }
            }
            
        except Exception as e:
            logger.error(f"Error running compliance checks: {e}")
            return {
                'individual_checks': {},
                'overall_statistics': {
                    'overall_score': 0.0,
                    'passed_checks': 0,
                    'total_checks': 0,
                    'compliance_level': 'Error'
                },
                'error': str(e)
            }
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get performance statistics"""
        return {
            'cache_size': len(self._cache),
            'excel_data_size': len(str(self.excel_data)),
            'document_data_size': len(str(self.document_data))
        }
