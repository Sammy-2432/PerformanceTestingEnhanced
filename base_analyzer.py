"""
Optimized Document Analyzer Base Class
High-performance document analysis with caching and efficient processing
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, Optional, List, Union
from pathlib import Path
import logging
from functools import lru_cache
from dataclasses import dataclass
import hashlib
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

from ..config import CACHE_SIZE

logger = logging.getLogger(__name__)


@dataclass
class AnalysisResult:
    """Structured analysis result for better performance"""
    success: bool
    data: Dict[str, Any]
    processing_time: float
    file_hash: str
    errors: List[str]


class BaseAnalyzer(ABC):
    """
    Base analyzer class with optimization features:
    - Result caching
    - Parallel processing support
    - Memory-efficient file handling
    - Performance monitoring
    """
    
    def __init__(self, file_input: Union[str, Path, bytes], cache_enabled: bool = True):
        """Initialize with optimization settings"""
        self.file_input = file_input
        self.cache_enabled = cache_enabled
        self._cache: Dict[str, AnalysisResult] = {}
        self._file_hash: Optional[str] = None
        
    def _get_file_hash(self) -> str:
        """Get file hash for caching"""
        if self._file_hash:
            return self._file_hash
            
        try:
            if isinstance(self.file_input, (str, Path)):
                with open(self.file_input, 'rb') as f:
                    content = f.read()
            else:
                content = self.file_input if isinstance(self.file_input, bytes) else self.file_input.getvalue()
            
            self._file_hash = hashlib.md5(content).hexdigest()
            return self._file_hash
        except Exception as e:
            logger.error(f"Error generating file hash: {e}")
            return f"error_{time.time()}"
    
    @lru_cache(maxsize=CACHE_SIZE)
    def _cached_analyze(self, file_hash: str) -> AnalysisResult:
        """Cached analysis method"""
        return self._perform_analysis()
    
    def _perform_analysis(self) -> AnalysisResult:
        """Perform the actual analysis"""
        start_time = time.time()
        errors = []
        
        try:
            data = self.analyze_content()
            processing_time = time.time() - start_time
            
            return AnalysisResult(
                success=True,
                data=data,
                processing_time=processing_time,
                file_hash=self._get_file_hash(),
                errors=errors
            )
        except Exception as e:
            processing_time = time.time() - start_time
            error_msg = f"Analysis failed: {str(e)}"
            errors.append(error_msg)
            logger.error(error_msg)
            
            return AnalysisResult(
                success=False,
                data={},
                processing_time=processing_time,
                file_hash=self._get_file_hash(),
                errors=errors
            )
    
    @abstractmethod
    def analyze_content(self) -> Dict[str, Any]:
        """Implement specific analysis logic"""
        pass
    
    def analyze(self) -> Dict[str, Any]:
        """Main analysis method with caching"""
        if self.cache_enabled:
            file_hash = self._get_file_hash()
            result = self._cached_analyze(file_hash)
        else:
            result = self._perform_analysis()
        
        if result.success:
            logger.info(f"Analysis completed in {result.processing_time:.2f}s")
        else:
            logger.error(f"Analysis failed: {result.errors}")
        
        return result.data
    
    def get_analysis_stats(self) -> Dict[str, Any]:
        """Get analysis performance statistics"""
        if self.cache_enabled:
            file_hash = self._get_file_hash()
            result = self._cached_analyze(file_hash)
        else:
            result = self._perform_analysis()
        
        return {
            'success': result.success,
            'processing_time': result.processing_time,
            'file_hash': result.file_hash,
            'errors': result.errors,
            'cache_enabled': self.cache_enabled
        }
    
    def clear_cache(self):
        """Clear analysis cache"""
        self._cache.clear()
        self._cached_analyze.cache_clear()


class ParallelAnalyzer:
    """Utility for running multiple analyzers in parallel"""
    
    @staticmethod
    def analyze_multiple(analyzers: List[BaseAnalyzer], max_workers: int = 4) -> Dict[str, Any]:
        """Run multiple analyzers in parallel"""
        results = {}
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all analysis tasks
            future_to_analyzer = {
                executor.submit(analyzer.analyze): i 
                for i, analyzer in enumerate(analyzers)
            }
            
            # Collect results as they complete
            for future in as_completed(future_to_analyzer):
                analyzer_index = future_to_analyzer[future]
                try:
                    result = future.result()
                    results[f'analyzer_{analyzer_index}'] = result
                except Exception as e:
                    logger.error(f"Analyzer {analyzer_index} failed: {e}")
                    results[f'analyzer_{analyzer_index}'] = {'error': str(e)}
        
        return results
