"""
Optimized Excel Reader Module
High-performance Excel file reading with caching and efficient data structures
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Optional, Any, Union, Tuple, Set
from functools import lru_cache
import logging
from dataclasses import dataclass
import weakref
from concurrent.futures import ThreadPoolExecutor
import gc

from src.config import (
    EXCEL_COLUMN_MAPPINGS, 
    get_column_mapping_cache, 
    CHUNK_SIZE,
    MAX_MEMORY_USAGE,
    CACHE_SIZE
)

logger = logging.getLogger(__name__)


@dataclass
class ColumnMapping:
    """Efficient column mapping storage"""
    release_col: Optional[str] = None
    project_col: Optional[str] = None
    business_app_id_col: Optional[str] = None
    enterprise_release_id_col: Optional[str] = None
    task_id_col: Optional[str] = None
    end_date_col: Optional[str] = None


class OptimizedExcelReader:
    """
    High-performance Excel reader with optimizations:
    - Memory-efficient data loading
    - Cached column mappings
    - Lazy evaluation
    - Optimized filtering using pandas vectorized operations
    """
    
    def __init__(self, file_path: Union[str, Path]):
        """Initialize with optimized data structures"""
        self.file_path = Path(file_path)
        self._df: Optional[pd.DataFrame] = None
        self._column_mapping: Optional[ColumnMapping] = None
        self._data_cache: Dict[str, Any] = {}
        self._is_loaded = False
        
        # Memory management
        self._memory_usage = 0
        self._cache_size = 0
        
        # Use weak references for cleanup
        self._weakref = weakref.ref(self)
    
    def __del__(self):
        """Cleanup resources"""
        self.clear_cache()
    
    def clear_cache(self):
        """Clear cache and free memory"""
        self._data_cache.clear()
        self._cache_size = 0
        if self._df is not None:
            del self._df
            self._df = None
        gc.collect()
    
    @property
    def memory_usage(self) -> int:
        """Get current memory usage in bytes"""
        if self._df is not None:
            return self._df.memory_usage(deep=True).sum()
        return 0
    
    def load_data(self, sheet_name: str = "Sheet1", optimize_memory: bool = True) -> bool:
        """
        Load Excel data with optimizations
        
        Args:
            sheet_name: Excel sheet name
            optimize_memory: Whether to optimize memory usage
        """
        try:
            if not self.file_path.exists():
                logger.error(f"Excel file not found: {self.file_path}")
                return False
            
            # Load with optimized dtypes
            dtype_mapping = self._get_optimal_dtypes() if optimize_memory else None
            
            # Use chunking for large files
            file_size = self.file_path.stat().st_size
            use_chunks = file_size > MAX_MEMORY_USAGE
            
            if use_chunks:
                self._df = self._load_in_chunks(sheet_name, dtype_mapping)
            else:
                self._df = pd.read_excel(
                    self.file_path,
                    sheet_name=sheet_name,
                    dtype=dtype_mapping,
                    engine='openpyxl'
                )
            
            if self._df is None or self._df.empty:
                logger.error("No data loaded from Excel file")
                return False
            
            # Optimize data types post-load
            if optimize_memory:
                self._optimize_dataframe()
            
            # Build column mapping once
            self._build_column_mapping()
            
            self._is_loaded = True
            logger.info(f"Loaded {len(self._df)} rows from {self.file_path}")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            return False
    
    def _get_optimal_dtypes(self) -> Dict[str, str]:
        """Get optimal data types for columns"""
        return {
            # Use category for string columns with limited values
            'Release': 'category',
            'Project Name': 'category',
            'Business Application ID': 'category',
            'Enterprise Release ID': 'category',
            'Task ID': 'string',
        }
    
    def _load_in_chunks(self, sheet_name: str, dtype_mapping: Optional[Dict]) -> pd.DataFrame:
        """Load large Excel file in chunks"""
        try:
            # For Excel files, we can't directly chunk, so we'll load and then optimize
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                dtype=dtype_mapping,
                engine='openpyxl'
            )
            return df
        except Exception as e:
            logger.error(f"Error loading Excel in chunks: {e}")
            return None
    
    def _optimize_dataframe(self):
        """Optimize DataFrame memory usage"""
        if self._df is None:
            return
        
        # Convert object columns to category where appropriate
        for col in self._df.select_dtypes(include=['object']):
            if self._df[col].nunique() / len(self._df) < 0.5:  # Less than 50% unique values
                self._df[col] = self._df[col].astype('category')
        
        # Optimize numeric columns
        for col in self._df.select_dtypes(include=['int64']):
            self._df[col] = pd.to_numeric(self._df[col], downcast='integer')
        
        for col in self._df.select_dtypes(include=['float64']):
            self._df[col] = pd.to_numeric(self._df[col], downcast='float')
    
    @lru_cache(maxsize=CACHE_SIZE)
    def _find_column_cached(self, mapping_key: str) -> Optional[str]:
        """Cached column finding for O(1) subsequent lookups"""
        if self._df is None:
            return None
        
        possible_names = EXCEL_COLUMN_MAPPINGS.get(mapping_key, ())
        
        # Exact match first (fastest)
        for col_name in possible_names:
            if col_name in self._df.columns:
                return col_name
        
        # Case-insensitive match
        for col in self._df.columns:
            for possible_name in possible_names:
                if possible_name.lower() in col.lower():
                    return col
        
        return None
    
    def _build_column_mapping(self):
        """Build column mapping once for efficient access"""
        self._column_mapping = ColumnMapping(
            release_col=self._find_column_cached('release'),
            project_col=self._find_column_cached('project_name'),
            business_app_id_col=self._find_column_cached('business_app_id'),
            enterprise_release_id_col=self._find_column_cached('enterprise_release_id'),
            task_id_col=self._find_column_cached('task_id'),
            end_date_col=self._find_column_cached('end_date')
        )
    
    @lru_cache(maxsize=CACHE_SIZE)
    def get_releases(self) -> List[str]:
        """Get unique releases with caching"""
        if not self._is_loaded or self._column_mapping.release_col is None:
            return []
        
        try:
            releases = self._df[self._column_mapping.release_col].dropna().unique()
            return sorted([str(r) for r in releases])
        except Exception as e:
            logger.error(f"Error getting releases: {e}")
            return []
    
    @lru_cache(maxsize=CACHE_SIZE)
    def get_projects_by_release(self, release: str) -> List[str]:
        """Get projects for a release with vectorized filtering"""
        if not self._is_loaded or not all([
            self._column_mapping.release_col,
            self._column_mapping.project_col
        ]):
            return []
        
        try:
            # Use vectorized operations for better performance
            mask = (self._df[self._column_mapping.release_col] == release)
            projects = self._df.loc[mask, self._column_mapping.project_col].dropna().unique()
            return sorted([str(p) for p in projects])
        except Exception as e:
            logger.error(f"Error getting projects for release {release}: {e}")
            return []
    
    @lru_cache(maxsize=CACHE_SIZE)
    def get_enterprise_release_ids_by_release_and_project(self, release: str, project: str) -> List[str]:
        """Get Enterprise Release IDs with optimized filtering"""
        if not self._is_loaded or not all([
            self._column_mapping.release_col,
            self._column_mapping.project_col,
            self._column_mapping.enterprise_release_id_col
        ]):
            return []
        
        try:
            # Use boolean indexing for faster filtering
            mask = (
                (self._df[self._column_mapping.release_col] == release) &
                (self._df[self._column_mapping.project_col] == project)
            )
            
            enterprise_ids = self._df.loc[mask, self._column_mapping.enterprise_release_id_col].dropna().unique()
            return sorted([str(eid) for eid in enterprise_ids])
        except Exception as e:
            logger.error(f"Error getting Enterprise Release IDs: {e}")
            return []
    
    @lru_cache(maxsize=CACHE_SIZE)
    def get_business_app_ids_by_release_and_project(self, release: str, project: str) -> List[str]:
        """Get Business App IDs with optimized filtering"""
        if not self._is_loaded or not all([
            self._column_mapping.release_col,
            self._column_mapping.project_col,
            self._column_mapping.business_app_id_col
        ]):
            return []
        
        try:
            mask = (
                (self._df[self._column_mapping.release_col] == release) &
                (self._df[self._column_mapping.project_col] == project)
            )
            
            app_ids = self._df.loc[mask, self._column_mapping.business_app_id_col].dropna().unique()
            return sorted([str(aid) for aid in app_ids])
        except Exception as e:
            logger.error(f"Error getting Business App IDs: {e}")
            return []
    
    def get_project_data_by_criteria(self, release: str, project: str, enterprise_release_id: str) -> Dict[str, Any]:
        """Get project data with optimized lookup"""
        if not self._is_loaded:
            return {}
        
        try:
            # Build filter conditions dynamically
            conditions = []
            
            if self._column_mapping.release_col:
                conditions.append(self._df[self._column_mapping.release_col] == release)
            if self._column_mapping.project_col:
                conditions.append(self._df[self._column_mapping.project_col] == project)
            if self._column_mapping.enterprise_release_id_col:
                conditions.append(self._df[self._column_mapping.enterprise_release_id_col] == enterprise_release_id)
            
            if not conditions:
                return {}
            
            # Combine conditions with boolean AND
            combined_mask = conditions[0]
            for condition in conditions[1:]:
                combined_mask &= condition
            
            filtered_df = self._df[combined_mask]
            
            if filtered_df.empty:
                return {}
            
            # Get first matching row
            row = filtered_df.iloc[0]
            
            # Build result efficiently
            result = {}
            mapping_to_column = {
                'Release': self._column_mapping.release_col,
                'Project Name': self._column_mapping.project_col,
                'Business Application ID': self._column_mapping.business_app_id_col,
                'Enterprise Release ID': self._column_mapping.enterprise_release_id_col,
                'Task ID': self._column_mapping.task_id_col,
                'End Date': self._column_mapping.end_date_col
            }
            
            for key, col in mapping_to_column.items():
                if col and col in row.index:
                    result[key] = row[col]
            
            return result
            
        except Exception as e:
            logger.error(f"Error getting project data: {e}")
            return {}
    
    def load_data_from_upload(self, uploaded_file, sheet_name: str = "Sheet1") -> bool:
        """Load data from uploaded file with optimization"""
        try:
            import io
            
            # Read into memory-efficient format
            file_content = uploaded_file.getvalue()
            
            # Check file size
            if len(file_content) > MAX_MEMORY_USAGE:
                logger.warning(f"Large file detected: {len(file_content)} bytes")
            
            with io.BytesIO(file_content) as buffer:
                self._df = pd.read_excel(
                    buffer,
                    sheet_name=sheet_name,
                    engine='openpyxl'
                )
            
            if self._df is None or self._df.empty:
                logger.error("No data loaded from uploaded file")
                return False
            
            # Optimize after loading
            self._optimize_dataframe()
            self._build_column_mapping()
            
            self._is_loaded = True
            logger.info(f"Loaded {len(self._df)} rows from uploaded file")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading uploaded Excel file: {e}")
            return False
    
    @property
    def data(self) -> Dict[str, List]:
        """Get data in legacy format for compatibility"""
        if not self._is_loaded:
            return {}
        
        # Use cached version if available
        cache_key = "legacy_data"
        if cache_key in self._data_cache:
            return self._data_cache[cache_key]
        
        try:
            result = {}
            mapping = {
                'release': self._column_mapping.release_col,
                'business_app_id': self._column_mapping.business_app_id_col,
                'enterprise_release_id': self._column_mapping.enterprise_release_id_col,
                'project_name': self._column_mapping.project_col,
                'task_id': self._column_mapping.task_id_col,
                'end_date': self._column_mapping.end_date_col
            }
            
            for key, col in mapping.items():
                if col and col in self._df.columns:
                    result[key] = self._df[col].dropna().tolist()
                else:
                    result[key] = []
            
            # Cache the result
            self._data_cache[cache_key] = result
            return result
            
        except Exception as e:
            logger.error(f"Error getting legacy data format: {e}")
            return {}
    
    def is_data_loaded(self) -> bool:
        """Check if data is loaded"""
        return self._is_loaded and self._df is not None and not self._df.empty
    
    def get_stats(self) -> Dict[str, Any]:
        """Get performance statistics"""
        return {
            'rows': len(self._df) if self._df is not None else 0,
            'columns': len(self._df.columns) if self._df is not None else 0,
            'memory_usage_mb': self.memory_usage / (1024 * 1024),
            'cache_size': len(self._data_cache),
            'is_loaded': self._is_loaded
        }


# Backward compatibility alias
ExcelReader = OptimizedExcelReader
