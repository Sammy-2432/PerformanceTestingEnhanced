"""
Optimized Configuration Module for Smart Test Plan Compliance Checker
Centralized configuration with efficient data structures and caching
"""

import os
from pathlib import Path
from typing import Dict, List, Tuple, Set
from dataclasses import dataclass
from functools import lru_cache


@dataclass(frozen=True)
class AppConfig:
    """Immutable application configuration using dataclass for better performance"""
    title: str = "Truist Smart Compliance Checker"
    icon: str = "📊"
    version: str = "2.0.0"
    port: int = 8501


@dataclass(frozen=True)
class UIConfig:
    """UI-specific configuration"""
    sidebar_width: int = 300
    max_file_upload_size: int = 200  # MB
    results_per_page: int = 10
    cache_ttl: int = 3600  # 1 hour


# Application configuration
APP = AppConfig()
UI = UIConfig()

# Paths configuration
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
TEMP_DIR = BASE_DIR / "temp"

# Excel file configuration
EXCEL_FILE_PATH = DATA_DIR / "sample_project_data.xlsx"
EXCEL_SHEET_NAME = "Sheet1"

# Alternative paths for Excel files (prioritized list)
ALTERNATIVE_EXCEL_PATHS = [
    r"\\shared\network\path\project_data.xlsx",
    r"c:\shared\data\project_data.xlsx",
    DATA_DIR / "project_data.xlsx"
]

# Time configuration
UPDATE_DAY = 2  # Wednesday (0=Monday, 1=Tuesday, etc.)

# Compliance configuration
COMPLIANCE_THRESHOLD = 0.6
PARTIAL_MATCH_WEIGHT = 0.5

# Required worksheets (using frozenset for O(1) lookups)
REQUIRED_WORKSHEETS: frozenset = frozenset([
    "Cover Page",
    "General Details", 
    "Business Scenario(s)",
    "Data Requirement",
    "Architecture",
    "Logs&Contacts"
])

# Optimized column mappings using tuples for immutability and better performance
EXCEL_COLUMN_MAPPINGS: Dict[str, Tuple[str, ...]] = {
    'release': ('Release', 'Release Version', 'Version'),
    'business_app_id': ('Business Application ID', 'Business App ID', 'App ID', 'Application ID'),
    'enterprise_release_id': ('Enterprise Release ID', 'Release ID', 'Enterprise ID'),
    'project_name': ('Project Name', 'Project', 'Name'),
    'task_id': ('Task ID', 'Task', 'ID'),
    'end_date': ('End Date', 'Completion Date', 'Target Date', 'Due Date')
}

# Caching configuration for column matching
@lru_cache(maxsize=128)
def get_column_mapping_cache(column_name: str, mapping_key: str) -> bool:
    """Cached column name matching for O(1) lookups after first match"""
    possible_names = EXCEL_COLUMN_MAPPINGS.get(mapping_key, ())
    return any(name.lower() in column_name.lower() for name in possible_names)


# File type configurations
SUPPORTED_DOCX_EXTENSIONS = frozenset(['docx'])
SUPPORTED_PPTX_EXTENSIONS = frozenset(['pptx'])
SUPPORTED_EXCEL_EXTENSIONS = frozenset(['xlsx', 'xls'])

# Regex patterns for validation (compiled once for performance)
import re

RELEASE_PATTERN = re.compile(r'^\d{4}\.M\d{2}$')  # Format: YYYY.MXX
TASK_ID_PATTERN = re.compile(r'^[A-Z]{2,4}\d{3,6}$')  # Format: ABC123456
ENTERPRISE_ID_PATTERN = re.compile(r'^REL\d{7}$')  # Format: REL1234567

# Performance settings
CHUNK_SIZE = 1000  # For processing large datasets
MAX_MEMORY_USAGE = 500 * 1024 * 1024  # 500MB max memory usage
CACHE_SIZE = 100  # Maximum cached items

# Logging configuration
LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
        },
        'detailed': {
            'format': '%(asctime)s [%(levelname)s] %(name)s:%(lineno)d: %(message)s'
        }
    },
    'handlers': {
        'default': {
            'level': 'INFO',
            'formatter': 'standard',
            'class': 'logging.StreamHandler',
        },
        'file': {
            'level': 'DEBUG',
            'formatter': 'detailed',
            'class': 'logging.FileHandler',
            'filename': LOGS_DIR / 'app.log',
            'mode': 'a',
        },
    },
    'loggers': {
        '': {
            'handlers': ['default', 'file'],
            'level': 'DEBUG',
            'propagate': False
        }
    }
}

# Create necessary directories
def ensure_directories():
    """Create required directories if they don't exist"""
    for directory in [DATA_DIR, LOGS_DIR, TEMP_DIR]:
        directory.mkdir(parents=True, exist_ok=True)


# Initialize directories on import
ensure_directories()
