"""WMG Feedback Generator

A Python library for generating student feedback documents from Excel data and Word templates.
"""

__version__ = "0.1.0"
__author__ = "Dr David Croft"
__email__ = "david.croft@warwick.ac.uk"

# Import main classes/functions for easy access
from .core import (
    find_columns,
    process_to_dicts,
    validate_row_data,
    gen_filename,
    default_validators
)
from .document_generator import generate

__all__ = [
    "find_columns",
    "process_to_dicts", 
    "validate_row_data",
    "gen_filename",
    "default_validators",
    "generate",
    "category"
]
