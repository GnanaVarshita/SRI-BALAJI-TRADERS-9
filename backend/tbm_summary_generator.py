"""
TBM Summary Generator Facade Module
Sri Balaji Traders Automation System

Re-exports from backend.tbm for 100% backward compatibility.
"""

from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from common.formatters import clean_str, format_date_val, format_date_xlrd, parse_num
from tbm.constants import FIELD_KEYWORDS
from tbm.reader import (
    extract_activities_from_sheet,
    extract_activities_from_xlrd_sheet,
    parse_slip_format,
)
from tbm.consolidator import (
    create_table_styles,
    write_table_to_sheet,
    create_tbm_amount_summary_sheet,
    generate_tbm_summary,
)

__all__ = [
    'clean_str',
    'format_date_val',
    'format_date_xlrd',
    'parse_num',
    'FIELD_KEYWORDS',
    'extract_activities_from_sheet',
    'extract_activities_from_xlrd_sheet',
    'parse_slip_format',
    'create_table_styles',
    'write_table_to_sheet',
    'create_tbm_amount_summary_sheet',
    'generate_tbm_summary',
]
