"""
TBM Formatter Facade Module
Sri Balaji Traders Automation System

Re-exports from backend.tbm for 100% backward compatibility.
"""

from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from common.formatters import clean_str, parse_num, parse_date_intelligent
from tbm.constants import FIELD_KEYWORDS
from tbm.reader import extract_activities_from_sheet, parse_slip_format
from tbm.formatter import (
    create_green_styles,
    build_grouped_tables,
    write_formatted_group_table,
    write_po_and_grand_totals_block,
    format_tbm_workbook,
    format_all_tbm_summaries_in_folder,
)

__all__ = [
    'clean_str',
    'parse_num',
    'parse_date_intelligent',
    'FIELD_KEYWORDS',
    'extract_activities_from_sheet',
    'parse_slip_format',
    'create_green_styles',
    'build_grouped_tables',
    'write_formatted_group_table',
    'write_po_and_grand_totals_block',
    'format_tbm_workbook',
    'format_all_tbm_summaries_in_folder',
]

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        target = sys.argv[1]
        p = Path(target)
        if p.is_dir():
            res = format_all_tbm_summaries_in_folder(p)
            print(res["message"])
        elif p.is_file():
            res = format_tbm_workbook(p)
            print(res["message"])
    else:
        print("Usage: python tbm_formatter.py <tbm_summary_folder_or_file>")
