"""
Details of Bills Generator Facade Module
Sri Balaji Traders Automation System

Re-exports from backend.details_of_bills and backend.common for 100% backward compatibility.
"""

from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from common.formatters import (
    clean_str,
    parse_num,
    parse_date_to_obj,
    format_date_str,
    calculate_receivable_date,
    normalize_po,
)
from common.styles import get_details_styles
from details_of_bills.extractor import extract_invoice_file_data
from details_of_bills.cards_updater import update_budget_po_summary_cards
from details_of_bills.manager import (
    create_or_load_details_of_bills_wb,
    scan_and_append_invoices,
)

__all__ = [
    'clean_str',
    'parse_num',
    'parse_date_to_obj',
    'format_date_str',
    'calculate_receivable_date',
    'normalize_po',
    'get_details_styles',
    'create_or_load_details_of_bills_wb',
    'extract_invoice_file_data',
    'update_budget_po_summary_cards',
    'scan_and_append_invoices',
]
