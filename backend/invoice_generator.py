"""
Invoice Generator Module (Facade)
Sri Balaji Traders Automation System

This module re-exports components from backend.invoices and backend.common
for clean backward compatibility with existing tests and scripts.
"""

from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from common.formatters import (
    clean_str,
    parse_num,
    format_date_val,
    parse_date_to_obj,
    format_date_str,
    normalize_po,
    format_short_iv,
    format_full_iv,
)
from common.currency import num_to_indian_words
from common.styles import (
    get_base_styles,
    setup_page_print_fit,
    write_sbt_header_block,
    write_bank_and_signature_block,
)
from invoices.extractor import (
    scan_pos_in_summary,
    extract_tables_for_po,
)
from invoices.corteva import (
    render_corteva_invoice_block,
    build_corteva_sheet1_invoice,
    build_corteva_sheet2_details,
    build_corteva_summary_sheet,
)
from invoices.fmc import (
    render_fmc_invoice_block,
    build_fmc_sheet1_invoice,
    build_fmc_sheet2_details,
)
from invoices.summary_table import render_details_cumulative_summary_table
from invoices.generator import generate_or_update_invoice

__all__ = [
    'clean_str',
    'parse_num',
    'format_date_val',
    'parse_date_to_obj',
    'format_date_str',
    'normalize_po',
    'format_short_iv',
    'format_full_iv',
    'num_to_indian_words',
    'get_base_styles',
    'setup_page_print_fit',
    'write_sbt_header_block',
    'write_bank_and_signature_block',
    'scan_pos_in_summary',
    'extract_tables_for_po',
    'render_corteva_invoice_block',
    'build_corteva_sheet1_invoice',
    'build_corteva_sheet2_details',
    'build_corteva_summary_sheet',
    'render_fmc_invoice_block',
    'build_fmc_sheet1_invoice',
    'build_fmc_sheet2_details',
    'render_details_cumulative_summary_table',
    'generate_or_update_invoice',
]
