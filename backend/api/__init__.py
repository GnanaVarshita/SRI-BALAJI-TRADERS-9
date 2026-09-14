"""
API Routes Package
Sri Balaji Traders Automation System
"""

from .routes_system import (
    WebLogRedirector,
    run_sync_task,
    handle_get_config,
    handle_save_config,
    handle_get_status,
    handle_start_sync,
    handle_reset_sync,
    handle_open_folder,
)
from .routes_dialogs import (
    handle_browse_file,
    handle_browse_folder,
)
from .static_server import (
    get_content_type,
    serve_static,
    scan_downloaded_files,
    handle_view_file,
)
from .routes_excel import (
    handle_process_excel,
    handle_generate_po_summary,
    handle_generate_corteva_master_summary,
    handle_generate_fmc_summary,
    handle_generate_fmc_step2,
    handle_format_tbm_summaries,
    handle_generate_tbm_summary,
    handle_sync_tbm_cards,
    handle_generate_invoices,
    handle_scan_pos_in_summary,
    handle_sync_details_of_bills,
)

__all__ = [
    'WebLogRedirector',
    'run_sync_task',
    'handle_get_config',
    'handle_save_config',
    'handle_get_status',
    'handle_start_sync',
    'handle_reset_sync',
    'handle_open_folder',
    'handle_browse_file',
    'handle_browse_folder',
    'get_content_type',
    'serve_static',
    'scan_downloaded_files',
    'handle_view_file',
    'handle_process_excel',
    'handle_generate_po_summary',
    'handle_generate_corteva_master_summary',
    'handle_generate_fmc_summary',
    'handle_generate_fmc_step2',
    'handle_format_tbm_summaries',
    'handle_generate_tbm_summary',
    'handle_sync_tbm_cards',
    'handle_generate_invoices',
    'handle_scan_pos_in_summary',
    'handle_sync_details_of_bills',
]
