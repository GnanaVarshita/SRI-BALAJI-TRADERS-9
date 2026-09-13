"""
Card Sync Engine Facade Module
Sri Balaji Traders Automation System

Re-exports from backend.card_sync for 100% backward compatibility.
"""

import sys
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from card_sync.constants import (
    ACTIVITY_NORMALIZATION,
    normalize_activity,
    match_activity_to_column,
)
from card_sync.extractor import (
    find_summary_files,
    extract_all_tbms_spent,
    mark_tbms_as_synced,
)
from card_sync.corteva_sync import (
    sync_corteva_cards_workbook,
)
from card_sync.fmc_sync import (
    write_sheet_end_summary_table,
    sync_fmc_cards_workbook,
)
from card_sync.engine import (
    sync_tbm_with_cards,
)

# Backward-compatible alias
def extract_tbm_spent_summary(tbm_summary_path):
    tbm_data, _, _, _ = extract_all_tbms_spent(tbm_summary_path)
    return tbm_data

__all__ = [
    'ACTIVITY_NORMALIZATION',
    'normalize_activity',
    'match_activity_to_column',
    'find_summary_files',
    'extract_all_tbms_spent',
    'extract_tbm_spent_summary',
    'mark_tbms_as_synced',
    'sync_corteva_cards_workbook',
    'write_sheet_end_summary_table',
    'sync_fmc_cards_workbook',
    'sync_tbm_with_cards',
]

if __name__ == "__main__":
    if len(sys.argv) > 2:
        c_path = sys.argv[1]
        t_path = sys.argv[2]
        sv_p = float(sys.argv[3]) if len(sys.argv) > 3 else 5.0
        res = sync_tbm_with_cards(c_path, t_path, service_charge_percent=sv_p)
        print(res["message"])
    else:
        print("Usage: python card_sync_engine.py <cards_excel_path> <tbm_summary_excel_path_or_folder> [service_charge_percent]")
