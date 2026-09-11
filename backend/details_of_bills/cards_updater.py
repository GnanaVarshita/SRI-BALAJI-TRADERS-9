"""
Budget PO Summary Cards Updater
Updates I.V NO and DATE in budget PO summary cards matching generated invoices.
"""

from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

try:
    from excel_parser import load_any_workbook
except ImportError:
    from ..excel_parser import load_any_workbook

from common.formatters import clean_str, normalize_po
from common.styles import get_details_styles


def update_budget_po_summary_cards(cards_excel_path, invoice_records_list):
    """
    Scans the Budget PO Summary Cards Excel (e.g. Nandyala FMC Budget.xlsx)
    and populates I.V NO (Column A) and DATE (Column B) for matching PO cards.
    """
    cards_file = Path(cards_excel_path).resolve()
    if not cards_file.exists():
        return {"success": False, "message": f"Cards file not found: {cards_excel_path}"}

    wb = load_any_workbook(cards_file)
    styles = get_details_styles()

    cards_updated = 0
    updated_pos = set()

    for inv_data in invoice_records_list:
        target_po = inv_data.get('po_number', '').strip()
        norm_target_po = normalize_po(target_po)
        short_iv = inv_data.get('short_iv', '')
        inv_date = inv_data.get('invoice_date', '')

        if not norm_target_po:
            continue

        for sname in wb.sheetnames:
            ws = wb[sname]
            max_r = ws.max_row

            for r in range(1, max_r + 1):
                c1_val = clean_str(ws.cell(r, 1).value)
                norm_c1 = normalize_po(c1_val)

                if norm_target_po in norm_c1 or norm_c1 == norm_target_po:
                    for tbl_r in range(r + 1, min(r + 15, max_r + 1)):
                        cell_a_hdr = clean_str(ws.cell(tbl_r, 1).value).upper().replace(".", "").replace(" ", "")
                        if "IVNO" in cell_a_hdr:
                            card_data_r = tbl_r + 1
                            
                            cell_iv = ws.cell(card_data_r, 1, short_iv)
                            cell_iv.font = styles['font_header_green']
                            cell_iv.alignment = styles['align_center']

                            cell_dt = ws.cell(card_data_r, 2, inv_date)
                            cell_dt.font = styles['font_header_green']
                            cell_dt.alignment = styles['align_center']

                            cards_updated += 1
                            updated_pos.add(target_po)
                            break

    wb.save(cards_file)
    wb.close()

    return {
        "success": True,
        "cardsUpdated": cards_updated,
        "updatedPos": list(updated_pos)
    }
