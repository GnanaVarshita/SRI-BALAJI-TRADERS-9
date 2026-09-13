"""
Details of Bills Manager
Orchestrates creating/loading Details of Bills master workbook and appending invoices.
"""

import re
from pathlib import Path
import sys
import openpyxl

BACKEND_DIR = Path(__file__).resolve().parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

try:
    from excel_parser import load_any_workbook
except ImportError:
    from ..excel_parser import load_any_workbook

from common.formatters import clean_str
from common.styles import get_details_styles
from .extractor import extract_invoice_file_data
from .cards_updater import update_budget_po_summary_cards


def get_details_of_bills_sheet(wb):
    """
    Finds and returns Sheet 1 (the bill entries worksheet) in Details of Bills workbook.
    NEVER returns Sheet 2 or secondary summary/dashboard sheets.
    """
    # 1. Exact match for 'Sheet1' or 'Sheet 1'
    for name in ['Sheet1', 'Sheet 1']:
        if name in wb.sheetnames:
            return wb[name]

    # 2. Find any sheet with Details of Bills header (e.g. 'I.V NO' or 'IV NO')
    # strictly ignoring 'Sheet2' or 'Sheet 2'
    for name in wb.sheetnames:
        if name.strip().lower() in ['sheet2', 'sheet 2']:
            continue
        ws_candidate = wb[name]
        for r in range(1, 10):
            for c in range(1, 10):
                val = clean_str(ws_candidate.cell(r, c).value).upper().replace(".", "").replace(" ", "")
                if "IVNO" in val:
                    return ws_candidate

    # 3. Fallback to first sheet if not Sheet2
    if len(wb.worksheets) > 0 and wb.worksheets[0].title.strip().lower() not in ['sheet2', 'sheet 2']:
        return wb.worksheets[0]

    # 4. If no valid Sheet1 exists, create one as the first sheet
    return wb.create_sheet(title='Sheet1', index=0)


def create_or_load_details_of_bills_wb(file_path, financial_year="APRIL 2026 to MARCH 2027"):
    """
    Initializes a new Details of Bills workbook if not present or empty, matching SS1 structure.
    Always operates strictly on Sheet 1 and never touches Sheet 2.
    """
    target = Path(file_path).resolve()
    styles = get_details_styles()

    if target.exists():
        try:
            wb = openpyxl.load_workbook(target, data_only=False)
        except Exception:
            wb = load_any_workbook(target)
        ws = get_details_of_bills_sheet(wb)
        # Check if ws already has headers (row 4 or rows 1-6)
        has_headers = False
        for r in range(1, 10):
            for c in range(1, 10):
                val = clean_str(ws.cell(r, c).value).upper().replace(".", "").replace(" ", "")
                if "IVNO" in val:
                    has_headers = True
                    break
            if has_headers:
                break
        if has_headers:
            return wb, ws, False
    else:
        wb = openpyxl.Workbook()
        ws = get_details_of_bills_sheet(wb)
        ws.title = "Sheet1"

    # Initialize full SS1 structure
    ws.views.sheetView[0].showGridLines = True

    col_widths = {
        'A': 8,   # I.V NO
        'B': 13,  # DATE
        'C': 12,  # AREA
        'D': 16,  # PO NUMBER
        'E': 10,  # BUDGET
        'F': 14,  # PRODUCT
        'G': 14,  # CROP
        'H': 14,  # ACTIVITY
        'I': 14,  # ZDGM
        'J': 16,  # Tbm
        'K': 10,  # NO.OF.Ac
        'L': 14,  # AMOUNT
        'M': 14,  # SERVICE IV
        'N': 15,  # TOTAL IV
        'O': 14,  # Grand Total
        'P': 15,  # RECEIVABLE DATE
        'Q': 14,  # RECEIVED
        'R': 12,  # TDS
        'S': 14,  # Received Date
        'T': 16,  # Receivable Amount
        'U': 16,  # Total Amount
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    ws.row_dimensions[1].height = 15

    # Row 2: Financial Year title
    ws.row_dimensions[2].height = 22
    ws.merge_cells('H2:K2')
    c_title = ws['H2']
    c_title.value = financial_year
    c_title.font = styles['font_title']
    c_title.alignment = styles['align_center']

    # Row 3: Top Summary Totals Formulas
    ws.row_dimensions[3].height = 20
    ws['L3'] = "=SUM(L5:L5)"
    ws['M3'] = "=SUM(M5:M5)"
    ws['N3'] = "=SUM(N5:N5)"
    ws['O3'] = "=SUM(O5:O5)"
    ws['Q3'] = "=SUM(Q5:Q5)"
    ws['R3'] = "=SUM(R5:R5)"
    ws['T3'] = "=O3-Q3-R3"
    ws['U3'] = "=Q3+R3+T3"

    for c_idx in [12, 13, 14, 15, 17, 18, 20, 21]:
        cell = ws.cell(3, c_idx)
        cell.font = styles['font_summary']
        cell.alignment = styles['align_right']
        cell.border = styles['box_border']
        cell.number_format = '#,##0.00'

    # Row 4: Column Headers
    ws.row_dimensions[4].height = 26
    headers_spec = [
        (1, "I.V NO", styles['font_header_green']),
        (2, "DATE", styles['font_header_red']),
        (3, "AREA", styles['font_header_brown']),
        (4, "PO NUMBER", styles['font_header_red']),
        (5, "BUDGET", styles['font_header_red']),
        (6, "PRODUCT", styles['font_header_red']),
        (7, "CROP", styles['font_header_red']),
        (8, "ACTIVITY", styles['font_header_red']),
        (9, "ZDGM", styles['font_header_red']),
        (10, "Tbm", styles['font_header_black']),
        (11, "NO.OF\n.Ac", styles['font_header_red']),
        (12, "AMOUNT", styles['font_header_red']),
        (13, "SERVICE IV", styles['font_header_red']),
        (14, "TOTAL IV", styles['font_header_red']),
        (15, "Grand Total", styles['font_header_red']),
        (16, "RECEIVABL\nE DATE", styles['font_header_red']),
        (17, "RECEIVED", styles['font_header_black']),
        (18, "TDS", styles['font_header_red']),
        (19, "Receive\nd Date", styles['font_header_red']),
        (20, "Receivable\nAmount", styles['font_header_red']),
        (21, "Total Amount", styles['font_header_black']),
    ]

    for c_idx, text, font_style in headers_spec:
        cell = ws.cell(4, c_idx, text)
        cell.font = font_style
        cell.alignment = styles['align_center_wrap']
        cell.border = styles['box_border']

    return wb, ws, True


def scan_and_append_invoices(
    details_excel_path,
    invoices_folder_path,
    budget_cards_path=None,
    financial_year="APRIL 2026 to MARCH 2027"
):
    """
    Main orchestration function:
    1. Loads or initializes Details of Bills workbook (Sheet1).
    2. Scans invoices folder and extracts bill details taking reference from Sheet 1.
    3. Appends new / missing invoices into Sheet1.
    4. Recalculates top summary formulas in Row 3.
    5. Syncs IV No and Date into Budget PO Summary Cards if provided.
    """
    invoices_folder = Path(invoices_folder_path).resolve()
    if not invoices_folder.exists() or not invoices_folder.is_dir():
        raise FileNotFoundError(f"Invoices folder not found: {invoices_folder_path}")

    wb, ws, is_new = create_or_load_details_of_bills_wb(details_excel_path, financial_year)
    styles = get_details_styles()

    existing_ivs = set()
    for r in range(5, ws.max_row + 1):
        iv_val = clean_str(ws.cell(r, 1).value)
        if iv_val:
            existing_ivs.add(iv_val.upper())
            m_num = re.search(r'(\d+)$', iv_val)
            if m_num:
                existing_ivs.add(str(int(m_num.group(1))))

    def get_invoice_sort_key(p):
        m = re.search(r'(\d+)', p.stem)
        return int(m.group(1)) if m else 999999

    invoice_files = sorted(list(invoices_folder.glob("*.xlsx")), key=get_invoice_sort_key)
    if not invoice_files:
        raise ValueError(f"No invoice Excel (.xlsx) files found in: {invoices_folder_path}")

    parsed_invoices = []
    appended_invoices = []
    skipped_invoices = []
    total_new_rows = 0

    current_r = 5
    for check_r in range(5, max(ws.max_row + 2, 6)):
        if not any(clean_str(ws.cell(check_r, c).value) for c in range(1, 15)):
            current_r = check_r
            break

    for inv_path in invoice_files:
        inv_data = extract_invoice_file_data(inv_path)
        if not inv_data:
            continue

        parsed_invoices.append(inv_data)
        short_iv = inv_data['short_iv']
        full_iv = inv_data['invoice_no']

        if short_iv.upper() in existing_ivs or full_iv.upper() in existing_ivs:
            skipped_invoices.append(short_iv)
            continue

        items = inv_data['items']
        num_items = len(items)

        for idx, item in enumerate(items, start=1):
            ws.row_dimensions[current_r].height = 18

            if idx == 1:
                ws.cell(current_r, 1, short_iv).font = styles['font_header_green']
                ws.cell(current_r, 1).alignment = styles['align_center']
            else:
                ws.cell(current_r, 1, "").alignment = styles['align_center']

            ws.cell(current_r, 2, item.get('date', '')).alignment = styles['align_center']
            ws.cell(current_r, 3, item.get('area', '')).alignment = styles['align_left']
            ws.cell(current_r, 4, item.get('po_number', '')).alignment = styles['align_center']
            ws.cell(current_r, 5, item.get('budget', 'Brand')).alignment = styles['align_center']
            ws.cell(current_r, 6, item.get('product', '')).alignment = styles['align_left']
            ws.cell(current_r, 7, item.get('crop', '')).alignment = styles['align_left']
            ws.cell(current_r, 8, item.get('activity', '')).alignment = styles['align_left']
            ws.cell(current_r, 9, item.get('zdgm', '')).alignment = styles['align_left']
            ws.cell(current_r, 10, item.get('tbm', '')).alignment = styles['align_left']

            c_qty = ws.cell(current_r, 11, item.get('qty', 1))
            c_qty.alignment = styles['align_center']
            c_qty.number_format = '#,##0'

            c_amt = ws.cell(current_r, 12, item.get('amount', 0.0))
            c_amt.alignment = styles['align_right']
            c_amt.number_format = '#,##0'

            c_siv = ws.cell(current_r, 13, item.get('service_iv', 0.0))
            c_siv.alignment = styles['align_right']
            c_siv.number_format = '#,##0.00'

            c_tiv = ws.cell(current_r, 14, item.get('total_iv', 0.0))
            c_tiv.alignment = styles['align_right']
            c_tiv.number_format = '#,##0.00'

            if idx == num_items:
                c_gt = ws.cell(current_r, 15, inv_data.get('grand_total', 0.0))
                c_gt.font = styles['font_bold']
                c_gt.alignment = styles['align_right']
                c_gt.number_format = '#,##0'

            ws.cell(current_r, 16, inv_data.get('receivable_date', '')).alignment = styles['align_center']

            for c in range(17, 22):
                ws.cell(current_r, c, "")

            for c in range(1, 22):
                cell = ws.cell(current_r, c)
                if c not in [1, 15]:
                    cell.font = styles['font_regular']
                cell.border = styles['box_border']

            current_r += 1
            total_new_rows += 1

        appended_invoices.append(short_iv)
        existing_ivs.add(short_iv.upper())

    final_max_r = max(5, current_r - 1)
    ws['L3'] = f"=SUM(L5:L{final_max_r})"
    ws['M3'] = f"=SUM(M5:M{final_max_r})"
    ws['N3'] = f"=SUM(N5:N{final_max_r})"
    ws['O3'] = f"=SUM(O5:O{final_max_r})"
    ws['Q3'] = f"=SUM(Q5:Q{final_max_r})"
    ws['R3'] = f"=SUM(R5:R{final_max_r})"
    ws['T3'] = "=O3-Q3-R3"
    ws['U3'] = "=Q3+R3+T3"

    # Always ensure Sheet1 is active when saving workbook so Sheet2 is never active or exposed
    wb.active = ws

    out_details_path = Path(details_excel_path).resolve()
    out_details_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_details_path)
    wb.close()

    cards_sync_res = None
    if budget_cards_path and str(budget_cards_path).strip():
        cards_sync_res = update_budget_po_summary_cards(budget_cards_path, parsed_invoices)

    return {
        "success": True,
        "isNew": is_new,
        "detailsExcelPath": str(out_details_path),
        "totalInvoicesFound": len(invoice_files),
        "appendedInvoices": appended_invoices,
        "skippedInvoices": skipped_invoices,
        "totalRowsAdded": total_new_rows,
        "totalRowsInSheet": final_max_r - 4,
        "cardsSync": cards_sync_res,
        "message": f"Successfully processed {len(invoice_files)} invoices. Appended {len(appended_invoices)} new invoices ({total_new_rows} rows) to Details of Bills."
    }
