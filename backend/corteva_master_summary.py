import os
import re
import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from excel_parser import load_any_workbook
except ImportError:
    try:
        from .excel_parser import load_any_workbook
    except ImportError:
        load_any_workbook = None


def _format_date(val):
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime('%d-%m-%Y')
    s = str(val).strip()
    # Try parsing common date formats
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m:
        y, mth, d = m.groups()
        return f"{int(d):02d}-{int(mth):02d}-{y}"
    m = re.search(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s)
    if m:
        d, mth, y = m.groups()
        return f"{int(d):02d}-{int(mth):02d}-{y}"
    return s.split(' ')[0] if ' ' in s else s


def _safe_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(',', '').replace('₹', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def extract_po_cards_from_folder(folder_path):
    """
    Scans all Excel workbooks in folder_path, iterates through all worksheets,
    and extracts PO summary card details.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Folder does not exist: {folder_path}")

    # Find all .xlsx and .xls files, ignoring temporary files
    excel_files = []
    for ext in ['*.xlsx', '*.xls']:
        for p in folder.glob(ext):
            base = p.name
            if base.startswith('~$'):
                continue
            # Ignore previously generated master summaries
            if 'master' in base.lower() or 'pos & expenditures' in base.lower():
                continue
            excel_files.append(p)

    excel_files.sort(key=lambda x: x.name.lower())

    detected_territory = None
    detected_zdgm = None
    po_cards = []

    for file_path in excel_files:
        try:
            wb_values = openpyxl.load_workbook(str(file_path), data_only=True)
        except Exception:
            if load_any_workbook:
                wb_values = load_any_workbook(str(file_path))
            else:
                continue

        try:
            wb_formulas = openpyxl.load_workbook(str(file_path), data_only=False)
        except Exception:
            wb_formulas = wb_values

        for sheet_name in wb_values.sheetnames:
            norm_name = sheet_name.strip().lower()
            if norm_name in ['sheet1', 'processed_emails', 'summary', 'master'] or norm_name.startswith('sheet'):
                continue

            ws_v = wb_values[sheet_name]
            ws_f = wb_formulas[sheet_name] if sheet_name in wb_formulas.sheetnames else ws_v

            # 1. Locate the summary table header containing "Activities" & "BUDGET"
            hdr_r, act_c = None, None
            for r in range(8, 26):
                for c in range(12, 28):
                    v = ws_v.cell(r, c).value
                    if v and 'activit' in str(v).lower():
                        # Verify adjacent column has budget or header
                        next_v = ws_v.cell(r, c + 1).value
                        if next_v and 'budget' in str(next_v).lower():
                            hdr_r, act_c = r, c
                            break
                        hdr_r, act_c = r, c
                        break
                if hdr_r:
                    break

            if not hdr_r:
                # No activities summary table in this sheet, skip
                continue

            # Identify column offsets for the summary table:
            # Expected relative to act_c:
            # 0: Activities, 1: BUDGET, 2: SPENT, 3: SV Charge, 4: TOTAL IV, 5: BALANCE
            offset_budget = 1
            offset_spent = 2
            offset_sv = 3
            offset_total_iv = 4
            offset_balance = 5

            for c_offset in range(1, 8):
                hdr_val = str(ws_v.cell(hdr_r, act_c + c_offset).value or '').lower()
                if 'budget' in hdr_val:
                    offset_budget = c_offset
                elif 'spent' in hdr_val:
                    offset_spent = c_offset
                elif 'total' in hdr_val and 'iv' in hdr_val:
                    offset_total_iv = c_offset
                elif 'balance' in hdr_val:
                    offset_balance = c_offset

            # 2. Extract activities
            activities = []
            for r in range(hdr_r + 1, hdr_r + 35):
                val_act = ws_v.cell(r, act_c).value
                if val_act is None:
                    val_act = ws_f.cell(r, act_c).value

                if val_act is None:
                    continue

                act_str = str(val_act).strip()
                if act_str.lower() in ['total', 'totals']:
                    break
                if act_str.lower() in ['', '0', '0.0', 'none']:
                    continue

                # If formula like =T3 or =R3, resolve cell reference
                if act_str.startswith('='):
                    ref = act_str[1:]
                    if ref in ws_v:
                        act_str = str(ws_v[ref].value or ws_f[ref].value or '').strip()
                if not act_str or act_str.lower() in ['total', 'activities', '0', 'none']:
                    continue

                def _get_val(c_offset):
                    v = ws_v.cell(r, act_c + c_offset).value
                    if v is None:
                        vf = ws_f.cell(r, act_c + c_offset).value
                        if vf and str(vf).startswith('='):
                            ref = str(vf)[1:]
                            if ref in ws_v:
                                v = ws_v[ref].value
                    return _safe_float(v)

                budget_amt = _get_val(offset_budget)
                spent_amt = _get_val(offset_spent)
                total_iv_amt = _get_val(offset_total_iv)
                bal_amt = _get_val(offset_balance)

                expenses = total_iv_amt if total_iv_amt > 0 else (spent_amt if spent_amt > 0 else 0.0)
                calculated_balance = budget_amt - expenses
                final_balance = bal_amt if (bal_amt > 0 or bal_amt < 0) else calculated_balance

                activities.append({
                    'activity': act_str,
                    'amount': budget_amt,
                    'expenses': expenses,
                    'balance': final_balance
                })

            if not activities:
                continue

            # 3. Extract PO card metadata
            # PO Number: A6 (some might not have PO number, leave blank)
            po_raw = ws_v.cell(6, 1).value
            po_number = ""
            if po_raw is not None:
                po_s = str(po_raw).strip()
                if po_s.lower() not in ['none', '0', '']:
                    if po_s.endswith('.0'):
                        po_s = po_s[:-2]
                    po_number = po_s

            # Budget Type: A7 (e.g. MA, MKTG)
            budget_type_raw = ws_v.cell(7, 1).value
            budget_type = str(budget_type_raw).strip() if budget_type_raw else ""
            if not budget_type:
                m_type = re.search(r'\((MA|MKTG)\)', sheet_name, re.I)
                if m_type:
                    budget_type = m_type.group(1).upper()
                elif 'mktg' in norm_name or 'mk' in norm_name:
                    budget_type = 'MKTG'
                elif 'ma' in norm_name:
                    budget_type = 'MA'
                else:
                    budget_type = 'MA'

            # Contact & Territory: D7 (e.g. "Bhaskar - Kurnool" or "R.Bhaskar - Kurnool")
            contact_raw = ws_v.cell(7, 4).value
            contact_str = str(contact_raw).strip() if contact_raw else ""
            if contact_str and '-' in contact_str:
                parts = [p.strip() for p in contact_str.split('-')]
                if len(parts) >= 2:
                    if not detected_zdgm:
                        detected_zdgm = parts[0]
                    if not detected_territory:
                        detected_territory = parts[1]

            # Product: F6 (or sheet title fallback)
            product_raw = ws_v.cell(6, 6).value
            product = str(product_raw).strip() if product_raw else sheet_name.strip()
            if product.lower() in ['product', 'produ', 'none', '']:
                product = sheet_name.strip()

            # Crop: J6
            crop_raw = ws_v.cell(6, 10).value
            crop = str(crop_raw).strip() if crop_raw and str(crop_raw).lower() != 'none' else ""

            # Date: M6
            date_raw = ws_v.cell(6, 13).value
            if date_raw is None:
                for col_idx in [12, 14, 15]:
                    v_dt = ws_v.cell(6, col_idx).value
                    if v_dt and isinstance(v_dt, (datetime.date, datetime.datetime)):
                        date_raw = v_dt
                        break
            po_date = _format_date(date_raw)

            po_cards.append({
                'file': file_path.name,
                'sheet': sheet_name,
                'po_number': po_number,
                'date': po_date,
                'budget_type': budget_type,
                'contact': contact_str,
                'product': product,
                'crop': crop,
                'activities': activities
            })

    # Territory fallback from folder name if not yet detected
    if not detected_territory:
        folder_str = str(folder).upper()
        for t in ['KURNOOL', 'NELLORE', 'SURYAPET', 'NANDYAL', 'NANDYALA']:
            if t in folder_str:
                detected_territory = t.title()
                break

    if not detected_territory:
        detected_territory = "Kurnool"

    if not detected_zdgm:
        detected_zdgm = "Bhaskar"

    return po_cards, detected_territory, detected_zdgm


def generate_master_po_summary(folder_path, output_name=None, territory=None, zdgm=None, budget_season="Kharif"):
    """
    Scans the folder containing PO summary card workbooks, extracts all PO sheets,
    and generates a master PO summary workbook in the same folder formatted identically to Screenshot 1.
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Input folder does not exist: {folder_path}")

    po_cards, auto_territory, auto_zdgm = extract_po_cards_from_folder(folder)

    if not po_cards:
        raise ValueError(f"No valid Corteva PO summary card sheets were found in: {folder_path}")

    final_territory = territory.strip() if territory and territory.strip() else auto_territory
    final_zdgm = zdgm.strip() if zdgm and zdgm.strip() else auto_zdgm
    zdgm_display = final_zdgm.replace('ZDGM:', '').replace('ZDGM', '').strip()

    if not output_name or not output_name.strip():
        output_name = f"{final_territory} Master PO Summary.xlsx"
    elif not output_name.lower().endswith('.xlsx'):
        output_name += ".xlsx"

    output_path = folder / output_name

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pos & Expenditures"
    ws.views.sheetView[0].showGridLines = True

    # --- Styling definition matching Screenshot 1 ---
    side_thin = Side(style='thin', color='000000')
    thin_border = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)

    header_fill = PatternFill(start_color='C2D69B', end_color='C2D69B', fill_type='solid')

    font_title_zdgm = Font(name='Calibri', size=11, bold=True, color='008000')  # Green bold
    font_title_banner = Font(name='Calibri', size=13, bold=True, color='FF0000') # Red bold
    font_header = Font(name='Calibri', size=10, bold=True, color='FF0000')       # Red bold
    font_green_bold = Font(name='Calibri', size=10, bold=True, color='008000')   # Green bold
    font_black_bold = Font(name='Calibri', size=10, bold=True, color='000000')   # Black bold

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # Column widths
    col_widths = {
        'A': 8,   # S.NO
        'B': 14,  # PO DATE
        'C': 15,  # PO NUMBER
        'D': 10,  # BUDGET
        'E': 14,  # BUDGET TYPE
        'F': 24,  # PRODUCT
        'G': 14,  # CROP
        'H': 24,  # ACTIVITY
        'I': 14,  # AMOUNT
        'J': 14,  # EXPENSES
        'K': 14,  # BALANCE
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # Row heights
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 10
    ws.row_dimensions[3].height = 26
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 26

    # --- Row 1: ZDGM Banner ---
    ws.cell(1, 8).value = f"ZDGM: {zdgm_display}"
    ws.cell(1, 8).font = font_title_zdgm
    ws.cell(1, 8).alignment = align_center

    # --- Row 3: Territory Pos & Expenditures Banner ---
    ws.merge_cells('F3:H3')
    ws['F3'].value = f"{final_territory} Pos & Expenditures"
    ws['F3'].font = font_title_banner
    ws['F3'].alignment = align_center

    # --- Row 5: Table Column Headers ---
    HEADERS = [
        'S.NO', 'PO DATE', 'PO NUMBER', 'BUDGET', 'BUDGET TYPE',
        'PRODUCT', 'CROP', 'ACTIVITY', 'AMOUNT', 'EXPENSES', 'BALANCE'
    ]
    for col_idx, hdr in enumerate(HEADERS, start=1):
        c = ws.cell(5, col_idx, hdr)
        c.font = font_header
        c.fill = header_fill
        c.alignment = align_center
        c.border = thin_border

    # --- Row 6+: Data Rows ---
    curr_row = 6
    s_no = 1
    total_amount_calc = 0.0
    total_expenses_calc = 0.0
    total_activities_count = 0

    for card in po_cards:
        activities = card['activities']
        for act_idx, act in enumerate(activities):
            r = curr_row
            ws.row_dimensions[r].height = 20

            if act_idx == 0:
                # S.NO
                cA = ws.cell(r, 1, s_no)
                cA.font = font_black_bold
                cA.alignment = align_center

                # PO DATE
                cB = ws.cell(r, 2, card['date'])
                cB.font = font_green_bold
                cB.alignment = align_center

                # PO NUMBER (leave blank if None/missing)
                cC = ws.cell(r, 3, card['po_number'] if card['po_number'] else "")
                cC.font = font_green_bold
                cC.alignment = align_center

                # BUDGET (e.g. Kharif)
                cD = ws.cell(r, 4, budget_season or "Kharif")
                cD.font = font_green_bold
                cD.alignment = align_center

                # BUDGET TYPE (e.g. MA, MKTG)
                cE = ws.cell(r, 5, card['budget_type'])
                cE.font = font_green_bold
                cE.alignment = align_center

                # PRODUCT
                cF = ws.cell(r, 6, card['product'])
                cF.font = font_green_bold
                cF.alignment = align_left

                # CROP
                cG = ws.cell(r, 7, card['crop'])
                cG.font = font_green_bold
                cG.alignment = align_left
            else:
                # Blank for subsequent activity rows of same sheet
                for c_idx in range(1, 8):
                    ws.cell(r, c_idx, None)

            # ACTIVITY
            cH = ws.cell(r, 8, act['activity'])
            cH.font = font_green_bold
            cH.alignment = align_left

            # AMOUNT
            cI = ws.cell(r, 9, act['amount'])
            cI.font = font_black_bold
            cI.alignment = align_right
            cI.number_format = '0.##'

            # EXPENSES
            cJ = ws.cell(r, 10, act['expenses'])
            cJ.font = font_black_bold
            cJ.alignment = align_right
            cJ.number_format = '0.##'

            # BALANCE formula: =I{r}-J{r}
            cK = ws.cell(r, 11, f'=I{r}-J{r}')
            cK.font = font_green_bold
            cK.alignment = align_right
            cK.number_format = '0.##'

            # Borders on all cells A to K
            for c_idx in range(1, 12):
                ws.cell(r, c_idx).border = thin_border

            total_amount_calc += act['amount']
            total_expenses_calc += act['expenses']
            total_activities_count += 1
            curr_row += 1

        s_no += 1

    last_data_row = curr_row - 1

    # --- Row 4: Summary Totals Row (Above table headers) ---
    for col_idx in range(1, 9):
        c = ws.cell(4, col_idx, None)
        c.border = thin_border

    # Amount Total: =SUM(I6:I...)
    cI4 = ws.cell(4, 9, f'=SUM(I6:I{last_data_row})' if last_data_row >= 6 else 0)
    cI4.font = font_black_bold
    cI4.alignment = align_right
    cI4.number_format = '0.##'
    cI4.border = thin_border

    # Expenses Total: =SUM(J6:J...)
    cJ4 = ws.cell(4, 10, f'=SUM(J6:J{last_data_row})' if last_data_row >= 6 else 0)
    cJ4.font = font_black_bold
    cJ4.alignment = align_right
    cJ4.number_format = '0.##'
    cJ4.border = thin_border

    # Balance Total: =I4-J4
    cK4 = ws.cell(4, 11, '=I4-J4')
    cK4.font = font_black_bold
    cK4.alignment = align_right
    cK4.number_format = '0.##'
    cK4.border = thin_border

    wb.save(output_path)
    wb.close()

    total_balance_calc = total_amount_calc - total_expenses_calc

    return {
        "success": True,
        "message": f"Corteva Master PO Summary generated successfully as '{output_name}'!",
        "outputPath": str(output_path),
        "folderPath": str(folder),
        "fileName": output_name,
        "territory": final_territory,
        "zdgm": zdgm_display,
        "totalCards": len(po_cards),
        "totalActivities": total_activities_count,
        "totals": {
            "totalBudget": total_amount_calc,
            "totalExpenses": total_expenses_calc,
            "totalBalance": total_balance_calc,
            "totalQty": total_activities_count
        }
    }
