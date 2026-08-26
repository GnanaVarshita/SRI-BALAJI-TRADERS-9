import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import re

def parse_pos_from_step1_sheet(ws):
    pos = []
    current_po = None
    
    for r in range(4, ws.max_row + 1):
        val_c5 = ws.cell(r, 5).value
        if val_c5 and str(val_c5).strip().lower() == 'total':
            break
            
        po_num = ws.cell(r, 2).value
        prod = ws.cell(r, 3).value
        crop = ws.cell(r, 4).value
        act = ws.cell(r, 5).value
        budget = ws.cell(r, 6).value
        
        if po_num and str(po_num).strip() != '':
            current_po = {
                'po_number': str(po_num).strip(),
                'product': str(prod).strip() if prod else '',
                'crop': str(crop).strip() if crop else '',
                'activities': []
            }
            pos.append(current_po)
            
        if current_po and act:
            current_po['activities'].append({
                'activity': str(act).strip(),
                'budget': budget or 0
            })
            
    return pos

def write_po_summary_block(ws, start_r, po_data, territory='Nandyal', am_name='Madhavareddy', date_str='01-08-2026'):
    # Styles
    red_bold = Font(name='Calibri', size=11, bold=True, color='FF0000')
    green_bold = Font(name='Calibri', size=11, bold=True, color='008000')
    blue_bold = Font(name='Calibri', size=11, bold=True, color='002060')
    green_header_font = Font(name='Calibri', size=10, bold=True, color='008000')
    regular_font = Font(name='Calibri', size=10)
    
    light_green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    def _center():
        return Alignment(horizontal='center', vertical='center')

    po_no = po_data['po_number']
    product = po_data['product']
    crop = po_data['crop']
    activities = po_data['activities']
    total_budget = sum(a['budget'] for a in activities)

    # Row 1 of Block (start_r)
    # A: PO Number
    cell = ws.cell(start_r, 1, po_no)
    cell.font = green_bold

    # C: PRODUCT header, E: Product value
    c_prod = ws.cell(start_r, 3, "PRODUCT")
    c_prod.font = red_bold
    
    v_prod = ws.cell(start_r, 5, product)
    v_prod.font = green_bold

    # F: CROP header, G: Crop value
    c_crop = ws.cell(start_r, 6, "CROP")
    c_crop.font = red_bold
    
    v_crop = ws.cell(start_r, 7, crop)
    v_crop.font = green_bold

    # H: DATE header, I: Date value
    c_date = ws.cell(start_r, 8, "DATE")
    c_date.font = red_bold
    
    v_date = ws.cell(start_r, 9, date_str)
    v_date.font = green_bold

    # List ALL activities in Col K & L starting at start_r
    for idx, act_item in enumerate(activities):
        r_pos = start_r + idx
        c_act = ws.cell(r_pos, 11, act_item['activity'])
        c_act.font = green_bold
        
        b_act = ws.cell(r_pos, 12, act_item['budget'])
        b_act.font = green_bold
        b_act.number_format = '#,##0'

    # Row 2 of Block (start_r + 1)
    # C: BB
    c_bb = ws.cell(start_r + 1, 3, "BB")
    c_bb.font = green_bold

    # Row 3 of Block (start_r + 2)
    # E: Territory
    c_terr = ws.cell(start_r + 2, 5, territory)
    c_terr.font = green_bold
    c_terr.alignment = _center()

    # Row 4 of Block (start_r + 3)
    # C: AM Name
    c_am = ws.cell(start_r + 3, 3, am_name)
    c_am.font = blue_bold

    # K: Total, L: Total Budget, M: WITH GST, N: Total with GST
    c_tot_lbl = ws.cell(start_r + 3, 11, "Total")
    c_tot_lbl.font = green_bold

    c_tot_val = ws.cell(start_r + 3, 12, total_budget)
    c_tot_val.font = green_bold
    c_tot_val.number_format = '#,##0'

    c_gst_lbl = ws.cell(start_r + 3, 13, "WITH GST")
    c_gst_lbl.font = green_bold

    c_gst_val = ws.cell(start_r + 3, 14, f"=L{start_r + 3}*1.18")
    c_gst_val.font = green_bold
    c_gst_val.number_format = '#,##0'

    # Row 5, 6, 7 of Block (start_r + 4 to start_r + 6): Main & Right Table Headers
    # Main Table Headers (Cols A to J)
    main_headers = ['I.V NO', 'DATE', 'AREA', 'PO NUMBER', 'BUDGET', 'Product', 'Crop', 'ACTIVITY', 'RBM', 'MIE']
    for c_idx, h_text in enumerate(main_headers, 1):
        cell = ws.cell(start_r + 6, c_idx, h_text)
        cell.font = red_bold if c_idx in [4, 5, 6, 7, 8, 9, 10] else green_bold
        cell.border = border
        cell.alignment = _center()

    # Activity Header in Main Table (Cols K, L, M)
    c_act_hdr = ws.cell(start_r + 5, 11, "NO.OF.Activities")
    c_act_hdr.font = red_bold
    c_act_hdr.fill = light_green_fill
    c_act_hdr.border = border
    c_act_hdr.alignment = _center()

    c_act_code = ws.cell(start_r + 5, 12, activities[0]['activity'] if activities else "FD")
    c_act_code.font = red_bold
    c_act_code.fill = light_green_fill
    c_act_code.border = border
    c_act_code.alignment = _center()

    c_act_amt = ws.cell(start_r + 6, 13, "Amount")
    c_act_amt.font = red_bold
    c_act_amt.fill = light_green_fill
    c_act_amt.border = border
    c_act_amt.alignment = _center()

    ws.cell(start_r + 5, 13, 0).border = border

    # Right Summary Table Headers (Cols K to P, Row start_r + 7)
    right_headers = ['Activities', 'BUDGET', 'SPENT', 'SV Charges', 'TOTAL IV', 'BALANCE']
    for c_idx, h_text in enumerate(right_headers, 11):
        cell = ws.cell(start_r + 7, c_idx, h_text)
        cell.font = green_header_font
        cell.border = border
        cell.alignment = _center()

    # Data Rows (Rows start_r + 7 to start_r + 16 in Main Table, Rows start_r + 8 to start_r + 16 in Right Table)
    # Main Table empty grid (Cols A to J, 10 rows)
    for r_offset in range(7, 17):
        for c in range(1, 11):
            ws.cell(start_r + r_offset, c).border = border

    # Right Summary Table Data Rows (Cols K to P)
    first_data_r = start_r + 8
    last_data_r = start_r + 15

    for idx in range(8):
        curr_r = first_data_r + idx
        if idx < len(activities):
            act_item = activities[idx]
            ws.cell(curr_r, 11, act_item['activity']).font = red_bold if idx == 0 else regular_font
            
            b_cell = ws.cell(curr_r, 12, act_item['budget'])
            b_cell.font = green_bold if idx == 0 else regular_font
            b_cell.number_format = '#,##0'
        else:
            ws.cell(curr_r, 11, 0).font = regular_font
            ws.cell(curr_r, 12, 0).font = regular_font

        ws.cell(curr_r, 13, 0).font = regular_font # SPENT
        ws.cell(curr_r, 14, 0).font = regular_font # SV Charges
        ws.cell(curr_r, 15, 0).font = regular_font # TOTAL IV
        
        bal_cell = ws.cell(curr_r, 16, f"=L{curr_r}-O{curr_r}")
        bal_cell.font = green_bold if idx == 0 else regular_font
        bal_cell.number_format = '#,##0'

        for c in range(11, 17):
            ws.cell(curr_r, c).border = border
            if c >= 12:
                ws.cell(curr_r, c).alignment = Alignment(horizontal='right', vertical='center')

    # Total Row for Right Summary Table (Row start_r + 16)
    tot_r = start_r + 16
    
    t_lbl = ws.cell(tot_r, 11, "TOTAL")
    t_lbl.font = red_bold
    t_lbl.border = border

    t_bud = ws.cell(tot_r, 12, f"=SUM(L{first_data_r}:L{last_data_r})")
    t_bud.font = green_bold
    t_bud.border = border
    t_bud.number_format = '#,##0'

    t_spent = ws.cell(tot_r, 13, f"=SUM(M{first_data_r}:M{last_data_r})")
    t_spent.font = regular_font
    t_spent.border = border

    t_sv = ws.cell(tot_r, 14, f"=SUM(N{first_data_r}:N{last_data_r})")
    t_sv.font = regular_font
    t_sv.border = border

    t_tot_iv = ws.cell(tot_r, 15, f"=SUM(O{first_data_r}:O{last_data_r})")
    t_tot_iv.font = regular_font
    t_tot_iv.border = border

    t_bal = ws.cell(tot_r, 16, f"=L{start_r + 3}-O{tot_r}")
    t_bal.font = red_bold
    t_bal.border = border
    t_bal.number_format = '#,##0'

    # Footer row zeros (Row start_r + 17)
    for c in [11, 12, 13]:
        cell = ws.cell(start_r + 17, c, 0)
        cell.font = red_bold

def generate_fmc_step2_summaries(excel_file_path, territory='Nandyal', am_name='Madhavareddy'):
    excel_path = Path(excel_file_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found at: {excel_file_path}")

    wb = openpyxl.load_workbook(excel_path)
    
    if 'Sheet1' not in wb.sheetnames:
        raise ValueError("Sheet1 not found in the uploaded Excel file. Please run Step 1 first.")
        
    ws_master = wb['Sheet1']
    pos = parse_pos_from_step1_sheet(ws_master)

    if not pos:
        raise ValueError("No PO data found in Sheet1.")

    # Split POs into chunks of 11
    chunk_size = 11
    po_chunks = [pos[i:i + chunk_size] for i in range(0, len(pos), chunk_size)]

    for chunk_idx, chunk in enumerate(po_chunks):
        # Name sheet based on PO number range (last 5 digits of first and last PO in chunk)
        first_po_short = chunk[0]['po_number'][-5:]
        last_po_short = chunk[-1]['po_number'][-5:]
        sheet_name = f"{first_po_short}-{last_po_short}"

        # If sheet already exists, remove or reuse
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
            
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True

        for b_idx, po_item in enumerate(chunk):
            start_r = 1 + b_idx * 19
            write_po_summary_block(ws, start_r, po_item, territory=territory, am_name=am_name)

        # Set column widths
        col_widths = {
            'A': 10, 'B': 10, 'C': 12, 'D': 15, 'E': 12,
            'F': 12, 'G': 12, 'H': 12, 'I': 10, 'J': 10,
            'K': 12, 'L': 12, 'M': 10, 'N': 10, 'O': 10, 'P': 12
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    wb.save(excel_path)
    wb.close()
    return len(pos), len(po_chunks)

