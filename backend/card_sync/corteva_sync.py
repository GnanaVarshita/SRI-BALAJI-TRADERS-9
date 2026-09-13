"""
Corteva PO Summary Cards Synchronization Module
Sri Balaji Traders Automation System
"""

from openpyxl.utils import get_column_letter
from .constants import (
    REGULAR_FONT,
    BOLD_FONT,
    RED_BOLD_FONT,
    normalize_activity,
    match_activity_to_column,
)


def is_corteva_row_occupied(ws, r, amount_col):
    """
    Checks if a row in Corteva card already contains valid data.
    """
    # Check key text columns: Col 4 (PO Number), Col 8 (Activity), Col 10 (TBM)
    c4 = str(ws.cell(r, 4).value or '').strip()
    c8 = str(ws.cell(r, 8).value or '').strip()
    c10 = str(ws.cell(r, 10).value or '').strip()
    if c4 or c8 or c10:
        return True

    # Check if any activity column (12 to amount_col - 1) has a numeric value > 0
    for c in range(12, amount_col):
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)) and v > 0:
            return True

    return False


def sync_corteva_cards_workbook(wb_cards, tbm_data, service_charge_percent=5.0):
    """
    Synchronizes Corteva PO Summary Cards (where each product is on its own worksheet).
    Appends new spending entries into available rows below existing ones (non-destructive),
    and updates bottom sums, Right Summary table, and overview sheets.
    """
    sv_rate = float(service_charge_percent) / 100.0
    updated_cards_count = 0
    card_activity_links = {}  # (po_val, norm_act) -> (sheet_name, summary_row_idx)

    for name in wb_cards.sheetnames:
        if name.strip().lower() in ['sheet1', 'processed_emails', 'pos & expenditures', 'overview']:
            continue

        ws = wb_cards[name]
        po_val = str(ws.cell(6, 1).value or '').strip().upper()
        if not po_val:
            continue

        # 1. Dynamically locate activity columns & Amount column in row 10 / 11
        amount_col = 18
        for c in range(12, 22):
            val_10 = str(ws.cell(10, c).value or '').strip().lower()
            val_11 = str(ws.cell(11, c).value or '').strip().lower()
            if val_10 == 'amount' or val_11 == 'amount':
                amount_col = c
                break

        # Build activity column map (from row 10 / 11 headers)
        act_map = {}
        for c in range(12, amount_col):
            h_val = ws.cell(10, c).value or ws.cell(11, c).value
            if h_val and str(h_val).strip() not in ['0', 'None', '']:
                act_map[normalize_activity(h_val)] = c

        # Fallback to rate grid (Cols T & U, rows 3..) if row 10 was incomplete
        if not act_map:
            for r_idx in range(3, 10):
                act_name = ws.cell(r_idx, 20).value
                if act_name and str(act_name).strip() not in ['TOTAL', '']:
                    act_map[normalize_activity(act_name)] = 12 + len(act_map)

        if not act_map:
            continue

        n_act = len(act_map)
        tbm_rows_for_po = tbm_data.get(po_val, [])
        if not tbm_rows_for_po:
            continue

        # 2. Non-Destructive Append: Find first available row between 12 and 28
        curr_row = 12
        while curr_row <= 28 and is_corteva_row_occupied(ws, curr_row, amount_col):
            curr_row += 1

        area_val = str(ws.cell(7, 4).value or '').split('-')[-1].strip()
        newly_added_count = 0

        for t_item in tbm_rows_for_po:
            if curr_row > 28:
                break  # Card reached maximum row capacity (row 28)

            ws.cell(curr_row, 1).value = None  # I.V NO
            ws.cell(curr_row, 2).value = None  # DATE
            ws.cell(curr_row, 3, area_val or t_item.get('area', '')).font = REGULAR_FONT
            ws.cell(curr_row, 4, po_val).font = REGULAR_FONT
            ws.cell(curr_row, 5, "Marketing").font = REGULAR_FONT
            ws.cell(curr_row, 6, t_item.get('product') or name).font = REGULAR_FONT
            ws.cell(curr_row, 7, t_item.get('crop') or "All Crops").font = REGULAR_FONT
            ws.cell(curr_row, 8, t_item.get('activity', '')).font = REGULAR_FONT
            ws.cell(curr_row, 9, t_item.get('zdgm') or "ZDGM").font = REGULAR_FONT
            ws.cell(curr_row, 10, t_item.get('tbm', '')).font = REGULAR_FONT
            ws.cell(curr_row, 11, t_item.get('num_activities', 1)).font = REGULAR_FONT

            target_col = match_activity_to_column(t_item.get('activity', ''), act_map)
            if not target_col or target_col >= amount_col:
                target_col = 12

            for c in range(12, amount_col):
                if c == target_col:
                    ws.cell(curr_row, c, t_item.get('total_amount', 0.0)).font = REGULAR_FONT
                    ws.cell(curr_row, c).number_format = '#,##0'
                else:
                    ws.cell(curr_row, c, None)

            # Row total formula
            parts = [f"{get_column_letter(c)}{curr_row}" for c in range(12, amount_col)]
            ws.cell(curr_row, amount_col, f"=SUM({','.join(parts)})").font = REGULAR_FONT
            ws.cell(curr_row, amount_col).number_format = '#,##0'

            curr_row += 1
            newly_added_count += 1

        if newly_added_count > 0:
            updated_cards_count += 1

        # 3. Bottom Totals row 29 (sums rows 12 to 28)
        for i in range(n_act):
            col = 12 + i
            cl = get_column_letter(col)
            cell = ws.cell(29, col, f"=SUM({cl}12:{cl}28)")
            cell.font = RED_BOLD_FONT
            cell.number_format = '#,##0'

        amt_cl = get_column_letter(amount_col)
        cell = ws.cell(29, amount_col, f"=SUM({amt_cl}12:{amt_cl}28)")
        cell.font = RED_BOLD_FONT
        cell.number_format = '#,##0'

        # 4. Right Summary block T16:Y (rows 16 to 16 + n_act)
        for i in range(n_act):
            r = 16 + i
            cl = get_column_letter(12 + i)
            act_name_c = ws.cell(r, 20).value
            if act_name_c and str(act_name_c).strip():
                card_activity_links[(po_val, normalize_activity(act_name_c))] = (name, r)

            # SPENT in Col V (22)
            ws.cell(r, 22, f"={cl}29").font = REGULAR_FONT
            ws.cell(r, 22).number_format = '#,##0'

            # SV Charges in Col W (23)
            ws.cell(r, 23, f"=V{r}*{sv_rate}").font = REGULAR_FONT
            ws.cell(r, 23).number_format = '#,##0.00'

            # TOTAL IV in Col X (24)
            ws.cell(r, 24, f"=V{r}+W{r}").font = REGULAR_FONT
            ws.cell(r, 24).number_format = '#,##0.00'

            # BALANCE in Col Y (25)
            ws.cell(r, 25, f"=U{r}-X{r}").font = REGULAR_FONT
            ws.cell(r, 25).number_format = '#,##0.00'

        tot_row = 16 + n_act
        ws.cell(tot_row, 20, "TOTAL").font = RED_BOLD_FONT
        for col_idx in [21, 22, 23, 24, 25]:
            cl = get_column_letter(col_idx)
            cell = ws.cell(tot_row, col_idx, f"=SUM({cl}16:{cl}{tot_row-1})")
            cell.font = RED_BOLD_FONT
            cell.number_format = '#,##0.00' if col_idx in [23, 24, 25] else '#,##0'

    # 5. Link overview sheets if present
    for sname in wb_cards.sheetnames:
        if sname.strip().lower() in ['pos & expenditures', 'sheet1']:
            ws_ov = wb_cards[sname]
            curr_po = ''
            for r in range(4, ws_ov.max_row + 1):
                po_cell_val = ws_ov.cell(r, 3).value
                if po_cell_val and str(po_cell_val).strip():
                    curr_po = str(po_cell_val).strip().upper()

                act_cell_val = ws_ov.cell(r, 8).value or ws_ov.cell(r, 6).value
                if act_cell_val and str(act_cell_val).strip():
                    norm_a = normalize_activity(act_cell_val)
                    key = (curr_po, norm_a)
                    if key in card_activity_links:
                        c_sheet, c_row = card_activity_links[key]
                        # Col 10 (J): Spent -> Link to card's TOTAL IV (Col X)
                        ws_ov.cell(r, 10, f"='{c_sheet}'!X{c_row}").number_format = '#,##0.00'
                        # Col 11 (K): Balance = Budget (Col 9 / I) - Spent
                        ws_ov.cell(r, 11, f"=I{r}-J{r}").number_format = '#,##0.00'

    return updated_cards_count
