"""
Details of Bills Invoice Extractor
Extracts metadata, particulars, and totals from generated invoice files.
"""

import re
from pathlib import Path
import sys

BACKEND_DIR = Path(__file__).resolve().parent.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

try:
    from excel_parser import load_any_workbook
except ImportError:
    from ..excel_parser import load_any_workbook

from common.formatters import (
    clean_str,
    parse_num,
    format_date_str,
    calculate_receivable_date,
    normalize_po,
)


def extract_invoice_file_data(invoice_file_path):
    """
    Parses a generated Invoice Excel workbook taking primary reference from Sheet 1,
    and grouping activities by TBM & Activity where NO.OF.Activities is the count of activities/days.
    """
    inv_file = Path(invoice_file_path).resolve()
    if not inv_file.exists():
        return None

    wb = load_any_workbook(inv_file)
    if "Sheet1" not in wb.sheetnames:
        wb.close()
        return None

    ws1 = wb["Sheet1"]
    
    invoice_no = ""
    invoice_date = ""
    po_number = ""
    area = ""
    product = ""
    crop = ""
    zdgm = ""
    grand_total = 0.0
    service_charge_pct = 5.0

    for r in range(6, min(30, ws1.max_row + 1)):
        g_val = clean_str(ws1.cell(r, 7).value).lower()
        i_val = clean_str(ws1.cell(r, 9).value)

        if "invoive no" in g_val or "invoice no" in g_val:
            invoice_no = i_val
        elif "invoice date" in g_val:
            invoice_date = format_date_str(i_val)
        elif "po no" in g_val or "po:" in g_val or "new gen po" in g_val or "corteva po" in g_val:
            po_number = i_val
        elif "area" in g_val:
            area = i_val
        elif "product" in g_val:
            product = i_val
        elif "crop" in g_val:
            crop = i_val
        elif "zdgm" in g_val or "amm" in g_val:
            zdgm = i_val

    if not invoice_no:
        invoice_no = inv_file.stem

    short_iv = invoice_no
    m_iv = re.search(r'SBT\d{4}(\d+)', invoice_no, re.I)
    if m_iv:
        short_iv = str(int(m_iv.group(1)))
    else:
        m_num = re.search(r'(\d+)$', invoice_no)
        if m_num:
            short_iv = str(int(m_num.group(1)))

    for r in range(25, min(45, ws1.max_row + 1)):
        e_val = clean_str(ws1.cell(r, 5).value)
        if "%" in e_val:
            try:
                service_charge_pct = float(e_val.replace("%", "").strip())
            except ValueError:
                pass

    norm_po = normalize_po(po_number)
    if norm_po.startswith("5"):
        budget_type = "Brand"
    elif norm_po.startswith("48"):
        budget_type = "MA"
    else:
        budget_type = "MKTG"

    items = []
    if "Sheet2" in wb.sheetnames:
        ws2 = wb["Sheet2"]
        tbm_act_groups = {}
        max_r = ws2.max_row
        r = 1
        while r <= max_r:
            c1_val = clean_str(ws2.cell(r, 1).value).lower()
            c2_val = clean_str(ws2.cell(r, 2).value).lower()

            if ("s.no" in c1_val or "si no" in c1_val or "sl" in c1_val) and "date" in c2_val:
                hdr_r = r
                data_r = hdr_r + 1
                while data_r <= max_r:
                    d_c1 = clean_str(ws2.cell(data_r, 1).value)
                    d_c15 = clean_str(ws2.cell(data_r, 15).value).lower()

                    if d_c1.lower() == 'total' or d_c15 == 'total' or 'total' in d_c1.lower():
                        break
                    if "activities expenses" in d_c1.lower() or "iv no" in d_c1.lower():
                        break

                    row_vals = [ws2.cell(data_r, c).value for c in range(1, 18)]
                    if not any(row_vals):
                        data_r += 1
                        continue

                    item_tbm = clean_str(ws2.cell(data_r, 4).value)
                    item_act = clean_str(ws2.cell(data_r, 9).value)
                    item_amt = parse_num(ws2.cell(data_r, 16).value)
                    if item_amt == 0.0:
                        calc_amt = sum(parse_num(ws2.cell(data_r, c).value) for c in range(12, 16))
                        if calc_amt > 0:
                            item_amt = calc_amt

                    if item_act or item_tbm or item_amt > 0:
                        grp_key = (item_tbm, item_act)
                        if grp_key not in tbm_act_groups:
                            tbm_act_groups[grp_key] = {
                                'tbm': item_tbm,
                                'activity': item_act,
                                'qty': 0,
                                'amount': 0.0
                            }
                        tbm_act_groups[grp_key]['qty'] += 1
                        tbm_act_groups[grp_key]['amount'] += item_amt

                    data_r += 1
                r = data_r
            else:
                r += 1

        for grp in tbm_act_groups.values():
            amt = grp['amount']
            s_iv = round(amt * (1 + service_charge_pct / 100.0), 2)
            t_iv = round(s_iv * 1.18, 2)
            items.append({
                'date': invoice_date,
                'area': area,
                'po_number': po_number,
                'budget': budget_type,
                'product': product,
                'crop': crop,
                'activity': grp['activity'],
                'zdgm': zdgm,
                'tbm': grp['tbm'],
                'qty': grp['qty'],
                'amount': amt,
                'service_iv': s_iv,
                'total_iv': t_iv
            })

    if not items:
        for r_part in range(25, 34):
            part_name = clean_str(ws1.cell(r_part, 2).value)
            part_qty = int(parse_num(ws1.cell(r_part, 9).value))
            part_amt = parse_num(ws1.cell(r_part, 10).value)

            if part_qty > 0 or part_amt > 0:
                act_clean = re.sub(r'(?i)\s*activities\s*expenses\s*', '', part_name).strip()
                if not act_clean:
                    act_clean = "Marketing Activities"

                if norm_po.startswith("48"):
                    s_iv = part_amt
                    raw_amt = round(s_iv / (1 + service_charge_pct / 100.0), 2)
                else:
                    raw_amt = part_amt
                    s_iv = round(raw_amt * (1 + service_charge_pct / 100.0), 2)
                
                t_iv = round(s_iv * 1.18, 2)
                items.append({
                    'date': invoice_date,
                    'area': area,
                    'po_number': po_number,
                    'budget': budget_type,
                    'product': product,
                    'crop': crop,
                    'activity': act_clean,
                    'zdgm': zdgm,
                    'tbm': "",
                    'qty': part_qty,
                    'amount': raw_amt,
                    'service_iv': s_iv,
                    'total_iv': t_iv
                })

    for r in range(25, min(50, ws1.max_row + 1)):
        g_val = clean_str(ws1.cell(r, 7).value).lower()
        if "grand total" in g_val:
            raw_gt = ws1.cell(r, 10).value
            if isinstance(raw_gt, (int, float)):
                grand_total = float(raw_gt)
            break

    if grand_total == 0.0 and items:
        tot_sub = sum(it.get('service_iv', 0.0) for it in items)
        tot_cgst = round(tot_sub * 0.09, 2)
        tot_sgst = round(tot_sub * 0.09, 2)
        grand_total = round(tot_sub + tot_cgst + tot_sgst)

    wb.close()

    receivable_date = calculate_receivable_date(invoice_date, days_credit=45)

    return {
        'invoice_no': invoice_no,
        'short_iv': short_iv,
        'invoice_date': invoice_date,
        'po_number': po_number,
        'area': area,
        'product': product,
        'crop': crop,
        'zdgm': zdgm,
        'grand_total': grand_total,
        'receivable_date': receivable_date,
        'service_charge_pct': service_charge_pct,
        'items': items
    }
