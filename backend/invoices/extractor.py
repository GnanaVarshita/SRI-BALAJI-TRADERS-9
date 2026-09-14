"""
Invoice Data Extractor
Extracts POs and activity tables from Consolidated TBM Summaries.
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

from common.formatters import clean_str, parse_num, format_date_val, normalize_po


def scan_pos_in_summary(tbm_summary_path):
    """
    Scans the All-TBMs-Summary.xlsx workbook and returns list of unique PO numbers found.
    Handles any length of PO number without trimming.
    """
    tbm_file = Path(tbm_summary_path).resolve()
    if not tbm_file.exists():
        return []

    wb = load_any_workbook(tbm_file)
    pos_found = set()

    for sname in wb.sheetnames:
        ws = wb[sname]
        for r in range(1, ws.max_row + 1):
            for c in range(1, min(ws.max_column + 1, 35)):
                val = clean_str(ws.cell(r, c).value)
                if val:
                    # Corteva PO: starts with 48 (e.g. 4800108503)
                    m_corteva = re.findall(r'\b(48\d{6,30})\b', val)
                    # FMC PO: starts with 5 (e.g. 500BB20260710172, 500BB2026018404)
                    m_fmc = re.findall(r'\b(5[A-Za-z0-9]{8,35})\b', val, re.I)
                    for m in m_corteva + m_fmc:
                        pos_found.add(m.upper())
                    
                    # Direct check on Column 17 or any PO-labeled column
                    if c == 17:
                        val_u = val.upper()
                        if val_u and val_u not in ["PO", "PO NUMBER", "PO NO", "P.O", "P.O.", "NONE", "TOTAL"]:
                            if re.match(r'^[A-Z0-9\-_]{6,35}$', val_u):
                                pos_found.add(val_u)
    wb.close()
    return sorted(list(pos_found))


def extract_tables_for_po(tbm_summary_path, target_po):
    """
    Extracts all activity expense records matching target_po from All-TBMs-Summary.xlsx.
    Performs strict exact PO matching on each row to prevent mixing activities from different POs.
    Returns: (records, metadata)
    """
    tbm_file = Path(tbm_summary_path).resolve()
    if not tbm_file.exists():
        raise FileNotFoundError(f"TBM Summary file not found at: {tbm_summary_path}")

    wb = load_any_workbook(tbm_file)
    norm_target_po = normalize_po(target_po)

    all_matched_records = []
    overall_metadata = {
        'product': '',
        'crop': '',
        'territory': '',
        'area': '',
        'zdgm': '',
        'amm': ''
    }

    for sname in wb.sheetnames:
        if sname.strip().lower() in ['tbm amount summary', 'amount summary']:
            continue

        ws = wb[sname]
        max_r = ws.max_row
        r = 1
        while r <= max_r:
            cell_val = clean_str(ws.cell(r, 1).value)
            if "ACTIVITIES EXPENSES" in cell_val.upper():
                hdr_row_idx = r + 1
                headers = [clean_str(ws.cell(hdr_row_idx, c).value).lower() for c in range(1, 18)]
                if not any('si' in h or 's.no' in h or 'sl' in h or 'date' in h for h in headers):
                    r += 1
                    continue

                data_r = hdr_row_idx + 1
                table_rows = []
                while data_r <= max_r:
                    c1_val = clean_str(ws.cell(data_r, 1).value).lower()
                    c15_val = clean_str(ws.cell(data_r, 15).value).lower()
                    
                    if c1_val == 'total' or c15_val == 'total' or 'total expenses' in c1_val or 'grand total' in c1_val:
                        break
                    if "ACTIVITIES EXPENSES" in c1_val.upper():
                        break

                    row_vals = [ws.cell(data_r, c).value for c in range(1, 18)]
                    if not any(row_vals):
                        data_r += 1
                        continue

                    po_in_row = clean_str(ws.cell(data_r, 17).value)
                    norm_row_po = normalize_po(po_in_row)

                    date_val = format_date_val(ws.cell(data_r, 2).value)
                    zdgm_val = clean_str(ws.cell(data_r, 3).value)
                    tbm_val = clean_str(ws.cell(data_r, 4).value)
                    mdo_val = clean_str(ws.cell(data_r, 5).value)
                    terr_val = clean_str(ws.cell(data_r, 6).value)
                    prod_val = clean_str(ws.cell(data_r, 7).value)
                    crop_val = clean_str(ws.cell(data_r, 8).value)
                    act_val = clean_str(ws.cell(data_r, 9).value)
                    vlg_val = clean_str(ws.cell(data_r, 10).value)
                    farm_val = int(parse_num(ws.cell(data_r, 11).value))
                    tent_val = parse_num(ws.cell(data_r, 12).value)
                    food_val = parse_num(ws.cell(data_r, 13).value)
                    trans_val = parse_num(ws.cell(data_r, 14).value)
                    oth_val = parse_num(ws.cell(data_r, 15).value)
                    tot_val = parse_num(ws.cell(data_r, 16).value)
                    calc_tot = tent_val + food_val + trans_val + oth_val
                    if tot_val == 0.0 and calc_tot > 0.0:
                        tot_val = calc_tot

                    is_match = False
                    if norm_target_po:
                        if norm_row_po:
                            # Strict exact match - prevents mixing activities belonging to different POs
                            if norm_row_po == norm_target_po:
                                is_match = True
                        else:
                            # If row cell is empty, check if the table header explicitly specifies this PO
                            tbl_header_norm = normalize_po(cell_val)
                            if norm_target_po in tbl_header_norm:
                                is_match = True
                    else:
                        is_match = True

                    if is_match and (act_val or prod_val or tot_val > 0):
                        rec = {
                            'date': date_val,
                            'zdgm': zdgm_val,
                            'tbm': tbm_val,
                            'mdo': mdo_val,
                            'territory': terr_val,
                            'product': prod_val,
                            'crop': crop_val,
                            'activity': act_val,
                            'village': vlg_val,
                            'farmers': farm_val,
                            'tent': tent_val,
                            'food': food_val,
                            'transport': trans_val,
                            'others': oth_val,
                            'total': tot_val,
                            'po_number': po_in_row or target_po
                        }
                        table_rows.append(rec)
                        
                        if prod_val and not overall_metadata['product']: overall_metadata['product'] = prod_val
                        if crop_val and not overall_metadata['crop']: overall_metadata['crop'] = crop_val
                        if terr_val and not overall_metadata['territory']: overall_metadata['territory'] = terr_val
                        if zdgm_val and not overall_metadata['zdgm']: overall_metadata['zdgm'] = zdgm_val
                        if zdgm_val and not overall_metadata['amm']: overall_metadata['amm'] = zdgm_val

                    data_r += 1

                if table_rows:
                    all_matched_records.extend(table_rows)

                r = data_r
            else:
                r += 1

    wb.close()
    return all_matched_records, overall_metadata
