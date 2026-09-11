"""
TBM Activity Readers
Extracts activity records from OpenPyXL worksheets, legacy xlrd sheets, and slip formats.
"""

import re
from common.formatters import clean_str, parse_num, parse_date_intelligent, format_date_xlrd
from .constants import FIELD_KEYWORDS


def parse_slip_format(ws, default_tbm_name="", default_territory=""):
    """Parses single-card / slip format where keys are in Col 2 and values in Col 4."""
    kv = {}
    for r in range(1, 25):
        k = clean_str(ws.cell(r, 2).value).lower()
        v = ws.cell(r, 4).value
        if k:
            if 'date' in k: kv['date'] = parse_date_intelligent(v)
            elif 'product' in k: kv['product'] = clean_str(v)
            elif 'crop' in k: kv['crop'] = clean_str(v)
            elif 'activity' in k: kv['activity'] = clean_str(v)
            elif 'place' in k or 'territory' in k: kv['territory'] = clean_str(v)
            elif 'farmers' in k: kv['farmers'] = int(parse_num(v))
            elif 'food' in k: kv['food'] = parse_num(v)
            elif 'chairs' in k or 'tent' in k: kv['tent'] = parse_num(v)
            elif 'auto' in k or 'transport' in k: kv['transport'] = parse_num(v)
            elif 'others' in k: kv['others'] = parse_num(v)
            elif 'tbm' in k: kv['tbm'] = clean_str(v)
            elif 'po' in k: kv['po_number'] = clean_str(v)

    if kv.get('activity') or kv.get('product') or kv.get('date'):
        tbm_val = kv.get('tbm') or default_tbm_name
        terr_val = kv.get('territory') or default_territory
        tent = kv.get('tent', 0.0)
        food = kv.get('food', 0.0)
        transport = kv.get('transport', 0.0)
        others = kv.get('others', 0.0)
        total = tent + food + transport + others
        return {
            'date': kv.get('date', ''),
            'zdgm': '',
            'tbm': tbm_val,
            'mdo': '',
            'territory': terr_val,
            'product': kv.get('product', ''),
            'crop': kv.get('crop', ''),
            'activity': kv.get('activity', ''),
            'village': '',
            'farmers': kv.get('farmers', 0),
            'tent': tent,
            'food': food,
            'transport': transport,
            'others': others,
            'total': total,
            'po_number': kv.get('po_number', '')
        }
    return None


def extract_activities_from_sheet(ws, default_tbm_name="", default_territory="", file_name=None):
    """
    Scans OpenPyXL worksheet for all tables / header rows and extracts activity records.
    """
    all_activities = []
    header_rows = []
    max_r = ws.max_row or 100
    max_c = min(ws.max_column or 35, 45)

    # 1. Detect all header rows across the sheet
    for r in range(1, max_r + 1):
        col_map = {}
        score = 0
        is_banner = False

        for c in range(1, max_c + 1):
            txt = clean_str(ws.cell(r, c).value).lower()
            if 'activities expenses by' in txt or 'marketing activities' in txt or 'activities expenses' in txt:
                is_banner = True
                break

        if is_banner:
            continue

        for c in range(1, max_c + 1):
            curr_txt = clean_str(ws.cell(r, c).value).lower()
            if not curr_txt:
                continue

            # Prioritize sl_no for column 1 or sl/sno labels
            if any(kw == curr_txt or kw in curr_txt for kw in FIELD_KEYWORDS['sl_no']):
                if 'sl_no' not in col_map:
                    col_map['sl_no'] = c
                    score += 1
                    continue

            for field, kw_list in FIELD_KEYWORDS.items():
                if field == 'sl_no':
                    continue
                if field not in col_map:
                    if any(kw == curr_txt or kw in curr_txt for kw in kw_list):
                        col_map[field] = c
                        score += 1
                        break

        if score >= 3 and any(k in col_map for k in ['date', 'product', 'activity', 'mdo', 'zdgm', 'territory', 'sl_no', 'farmers']):
            header_rows.append((r, col_map))

    if not header_rows:
        slip_act = parse_slip_format(ws, default_tbm_name, default_territory)
        if slip_act:
            if file_name:
                slip_act['source_file'] = file_name
            return [slip_act]
        return []

    # 2. Extract data rows for each detected table header
    for i, (hr, cmap) in enumerate(header_rows):
        next_hr = header_rows[i + 1][0] if i + 1 < len(header_rows) else max_r + 1

        for r in range(hr + 1, next_hr):
            row_vals = [ws.cell(r, c).value for c in range(1, max_c + 1)]
            if not any(row_vals):
                continue

            row_str = ' '.join(clean_str(v).lower() for v in row_vals)
            has_data_content = any(
                clean_str(ws.cell(r, cmap.get(k, 0)).value)
                for k in ['product', 'activity', 'date', 'village']
                if k in cmap and cmap.get(k, 0) <= max_c
            )
            if 'total' in row_str and not has_data_content:
                break

            raw_date = ws.cell(r, cmap.get('date', 0)).value if 'date' in cmap else None
            date_val = parse_date_intelligent(raw_date)

            prod_val = clean_str(ws.cell(r, cmap.get('product', 0)).value) if 'product' in cmap else ''
            act_val = clean_str(ws.cell(r, cmap.get('activity', 0)).value) if 'activity' in cmap else ''
            crop_val = clean_str(ws.cell(r, cmap.get('crop', 0)).value) if 'crop' in cmap else ''
            vlg_val = clean_str(ws.cell(r, cmap.get('village', 0)).value) if 'village' in cmap else ''
            zdgm_val = clean_str(ws.cell(r, cmap.get('zdgm', 0)).value) if 'zdgm' in cmap else ''
            mdo_val = clean_str(ws.cell(r, cmap.get('mdo', 0)).value) if 'mdo' in cmap else ''
            terr_val = clean_str(ws.cell(r, cmap.get('territory', 0)).value) if 'territory' in cmap else ''
            if not terr_val:
                terr_val = default_territory

            tbm_val = clean_str(ws.cell(r, cmap.get('tbm', 0)).value) if 'tbm' in cmap else ''
            if not tbm_val or tbm_val.isdigit():
                tbm_val = default_tbm_name

            farm_val = parse_num(ws.cell(r, cmap.get('farmers', 0)).value) if 'farmers' in cmap else 0.0
            tent_val = parse_num(ws.cell(r, cmap.get('tent', 0)).value) if 'tent' in cmap else 0.0
            food_val = parse_num(ws.cell(r, cmap.get('food', 0)).value) if 'food' in cmap else 0.0
            trans_val = parse_num(ws.cell(r, cmap.get('transport', 0)).value) if 'transport' in cmap else 0.0
            oth_val = parse_num(ws.cell(r, cmap.get('others', 0)).value) if 'others' in cmap else 0.0
            tot_val = parse_num(ws.cell(r, cmap.get('total', 0)).value) if 'total' in cmap else 0.0
            if tot_val == 0.0:
                tot_val = tent_val + food_val + trans_val + oth_val

            if not prod_val and not act_val and not date_val and not vlg_val and not crop_val and tot_val == 0.0:
                continue
            if not prod_val and not act_val and not date_val and not vlg_val and not crop_val and not (zdgm_val and terr_val):
                continue

            po_val = clean_str(ws.cell(r, cmap.get('po_number', 0)).value) if 'po_number' in cmap else ''
            if not po_val:
                for c in range(1, max_c + 1):
                    v = clean_str(ws.cell(r, c).value)
                    m = re.search(r'5\d{2}[A-Z0-9]{8,20}', v, re.I)
                    if m:
                        po_val = m.group(0).upper()
                        break

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
                'farmers': int(farm_val),
                'tent': tent_val,
                'food': food_val,
                'transport': trans_val,
                'others': oth_val,
                'total': tot_val,
                'po_number': po_val
            }
            if file_name:
                rec['source_file'] = file_name
            all_activities.append(rec)

    return all_activities


def extract_activities_from_xlrd_sheet(sh, datemode, tbm_folder_name, file_name):
    """
    Extracts activity records from legacy .xls sheets using xlrd.
    """
    best_row = None
    best_map = {}
    best_score = 0

    max_r = min(sh.nrows, 25)
    for r in range(max_r):
        curr_map = {}
        score = 0
        for c in range(sh.ncols):
            top_txt = clean_str(sh.cell_value(r - 1, c)).lower() if r > 0 else ""
            curr_txt = clean_str(sh.cell_value(r, c)).lower()
            comb_txt = f"{top_txt} {curr_txt}".strip()
            if not comb_txt:
                continue

            for field, kw_list in FIELD_KEYWORDS.items():
                if field not in curr_map:
                    if any(kw in comb_txt for kw in kw_list):
                        curr_map[field] = c
                        score += 1
                        break
        if score > best_score and score >= 3:
            best_score = score
            best_row = r
            best_map = curr_map

    if best_row is None:
        return []

    sheet_po_number = ""
    for r in range(min(10, sh.nrows)):
        for c in range(min(20, sh.ncols)):
            val = clean_str(sh.cell_value(r, c))
            m = re.search(r'5\d{2}[A-Z0-9]{8,12}', val, re.I)
            if m:
                sheet_po_number = m.group(0).upper()
                break
        if sheet_po_number:
            break

    activities = []
    for r in range(best_row + 1, sh.nrows):
        row_vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if not any(row_vals):
            continue

        c1_val = clean_str(row_vals[0]).lower()
        c2_val = clean_str(row_vals[1]).lower() if len(row_vals) > 1 else ""
        date_col_idx = best_map.get('date', 2)
        date_c_val = clean_str(sh.cell_value(r, date_col_idx)).lower() if date_col_idx < len(row_vals) else ""

        if c1_val == 'total' or c2_val == 'total' or date_c_val == 'total':
            break

        date_val = format_date_xlrd(sh.cell_value(r, date_col_idx), datemode) if date_col_idx < len(row_vals) else ""
        
        tbm_idx = best_map.get('tbm', 0)
        tbm_val = clean_str(sh.cell_value(r, tbm_idx)) if tbm_idx < len(row_vals) else ""
        if not tbm_val:
            tbm_val = tbm_folder_name

        zdgm_idx = best_map.get('zdgm', 0)
        zdgm_val = clean_str(sh.cell_value(r, zdgm_idx)) if zdgm_idx < len(row_vals) else ""

        mdo_idx = best_map.get('mdo', 0)
        mdo_val = clean_str(sh.cell_value(r, mdo_idx)) if mdo_idx < len(row_vals) else ""

        terr_idx = best_map.get('territory', 0)
        territory_val = clean_str(sh.cell_value(r, terr_idx)) if terr_idx < len(row_vals) else ""

        prod_idx = best_map.get('product', 0)
        product_val = clean_str(sh.cell_value(r, prod_idx)) if prod_idx < len(row_vals) else ""

        crop_idx = best_map.get('crop', 0)
        crop_val = clean_str(sh.cell_value(r, crop_idx)) if crop_idx < len(row_vals) else ""

        act_idx = best_map.get('activity', 0)
        activity_val = clean_str(sh.cell_value(r, act_idx)) if act_idx < len(row_vals) else ""

        vlg_idx = best_map.get('village', 0)
        village_val = clean_str(sh.cell_value(r, vlg_idx)) if vlg_idx < len(row_vals) else ""

        farm_idx = best_map.get('farmers', 0)
        farmers_val = parse_num(sh.cell_value(r, farm_idx)) if farm_idx < len(row_vals) else 0.0

        tent_idx = best_map.get('tent', 0)
        tent_val = parse_num(sh.cell_value(r, tent_idx)) if tent_idx < len(row_vals) else 0.0

        food_idx = best_map.get('food', 0)
        food_val = parse_num(sh.cell_value(r, food_idx)) if food_idx < len(row_vals) else 0.0

        trans_idx = best_map.get('transport', 0)
        transport_val = parse_num(sh.cell_value(r, trans_idx)) if trans_idx < len(row_vals) else 0.0

        oth_idx = best_map.get('others', 0)
        others_val = parse_num(sh.cell_value(r, oth_idx)) if oth_idx < len(row_vals) else 0.0

        tot_idx = best_map.get('total', 0)
        total_val = parse_num(sh.cell_value(r, tot_idx)) if tot_idx < len(row_vals) else 0.0
        calc_total = tent_val + food_val + transport_val + others_val
        if total_val == 0.0 and calc_total > 0.0:
            total_val = calc_total

        po_idx = best_map.get('po_number', 0)
        po_val = clean_str(sh.cell_value(r, po_idx)) if po_idx < len(row_vals) else ""
        if not po_val and sheet_po_number:
            po_val = sheet_po_number

        if not activity_val and not product_val and total_val == 0.0 and not date_val:
            continue

        activities.append({
            'date': date_val,
            'zdgm': zdgm_val,
            'tbm': tbm_val,
            'mdo': mdo_val,
            'territory': territory_val,
            'product': product_val,
            'crop': crop_val,
            'activity': activity_val,
            'village': village_val,
            'farmers': int(farmers_val),
            'tent': tent_val,
            'food': food_val,
            'transport': transport_val,
            'others': others_val,
            'total': total_val,
            'po_number': po_val,
            'source_file': file_name
        })

    return activities
