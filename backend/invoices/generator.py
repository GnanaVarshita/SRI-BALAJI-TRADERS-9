"""
Invoice Generator Coordinator
Orchestrates extracting PO activity records, building company-specific sheets,
and writing formatted Excel workbooks.
"""

import datetime
from pathlib import Path
import openpyxl

from common.formatters import clean_str, format_short_iv, format_full_iv
from common.currency import num_to_indian_words
from common.styles import get_base_styles
from .extractor import extract_tables_for_po
from .corteva import (
    build_corteva_sheet1_invoice,
    build_corteva_sheet2_details,
    build_corteva_summary_sheet,
)
from .fmc import (
    build_fmc_sheet1_invoice,
    build_fmc_sheet2_details,
)


def generate_or_update_invoice(
    company,
    tbm_summary_path,
    save_folder_path,
    invoice_number,
    po_number,
    service_charge_pct=5.0,
    invoice_date=None,
    po_value=None,
    requester_name=None,
    area=None
):
    """
    Main function to generate or update a PO Tax Invoice workbook.
    """
    if not po_number or not str(po_number).strip():
        raise ValueError("PO Number is a mandatory field.")

    if not invoice_number or not str(invoice_number).strip():
        raise ValueError("Invoice Number is a mandatory field.")

    target_po = str(po_number).strip()
    inv_num_raw = str(invoice_number).strip()
    company_clean = str(company).strip().title()  # "Corteva" or "Fmc"

    short_iv = format_short_iv(inv_num_raw)
    full_iv = format_full_iv(inv_num_raw)

    if not invoice_date:
        invoice_date = datetime.date.today().strftime("%d-%m-%Y")

    save_dir = Path(save_folder_path).resolve()
    save_dir.mkdir(parents=True, exist_ok=True)

    # 1. Extract activity records for this PO from All-TBMs-Summary.xlsx
    records, metadata = extract_tables_for_po(tbm_summary_path, target_po)
    if not records:
        raise ValueError(f"No activity expense records found for PO {target_po} in {Path(tbm_summary_path).name}")

    if requester_name:
        metadata['zdgm'] = requester_name
        metadata['amm'] = requester_name

    if area and str(area).strip():
        metadata['area'] = str(area).strip()
        metadata['territory'] = str(area).strip()

    # Group by Activity for Sheet1 Particulars
    activity_groups = {}
    for r in records:
        act = clean_str(r.get('activity', '')).upper() or "GENERAL"
        if act not in activity_groups:
            activity_groups[act] = {'qty': 0, 'raw_amount': 0.0, 'rows': []}
        activity_groups[act]['qty'] += 1
        activity_groups[act]['raw_amount'] += float(r.get('total', 0.0))
        activity_groups[act]['rows'].append(r)

    styles = get_base_styles()

    # 2. Check if an invoice file for this invoice number / PO already exists
    existing_file = None
    for f in save_dir.glob("*.xlsx"):
        if f.stem == short_iv or f.stem == full_iv or f.stem.lower() == inv_num_raw.lower() or target_po.upper() in f.stem.upper():
            existing_file = f
            break

    is_update = existing_file is not None and existing_file.exists()
    out_file_name = f"{short_iv}.xlsx" if not is_update else existing_file.name
    out_path = save_dir / out_file_name

    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    ws_sheet1 = wb.create_sheet(title="Sheet1")
    ws_sheet2 = wb.create_sheet(title="Sheet2")

    if company_clean.startswith("Corteva"):
        r_subtotal, r_grand = build_corteva_sheet1_invoice(
            ws_sheet1, full_iv, invoice_date, target_po, metadata, activity_groups, service_charge_pct, styles
        )
        build_corteva_sheet2_details(ws_sheet2, short_iv, records, service_charge_pct, styles)
        ws_summary = wb.create_sheet(title="Sheet4")
        build_corteva_summary_sheet(
            ws_summary, full_iv, invoice_date, target_po, metadata, po_value, r_subtotal, r_grand, styles
        )
    else:
        # FMC / New Gen
        r_subtotal, r_grand = build_fmc_sheet1_invoice(
            ws_sheet1, full_iv, invoice_date, target_po, metadata, activity_groups, service_charge_pct, styles
        )
        build_fmc_sheet2_details(ws_sheet2, short_iv, records, styles, service_charge_pct)

    wb.save(out_path)
    wb.close()

    total_activities = len(records)
    total_raw_amount = sum(float(r.get('total', 0.0)) for r in records)
    total_with_sc = total_raw_amount * (1 + service_charge_pct / 100.0)
    cgst_amt = round(total_with_sc * 0.09, 2)
    sgst_amt = round(total_with_sc * 0.09, 2)
    grand_total = round(total_with_sc + cgst_amt + sgst_amt)

    return {
        "success": True,
        "isUpdate": is_update,
        "message": f"Invoice {full_iv} {'updated & appended' if is_update else 'generated successfully'} for PO {target_po} in {out_path.name}!",
        "outputPath": str(out_path),
        "invoiceNo": full_iv,
        "shortInvoiceNo": short_iv,
        "poNumber": target_po,
        "company": company_clean,
        "area": metadata.get('area') or metadata.get('territory', ''),
        "invoiceDate": invoice_date,
        "totalActivities": total_activities,
        "subTotalExcGst": round(total_with_sc, 2),
        "grandTotalIncGst": grand_total,
        "grandTotalWords": num_to_indian_words(grand_total)
    }
