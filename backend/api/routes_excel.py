"""
Excel & Business Process API Routes
Handles quotations, summaries, TBM formatting, card synchronization, and invoice generation.
"""

import re
import datetime
from pathlib import Path

import excel_processor
import corteva_master_summary
import fmc_summary_generator
import fmc_step2_generator
import tbm_formatter
import tbm_summary_generator
import card_sync_engine
import invoice_generator
import details_of_bills_generator


def handle_process_excel(data):
    """Handles POST /api/process-excel."""
    try:
        file_path_str = data.get("filePath", "").strip()
        if not file_path_str:
            return 400, {"success": False, "message": "No file path provided"}

        file_path = Path(file_path_str)
        if not file_path.exists() or not file_path.is_file():
            return 400, {"success": False, "message": f"File does not exist at: {file_path_str}"}

        if not any(file_path_str.lower().endswith(ext) for ext in ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv']):
            return 400, {"success": False, "message": "Please select a valid Excel file (.xlsx, .xls, .xlsm, .xlsb, .csv)"}

        company = data.get('company', 'Corteva Agriscience').strip()
        contact = data.get('contact', '').strip()
        designation = data.get('designation', '').strip()
        territory = data.get('territory', '').strip()
        date_str = data.get('date', '').strip()
        if not date_str:
            date_str = datetime.date.today().strftime('%d-%m-%Y')

        validation_res = excel_processor.validate_budget_sheet(file_path)
        if not validation_res["rows"]:
            return 400, {
                "success": False,
                "message": "The budget spreadsheet does not contain any valid data rows starting from row 12.",
                "errors": validation_res["errors"]
            }

        excel_processor.generate_quotations(file_path, company, contact, designation, territory, date_str)

        return 200, {
            "success": True,
            "message": f"Quotation sheets generated and appended directly in-place to {file_path.name}!",
            "valid": validation_res["success"],
            "errors": validation_res["errors"],
            "totals": validation_res["totals"]
        }
    except Exception as e:
        return 500, {"success": False, "message": f"Excel generation failed: {e}"}


def handle_generate_po_summary(data):
    """Handles POST /api/generate-summary."""
    try:
        input_path_str = data.get("inputPath", "").strip()
        save_folder_str = data.get("saveFolderPath", "").strip()
        output_name = data.get("outputName", "").strip()
        po_number = data.get("poNumber", "").strip()
        date_str = data.get("date", "").strip()
        contact = data.get("contact", "").strip()
        territory = data.get("territory", "").strip()
        
        if not input_path_str:
            return 400, {"success": False, "message": "Input Quotation file path is required"}
        if not save_folder_str:
            return 400, {"success": False, "message": "Save folder path is required"}
        if not output_name:
            return 400, {"success": False, "message": "Output PO Summary file name is required"}
            
        input_path = Path(input_path_str)
        if not input_path.exists() or not input_path.is_file():
            return 400, {"success": False, "message": f"Input file does not exist at: {input_path_str}"}
            
        save_folder = Path(save_folder_str)
        if not save_folder.exists() or not save_folder.is_dir():
            return 400, {"success": False, "message": f"Save folder does not exist at: {save_folder_str}"}
            
        if not output_name.lower().endswith('.xlsx'):
            output_name += ".xlsx"
            
        output_path = save_folder / output_name
        if not date_str:
            date_str = datetime.date.today().strftime('%d-%m-%Y')
            
        excel_processor.generate_po_summary(input_path, output_path, po_number, date_str, contact, territory)
        
        return 200, {
            "success": True,
            "message": f"PO Summary sheet generated successfully as {output_name}!",
            "outputPath": str(output_path)
        }
    except Exception as e:
        return 500, {"success": False, "message": f"PO Summary generation failed: {e}"}


def handle_generate_corteva_master_summary(data):
    """Handles POST /api/generate-corteva-master-summary."""
    try:
        folder_path_str = data.get("folderPath", "").strip()
        output_name = data.get("outputName", "").strip()
        territory = data.get("territory", "").strip()
        zdgm = data.get("zdgm", "").strip()
        budget_season = data.get("budgetSeason", "").strip() or "Kharif"
        
        if not folder_path_str:
            return 400, {"success": False, "message": "Input folder path containing PO summary cards is required"}
            
        folder_path = Path(folder_path_str)
        if not folder_path.exists() or not folder_path.is_dir():
            return 400, {"success": False, "message": f"Input folder does not exist at: {folder_path_str}"}
            
        result = corteva_master_summary.generate_master_po_summary(
            folder_path=folder_path,
            output_name=output_name,
            territory=territory,
            zdgm=zdgm,
            budget_season=budget_season
        )
        return 200, result
    except Exception as e:
        return 500, {"success": False, "message": f"Corteva Master PO Summary generation failed: {e}"}


def handle_generate_fmc_summary(data):
    """Handles POST /api/generate-fmc-summary."""
    try:
        input_folder_str = data.get("inputFolderPath", "").strip()
        save_folder_str = data.get("saveFolderPath", "").strip()
        output_name = data.get("outputName", "").strip()
        territory = data.get("territory", "").strip()
        am_name = data.get("amName", "").strip() or "Madhavareddy"
        
        if not input_folder_str:
            return 400, {"success": False, "message": "Input FMC PDF folder path is required"}
        if not save_folder_str:
            return 400, {"success": False, "message": "Save folder path is required"}
            
        input_folder = Path(input_folder_str)
        if not input_folder.exists() or not input_folder.is_dir():
            return 400, {"success": False, "message": f"Input folder does not exist at: {input_folder_str}"}
            
        save_folder = Path(save_folder_str)
        if not save_folder.exists() or not save_folder.is_dir():
            return 400, {"success": False, "message": f"Save folder does not exist at: {save_folder_str}"}
            
        if not territory:
            for part in input_folder.parts:
                clean = part.upper().replace('-FMC', '').replace(' POS', '').replace(' PO', '').strip()
                if clean in ['NANDYALA', 'NANDYAL', 'NELLORE', 'SURYAPET', 'KURNOOL']:
                    territory = clean.title()
                    break
            if not territory:
                match = re.search(r'(nandyala|nandyal|nellore|suryapet|kurnool)', str(input_folder), re.I)
                if match:
                    territory = match.group(1).title()
                else:
                    territory = "Nandyala"

        if not output_name:
            output_name = f"{territory} FMC Budget.xlsx"
            
        if not output_name.lower().endswith('.xlsx'):
            output_name += ".xlsx"
            
        output_path = save_folder / output_name
        new_added = fmc_summary_generator.generate_fmc_summary(input_folder, output_path, territory, am_name)
        
        if new_added > 0:
            msg = f"FMC Master Budget sheet generated/updated as {output_name}! ({new_added} new PO(s) appended to Sheet1)"
        else:
            msg = f"FMC Master Budget sheet ({output_name}) is already up to date! (0 new POs to add)"

        return 200, {
            "success": True,
            "message": msg,
            "outputPath": str(output_path)
        }
    except Exception as e:
        return 500, {"success": False, "message": f"FMC PO Summary generation failed: {e}"}


def handle_generate_fmc_step2(data):
    """Handles POST /api/generate-fmc-step2."""
    try:
        excel_path_str = data.get("excelPath", "").strip()
        territory = data.get("territory", "").strip()
        am_name = data.get("amName", "").strip() or "Madhavareddy"
        
        if not excel_path_str:
            return 400, {"success": False, "message": "Excel file path is required"}
            
        excel_path = Path(excel_path_str)
        if not excel_path.exists() or not excel_path.is_file():
            return 400, {"success": False, "message": f"Excel file does not exist at: {excel_path_str}"}

        if not territory:
            for part in excel_path.parts:
                clean = part.upper().replace('-FMC', '').replace(' POS', '').replace(' PO', '').strip()
                if clean in ['NANDYALA', 'NANDYAL', 'NELLORE', 'SURYAPET', 'KURNOOL']:
                    territory = clean.title()
                    break
            if not territory:
                match = re.search(r'(nandyala|nandyal|nellore|suryapet|kurnool)', excel_path.name, re.I)
                if match:
                    territory = match.group(1).title()
                else:
                    territory = "Nandyala"

        n_new_cards, n_total_sheets = fmc_step2_generator.generate_fmc_step2_summaries(excel_path, territory, am_name)
        
        if n_new_cards > 0:
            msg = f"Generated {n_new_cards} new PO summary card(s) across {n_total_sheets} card sheet(s) in {excel_path.name}!"
        else:
            msg = f"All PO summary cards in {excel_path.name} are already up to date! (0 new PO cards to generate)"

        return 200, {
            "success": True,
            "message": msg,
            "outputPath": str(excel_path)
        }
    except Exception as e:
        return 500, {"success": False, "message": f"FMC Step 2 Summary generation failed: {e}"}


def handle_format_tbm_summaries(data):
    """Handles POST /api/format-tbm-summaries."""
    try:
        folder_or_file_str = data.get("path", "").strip() or data.get("tbmFolderPath", "").strip()
        if not folder_or_file_str:
            return 400, {"success": False, "message": "TBM folder or file path is required"}

        target_path = Path(folder_or_file_str)
        if not target_path.exists():
            return 400, {"success": False, "message": f"Path does not exist at: {folder_or_file_str}"}

        if target_path.is_file():
            res = tbm_formatter.format_tbm_workbook(target_path)
        else:
            res = tbm_formatter.format_all_tbm_summaries_in_folder(target_path)

        status_code = 200 if res.get("success") else 400
        return status_code, res
    except Exception as e:
        return 500, {"success": False, "message": f"TBM Formatting failed: {e}"}


def handle_generate_tbm_summary(data):
    """Handles POST /api/generate-tbm-summary."""
    try:
        input_folder_str = data.get("tbmFolderPath", "").strip()
        output_path_str = data.get("outputPath", "").strip()
        priority_po_list = data.get("priorityPoList", None)

        if not input_folder_str:
            return 400, {"success": False, "message": "TBM Summary Folder path is required"}

        input_folder = Path(input_folder_str)
        if not input_folder.exists() or not input_folder.is_dir():
            return 400, {"success": False, "message": f"TBM Summary folder does not exist at: {input_folder_str}"}

        res = tbm_summary_generator.generate_tbm_summary(
            tbm_folder_path=input_folder,
            output_path=output_path_str if output_path_str else None,
            priority_po_list=priority_po_list
        )
        status_code = 200 if res.get("success") else 400
        return status_code, res
    except Exception as e:
        return 500, {"success": False, "message": f"TBM Summary generation failed: {e}"}


def handle_sync_tbm_cards(data):
    """Handles POST /api/sync-tbm-cards."""
    try:
        cards_path_str = data.get("cardsExcelPath", "").strip() or data.get("cardsSummaryPath", "").strip()
        tbm_path_str = data.get("tbmSummaryPath", "").strip() or data.get("tbmExcelPath", "").strip()
        output_path_str = data.get("outputPath", "").strip()
        sv_percent = float(data.get("serviceChargePercent", 5.0) or 5.0)

        if not cards_path_str:
            return 400, {"success": False, "message": "PO Cards Summary Excel file path is required"}
        if not tbm_path_str:
            return 400, {"success": False, "message": "Consolidated TBM Summary Excel file path is required"}

        cards_path = Path(cards_path_str)
        tbm_path = Path(tbm_path_str)

        if not cards_path.exists() or not cards_path.is_file():
            return 400, {"success": False, "message": f"Cards summary file does not exist at: {cards_path_str}"}
        if not tbm_path.exists() or not tbm_path.is_file():
            return 400, {"success": False, "message": f"TBM summary file does not exist at: {tbm_path_str}"}

        res = card_sync_engine.sync_tbm_with_cards(
            cards_excel_path=cards_path,
            tbm_summary_excel_path=tbm_path,
            output_path=output_path_str if output_path_str else None,
            service_charge_percent=sv_percent
        )
        status_code = 200 if res.get("success") else 400
        return status_code, res
    except Exception as e:
        return 500, {"success": False, "message": f"Cards synchronization failed: {e}"}


def handle_generate_invoices(data):
    """Handles POST /api/generate-invoices."""
    try:
        company = data.get("company", "Corteva").strip()
        tbm_summary_path = data.get("tbmSummaryPath", "").strip()
        save_folder_path = data.get("saveFolderPath", "").strip()
        invoice_number = data.get("invoiceNumber", "").strip()
        po_number = data.get("poNumber", "").strip()
        service_charge_pct = float(data.get("serviceChargePercent", 5.0) or 5.0)
        invoice_date = data.get("invoiceDate", "").strip() or None
        po_value = float(data.get("poValue", 0.0) or 0.0) if data.get("poValue") else None
        requester_name = data.get("requesterName", "").strip() or None
        area = data.get("area", "").strip() or None

        if not tbm_summary_path:
            return 400, {"success": False, "message": "All-TBMs Summary Excel file path is required"}
        if not save_folder_path:
            return 400, {"success": False, "message": "Save folder path is required"}
        if not invoice_number:
            return 400, {"success": False, "message": "Invoice Number is required"}
        if not po_number:
            return 400, {"success": False, "message": "PO Number is a mandatory field"}

        res = invoice_generator.generate_or_update_invoice(
            company=company,
            tbm_summary_path=tbm_summary_path,
            save_folder_path=save_folder_path,
            invoice_number=invoice_number,
            po_number=po_number,
            service_charge_pct=service_charge_pct,
            invoice_date=invoice_date,
            po_value=po_value,
            requester_name=requester_name,
            area=area
        )
        return 200, res
    except Exception as e:
        return 500, {"success": False, "message": f"Invoice generation failed: {e}"}


def handle_scan_pos_in_summary(data):
    """Handles POST /api/scan-pos-in-summary."""
    try:
        tbm_summary_path = data.get("tbmSummaryPath", "").strip()
        if not tbm_summary_path:
            return 400, {"success": False, "message": "TBM Summary path is required"}
        pos = invoice_generator.scan_pos_in_summary(tbm_summary_path)
        return 200, {"success": True, "pos": pos}
    except Exception as e:
        return 500, {"success": False, "message": f"Scan POs failed: {e}"}


def handle_sync_details_of_bills(data):
    """Handles POST /api/sync-details-of-bills."""
    try:
        details_excel_path = data.get("detailsExcelPath", "").strip()
        invoices_folder_path = data.get("invoicesFolderPath", "").strip()
        budget_cards_path = data.get("budgetCardsPath", "").strip() or None
        financial_year = data.get("financialYear", "").strip() or "APRIL 2026 to MARCH 2027"

        if not details_excel_path:
            return 400, {"success": False, "message": "Details of Bills Excel file path is required"}
        if not invoices_folder_path:
            return 400, {"success": False, "message": "Invoices Folder path is required"}

        res = details_of_bills_generator.scan_and_append_invoices(
            details_excel_path=details_excel_path,
            invoices_folder_path=invoices_folder_path,
            budget_cards_path=budget_cards_path,
            financial_year=financial_year
        )
        return 200, res
    except Exception as e:
        return 500, {"success": False, "message": f"Details of Bills synchronization failed: {e}"}
