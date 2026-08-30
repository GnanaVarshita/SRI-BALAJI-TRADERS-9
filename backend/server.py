import os
import sys
import json
import threading
import queue
import urllib.parse
import re
import datetime
import shutil
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer

# Add current directory to path so we can import modules
BACKEND_DIR = Path(__file__).resolve().parent
sys.path.append(str(BACKEND_DIR))

import download_attachments
import excel_processor
from config import WORKSPACE_DIR, ENV_PATH, PROCESSED_DB_PATH
import state

DIST_DIR = WORKSPACE_DIR / "frontend/dist"
PORT = 5000

class WebLogRedirector:
    def __init__(self, old_stream):
        self.old_stream = old_stream
        self.line_buffer = ""
    def write(self, string):
        self.old_stream.write(string)
        self.line_buffer += string
        while "\n" in self.line_buffer:
            line, self.line_buffer = self.line_buffer.split("\n", 1)
            line = line.strip()
            if line:
                with state.LOGS_LOCK:
                    state.LOGS_BUFFER.append(line)
    def flush(self):
        self.old_stream.flush()

# Setup stdout/stderr redirection
sys.stdout = WebLogRedirector(sys.stdout)
sys.stderr = WebLogRedirector(sys.stderr)

def run_sync_task():
    try:
        download_attachments.main()
    except Exception as e:
        print(f"Sync thread crashed: {e}")
    finally:
        state.is_syncing = False

def get_content_type(file_path):
    suffix = file_path.suffix.lower()
    types = {
        '.html': 'text/html; charset=utf-8',
        '.css': 'text/css; charset=utf-8',
        '.js': 'application/javascript; charset=utf-8',
        '.mjs': 'application/javascript; charset=utf-8',
        '.json': 'application/json; charset=utf-8',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.ico': 'image/x-icon'
    }
    return types.get(suffix, 'application/octet-stream')

class APIHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS for local development (Vite dev server runs on port 5173)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        # 1. Handle API requests
        if path.startswith('/api/'):
            self.handle_api_get(path, parsed_url)
        else:
            # 2. Serve Static React files
            self.handle_static_get(path)

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if path.startswith('/api/'):
            content_length = int(self.headers.get('Content-Length', 0))
            content_type = self.headers.get('Content-Type', '')
            body_bytes = self.rfile.read(content_length) if content_length > 0 else b""
            self.handle_api_post(path, body_bytes, content_type)
        else:
            self.send_error(404, "Not Found")

    def handle_static_get(self, path):
        # Clean path to prevent directory traversal
        clean_path = path.lstrip('/')
        if not clean_path or clean_path == "":
            clean_path = "index.html"

        file_path = DIST_DIR / clean_path

        # If file doesn't exist, or if it maps to index.html (fallback for React routing)
        if not file_path.exists() or file_path.is_dir():
            file_path = DIST_DIR / "index.html"

        if not file_path.exists():
            # If the index.html still doesn't exist (frontend not built yet)
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(b"<h1>Sri Balaji Traders Dashboard</h1><p>Frontend is not compiled yet. Please run <code>npm run build</code> in the frontend folder, or run <code>run_dashboard.ps1</code>.</p>")
            return

        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', get_content_type(file_path))
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, f"Internal server error: {e}")

    def scan_downloads(self):
        files = []
        target_dirs = ["Corteva POs", "New Gen POs", "FMC POs", "CORTEVA", "FMC", "NEW GEN"]
        for t_dir in target_dirs:
            root_dir = WORKSPACE_DIR / t_dir
            if not root_dir.exists():
                continue
            for path in root_dir.rglob("*"):
                if path.is_file():
                    rel_path = path.relative_to(WORKSPACE_DIR)
                    parts = rel_path.parts
                    
                    company = parts[0].replace(" POs", "").strip() if len(parts) > 0 else "Unknown"
                    
                    year = "Unknown"
                    year_idx = -1
                    for idx, p in enumerate(parts):
                        year_match = re.search(r'(\d{4}-\d{4})', p)
                        if year_match:
                            year = year_match.group(1)
                            year_idx = idx
                            break
                    
                    area = "Unknown"
                    if year_idx > 1:
                        area = parts[1].strip()
                    elif year_idx == 1 and len(parts) > 2:
                        area = parts[2].replace(" POs", "").replace(" Pos", "").strip()
                    elif len(parts) > 1:
                        area = parts[1].strip()
                    
                    stat = path.stat()
                    mod_time = datetime.datetime.fromtimestamp(stat.st_mtime).isoformat()
                    
                    files.append({
                        "company": company,
                        "year": year,
                        "area": area,
                        "filename": path.name,
                        "sizeBytes": stat.st_size,
                        "modified": mod_time,
                        "relativePath": str(rel_path).replace("\\", "/")
                    })
        return files

    def handle_view_file(self, parsed_url):
        parsed_query = urllib.parse.parse_qs(parsed_url.query)
        file_rel_path = parsed_query.get('path', [None])[0]
        if not file_rel_path:
            self.send_error(400, "Bad Request: Missing path")
            return

        # Secure path resolution to prevent directory traversal
        file_abs_path = (WORKSPACE_DIR / file_rel_path).resolve()
        if not str(file_abs_path).startswith(str(WORKSPACE_DIR.resolve())) or not file_abs_path.is_file():
            self.send_error(403, "Forbidden: Invalid file path")
            return

        allowed = False
        for allowed_dir in ["Corteva POs", "New Gen POs", "FMC POs", "CORTEVA", "FMC", "NEW GEN"]:
            allowed_abs = (WORKSPACE_DIR / allowed_dir).resolve()
            if str(file_abs_path).startswith(str(allowed_abs)):
                allowed = True
                break

        if not allowed:
            self.send_error(403, "Forbidden: Access to this directory is not allowed")
            return

        try:
            with open(file_abs_path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', get_content_type(file_abs_path))
            self.send_header('Content-Length', str(len(content)))
            # Force browser download
            self.send_header('Content-Disposition', f'attachment; filename="{file_abs_path.name}"')
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, f"Internal server error: {e}")

    def handle_api_get(self, path, parsed_url):
        if path == '/api/config':
            env_vars = download_attachments.load_env(ENV_PATH)
            email_val = env_vars.get("GMAIL_EMAIL", "")
            has_pass = "GMAIL_APP_PASSWORD" in env_vars and len(env_vars["GMAIL_APP_PASSWORD"]) > 0
            
            response = {
                "email": email_val,
                "hasPassword": has_pass
            }
            self.send_json(200, response)

        elif path == '/api/status':
            with state.LOGS_LOCK:
                # Copy logs list
                current_logs = list(state.LOGS_BUFFER)
            
            # Simple count of downloaded files
            downloads_count = 0
            if PROCESSED_DB_PATH.exists():
                try:
                    with open(PROCESSED_DB_PATH, "r", encoding="utf-8") as f:
                        db = json.load(f)
                        downloads_count = sum(len(uids) for uids in db.values())
                except Exception:
                    pass

            response = {
                "isSyncing": state.is_syncing,
                "logs": current_logs,
                "totalSynced": downloads_count
            }
            self.send_json(200, response)

        elif path == '/api/downloads':
            files = self.scan_downloads()
            self.send_json(200, files)

        elif path == '/api/view-file':
            self.handle_view_file(parsed_url)

        else:
            self.send_error(404, "Not Found")

    def handle_browse_file(self):
        try:
            q = queue.Queue()
            def run_dialog():
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    path = filedialog.askopenfilename(
                        initialdir=str(WORKSPACE_DIR),
                        title="Select Budget Excel File",
                        filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb;*.csv"), ("All files", "*.*")]
                    )
                    root.destroy()
                    q.put(path)
                except Exception as e:
                    q.put(e)
            
            t = threading.Thread(target=run_dialog)
            t.start()
            t.join()
            res = q.get()
            if isinstance(res, Exception):
                raise res
                
            self.send_json(200, {"success": True, "filePath": res})
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Browse failed: {e}"})

    def handle_browse_folder(self):
        try:
            q = queue.Queue()
            def run_dialog():
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    path = filedialog.askdirectory(
                        initialdir=str(WORKSPACE_DIR),
                        title="Select Save Folder"
                    )
                    root.destroy()
                    q.put(path)
                except Exception as e:
                    q.put(e)
            
            t = threading.Thread(target=run_dialog)
            t.start()
            t.join()
            res = q.get()
            if isinstance(res, Exception):
                raise res
                
            self.send_json(200, {"success": True, "folderPath": res})
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Browse folder failed: {e}"})

    def handle_process_excel_local(self, data):
        try:
            file_path_str = data.get("filePath", "").strip()
            if not file_path_str:
                self.send_json(400, {"success": False, "message": "No file path provided"})
                return

            file_path = Path(file_path_str)
            if not file_path.exists() or not file_path.is_file():
                self.send_json(400, {"success": False, "message": f"File does not exist at: {file_path_str}"})
                return

            if not any(file_path_str.lower().endswith(ext) for ext in ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv']):
                self.send_json(400, {"success": False, "message": "Please select a valid Excel file (.xlsx, .xls, .xlsm, .xlsb, .csv)"})
                return

            company = data.get('company', 'Corteva Agriscience').strip()
            contact = data.get('contact', '').strip()
            designation = data.get('designation', '').strip()
            territory = data.get('territory', '').strip()
            date_str = data.get('date', '').strip()
            if not date_str:
                date_str = datetime.date.today().strftime('%d-%m-%Y')

            # 1. Validate the sheet (in-place)
            validation_res = excel_processor.validate_budget_sheet(file_path)
            
            if not validation_res["rows"]:
                self.send_json(400, {
                    "success": False,
                    "message": "The budget spreadsheet does not contain any valid data rows starting from row 12.",
                    "errors": validation_res["errors"]
                })
                return

            # 2. Append quotation sheets to the original file in-place
            excel_processor.generate_quotations(file_path, company, contact, designation, territory, date_str)

            self.send_json(200, {
                "success": True,
                "message": f"Quotation sheets generated and appended directly in-place to {file_path.name}!",
                "valid": validation_res["success"],
                "errors": validation_res["errors"],
                "totals": validation_res["totals"]
            })
            
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Excel generation failed: {e}"})

    def handle_generate_po_summary_local(self, data):
        try:
            input_path_str = data.get("inputPath", "").strip()
            save_folder_str = data.get("saveFolderPath", "").strip()
            output_name = data.get("outputName", "").strip()
            po_number = data.get("poNumber", "").strip()
            date_str = data.get("date", "").strip()
            contact = data.get("contact", "").strip()
            territory = data.get("territory", "").strip()
            
            if not input_path_str:
                self.send_json(400, {"success": False, "message": "Input Quotation file path is required"})
                return
            if not save_folder_str:
                self.send_json(400, {"success": False, "message": "Save folder path is required"})
                return
            if not output_name:
                self.send_json(400, {"success": False, "message": "Output PO Summary file name is required"})
                return
                
            input_path = Path(input_path_str)
            if not input_path.exists() or not input_path.is_file():
                self.send_json(400, {"success": False, "message": f"Input file does not exist at: {input_path_str}"})
                return
                
            save_folder = Path(save_folder_str)
            if not save_folder.exists() or not save_folder.is_dir():
                self.send_json(400, {"success": False, "message": f"Save folder does not exist at: {save_folder_str}"})
                return
                
            if not output_name.lower().endswith('.xlsx'):
                output_name += ".xlsx"
                
            output_path = save_folder / output_name
            
            if not date_str:
                date_str = datetime.date.today().strftime('%d-%m-%Y')
                
            excel_processor.generate_po_summary(input_path, output_path, po_number, date_str, contact, territory)
            
            self.send_json(200, {
                "success": True,
                "message": f"PO Summary sheet generated successfully as {output_name}!",
                "outputPath": str(output_path)
            })
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"PO Summary generation failed: {e}"})

    def handle_generate_fmc_summary_local(self, data):
        try:
            import fmc_summary_generator
            input_folder_str = data.get("inputFolderPath", "").strip()
            save_folder_str = data.get("saveFolderPath", "").strip()
            output_name = data.get("outputName", "").strip()
            territory = data.get("territory", "").strip()
            am_name = data.get("amName", "").strip() or "Madhavareddy"
            
            if not input_folder_str:
                self.send_json(400, {"success": False, "message": "Input FMC PDF folder path is required"})
                return
            if not save_folder_str:
                self.send_json(400, {"success": False, "message": "Save folder path is required"})
                return
                
            input_folder = Path(input_folder_str)
            if not input_folder.exists() or not input_folder.is_dir():
                self.send_json(400, {"success": False, "message": f"Input folder does not exist at: {input_folder_str}"})
                return
                
            save_folder = Path(save_folder_str)
            if not save_folder.exists() or not save_folder.is_dir():
                self.send_json(400, {"success": False, "message": f"Save folder does not exist at: {save_folder_str}"})
                return
                
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

            self.send_json(200, {
                "success": True,
                "message": msg,
                "outputPath": str(output_path)
            })
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"FMC PO Summary generation failed: {e}"})

    def handle_generate_fmc_step2_local(self, data):
        try:
            import fmc_step2_generator
            excel_path_str = data.get("excelPath", "").strip()
            territory = data.get("territory", "").strip()
            am_name = data.get("amName", "").strip() or "Madhavareddy"
            
            if not excel_path_str:
                self.send_json(400, {"success": False, "message": "Excel file path is required"})
                return
                
            excel_path = Path(excel_path_str)
            if not excel_path.exists() or not excel_path.is_file():
                self.send_json(400, {"success": False, "message": f"Excel file does not exist at: {excel_path_str}"})
                return

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

            self.send_json(200, {
                "success": True,
                "message": msg,
                "outputPath": str(excel_path)
            })
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"FMC Step 2 Summary generation failed: {e}"})

    def handle_format_tbm_summaries_local(self, data):
        try:
            import tbm_formatter
            folder_or_file_str = data.get("path", "").strip() or data.get("tbmFolderPath", "").strip()
            if not folder_or_file_str:
                self.send_json(400, {"success": False, "message": "TBM folder or file path is required"})
                return

            target_path = Path(folder_or_file_str)
            if not target_path.exists():
                self.send_json(400, {"success": False, "message": f"Path does not exist at: {folder_or_file_str}"})
                return

            if target_path.is_file():
                res = tbm_formatter.format_tbm_workbook(target_path)
            else:
                res = tbm_formatter.format_all_tbm_summaries_in_folder(target_path)

            if res.get("success"):
                self.send_json(200, res)
            else:
                self.send_json(400, res)
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"TBM Formatting failed: {e}"})

    def handle_generate_tbm_summary_local(self, data):
        try:
            import tbm_summary_generator
            input_folder_str = data.get("tbmFolderPath", "").strip()
            output_path_str = data.get("outputPath", "").strip()
            priority_po_list = data.get("priorityPoList", None)

            if not input_folder_str:
                self.send_json(400, {"success": False, "message": "TBM Summary Folder path is required"})
                return

            input_folder = Path(input_folder_str)
            if not input_folder.exists() or not input_folder.is_dir():
                self.send_json(400, {"success": False, "message": f"TBM Summary folder does not exist at: {input_folder_str}"})
                return

            res = tbm_summary_generator.generate_tbm_summary(
                tbm_folder_path=input_folder,
                output_path=output_path_str if output_path_str else None,
                priority_po_list=priority_po_list
            )

            if res.get("success"):
                self.send_json(200, res)
            else:
                self.send_json(400, res)
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"TBM Summary generation failed: {e}"})

    def handle_sync_tbm_cards_local(self, data):
        try:
            import card_sync_engine
            cards_path_str = data.get("cardsExcelPath", "").strip() or data.get("cardsSummaryPath", "").strip()
            tbm_path_str = data.get("tbmSummaryPath", "").strip() or data.get("tbmExcelPath", "").strip()
            output_path_str = data.get("outputPath", "").strip()
            sv_percent = float(data.get("serviceChargePercent", 5.0) or 5.0)

            if not cards_path_str:
                self.send_json(400, {"success": False, "message": "PO Cards Summary Excel file path is required"})
                return
            if not tbm_path_str:
                self.send_json(400, {"success": False, "message": "Consolidated TBM Summary Excel file path is required"})
                return

            cards_path = Path(cards_path_str)
            tbm_path = Path(tbm_path_str)

            if not cards_path.exists() or not cards_path.is_file():
                self.send_json(400, {"success": False, "message": f"Cards summary file does not exist at: {cards_path_str}"})
                return
            if not tbm_path.exists() or not tbm_path.is_file():
                self.send_json(400, {"success": False, "message": f"TBM summary file does not exist at: {tbm_path_str}"})
                return

            res = card_sync_engine.sync_tbm_with_cards(
                cards_excel_path=cards_path,
                tbm_summary_excel_path=tbm_path,
                output_path=output_path_str if output_path_str else None,
                service_charge_percent=sv_percent
            )

            if res.get("success"):
                self.send_json(200, res)
            else:
                self.send_json(400, res)
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Cards synchronization failed: {e}"})

    def handle_generate_invoices_local(self, data):
        try:
            import invoice_generator
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
                self.send_json(400, {"success": False, "message": "All-TBMs Summary Excel file path is required"})
                return
            if not save_folder_path:
                self.send_json(400, {"success": False, "message": "Save folder path is required"})
                return
            if not invoice_number:
                self.send_json(400, {"success": False, "message": "Invoice Number is required"})
                return
            if not po_number:
                self.send_json(400, {"success": False, "message": "PO Number is a mandatory field"})
                return

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
            self.send_json(200, res)
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Invoice generation failed: {e}"})

    def handle_scan_pos_in_summary_local(self, data):
        try:
            import invoice_generator
            tbm_summary_path = data.get("tbmSummaryPath", "").strip()
            if not tbm_summary_path:
                self.send_json(400, {"success": False, "message": "TBM Summary path is required"})
                return
            pos = invoice_generator.scan_pos_in_summary(tbm_summary_path)
            self.send_json(200, {"success": True, "pos": pos})
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Scan POs failed: {e}"})

    def handle_sync_details_of_bills_local(self, data):
        try:
            import details_of_bills_generator
            details_excel_path = data.get("detailsExcelPath", "").strip()
            invoices_folder_path = data.get("invoicesFolderPath", "").strip()
            budget_cards_path = data.get("budgetCardsPath", "").strip() or None
            financial_year = data.get("financialYear", "").strip() or "APRIL 2026 to MARCH 2027"

            if not details_excel_path:
                self.send_json(400, {"success": False, "message": "Details of Bills Excel file path is required"})
                return
            if not invoices_folder_path:
                self.send_json(400, {"success": False, "message": "Invoices Folder path is required"})
                return

            res = details_of_bills_generator.scan_and_append_invoices(
                details_excel_path=details_excel_path,
                invoices_folder_path=invoices_folder_path,
                budget_cards_path=budget_cards_path,
                financial_year=financial_year
            )
            self.send_json(200, res)
        except Exception as e:
            self.send_json(500, {"success": False, "message": f"Details of Bills synchronization failed: {e}"})

    def handle_api_post(self, path, body_bytes, content_type):
        global is_syncing, sync_thread
        
        # Parse JSON if request content type is JSON
        data = {}
        if 'application/json' in content_type.lower() or not content_type:
            try:
                data = json.loads(body_bytes.decode('utf-8')) if body_bytes else {}
            except Exception:
                self.send_error(400, "Bad Request: Invalid JSON")
                return

        if path == '/api/browse-file':
            self.handle_browse_file()
            return

        if path == '/api/browse-folder':
            self.handle_browse_folder()
            return

        if path == '/api/process-excel':
            self.handle_process_excel_local(data)
            return

        if path == '/api/generate-summary':
            self.handle_generate_po_summary_local(data)
            return

        if path == '/api/generate-fmc-summary':
            self.handle_generate_fmc_summary_local(data)
            return

        if path == '/api/generate-fmc-step2':
            self.handle_generate_fmc_step2_local(data)
            return

        if path == '/api/format-tbm-summaries':
            self.handle_format_tbm_summaries_local(data)
            return

        if path == '/api/generate-tbm-summary':
            self.handle_generate_tbm_summary_local(data)
            return

        if path == '/api/sync-tbm-cards':
            self.handle_sync_tbm_cards_local(data)
            return

        if path == '/api/generate-invoices':
            self.handle_generate_invoices_local(data)
            return

        if path == '/api/scan-pos-in-summary':
            self.handle_scan_pos_in_summary_local(data)
            return

        if path == '/api/sync-details-of-bills':
            self.handle_sync_details_of_bills_local(data)
            return

        if path == '/api/config':
            email_val = data.get("email", "").strip()
            pass_val = data.get("password", "").strip()

            if not email_val:
                self.send_json(400, {"success": False, "message": "Email is required"})
                return

            try:
                env_vars = download_attachments.load_env(ENV_PATH)
                final_pass = pass_val if pass_val else env_vars.get("GMAIL_APP_PASSWORD", "")

                with open(ENV_PATH, "w", encoding="utf-8") as f:
                    f.write(f"GMAIL_EMAIL={email_val}\n")
                    f.write(f"GMAIL_APP_PASSWORD={final_pass}\n")
                self.send_json(200, {"success": True, "message": "Credentials saved"})
            except Exception as e:
                self.send_json(500, {"success": False, "message": f"Failed to save: {e}"})

        elif path == '/api/sync':
            if state.is_syncing:
                self.send_json(400, {"success": False, "message": "Sync is already running"})
                return

            with state.LOGS_LOCK:
                state.LOGS_BUFFER.clear()
                state.LOGS_BUFFER.append("Initializing sync request...")

            state.is_syncing = True
            state.sync_thread = threading.Thread(target=run_sync_task, daemon=True)
            state.sync_thread.start()
            self.send_json(200, {"success": True, "message": "Sync started"})

        elif path == '/api/reset':
            if state.is_syncing:
                self.send_json(400, {"success": False, "message": "Cannot reset history while sync is running"})
                return

            if not PROCESSED_DB_PATH.exists():
                self.send_json(200, {"success": True, "message": "Sync history is already empty"})
                return

            try:
                os.remove(PROCESSED_DB_PATH)
                with state.LOGS_LOCK:
                    state.LOGS_BUFFER.append(">>> Sync history reset successfully. Next run will process all emails.")
                self.send_json(200, {"success": True, "message": "Sync history cleared"})
            except Exception as e:
                self.send_json(500, {"success": False, "message": f"Failed to reset: {e}"})

        elif path == '/api/open-folder':
            folder_type = data.get("folder", "")
            if folder_type == "corteva":
                folder_path = WORKSPACE_DIR / "Corteva POs"
            elif folder_type == "newgen":
                folder_path = WORKSPACE_DIR / "New Gen POs"
            elif folder_type == "fmc":
                folder_path = WORKSPACE_DIR / "FMC POs"
            else:
                self.send_json(400, {"success": False, "message": "Invalid folder type"})
                return

            try:
                os.makedirs(folder_path, exist_ok=True)
                os.startfile(folder_path)
                self.send_json(200, {"success": True, "message": f"Opened folder: {folder_path.name}"})
            except Exception as e:
                self.send_json(500, {"success": False, "message": f"Failed to open folder: {e}"})
        else:
            self.send_error(404, "Not Found")

    def send_json(self, status, data):
        content = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format, *args):
        # Suppress logging every single GET/POST request to avoid cluttering stderr
        pass

def main():
    print(f"Starting server on http://localhost:{PORT}")
    server = HTTPServer(('0.0.0.0', PORT), APIHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server.")
        server.server_close()

if __name__ == "__main__":
    main()
