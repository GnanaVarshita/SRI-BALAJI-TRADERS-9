import os
import sys
import json
import threading
import urllib.parse
import re
import datetime
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer

# Add current directory to path so we can import download_attachments
BACKEND_DIR = Path(__file__).resolve().parent
sys.path.append(str(BACKEND_DIR))

import download_attachments

# Configurations
WORKSPACE_DIR = Path("D:/SRI BALAJI TRADERS")
ENV_PATH = WORKSPACE_DIR / ".env"
PROCESSED_DB_PATH = WORKSPACE_DIR / "processed_emails.json"
DIST_DIR = WORKSPACE_DIR / "frontend/dist"

PORT = 5000

# Global state
LOGS_BUFFER = []
LOGS_LOCK = threading.Lock()
is_syncing = False
sync_thread = None

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
                with LOGS_LOCK:
                    LOGS_BUFFER.append(line)
                    if len(LOGS_BUFFER) > 1000:
                        LOGS_BUFFER.pop(0)
    def flush(self):
        self.old_stream.flush()

# Setup stdout/stderr redirection
sys.stdout = WebLogRedirector(sys.stdout)
sys.stderr = WebLogRedirector(sys.stderr)

def run_sync_task():
    global is_syncing
    try:
        download_attachments.main()
    except Exception as e:
        print(f"Sync thread crashed: {e}")
    finally:
        is_syncing = False

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
            # Read request body
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8') if content_length > 0 else ""
            self.handle_api_post(path, body)
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
        target_dirs = ["Corteva POs", "New Gen POs"]
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
                    area = "Unknown"
                    if len(parts) > 1:
                        year_match = re.search(r'(\d{4}-\d{4})', parts[1])
                        if year_match:
                            year = year_match.group(1)
                    if len(parts) > 2:
                        area_match = re.match(r'([a-zA-Z\s]+)', parts[2])
                        if area_match:
                            area = area_match.group(1).strip()
                    
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
        for allowed_dir in ["Corteva POs", "New Gen POs"]:
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
            global is_syncing
            with LOGS_LOCK:
                # Copy logs list
                current_logs = list(LOGS_BUFFER)
            
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
                "isSyncing": is_syncing,
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

    def handle_api_post(self, path, body):
        global is_syncing, sync_thread
        
        try:
            data = json.loads(body) if body else {}
        except Exception:
            self.send_error(400, "Bad Request: Invalid JSON")
            return

        if path == '/api/config':
            email_val = data.get("email", "").strip()
            pass_val = data.get("password", "").strip()

            if not email_val:
                self.send_json(400, {"success": False, "message": "Email is required"})
                return

            try:
                # Load current env first, in case they left password blank (meaning no change)
                env_vars = download_attachments.load_env(ENV_PATH)
                final_pass = pass_val if pass_val else env_vars.get("GMAIL_APP_PASSWORD", "")

                with open(ENV_PATH, "w", encoding="utf-8") as f:
                    f.write(f"GMAIL_EMAIL={email_val}\n")
                    f.write(f"GMAIL_APP_PASSWORD={final_pass}\n")
                self.send_json(200, {"success": True, "message": "Credentials saved"})
            except Exception as e:
                self.send_json(500, {"success": False, "message": f"Failed to save: {e}"})

        elif path == '/api/sync':
            if is_syncing:
                self.send_json(400, {"success": False, "message": "Sync is already running"})
                return

            # Clear logs buffer
            with LOGS_LOCK:
                LOGS_BUFFER.clear()
                LOGS_BUFFER.append("Initializing sync request...")

            is_syncing = True
            sync_thread = threading.Thread(target=run_sync_task, daemon=True)
            sync_thread.start()
            self.send_json(200, {"success": True, "message": "Sync started"})

        elif path == '/api/reset':
            if is_syncing:
                self.send_json(400, {"success": False, "message": "Cannot reset history while sync is running"})
                return

            if not PROCESSED_DB_PATH.exists():
                self.send_json(200, {"success": True, "message": "Sync history is already empty"})
                return

            try:
                os.remove(PROCESSED_DB_PATH)
                with LOGS_LOCK:
                    LOGS_BUFFER.append(">>> Sync history reset successfully. Next run will process all emails.")
                self.send_json(200, {"success": True, "message": "Sync history cleared"})
            except Exception as e:
                self.send_json(500, {"success": False, "message": f"Failed to reset: {e}"})

        elif path == '/api/open-folder':
            folder_type = data.get("folder", "")
            if folder_type == "corteva":
                folder_path = WORKSPACE_DIR / "Corteva POs"
            elif folder_type == "newgen":
                folder_path = WORKSPACE_DIR / "New Gen POs"
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
