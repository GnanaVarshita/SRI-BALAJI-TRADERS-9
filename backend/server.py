"""
Sri Balaji Traders PO Web Dashboard Server
Modular Python HTTP & API Server
"""

import sys
import json
import urllib.parse
from pathlib import Path
from http.server import BaseHTTPRequestHandler, HTTPServer

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.append(str(BACKEND_DIR))

from api import (
    WebLogRedirector,
    serve_static,
    scan_downloaded_files,
    handle_view_file,
    handle_get_config,
    handle_save_config,
    handle_get_status,
    handle_start_sync,
    handle_reset_sync,
    handle_open_folder,
    handle_browse_file,
    handle_browse_folder,
    handle_process_excel,
    handle_generate_po_summary,
    handle_generate_corteva_master_summary,
    handle_generate_fmc_summary,
    handle_generate_fmc_step2,
    handle_format_tbm_summaries,
    handle_generate_tbm_summary,
    handle_sync_tbm_cards,
    handle_generate_invoices,
    handle_scan_pos_in_summary,
    handle_sync_details_of_bills,
)

PORT = 5000

# Setup web-visible stdout/stderr redirection
sys.stdout = WebLogRedirector(sys.stdout)
sys.stderr = WebLogRedirector(sys.stderr)


class APIHandler(BaseHTTPRequestHandler):
    """HTTP request handler with CORS support and modular endpoint routing."""

    def end_headers(self):
        # Allow cross-origin requests for Vite dev server (e.g. port 5173)
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

        if path == '/api/config':
            code, resp = handle_get_config()
            self.send_json(code, resp)
        elif path == '/api/status':
            code, resp = handle_get_status()
            self.send_json(code, resp)
        elif path == '/api/downloads':
            self.send_json(200, scan_downloaded_files())
        elif path == '/api/view-file':
            handle_view_file(self, parsed_url)
        elif path.startswith('/api/'):
            self.send_error(404, "Not Found")
        else:
            serve_static(self, path)

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if not path.startswith('/api/'):
            self.send_error(404, "Not Found")
            return

        content_length = int(self.headers.get('Content-Length', 0))
        content_type = self.headers.get('Content-Type', '')
        body_bytes = self.rfile.read(content_length) if content_length > 0 else b""

        data = {}
        if 'application/json' in content_type.lower() or not content_type:
            try:
                data = json.loads(body_bytes.decode('utf-8')) if body_bytes else {}
            except Exception:
                self.send_error(400, "Bad Request: Invalid JSON")
                return

        # Route table mapping endpoints to their respective handler functions
        routes = {
            '/api/browse-file': lambda: handle_browse_file(),
            '/api/browse-folder': lambda: handle_browse_folder(),
            '/api/process-excel': lambda: handle_process_excel(data),
            '/api/generate-summary': lambda: handle_generate_po_summary(data),
            '/api/generate-corteva-master-summary': lambda: handle_generate_corteva_master_summary(data),
            '/api/generate-fmc-summary': lambda: handle_generate_fmc_summary(data),
            '/api/generate-fmc-step2': lambda: handle_generate_fmc_step2(data),
            '/api/format-tbm-summaries': lambda: handle_format_tbm_summaries(data),
            '/api/generate-tbm-summary': lambda: handle_generate_tbm_summary(data),
            '/api/sync-tbm-cards': lambda: handle_sync_tbm_cards(data),
            '/api/generate-invoices': lambda: handle_generate_invoices(data),
            '/api/scan-pos-in-summary': lambda: handle_scan_pos_in_summary(data),
            '/api/sync-details-of-bills': lambda: handle_sync_details_of_bills(data),
            '/api/config': lambda: handle_save_config(data),
            '/api/sync': lambda: handle_start_sync(),
            '/api/reset': lambda: handle_reset_sync(),
            '/api/open-folder': lambda: handle_open_folder(data),
        }

        handler = routes.get(path)
        if handler:
            status_code, response_data = handler()
            self.send_json(status_code, response_data)
        else:
            self.send_error(404, "Not Found")

    def send_json(self, status, data):
        """Helper to send JSON response with correct headers."""
        content = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format, *args):
        """Suppress default per-request HTTP access logging in terminal."""
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
