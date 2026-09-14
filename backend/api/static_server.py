"""
Static File Serving & Downloads
Handles React production bundle serving, MIME types, and secure file downloads.
"""

import os
import re
import datetime
import urllib.parse
from pathlib import Path
from config import WORKSPACE_DIR

DIST_DIR = WORKSPACE_DIR / "frontend/dist"
ALLOWED_DOWNLOAD_DIRS = ["Corteva POs", "New Gen POs", "FMC POs", "CORTEVA", "FMC", "NEW GEN"]


def get_content_type(file_path):
    """Returns the MIME content type based on file extension."""
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


def serve_static(req_handler, path):
    """Serves compiled frontend static files or fallback index.html for React Router."""
    clean_path = path.lstrip('/')
    if not clean_path:
        clean_path = "index.html"

    file_path = DIST_DIR / clean_path

    if not file_path.exists() or file_path.is_dir():
        file_path = DIST_DIR / "index.html"

    if not file_path.exists():
        req_handler.send_response(200)
        req_handler.send_header('Content-Type', 'text/html; charset=utf-8')
        req_handler.end_headers()
        req_handler.wfile.write(b"<h1>Sri Balaji Traders Dashboard</h1><p>Frontend is not compiled yet. Please run <code>npm run build</code> in the frontend folder, or run <code>run_dashboard.ps1</code>.</p>")
        return

    try:
        with open(file_path, 'rb') as f:
            content = f.read()
        req_handler.send_response(200)
        req_handler.send_header('Content-Type', get_content_type(file_path))
        req_handler.send_header('Content-Length', str(len(content)))
        req_handler.end_headers()
        req_handler.wfile.write(content)
    except Exception as e:
        req_handler.send_error(500, f"Internal server error: {e}")


def scan_downloaded_files():
    """Scans all PO download directories and returns list of file info objects."""
    files = []
    for t_dir in ALLOWED_DOWNLOAD_DIRS:
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


def handle_view_file(req_handler, parsed_url):
    """Streams a requested PO file with security path traversal checks."""
    parsed_query = urllib.parse.parse_qs(parsed_url.query)
    file_rel_path = parsed_query.get('path', [None])[0]
    if not file_rel_path:
        req_handler.send_error(400, "Bad Request: Missing path")
        return

    file_abs_path = (WORKSPACE_DIR / file_rel_path).resolve()
    if not str(file_abs_path).startswith(str(WORKSPACE_DIR.resolve())) or not file_abs_path.is_file():
        req_handler.send_error(403, "Forbidden: Invalid file path")
        return

    allowed = False
    for allowed_dir in ALLOWED_DOWNLOAD_DIRS:
        allowed_abs = (WORKSPACE_DIR / allowed_dir).resolve()
        if str(file_abs_path).startswith(str(allowed_abs)):
            allowed = True
            break

    if not allowed:
        req_handler.send_error(403, "Forbidden: Access to this directory is not allowed")
        return

    try:
        with open(file_abs_path, 'rb') as f:
            content = f.read()
        req_handler.send_response(200)
        req_handler.send_header('Content-Type', get_content_type(file_abs_path))
        req_handler.send_header('Content-Length', str(len(content)))
        req_handler.send_header('Content-Disposition', f'attachment; filename="{file_abs_path.name}"')
        req_handler.end_headers()
        req_handler.wfile.write(content)
    except Exception as e:
        req_handler.send_error(500, f"Internal server error: {e}")
