"""
System API Routes
Handles configuration, server logs, sync status, history reset, and folder openers.
"""

import os
import json
import threading
from pathlib import Path
from config import WORKSPACE_DIR, ENV_PATH, PROCESSED_DB_PATH
import state
import download_attachments


class WebLogRedirector:
    """Redirects stdout and stderr to in-memory ring buffer for web UI log streaming."""
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


def run_sync_task():
    """Background task to run Gmail PO attachment downloader."""
    try:
        download_attachments.main()
    except Exception as e:
        print(f"Sync thread crashed: {e}")
    finally:
        state.is_syncing = False


def handle_get_config():
    """Handles GET /api/config."""
    env_vars = download_attachments.load_env(ENV_PATH)
    email_val = env_vars.get("GMAIL_EMAIL", "")
    has_pass = "GMAIL_APP_PASSWORD" in env_vars and len(env_vars["GMAIL_APP_PASSWORD"]) > 0
    return 200, {
        "email": email_val,
        "hasPassword": has_pass
    }


def handle_save_config(data):
    """Handles POST /api/config."""
    email_val = data.get("email", "").strip()
    pass_val = data.get("password", "").strip()

    if not email_val:
        return 400, {"success": False, "message": "Email is required"}

    try:
        env_vars = download_attachments.load_env(ENV_PATH)
        final_pass = pass_val if pass_val else env_vars.get("GMAIL_APP_PASSWORD", "")

        with open(ENV_PATH, "w", encoding="utf-8") as f:
            f.write(f"GMAIL_EMAIL={email_val}\n")
            f.write(f"GMAIL_APP_PASSWORD={final_pass}\n")
        return 200, {"success": True, "message": "Credentials saved"}
    except Exception as e:
        return 500, {"success": False, "message": f"Failed to save: {e}"}


def handle_get_status():
    """Handles GET /api/status."""
    with state.LOGS_LOCK:
        current_logs = list(state.LOGS_BUFFER)

    downloads_count = 0
    if PROCESSED_DB_PATH.exists():
        try:
            with open(PROCESSED_DB_PATH, "r", encoding="utf-8") as f:
                db = json.load(f)
                downloads_count = sum(len(uids) for uids in db.values())
        except Exception:
            pass

    return 200, {
        "isSyncing": state.is_syncing,
        "logs": current_logs,
        "totalSynced": downloads_count
    }


def handle_start_sync():
    """Handles POST /api/sync."""
    if state.is_syncing:
        return 400, {"success": False, "message": "Sync is already running"}

    with state.LOGS_LOCK:
        state.LOGS_BUFFER.clear()
        state.LOGS_BUFFER.append("Initializing sync request...")

    state.is_syncing = True
    state.sync_thread = threading.Thread(target=run_sync_task, daemon=True)
    state.sync_thread.start()
    return 200, {"success": True, "message": "Sync started"}


def handle_reset_sync():
    """Handles POST /api/reset."""
    if state.is_syncing:
        return 400, {"success": False, "message": "Cannot reset history while sync is running"}

    if not PROCESSED_DB_PATH.exists():
        return 200, {"success": True, "message": "Sync history is already empty"}

    try:
        os.remove(PROCESSED_DB_PATH)
        with state.LOGS_LOCK:
            state.LOGS_BUFFER.append(">>> Sync history reset successfully. Next run will process all emails.")
        return 200, {"success": True, "message": "Sync history cleared"}
    except Exception as e:
        return 500, {"success": False, "message": f"Failed to reset: {e}"}


def handle_open_folder(data):
    """Handles POST /api/open-folder."""
    folder_type = data.get("folder", "")
    if folder_type == "corteva":
        folder_path = WORKSPACE_DIR / "Corteva POs"
    elif folder_type == "newgen":
        folder_path = WORKSPACE_DIR / "New Gen POs"
    elif folder_type == "fmc":
        folder_path = WORKSPACE_DIR / "FMC POs"
    else:
        return 400, {"success": False, "message": "Invalid folder type"}

    try:
        os.makedirs(folder_path, exist_ok=True)
        os.startfile(folder_path)
        return 200, {"success": True, "message": f"Opened folder: {folder_path.name}"}
    except Exception as e:
        return 500, {"success": False, "message": f"Failed to open folder: {e}"}
