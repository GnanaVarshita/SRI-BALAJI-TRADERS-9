"""
Dialog API Handlers
Handles native file and folder picker requests.
"""

from common.dialogs import browse_file_dialog, browse_folder_dialog


def handle_browse_file():
    """Handles POST /api/browse-file."""
    try:
        file_path = browse_file_dialog()
        return 200, {"success": True, "filePath": file_path}
    except Exception as e:
        return 500, {"success": False, "message": f"Browse failed: {e}"}


def handle_browse_folder():
    """Handles POST /api/browse-folder."""
    try:
        folder_path = browse_folder_dialog()
        return 200, {"success": True, "folderPath": folder_path}
    except Exception as e:
        return 500, {"success": False, "message": f"Browse folder failed: {e}"}
