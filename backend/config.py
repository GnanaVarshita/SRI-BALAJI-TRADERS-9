import os
from pathlib import Path

# Workspace Root Paths
WORKSPACE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = WORKSPACE_DIR / ".env"
PROCESSED_DB_PATH = WORKSPACE_DIR / "processed_emails.json"

# Fallback defaults
DEFAULT_USER_EMAIL = "sribalajitraderspdtr@gmail.com"
