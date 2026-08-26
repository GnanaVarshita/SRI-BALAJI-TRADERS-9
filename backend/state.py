import threading
from collections import deque

# Global logs buffers and syncing states
LOGS_LOCK = threading.Lock()
LOGS_BUFFER = deque(maxlen=200)

is_syncing = False
sync_thread = None

def add_log_message(msg):
    """
    Appends a timestamped log line to the global logs buffer in a thread-safe way.
    """
    import datetime
    timestamp = datetime.datetime.now().strftime('%H:%M:%S')
    with LOGS_LOCK:
        LOGS_BUFFER.append(f"[{timestamp}] {msg}")
    print(f"[{timestamp}] {msg}")
