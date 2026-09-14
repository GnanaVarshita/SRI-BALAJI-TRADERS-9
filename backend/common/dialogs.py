"""
Thread-Safe GUI Dialog Utilities
Sri Balaji Traders Automation System
"""

import queue
import threading
from pathlib import Path
from config import WORKSPACE_DIR


def browse_file_dialog(initial_dir=None, title="Select Budget Excel File", filetypes=None):
    """
    Opens a thread-safe native Tkinter file picker dialog.
    """
    if initial_dir is None:
        initial_dir = str(WORKSPACE_DIR)
    if filetypes is None:
        filetypes = [
            ("Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb;*.csv"),
            ("All files", "*.*")
        ]

    q = queue.Queue()

    def _run():
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askopenfilename(
                initialdir=initial_dir,
                title=title,
                filetypes=filetypes
            )
            root.destroy()
            q.put(path)
        except Exception as e:
            q.put(e)

    t = threading.Thread(target=_run)
    t.start()
    t.join()
    res = q.get()
    if isinstance(res, Exception):
        raise res
    return res


def browse_folder_dialog(initial_dir=None, title="Select Save Folder"):
    """
    Opens a thread-safe native Tkinter directory picker dialog.
    """
    if initial_dir is None:
        initial_dir = str(WORKSPACE_DIR)

    q = queue.Queue()

    def _run():
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askdirectory(
                initialdir=initial_dir,
                title=title
            )
            root.destroy()
            q.put(path)
        except Exception as e:
            q.put(e)

    t = threading.Thread(target=_run)
    t.start()
    t.join()
    res = q.get()
    if isinstance(res, Exception):
        raise res
    return res
