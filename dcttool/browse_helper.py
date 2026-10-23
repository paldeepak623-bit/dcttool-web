# -*- coding: utf-8 -*-
"""
Standalone file picker helper that executes safely on its own process main thread.
Avoids RuntimeError: main thread is not in main loop in multi-threaded web servers.
"""
import sys
import tkinter as tk
from tkinter import filedialog


def main():
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select PACS Master Excel Sheet",
            filetypes=[("Excel Workbooks", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )
        root.destroy()
        if path:
            sys.stdout.write(path)
            sys.stdout.flush()
    except Exception as ex:
        sys.stderr.write(str(ex))
        sys.stderr.flush()


if __name__ == "__main__":
    main()
