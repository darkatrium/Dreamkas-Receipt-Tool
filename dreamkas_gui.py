#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple GUI launcher for Dreamkas Receipt Tool.

This GUI does not replace the fiscalization logic. It launches the proven
console script in a separate terminal, opens Excel, and opens work folders.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


APP_DIR = Path(__file__).resolve().parent
SCRIPT = APP_DIR / "dreamkas_receipt.py"
DEFAULT_EXCEL = APP_DIR / "dreamkas_receipt_template.xlsx"


def find_python() -> str:
    candidates = [
        r"C:\Python314\python.exe",
        sys.executable,
        "py",
        "python",
    ]
    for cmd in candidates:
        try:
            if Path(cmd).exists():
                return cmd
            subprocess.run([cmd, "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
            return cmd
        except Exception:
            continue
    return "python"


def open_path(path: Path) -> None:
    path = path.resolve()
    if not path.exists():
        messagebox.showwarning("Dreamkas", f"Path not found:\n{path}")
        return
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Dreamkas Receipt Tool")
        self.geometry("720x360")
        self.resizable(False, False)

        self.excel_var = tk.StringVar(value=str(DEFAULT_EXCEL))
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        title = tk.Label(self, text="Dreamkas Receipt Tool v6.1", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", padx=14, pady=(14, 4))

        frame = tk.Frame(self)
        frame.pack(fill="x", **pad)

        tk.Label(frame, text="Excel file:").grid(row=0, column=0, sticky="w")
        entry = tk.Entry(frame, textvariable=self.excel_var, width=76)
        entry.grid(row=1, column=0, sticky="we", padx=(0, 8))
        tk.Button(frame, text="Choose...", command=self.choose_excel).grid(row=1, column=1)

        btns = tk.Frame(self)
        btns.pack(fill="x", **pad)

        tk.Button(btns, text="Open Excel", width=18, command=self.open_excel).grid(row=0, column=0, padx=4, pady=4)
        tk.Button(btns, text="Start console tool", width=18, command=self.start_console).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(btns, text="Open receipts", width=18, command=lambda: open_path(APP_DIR / "receipts_pdf")).grid(row=0, column=2, padx=4, pady=4)
        tk.Button(btns, text="Open DB folder", width=18, command=lambda: open_path(APP_DIR / "db")).grid(row=1, column=0, padx=4, pady=4)
        tk.Button(btns, text="Open logs", width=18, command=lambda: open_path(APP_DIR / "logs")).grid(row=1, column=1, padx=4, pady=4)
        tk.Button(btns, text="Open settings", width=18, command=lambda: open_path(APP_DIR / "settings.txt")).grid(row=1, column=2, padx=4, pady=4)

        info = (
            "The console window remains the main safe workflow: shop/device selection, "
            "SALE/REFUND mode, precheck confirmation, SQLite journal, PDF/TXT/QR output."
        )
        tk.Label(self, text=info, wraplength=680, justify="left").pack(anchor="w", padx=14, pady=(10, 4))

        status = tk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken")
        status.pack(side="bottom", fill="x")

    def choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Excel receipt template",
            initialdir=str(APP_DIR),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.excel_var.set(path)

    def open_excel(self) -> None:
        open_path(Path(self.excel_var.get()))

    def start_console(self) -> None:
        if not SCRIPT.exists():
            messagebox.showerror("Dreamkas", "dreamkas_receipt.py not found")
            return
        excel = Path(self.excel_var.get())
        py = find_python()

        if os.name == "nt":
            cmd = f'cd /d "{APP_DIR}" && "{py}" "dreamkas_receipt.py" --excel "{excel}" && pause'
            subprocess.Popen(["cmd.exe", "/k", cmd])
        else:
            subprocess.Popen([py, str(SCRIPT), "--excel", str(excel)], cwd=str(APP_DIR))

        self.status_var.set("Console tool started")


if __name__ == "__main__":
    App().mainloop()
