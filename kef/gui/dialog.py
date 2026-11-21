"""[10.] Megerősítés / üzenet párbeszédablak – minden képernyő ezt használja."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from kef.gui.base import ERR_RED, OK_GREEN, PRIMARY

KINDS = {
    "confirm": ("?", PRIMARY, "Megerősítés"),
    "info": ("i", OK_GREEN, "Információ"),
    "warning": ("!", "#B26A00", "Figyelmeztetés"),
    "error": ("×", ERR_RED, "Hiba"),
}


class MessageDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, message: str, kind: str = "info", title: str | None = None,
                 ok_text: str = "OK", cancel_text: str = "Mégse") -> None:
        super().__init__(parent)
        icon, color, default_title = KINDS[kind]
        self.result = False
        self.title(f"[10.] {title or default_title}")
        self.configure(bg="white")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())

        top = tk.Frame(self, bg="white", padx=20, pady=18)
        top.pack(fill="both", expand=True)
        tk.Label(top, text=icon, bg=color, fg="white", font=("Segoe UI Semibold", 18), width=2).pack(
            side="left", anchor="n", padx=(0, 16))
        tk.Label(top, text=message, bg="white", justify="left", wraplength=420,
                 font=("Segoe UI", 10)).pack(side="left", fill="both")

        bar = tk.Frame(self, bg="#EEF1F5", padx=12, pady=10)
        bar.pack(fill="x")
        ok = ttk.Button(bar, text=ok_text, command=self._ok, style="Primary.TButton")
        if kind == "confirm":
            ttk.Button(bar, text=cancel_text, command=self._cancel).pack(side="right")
            ok.pack(side="right", padx=(0, 8))
        else:
            ok.pack(side="right")
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.update_idletasks()
        p = parent.winfo_toplevel()
        x = p.winfo_rootx() + (p.winfo_width() - self.winfo_width()) // 2
        y = p.winfo_rooty() + (p.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        ok.focus_set()

    def _ok(self) -> None:
        self.result = True
        self.destroy()

    def _cancel(self) -> None:
        self.result = False
        self.destroy()

    def show(self) -> bool:
        self.grab_set()
        self.wait_window()
        return self.result


def confirm(parent, message: str, title: str | None = None, ok_text: str = "Igen") -> bool:
    return MessageDialog(parent, message, "confirm", title, ok_text=ok_text).show()


def info(parent, message: str, title: str | None = None) -> None:
    MessageDialog(parent, message, "info", title).show()


def warning(parent, message: str, title: str | None = None) -> None:
    MessageDialog(parent, message, "warning", title).show()


def error(parent, message: str, title: str | None = None) -> None:
    MessageDialog(parent, message, "error", title).show()
