"""Közös GUI-elemek: képernyő-alaposztály, táblázat, dátumkezelés, stílus."""
from __future__ import annotations

import tkinter as tk
from datetime import date, datetime
from tkinter import ttk
from typing import TYPE_CHECKING, Iterable, Sequence

if TYPE_CHECKING:
    from kef.gui.app import App

PRIMARY = "#1F4E78"
ACCENT = "#2E75B6"
BG = "#F4F6F9"
MUTED = "#5F6B7A"
OK_GREEN = "#2E7D32"
ERR_RED = "#C62828"


def setup_style(root: tk.Tk) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except tk.TclError:
        style.theme_use("clam")
    root.configure(bg=BG)
    default = ("Segoe UI", 10)
    root.option_add("*Font", default)
    style.configure(".", font=default)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG)
    style.configure("TCheckbutton", background=BG)
    style.configure("TRadiobutton", background=BG)
    style.configure("TLabelframe", background=BG)
    style.configure("TLabelframe.Label", background=BG, foreground=PRIMARY, font=("Segoe UI Semibold", 10))
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("Big.TLabel", font=("Segoe UI Semibold", 14), foreground=PRIMARY)
    style.configure("Value.TLabel", font=("Segoe UI Semibold", 11))
    style.configure("Ok.TLabel", foreground=OK_GREEN, font=("Segoe UI Semibold", 10))
    style.configure("Err.TLabel", foreground=ERR_RED, font=("Segoe UI Semibold", 10))
    style.configure("Primary.TButton", font=("Segoe UI Semibold", 10))
    style.configure("Menu.TButton", font=("Segoe UI", 11), padding=(10, 12))
    style.configure("Treeview", rowheight=22)
    style.configure("Treeview.Heading", font=("Segoe UI Semibold", 9))


class Screen(ttk.Frame):
    """Egy „ablak” a dokumentációban: sorszám + cím + fejléc a bejelentkezett felhasználóval."""

    number: int = 0
    title: str = ""

    def __init__(self, app: "App") -> None:
        super().__init__(app.container)
        self.app = app
        header = tk.Frame(self, bg=PRIMARY)
        header.pack(fill="x")
        self._title_lbl = tk.Label(header, text=f"[{self.number}.]  {self.title}", bg=PRIMARY, fg="white",
                                   font=("Segoe UI Semibold", 13), padx=14, pady=10)
        self._title_lbl.pack(side="left")
        self._user_lbl = tk.Label(header, text="", bg=PRIMARY, fg="#D6E4F0", font=("Segoe UI", 9), padx=14)
        self._user_lbl.pack(side="right")
        self.body = ttk.Frame(self, padding=(18, 14))
        self.body.pack(fill="both", expand=True)

    def refresh_header(self) -> None:
        u = self.app.user
        self._user_lbl.configure(text=f"Bejelentkezve: {u.nev} ({u.azonosito}) – {u.role_name}" if u else "")

    def on_show(self, **kwargs) -> None:  # felülírható
        pass


class Table(ttk.Frame):
    """Treeview görgetősávokkal."""

    def __init__(self, parent, columns: Sequence[tuple[str, str, int]], height: int = 10, selectmode="browse"):
        super().__init__(parent)
        self.keys = [c[0] for c in columns]
        self.tree = ttk.Treeview(self, columns=self.keys, show="headings", height=height, selectmode=selectmode)
        for key, text, width in columns:
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, minwidth=40, stretch=width >= 150,
                             anchor="e" if text.endswith(("(Ft)", "sor", "Sorok", "Késésben", "Fájlok")) else "w")
        ys = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure("late", background="#FBE3D6")
        self.tree.tag_configure("bad", foreground=ERR_RED)

    def set_rows(self, rows: Iterable[Sequence], tags: Iterable[str] | None = None, iids: Iterable | None = None):
        self.tree.delete(*self.tree.get_children())
        tags = list(tags) if tags is not None else None
        iids = list(iids) if iids is not None else None
        for i, row in enumerate(rows):
            kw = {}
            if tags is not None and tags[i]:
                kw["tags"] = (tags[i],)
            if iids is not None:
                kw["iid"] = str(iids[i])
            self.tree.insert("", "end", values=[_fmt(v) for v in row], **kw)

    def selected_iid(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None


def _fmt(v) -> str:
    if v is None:
        return ""
    try:
        import pandas as pd
        if v is pd.NaT or (isinstance(v, float) and v != v):
            return ""
        if isinstance(v, pd.Timestamp):
            return v.strftime("%Y.%m.%d.")
    except ImportError:  # pragma: no cover
        pass
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y.%m.%d.")
    if isinstance(v, int) and not isinstance(v, bool) and abs(v) >= 1000:
        return f"{v:,}".replace(",", " ")
    return str(v)


def fmt_int(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def parse_hu_date(text: str, field: str) -> date | None:
    """'2025.03.01.', '2025.03.01' vagy '2025-03-01' -> date; üres -> None; hibás -> ValueError."""
    t = (text or "").strip().rstrip(".")
    if not t:
        return None
    for f in ("%Y.%m.%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, f).date()
        except ValueError:
            pass
    raise ValueError(f"Érvénytelen dátum a(z) „{field}” mezőben (formátum: ÉÉÉÉ.HH.NN.).")


def hu_date(d: date | None) -> str:
    return d.strftime("%Y.%m.%d.") if d else ""


def button_row(parent, *buttons: tuple[str, object, str | None]) -> ttk.Frame:
    row = ttk.Frame(parent)
    for text, cmd, style in buttons:
        ttk.Button(row, text=text, command=cmd, style=style or "TButton").pack(side="left", padx=(0, 8))
    return row
