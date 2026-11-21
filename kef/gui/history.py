"""[8.] Futtatási előzmények."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from kef.gui import dialog
from kef.gui.base import Screen, Table, parse_hu_date
from kef.gui.result import open_path

STATUS_TEXT = {"sikeres": "sikeres", "hibas": "hibás", "megszakitva": "megszakítva", "fut": "fut"}


class HistoryScreen(Screen):
    number = 8
    title = "Futtatási előzmények"

    def __init__(self, app) -> None:
        super().__init__(app)
        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", side="bottom", pady=(12, 0))
        ttk.Button(bottom, text="Kimeneti mappa megnyitása", command=self.open_dir).pack(side="left")
        ttk.Button(bottom, text="‹ Főmenü", command=lambda: app.show("menu")).pack(side="right")

        flt = ttk.LabelFrame(self.body, text="Szűrés", padding=10)
        flt.pack(fill="x")
        self.from_var, self.to_var, self.user_var = tk.StringVar(), tk.StringVar(), tk.StringVar(value="Mindenki")
        ttk.Label(flt, text="Indítás -tól:").pack(side="left")
        ttk.Entry(flt, textvariable=self.from_var, width=12).pack(side="left", padx=(4, 12))
        ttk.Label(flt, text="-ig:").pack(side="left")
        ttk.Entry(flt, textvariable=self.to_var, width=12).pack(side="left", padx=(4, 12))
        ttk.Label(flt, text="Felhasználó:").pack(side="left")
        self.user_cb = ttk.Combobox(flt, textvariable=self.user_var, state="readonly", width=18)
        self.user_cb.pack(side="left", padx=(4, 12))
        ttk.Button(flt, text="Lista frissítése", command=self.refresh).pack(side="left")
        ttk.Button(flt, text="Szűrők törlése", command=self.clear).pack(side="left", padx=6)

        self.runs = Table(self.body, [
            ("id", "#", 45), ("inditas", "Indítás", 140), ("felh", "Felhasználó", 100),
            ("forras", "Forrásfájl", 210), ("sorok", "Sorok", 75), ("hibas", "Hibás", 60),
            ("szempont", "Szempont", 80), ("fajlok", "Fájlok", 60), ("statusz", "Státusz", 95),
            ("uzenet", "Üzenet", 260)], height=9)
        self.runs.pack(fill="both", expand=True, pady=(10, 0))
        self.runs.tree.bind("<<TreeviewSelect>>", lambda e: self.show_files())
        ttk.Label(self.body, text="A kijelölt futtatás fájljai:", style="Muted.TLabel").pack(anchor="w", pady=(8, 2))
        self.files = Table(self.body, [("kulcs", "Csoport", 90), ("nev", "Név", 180), ("fajl", "Fájlnév", 400),
                                       ("sorok", "Sorok", 80), ("late", "Késésben", 80), ("dij", "Díj (Ft)", 100)],
                           height=6)
        self.files.pack(fill="both", expand=True)
        self._users: dict[str, int] = {}
        self._runs: dict[str, object] = {}

    def on_show(self, **kwargs) -> None:
        self._users = {r["azonosito"]: r["id"] for r in self.app.db.list_users()}
        self.user_cb["values"] = ["Mindenki", *self._users]
        self.refresh()

    def clear(self) -> None:
        self.from_var.set("")
        self.to_var.set("")
        self.user_var.set("Mindenki")
        self.refresh()

    def refresh(self) -> None:
        try:
            tol = parse_hu_date(self.from_var.get(), "-tól")
            ig = parse_hu_date(self.to_var.get(), "-ig")
        except ValueError as exc:
            dialog.error(self, str(exc))
            return
        if tol and ig and tol > ig:
            dialog.error(self, "Érvénytelen időszak: a kezdő dátum későbbi a záró dátumnál.")
            return
        uid = self._users.get(self.user_var.get())
        rows = self.app.db.list_runs(tol.isoformat() if tol else None, ig.isoformat() if ig else None, uid)
        self._runs = {str(r["id"]): r for r in rows}
        self.runs.set_rows(
            [(r["id"], r["inditas"], r["felhasznalo"], Path(r["forrasfajl"]).name, r["sorok_be"], r["hibas_sorok"],
              r["szempont"], r["fajlok"], STATUS_TEXT.get(r["statusz"], r["statusz"]), r["uzenet"] or "")
             for r in rows],
            tags=["" if r["statusz"] == "sikeres" else "bad" for r in rows], iids=[r["id"] for r in rows])
        self.files.set_rows([])
        if rows:
            self.runs.tree.selection_set(str(rows[0]["id"]))

    def show_files(self) -> None:
        iid = self.runs.selected_iid()
        if not iid:
            return
        files = self.app.db.list_output_files(int(iid))
        self.files.set_rows([(f["csoport_kulcs"], f["csoport_nev"], f["fajlnev"], f["sorok_szama"], f["kesesben"],
                              f["osszeg_dij"]) for f in files])

    def open_dir(self) -> None:
        iid = self.runs.selected_iid()
        if not iid:
            dialog.warning(self, "Előbb jelölj ki egy futtatást!")
            return
        try:
            open_path(Path(self._runs[iid]["kimeneti_mappa"]))
        except OSError as exc:
            dialog.error(self, str(exc))
