"""[7.] Eredmény összesítő."""
from __future__ import annotations

import os
from pathlib import Path
from tkinter import ttk

from kef.gui import dialog
from kef.gui.base import Screen, Table, fmt_int


def open_path(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Nem található: {path}")
    os.startfile(str(path))  # type: ignore[attr-defined]  # Windows


class ResultScreen(Screen):
    number = 7
    title = "Eredmény összesítő"

    def __init__(self, app) -> None:
        super().__init__(app)
        self.head = ttk.Label(self.body, text="", style="Big.TLabel")
        self.head.pack(anchor="w")
        stats = ttk.Frame(self.body)
        stats.pack(fill="x", pady=(10, 12))
        self.vals: dict[str, ttk.Label] = {}
        for i, (k, label) in enumerate([("files", "Létrehozott fájlok"), ("rows", "Kiírt sorok"),
                                        ("late", "Késésben"), ("fee", "Késedelmi díj összesen"),
                                        ("errors", "Kihagyott hibás sor"), ("time", "Futásidő")]):
            f = ttk.Frame(stats, padding=(0, 0, 30, 0))
            f.grid(row=0, column=i, sticky="w")
            ttk.Label(f, text=label, style="Muted.TLabel").pack(anchor="w")
            v = ttk.Label(f, text="–", style="Value.TLabel")
            v.pack(anchor="w")
            self.vals[k] = v
        self.table = Table(self.body, [("kulcs", "Csoport", 90), ("nev", "Név", 170), ("fajl", "Fájlnév", 380),
                                       ("sorok", "Sorok", 80), ("late", "Késésben", 80), ("dij", "Díj (Ft)", 100)],
                           height=14)
        self.table.pack(fill="both", expand=True)
        self.table.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.dir_lbl = ttk.Label(self.body, text="", style="Muted.TLabel")
        self.dir_lbl.pack(anchor="w", pady=(6, 0))

        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Button(bottom, text="Mappa megnyitása", command=self.open_dir).pack(side="left")
        ttk.Button(bottom, text="Kijelölt fájl megnyitása", command=self.open_selected).pack(side="left", padx=8)
        ttk.Button(bottom, text="Főmenü", command=lambda: app.show("menu")).pack(side="right")
        ttk.Button(bottom, text="Új feldolgozás", style="Primary.TButton",
                   command=self.new_run).pack(side="right", padx=8)

    def on_show(self, **kwargs) -> None:
        r = self.app.last_run
        cfg = self.app.split_config
        lr = self.app.load_result
        if r is None:
            return
        self.head.configure(text=f"A feldolgozás sikeresen befejeződött (futtatás #{r.run_id})")
        self.vals["files"].configure(text=f"{len(r.files)}" + (" + összesítő" if cfg.osszesito_fajl else ""))
        self.vals["rows"].configure(text=fmt_int(r.total_rows))
        self.vals["late"].configure(text=fmt_int(sum(f.kesesben for f in r.files)))
        self.vals["fee"].configure(text=f"{fmt_int(r.total_fee)} Ft")
        self.vals["errors"].configure(text=fmt_int(lr.error_rows))
        self.vals["time"].configure(text=f"{r.seconds:.1f} mp")
        self.table.set_rows([(f.csoport_kulcs, f.csoport_nev, f.fajlnev, f.sorok_szama, f.kesesben, f.osszeg_dij)
                             for f in r.files], iids=[f.fajlnev for f in r.files])
        self.dir_lbl.configure(text=f"Kimeneti mappa: {cfg.kimeneti_mappa}")

    def open_dir(self) -> None:
        try:
            open_path(Path(self.app.split_config.kimeneti_mappa))
        except OSError as exc:
            dialog.error(self, str(exc))

    def open_selected(self) -> None:
        iid = self.table.selected_iid()
        if not iid:
            dialog.warning(self, "Előbb jelölj ki egy fájlt a listában!")
            return
        try:
            open_path(Path(self.app.split_config.kimeneti_mappa) / iid)
        except OSError as exc:
            dialog.error(self, str(exc))

    def new_run(self) -> None:
        self.app.reset_workflow()
        self.app.show("load")
