"""[5.] Szétbontás beállításai (a valós megoldásban: Power Automate + Office Script)."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from kef import config
from kef.gui import dialog
from kef.gui.base import Screen, Table, fmt_int, hu_date, parse_hu_date
from kef.pipeline import OutputError, SplitConfig
from kef.pipeline.splitter import build_filename, ensure_writable, filter_period, group_counts, period_label


class SplitScreen(Screen):
    number = 5
    title = "Szétbontás beállításai"

    def __init__(self, app) -> None:
        super().__init__(app)
        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", side="bottom", pady=(12, 0))
        ttk.Button(bottom, text="‹ Vissza", command=lambda: app.show("rules", wizard=True)).pack(side="left")
        ttk.Button(bottom, text="Feldolgozás indítása", style="Primary.TButton", command=self.start).pack(side="right")

        left = ttk.Frame(self.body)
        left.pack(side="left", fill="y", padx=(0, 16))

        key_box = ttk.LabelFrame(left, text="Szétbontás szempontja", padding=10)
        key_box.pack(fill="x")
        self.key_var = tk.StringVar(value="fiok")
        for key, (_, _, label) in config.SPLIT_KEYS.items():
            ttk.Radiobutton(key_box, text=label, value=key, variable=self.key_var,
                            command=self.update_preview).pack(anchor="w", pady=2)

        per = ttk.LabelFrame(left, text="Időszak (kölcsönzés dátuma, nem kötelező)", padding=10)
        per.pack(fill="x", pady=(12, 0))
        self.from_var, self.to_var = tk.StringVar(), tk.StringVar()
        ttk.Label(per, text="-tól:").grid(row=0, column=0, sticky="w")
        ttk.Entry(per, textvariable=self.from_var, width=13).grid(row=0, column=1, padx=6, pady=3)
        ttk.Label(per, text="-ig:").grid(row=1, column=0, sticky="w")
        ttk.Entry(per, textvariable=self.to_var, width=13).grid(row=1, column=1, padx=6, pady=3)
        ttk.Label(per, text="ÉÉÉÉ.HH.NN.", style="Muted.TLabel").grid(row=0, column=2, rowspan=2)

        out = ttk.LabelFrame(left, text="Kimenet", padding=10)
        out.pack(fill="x", pady=(12, 0))
        ttk.Label(out, text="Kimeneti mappa:").grid(row=0, column=0, sticky="w")
        self.dir_var = tk.StringVar()
        ttk.Entry(out, textvariable=self.dir_var, width=38).grid(row=1, column=0, sticky="ew")
        ttk.Button(out, text="…", width=3, command=self.browse).grid(row=1, column=1, padx=(4, 0))
        ttk.Label(out, text="Fájlnév-minta:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.pattern_var = tk.StringVar()
        ttk.Entry(out, textvariable=self.pattern_var, width=38).grid(row=3, column=0, columnspan=2, sticky="ew")
        ttk.Label(out, text="Helyőrzők: {idoszak} {kulcs} {nev} {datum}", style="Muted.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="w")
        self.summary_var = tk.BooleanVar(value=True)
        self.errors_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(out, text="Összesítő fájl készítése", variable=self.summary_var).grid(
            row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(out, text="Hibás sorok lap az összesítőben", variable=self.errors_var).grid(
            row=6, column=0, sticky="w")

        right = ttk.LabelFrame(self.body, text="Előnézet – létrejövő fájlok", padding=10)
        right.pack(side="left", fill="both", expand=True)
        self.preview_lbl = ttk.Label(right, text="", style="Value.TLabel")
        self.preview_lbl.pack(anchor="w")
        self.example_lbl = ttk.Label(right, text="", style="Muted.TLabel")
        self.example_lbl.pack(anchor="w", pady=(2, 8))
        self.table = Table(right, [("kulcs", "Csoport", 150), ("fajl", "Fájlnév", 420), ("sorok", "Sorok", 80)],
                           height=14)
        self.table.pack(fill="both", expand=True)
        ttk.Button(right, text="Előnézet frissítése", command=self.update_preview).pack(anchor="e", pady=(8, 0))

    def on_show(self, **kwargs) -> None:
        cfg = self.app.split_config
        if cfg is None:
            s = self.app.db.get_settings()
            cfg = SplitConfig(kimeneti_mappa=Path(s["kimeneti_mappa"]), fajlnev_minta=s["fajlnev_minta"])
        self.key_var.set(cfg.szempont)
        self.from_var.set(hu_date(cfg.idoszak_tol))
        self.to_var.set(hu_date(cfg.idoszak_ig))
        self.dir_var.set(str(cfg.kimeneti_mappa))
        self.pattern_var.set(cfg.fajlnev_minta)
        self.summary_var.set(cfg.osszesito_fajl)
        self.errors_var.set(cfg.hibas_sorok_fajl)
        self.update_preview()

    def browse(self) -> None:
        d = filedialog.askdirectory(title="Kimeneti mappa", initialdir=self.dir_var.get() or str(config.BASE_DIR))
        if d:
            self.dir_var.set(d)

    def read(self) -> SplitConfig:
        cfg = SplitConfig(
            szempont=self.key_var.get(),
            kimeneti_mappa=Path(self.dir_var.get().strip() or config.DEFAULT_OUTPUT_DIR),
            fajlnev_minta=self.pattern_var.get().strip(),
            idoszak_tol=parse_hu_date(self.from_var.get(), "-tól"),
            idoszak_ig=parse_hu_date(self.to_var.get(), "-ig"),
            hibas_sorok_fajl=self.errors_var.get(),
            osszesito_fajl=self.summary_var.get(),
        )
        cfg.validate()
        return cfg

    def update_preview(self) -> None:
        lr = self.app.load_result
        if lr is None:
            return
        try:
            cfg = self.read()
        except (ValueError, OutputError) as exc:
            self.preview_lbl.configure(text=str(exc), style="Err.TLabel")
            self.table.set_rows([])
            return
        counts = group_counts(lr.valid, cfg)
        key_col, name_col, _ = config.SPLIT_KEYS[cfg.szempont]
        data = filter_period(lr.valid, cfg.idoszak_tol, cfg.idoszak_ig)
        names = data.drop_duplicates(key_col).set_index(key_col)[name_col] if name_col else None
        idoszak = period_label(data, cfg) if len(data) else "ures"
        rows = [(k, build_filename(cfg, idoszak, str(k), str(names[k]) if names is not None else ""), int(n))
                for k, n in counts.head(300).items()]
        self.table.set_rows(rows)
        extra = " + 1 összesítő" if cfg.osszesito_fajl else ""
        self.preview_lbl.configure(
            text=f"{fmt_int(len(counts))} fájl{extra}, összesen {fmt_int(int(counts.sum()))} sor", style="Value.TLabel")
        self.example_lbl.configure(text=f"Mappa: {cfg.kimeneti_mappa}" +
                                   ("   (az előnézet az első 300 csoportot mutatja)" if len(counts) > 300 else ""))

    def start(self) -> None:
        lr = self.app.load_result
        try:
            cfg = self.read()
            ensure_writable(Path(cfg.kimeneti_mappa))
        except (ValueError, OutputError) as exc:
            dialog.error(self, str(exc), "Hibás beállítás")
            return
        n = len(group_counts(lr.valid, cfg))
        if n == 0:
            dialog.error(self, "A megadott időszakban nincs kölcsönzés.", "Hibás beállítás")
            return
        if n > config.MAX_OUTPUT_FILES_WARNING and not dialog.confirm(
                self, f"A választott szempont szerint {fmt_int(n)} fájl jönne létre, ami sok lehet. Folytatod?",
                ok_text="Folytatás"):
            return
        existing = [p for p in Path(cfg.kimeneti_mappa).glob("*.xlsx")]
        msg = (f"Biztosan elindítod a feldolgozást?\n\n{fmt_int(n)} fájl"
               f"{' + összesítő' if cfg.osszesito_fajl else ''} jön létre ide:\n{cfg.kimeneti_mappa}")
        if existing:
            msg += f"\n\nA mappában már van {len(existing)} .xlsx fájl; az azonos nevűek felülíródnak."
        if not dialog.confirm(self, msg, ok_text="Indítás"):
            return
        self.app.split_config = cfg
        if self.app.user.can("beallitas_modositas"):
            self.app.db.save_settings({"kimeneti_mappa": str(cfg.kimeneti_mappa),
                                       "fajlnev_minta": cfg.fajlnev_minta}, self.app.user.id)
        self.app.show("progress", start=True)
