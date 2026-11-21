"""[3.] Adatbetöltés és validáció."""
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from kef import config
from kef.auth import PermissionDenied, require
from kef.gui import dialog
from kef.gui.base import Screen, Table, fmt_int
from kef.pipeline import ValidationError, load_file

PREVIEW_COLS = ["kolcsonzes_id", "olvaso_id", "olvaso_tipus", "cim", "kategoria", "fiok_kod",
                "kolcsonzes_datum", "lejarat_datum", "visszahozas_datum", "hosszabbitasok_szama"]


class LoadScreen(Screen):
    number = 3
    title = "Adatbetöltés és validáció"

    def __init__(self, app) -> None:
        super().__init__(app)
        top = ttk.LabelFrame(self.body, text="Forrásfájl (a könyvtári rendszer exportja: .xlsx vagy .csv)", padding=10)
        top.pack(fill="x")
        self.path_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.path_var).pack(side="left", fill="x", expand=True)
        ttk.Button(top, text="Tallózás…", command=self.browse).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Mintaadat", command=self.use_sample).pack(side="left", padx=(8, 0))
        self.load_btn = ttk.Button(top, text="Betöltés", style="Primary.TButton", command=self.load)
        self.load_btn.pack(side="left", padx=(8, 0))

        self.pbar = ttk.Progressbar(self.body, mode="indeterminate")
        self.status = ttk.Label(self.body, text="Válaszd ki a fájlt, majd kattints a Betöltés gombra.",
                                style="Muted.TLabel")
        self.status.pack(anchor="w", pady=(10, 4))

        stats = ttk.Frame(self.body)
        stats.pack(fill="x", pady=(0, 8))
        self.stat_vals: dict[str, ttk.Label] = {}
        for i, (key, label) in enumerate([("total", "Összes sor"), ("valid", "Érvényes"), ("errors", "Hibás"),
                                          ("branches", "Fiókok"), ("period", "Időszak")]):
            f = ttk.Frame(stats, padding=(0, 0, 28, 0))
            f.grid(row=0, column=i, sticky="w")
            ttk.Label(f, text=label, style="Muted.TLabel").pack(anchor="w")
            v = ttk.Label(f, text="–", style="Value.TLabel")
            v.pack(anchor="w")
            self.stat_vals[key] = v

        self.nb = ttk.Notebook(self.body)
        self.nb.pack(fill="both", expand=True)
        self.preview = Table(self.nb, [(c, config.DISPLAY_NAMES[c], 140 if c == "cim" else 105) for c in PREVIEW_COLS],
                             height=12)
        self.err_table = Table(self.nb, [("sor", "Sor", 60), ("kolcsonzes_id", "Kölcsönzés azonosító", 150),
                                         ("hiba", "Hiba", 620)], height=12)
        self.nb.add(self.preview, text="Előnézet (első 100 érvényes sor)")
        self.nb.add(self.err_table, text="Hibás sorok")

        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Button(bottom, text="‹ Főmenü", command=lambda: app.show("menu")).pack(side="left")
        self.next_btn = ttk.Button(bottom, text="Tovább: formázási szabályok ›", style="Primary.TButton",
                                   command=lambda: app.show("rules", wizard=True))
        self.next_btn.pack(side="right")

    def on_show(self, **kwargs) -> None:
        if self.app.load_result is None:
            self._clear()

    def _clear(self) -> None:
        for v in self.stat_vals.values():
            v.configure(text="–")
        self.preview.set_rows([])
        self.err_table.set_rows([])
        self.nb.tab(1, text="Hibás sorok")
        self.next_btn.state(["disabled"])
        self.status.configure(text="Válaszd ki a fájlt, majd kattints a Betöltés gombra.", style="Muted.TLabel")

    def browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Kölcsönzési export kiválasztása", initialdir=str(config.SAMPLE_DIR),
            filetypes=[("Excel és CSV", "*.xlsx *.xlsm *.csv"), ("Minden fájl", "*.*")])
        if path:
            self.path_var.set(path)

    def use_sample(self) -> None:
        self.path_var.set(str(config.SAMPLE_DIR / "kolcsonzesek_export.xlsx"))

    def load(self) -> None:
        try:
            require(self.app.user, "feldolgozas")
        except PermissionDenied as exc:
            dialog.error(self, str(exc))
            return
        path = self.path_var.get().strip()
        if not path:
            dialog.warning(self, "Előbb válassz ki egy forrásfájlt!")
            return
        self._clear()
        self.load_btn.state(["disabled"])
        self.status.configure(text=f"Betöltés folyamatban: {Path(path).name} …", style="Muted.TLabel")
        self.pbar.pack(fill="x", before=self.status)
        self.pbar.start(12)
        self.app.busy = True
        self.app.run_in_thread(lambda: load_file(path), self._loaded, self._failed)

    def _done_common(self) -> None:
        self.app.busy = False
        self.pbar.stop()
        self.pbar.pack_forget()
        self.load_btn.state(["!disabled"])

    def _failed(self, exc: Exception) -> None:
        self._done_common()
        self.status.configure(text="A betöltés sikertelen.", style="Err.TLabel")
        msg = str(exc) if isinstance(exc, ValidationError) else f"Váratlan hiba: {exc}"
        dialog.error(self, msg, "Sikertelen betöltés")

    def _loaded(self, result) -> None:
        self._done_common()
        self.app.load_result = result
        self.app.transform_settings = None
        self.app.split_config = None
        self.show_result(result)

    def show_result(self, r) -> None:
        v = r.valid
        self.stat_vals["total"].configure(text=fmt_int(r.total_rows))
        self.stat_vals["valid"].configure(text=fmt_int(r.valid_rows))
        self.stat_vals["errors"].configure(text=fmt_int(r.error_rows))
        self.stat_vals["branches"].configure(text=str(v["fiok_kod"].nunique()) if len(v) else "0")
        if len(v):
            self.stat_vals["period"].configure(
                text=f"{v['kolcsonzes_datum'].min():%Y.%m.%d.} – {v['kolcsonzes_datum'].max():%Y.%m.%d.}")
        head = v.head(100)
        self.preview.set_rows(head[PREVIEW_COLS].itertuples(index=False, name=None))
        self.err_table.set_rows(r.errors.itertuples(index=False, name=None), tags=["bad"] * len(r.errors))
        self.nb.tab(1, text=f"Hibás sorok ({fmt_int(r.error_rows)})")
        if r.valid_rows == 0:
            self.status.configure(text="A fájlban nincs érvényes sor – a feldolgozás nem folytatható.", style="Err.TLabel")
            self.next_btn.state(["disabled"])
            return
        extra = f"  Figyelmen kívül hagyott plusz oszlop(ok): {', '.join(r.extra_columns)}." if r.extra_columns else ""
        if r.error_rows:
            self.status.configure(text=(f"Betöltve: {r.path.name}. {fmt_int(r.error_rows)} hibás sor kimarad a "
                                        f"feldolgozásból (az összesítő fájl „Hibás sorok” lapjára kerül).{extra}"),
                                  style="Err.TLabel")
        else:
            self.status.configure(text=f"Betöltve: {r.path.name}. Minden sor érvényes.{extra}", style="Ok.TLabel")
        self.next_btn.state(["!disabled"])
