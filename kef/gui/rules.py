"""[4.] Formázási szabályok – számított oszlopok és késedelmi díj (a „VBA-makró” beállításai)."""
from __future__ import annotations

import tkinter as tk
from datetime import date
from tkinter import ttk

from kef import config
from kef.auth import PermissionDenied, require
from kef.gui import dialog
from kef.gui.base import Screen, hu_date, parse_hu_date
from kef.pipeline import TransformSettings

COLUMN_HELP = {
    "kolcsonzesi_napok": "kölcsönzés napjától a visszahozásig (vagy a viszonyítási napig)",
    "keses_napok": "a lejárat utáni napok száma, ha később hozták vissza / még kint van",
    "kesedelmi_dij": "késés × díj/nap, legfeljebb a plafonig; diák és nyugdíjas kedvezménnyel",
    "statusz": "visszahozva / kint / késésben",
    "ev_honap": "a kölcsönzés éve és hónapja (ÉÉÉÉ-HH) – kimutatásokhoz",
    "korcsoport": "0–14, 15–18, 19–35, 36–64, 65+ a születési év alapján",
}


class RulesScreen(Screen):
    number = 4
    title = "Formázási szabályok"

    def __init__(self, app) -> None:
        super().__init__(app)
        self.wizard = False
        self.info = ttk.Label(self.body, text="", style="Muted.TLabel", wraplength=900, justify="left")
        self.info.pack(anchor="w", pady=(0, 10))

        cols = ttk.Frame(self.body)
        cols.pack(fill="x")
        fee = ttk.LabelFrame(cols, text="Késedelmi díj", padding=12)
        fee.pack(side="left", fill="y", padx=(0, 16))
        self.vars = {k: tk.StringVar() for k in ("dij_per_nap", "dij_plafon", "kedvezmeny_szazalek", "ref")}
        self.inputs: list[ttk.Widget] = []
        for r, (key, label, unit) in enumerate([("dij_per_nap", "Díj / késedelmes nap:", "Ft"),
                                                ("dij_plafon", "Díjplafon kölcsönzésenként:", "Ft"),
                                                ("kedvezmeny_szazalek", "Kedvezmény (diák, nyugdíjas):", "%")]):
            ttk.Label(fee, text=label).grid(row=r, column=0, sticky="w", pady=5)
            e = ttk.Spinbox(fee, textvariable=self.vars[key], from_=0, to=100_000, increment=10, width=10)
            e.grid(row=r, column=1, padx=8)
            ttk.Label(fee, text=unit).grid(row=r, column=2, sticky="w")
            self.inputs.append(e)
        ttk.Separator(fee).grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
        ttk.Label(fee, text="Viszonyítási dátum:").grid(row=4, column=0, sticky="w")
        ttk.Entry(fee, textvariable=self.vars["ref"], width=12).grid(row=4, column=1, padx=8)
        ttk.Label(fee, text="ÉÉÉÉ.HH.NN.", style="Muted.TLabel").grid(row=4, column=2, sticky="w")
        ttk.Label(fee, text="(a még kint lévő könyvek késését eddig a napig számoljuk)",
                  style="Muted.TLabel").grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))

        cc = ttk.LabelFrame(cols, text="Hozzáadandó számított oszlopok", padding=12)
        cc.pack(side="left", fill="both", expand=True)
        self.col_vars: dict[str, tk.BooleanVar] = {}
        for r, name in enumerate(config.COMPUTED_COLUMNS):
            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(cc, text=config.DISPLAY_NAMES[name], variable=var)
            cb.grid(row=r, column=0, sticky="w", pady=3)
            ttk.Label(cc, text=COLUMN_HELP[name], style="Muted.TLabel").grid(row=r, column=1, sticky="w", padx=10)
            self.col_vars[name] = var
            self.inputs.append(cb)

        fixed = ttk.LabelFrame(self.body, text="Mindig elvégzett formázás", padding=12)
        fixed.pack(fill="x", pady=(14, 0))
        ttk.Label(fixed, justify="left", text=(
            "• technikai mezőnevek cseréje olvasható magyar fejlécekre   • dátumok ÉÉÉÉ.HH.NN. formátumban\n"
            "• fejléc kiemelése, rögzített első sor, automatikus szűrő, oszlopszélesség   • késésben lévő sorok "
            "színezése\n• felesleges szóközök eltávolítása, ISBN egységesítése   • rendezés fiók és kölcsönzési dátum "
            "szerint   • „Összesítő” munkalap minden fájlban")).pack(anchor="w")

        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", side="bottom", pady=(12, 0))
        self.back_btn = ttk.Button(bottom, text="‹ Vissza", command=self.back)
        self.back_btn.pack(side="left")
        self.save_btn = ttk.Button(bottom, text="Mentés alapértelmezettként", command=self.save_defaults)
        self.save_btn.pack(side="left", padx=8)
        self.reset_btn = ttk.Button(bottom, text="Gyári értékek", command=self.factory_reset)
        self.reset_btn.pack(side="left")
        self.next_btn = ttk.Button(bottom, text="Tovább: szétbontás ›", style="Primary.TButton", command=self.next)
        self.next_btn.pack(side="right")

    # --- megjelenítés ---
    def on_show(self, wizard: bool | None = None, **kwargs) -> None:
        if wizard is not None:
            self.wizard = wizard
        s = self.app.transform_settings if self.wizard and self.app.transform_settings else \
            TransformSettings.from_store(self.app.db.get_settings())
        self._fill(s)
        admin = self.app.user.can("beallitas_modositas")
        for w in self.inputs:
            w.state(["!disabled"] if admin else ["disabled"])
        self.save_btn.state(["!disabled"] if admin else ["disabled"])
        self.reset_btn.state(["!disabled"] if admin else ["disabled"])
        if self.wizard:
            self.next_btn.pack(side="right")
        else:
            self.next_btn.pack_forget()
        txt = ("A betöltött adatokhoz ezek a szabályok szerint adja hozzá a program a számított oszlopokat "
               "(a valós megoldásban ezt a VBA-makró végezte). ")
        txt += ("Adminisztrátorként módosíthatod és elmentheted az alapértelmezett értékeket."
                if admin else "A szabályokat csak az adminisztrátor módosíthatja; a viszonyítási dátum "
                              "futtatásonként megadható.")
        self.info.configure(text=txt)

    def _fill(self, s: TransformSettings) -> None:
        self.vars["dij_per_nap"].set(str(s.dij_per_nap))
        self.vars["dij_plafon"].set(str(s.dij_plafon))
        self.vars["kedvezmeny_szazalek"].set(str(s.kedvezmeny_szazalek))
        self.vars["ref"].set(hu_date(s.viszonyitasi_datum))
        for name, var in self.col_vars.items():
            var.set(name in s.szamitott_oszlopok)

    def read(self) -> TransformSettings:
        def num(key: str, label: str) -> int:
            try:
                return int(self.vars[key].get().strip())
            except ValueError:
                raise ValueError(f"A(z) „{label}” mezőbe egész számot adj meg.") from None

        s = TransformSettings(
            dij_per_nap=num("dij_per_nap", "Díj / késedelmes nap"),
            dij_plafon=num("dij_plafon", "Díjplafon"),
            kedvezmeny_szazalek=num("kedvezmeny_szazalek", "Kedvezmény"),
            viszonyitasi_datum=parse_hu_date(self.vars["ref"].get(), "Viszonyítási dátum") or date.today(),
            szamitott_oszlopok=[n for n, v in self.col_vars.items() if v.get()],
        )
        s.validate()
        return s

    # --- műveletek ---
    def back(self) -> None:
        self.app.show("load" if self.wizard else "menu")

    def next(self) -> None:
        try:
            self.app.transform_settings = self.read()
        except ValueError as exc:
            dialog.error(self, str(exc), "Hibás beállítás")
            return
        self.app.show("split")

    def save_defaults(self) -> None:
        try:
            require(self.app.user, "beallitas_modositas")
            s = self.read()
        except (ValueError, PermissionDenied) as exc:
            dialog.error(self, str(exc))
            return
        if not dialog.confirm(self, "Biztosan elmented ezeket az értékeket alapértelmezettként?\n"
                                    "A következő feldolgozások már ezekkel indulnak.", ok_text="Mentés"):
            return
        self.app.db.save_settings({
            "dij_per_nap": s.dij_per_nap, "dij_plafon": s.dij_plafon,
            "kedvezmeny_szazalek": s.kedvezmeny_szazalek, "szamitott_oszlopok": ",".join(s.szamitott_oszlopok),
        }, self.app.user.id)
        dialog.info(self, "A beállítások mentése megtörtént.")

    def factory_reset(self) -> None:
        self._fill(TransformSettings.from_store(dict(config.DEFAULT_SETTINGS)))
