"""[9.] Felhasználók és jelszó – saját jelszó módosítása; adminisztrátornak felhasználókezelés is."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from kef.auth import (PASSWORD_RULE, ROLE_NAMES, AuthError, PermissionDenied, change_password, create_user,
                      reset_password, update_user)
from kef.gui import dialog
from kef.gui.base import Screen, Table

ROLE_BY_NAME = {v: k for k, v in ROLE_NAMES.items()}


class UsersScreen(Screen):
    number = 9
    title = "Felhasználók és jelszó"

    def __init__(self, app) -> None:
        super().__init__(app)
        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", side="bottom", pady=(12, 0))
        ttk.Button(bottom, text="‹ Főmenü", command=lambda: app.show("menu")).pack(side="right")

        pw = ttk.LabelFrame(self.body, text="Saját jelszó módosítása", padding=12)
        pw.pack(fill="x")
        self.pw_vars = [tk.StringVar() for _ in range(3)]
        for i, label in enumerate(["Régi jelszó:", "Új jelszó:", "Új jelszó még egyszer:"]):
            ttk.Label(pw, text=label).grid(row=0, column=2 * i, sticky="w", padx=(0 if i == 0 else 14, 4))
            ttk.Entry(pw, textvariable=self.pw_vars[i], show="•", width=18).grid(row=0, column=2 * i + 1)
        ttk.Button(pw, text="Jelszó mentése", style="Primary.TButton", command=self.change_pw).grid(
            row=0, column=6, padx=(14, 0))
        ttk.Label(pw, text=PASSWORD_RULE, style="Muted.TLabel").grid(row=1, column=0, columnspan=7, sticky="w",
                                                                     pady=(6, 0))

        self.admin_box = ttk.LabelFrame(self.body, text="Felhasználók kezelése (csak adminisztrátor)", padding=12)
        self.admin_box.pack(fill="both", expand=True, pady=(14, 0))
        self.table = Table(self.admin_box, [("azon", "Azonosító", 120), ("nev", "Név", 180),
                                            ("szerep", "Szerepkör", 120), ("aktiv", "Aktív", 60),
                                            ("letrehozva", "Létrehozva", 140), ("belepes", "Utolsó belépés", 150)],
                           height=7)
        self.table.pack(fill="both", expand=True)
        self.table.tree.bind("<<TreeviewSelect>>", lambda e: self._selected())

        forms = ttk.Frame(self.admin_box)
        forms.pack(fill="x", pady=(10, 0))
        edit = ttk.LabelFrame(forms, text="Kijelölt felhasználó", padding=10)
        edit.pack(side="left", fill="y")
        self.sel_lbl = ttk.Label(edit, text="–", style="Value.TLabel")
        self.sel_lbl.grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(edit, text="Szerepkör:").grid(row=1, column=0, sticky="w", pady=4)
        self.edit_role = tk.StringVar()
        ttk.Combobox(edit, textvariable=self.edit_role, values=list(ROLE_BY_NAME), state="readonly",
                     width=16).grid(row=1, column=1, padx=6)
        self.edit_active = tk.BooleanVar()
        ttk.Checkbutton(edit, text="Aktív", variable=self.edit_active).grid(row=2, column=1, sticky="w")
        ttk.Button(edit, text="Módosítás mentése", command=self.save_user).grid(row=3, column=0, pady=(6, 0))
        ttk.Button(edit, text="Jelszó visszaállítása", command=self.reset_pw).grid(row=3, column=1, pady=(6, 0))

        new = ttk.LabelFrame(forms, text="Új felhasználó", padding=10)
        new.pack(side="left", fill="both", expand=True, padx=(14, 0))
        self.new_vars = {k: tk.StringVar() for k in ("azon", "nev", "jelszo")}
        self.new_role = tk.StringVar(value=ROLE_NAMES["feldolgozo"])
        for r, (k, label) in enumerate([("azon", "Azonosító:"), ("nev", "Név:"), ("jelszo", "Kezdő jelszó:")]):
            ttk.Label(new, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(new, textvariable=self.new_vars[k], width=22, show="•" if k == "jelszo" else "").grid(
                row=r, column=1, padx=6)
        ttk.Label(new, text="Szerepkör:").grid(row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Combobox(new, textvariable=self.new_role, values=list(ROLE_BY_NAME), state="readonly",
                     width=16).grid(row=0, column=3, padx=6)
        ttk.Button(new, text="Felhasználó létrehozása", style="Primary.TButton", command=self.add_user).grid(
            row=2, column=2, columnspan=2, sticky="e")

    def on_show(self, **kwargs) -> None:
        for v in self.pw_vars:
            v.set("")
        if self.app.user.can("felhasznalo_kezeles"):
            self.admin_box.pack(fill="both", expand=True, pady=(14, 0))
            self.load_users()
        else:
            self.admin_box.pack_forget()

    def load_users(self) -> None:
        rows = self.app.db.list_users()
        self._rows = {str(r["id"]): r for r in rows}
        self.table.set_rows([(r["azonosito"], r["nev"], ROLE_NAMES[r["szerepkor"]], "igen" if r["aktiv"] else "nem",
                              r["letrehozva"], r["utolso_belepes"] or "") for r in rows],
                            tags=["" if r["aktiv"] else "bad" for r in rows], iids=[r["id"] for r in rows])
        self.sel_lbl.configure(text="–")

    def _selected(self):
        iid = self.table.selected_iid()
        if not iid:
            return None
        r = self._rows[iid]
        self.sel_lbl.configure(text=f"{r['azonosito']} – {r['nev']}")
        self.edit_role.set(ROLE_NAMES[r["szerepkor"]])
        self.edit_active.set(bool(r["aktiv"]))
        return r

    def change_pw(self) -> None:
        old, new, again = (v.get() for v in self.pw_vars)
        try:
            change_password(self.app.db, self.app.user, old, new, again)
        except AuthError as exc:
            dialog.error(self, str(exc), "Sikertelen jelszómódosítás")
            return
        for v in self.pw_vars:
            v.set("")
        dialog.info(self, "Az új jelszó mentése megtörtént.")

    def save_user(self) -> None:
        r = self._selected()
        if r is None:
            dialog.warning(self, "Előbb jelölj ki egy felhasználót!")
            return
        if not dialog.confirm(self, f"Biztosan módosítod {r['azonosito']} adatait?", ok_text="Mentés"):
            return
        try:
            update_user(self.app.db, self.app.user, r["id"], ROLE_BY_NAME[self.edit_role.get()],
                        self.edit_active.get())
        except (AuthError, PermissionDenied) as exc:
            dialog.error(self, str(exc))
            return
        self.load_users()
        dialog.info(self, "Az adatok mentése megtörtént.")

    def reset_pw(self) -> None:
        r = self._selected()
        if r is None:
            dialog.warning(self, "Előbb jelölj ki egy felhasználót!")
            return
        if not dialog.confirm(self, f"Biztosan visszaállítod {r['azonosito']} jelszavát az alapértelmezettre?",
                              ok_text="Visszaállítás"):
            return
        try:
            pw = reset_password(self.app.db, self.app.user, r["id"])
        except PermissionDenied as exc:
            dialog.error(self, str(exc))
            return
        dialog.info(self, f"A jelszó visszaállítva. Ideiglenes jelszó: {pw}\nKérd meg a felhasználót, "
                          "hogy belépés után változtassa meg.")

    def add_user(self) -> None:
        v = self.new_vars
        try:
            create_user(self.app.db, self.app.user, v["azon"].get(), v["nev"].get(),
                        ROLE_BY_NAME[self.new_role.get()], v["jelszo"].get())
        except (AuthError, PermissionDenied) as exc:
            dialog.error(self, str(exc), "Sikertelen létrehozás")
            return
        for var in v.values():
            var.set("")
        self.load_users()
        dialog.info(self, "Az új felhasználó létrejött.")
