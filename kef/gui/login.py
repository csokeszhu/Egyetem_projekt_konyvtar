"""[1.] Bejelentkezés."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from kef import config
from kef.auth import AuthError, login
from kef.gui import dialog
from kef.gui.base import PRIMARY, Screen


class LoginScreen(Screen):
    number = 1
    title = "Bejelentkezés"

    def __init__(self, app) -> None:
        super().__init__(app)
        card = tk.Frame(self.body, bg="white", padx=36, pady=30, highlightthickness=1, highlightbackground="#D5DCE4")
        card.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(card, text=config.APP_NAME, bg="white", fg=PRIMARY, font=("Segoe UI Semibold", 17)).grid(
            row=0, column=0, columnspan=2, pady=(0, 2))
        tk.Label(card, text="Városi Könyvtárhálózat – kölcsönzési exportok feldolgozása és szétbontása",
                 bg="white", fg="#5F6B7A").grid(row=1, column=0, columnspan=2, pady=(0, 22))

        tk.Label(card, text="Felhasználói azonosító:", bg="white").grid(row=2, column=0, sticky="w", pady=6)
        self.user_var = tk.StringVar()
        self.user_entry = ttk.Entry(card, textvariable=self.user_var, width=30)
        self.user_entry.grid(row=2, column=1, pady=6, padx=(12, 0))
        tk.Label(card, text="Jelszó:", bg="white").grid(row=3, column=0, sticky="w", pady=6)
        self.pw_var = tk.StringVar()
        self.pw_entry = ttk.Entry(card, textvariable=self.pw_var, show="•", width=30)
        self.pw_entry.grid(row=3, column=1, pady=6, padx=(12, 0))

        btns = tk.Frame(card, bg="white")
        btns.grid(row=4, column=0, columnspan=2, pady=(22, 0), sticky="e")
        ttk.Button(btns, text="Bejelentkezés", style="Primary.TButton", command=self.do_login).pack(side="left", padx=6)
        ttk.Button(btns, text="Kilépés", command=app.quit_app).pack(side="left")
        tk.Label(card, text="Elfelejtett jelszó esetén az adminisztrátor tudja visszaállítani.",
                 bg="white", fg="#8A94A3", font=("Segoe UI", 8)).grid(row=5, column=0, columnspan=2, pady=(18, 0))

        self.user_entry.bind("<Return>", lambda e: self.pw_entry.focus_set())
        self.pw_entry.bind("<Return>", lambda e: self.do_login())

    def on_show(self, **kwargs) -> None:
        self.pw_var.set("")
        self.user_entry.focus_set()

    def do_login(self) -> None:
        try:
            user = login(self.app.db, self.user_var.get(), self.pw_var.get())
        except AuthError as exc:
            self.pw_var.set("")
            dialog.error(self, str(exc), "Sikertelen bejelentkezés")
            self.pw_entry.focus_set()
            return
        self.app.user = user
        self.app.show("menu")
