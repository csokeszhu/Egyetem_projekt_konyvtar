"""[2.] Főmenü – a szerepkörnek megfelelő funkciók érhetők el."""
from __future__ import annotations

from tkinter import ttk

from kef.gui.base import Screen, fmt_int


class MenuScreen(Screen):
    number = 2
    title = "Főmenü"

    def __init__(self, app) -> None:
        super().__init__(app)
        self.welcome = ttk.Label(self.body, text="", style="Big.TLabel")
        self.welcome.pack(anchor="w")
        self.role_lbl = ttk.Label(self.body, text="", style="Muted.TLabel")
        self.role_lbl.pack(anchor="w", pady=(2, 18))

        grid = ttk.Frame(self.body)
        grid.pack(anchor="w")
        self.buttons: dict[str, ttk.Button] = {}
        items = [
            ("new", "Új feldolgozás\n(betöltés → formázás → szétbontás)", lambda: self._new()),
            ("rules", "Formázási szabályok", lambda: app.show("rules", wizard=False)),
            ("history", "Futtatási előzmények", lambda: app.show("history")),
            ("users", "Felhasználók és jelszó", lambda: app.show("users")),
            ("logout", "Kijelentkezés", app.logout),
            ("quit", "Kilépés", app.quit_app),
        ]
        for i, (key, text, cmd) in enumerate(items):
            b = ttk.Button(grid, text=text, command=cmd, style="Menu.TButton", width=34)
            b.grid(row=i // 2, column=i % 2, padx=(0, 14), pady=7, sticky="ew")
            self.buttons[key] = b

        self.hint = ttk.Label(self.body, text="", style="Muted.TLabel", wraplength=760, justify="left")
        self.hint.pack(anchor="w", pady=(18, 0))
        box = ttk.LabelFrame(self.body, text="Legutóbbi futtatás", padding=12)
        box.pack(anchor="w", fill="x", pady=(16, 0))
        self.last_lbl = ttk.Label(box, text="", justify="left")
        self.last_lbl.pack(anchor="w")

    def _new(self) -> None:
        self.app.reset_workflow()
        self.app.show("load")

    def on_show(self, **kwargs) -> None:
        u = self.app.user
        self.welcome.configure(text=f"Üdvözlünk, {u.nev}!")
        self.role_lbl.configure(text=f"Szerepkör: {u.role_name}")
        self.buttons["new"].state(["!disabled"] if u.can("feldolgozas") else ["disabled"])
        hints = []
        if not u.can("feldolgozas"):
            hints.append("Megtekintő szerepkörrel új feldolgozást nem indíthatsz; az előzmények és a "
                         "kimeneti fájlok megtekinthetők.")
        if not u.can("beallitas_modositas"):
            hints.append("A formázási szabályokat csak az adminisztrátor módosíthatja, te megtekintheted őket.")
        self.hint.configure(text="\n".join(hints))
        runs = self.app.db.list_runs()
        if runs:
            r = runs[0]
            self.last_lbl.configure(text=(
                f"#{r['id']}  {r['inditas']}  –  {r['felhasznalo']}  –  státusz: {r['statusz']}\n"
                f"{r['fajlok']} fájl, {fmt_int(r['ervenyes_sorok'])} érvényes / {fmt_int(r['hibas_sorok'])} hibás sor"))
        else:
            self.last_lbl.configure(text="Még nem volt feldolgozás.")
