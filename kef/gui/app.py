"""A főablak: képernyők közötti navigáció, közös állapot, háttérszálak kezelése."""
from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from tkinter import ttk
from typing import Any, Callable

from kef import config
from kef.auth import User
from kef.db import Database
from kef.gui import dialog
from kef.gui.base import setup_style


def enable_dpi_awareness() -> None:
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # pragma: no cover
            pass


class App(tk.Tk):
    def __init__(self, db: Database | None = None) -> None:
        enable_dpi_awareness()
        super().__init__()
        self.title(f"{config.APP_NAME} ({config.APP_SHORT} {config.APP_VERSION})")
        self.geometry("1180x760")
        self.minsize(1000, 660)
        setup_style(self)

        self.db = db or Database()
        self.user: User | None = None
        # a varázsló lépései között megőrzött állapot
        self.load_result = None
        self.transform_settings = None
        self.split_config = None
        self.last_run = None
        self.busy = False

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)
        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        from kef.gui.history import HistoryScreen
        from kef.gui.load import LoadScreen
        from kef.gui.login import LoginScreen
        from kef.gui.menu import MenuScreen
        from kef.gui.progress import ProgressScreen
        from kef.gui.result import ResultScreen
        from kef.gui.rules import RulesScreen
        from kef.gui.splitscreen import SplitScreen
        from kef.gui.users import UsersScreen

        self.screens: dict[str, Any] = {}
        for key, cls in [("login", LoginScreen), ("menu", MenuScreen), ("load", LoadScreen),
                         ("rules", RulesScreen), ("split", SplitScreen), ("progress", ProgressScreen),
                         ("result", ResultScreen), ("history", HistoryScreen), ("users", UsersScreen)]:
            screen = cls(self)
            screen.grid(row=0, column=0, sticky="nsew")
            self.screens[key] = screen
        self.current = ""

        self._queue: queue.Queue = queue.Queue()
        self.after(80, self._poll)
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.show("login")

    # --- navigáció ---
    def show(self, key: str, **kwargs) -> None:
        screen = self.screens[key]
        screen.refresh_header()
        screen.tkraise()
        self.current = key
        self.title(f"[{screen.number}.] {screen.title} – {config.APP_NAME}")
        screen.on_show(**kwargs)

    def reset_workflow(self) -> None:
        self.load_result = None
        self.transform_settings = None
        self.split_config = None

    def logout(self) -> None:
        if dialog.confirm(self, "Biztosan kijelentkezel?", ok_text="Kijelentkezés"):
            self.user = None
            self.reset_workflow()
            self.last_run = None
            self.show("login")

    def quit_app(self) -> None:
        msg = "Biztosan kilépsz a programból?"
        if self.busy:
            msg = "Feldolgozás van folyamatban! Kilépéskor a részleges fájlok törlődnek. Biztosan kilépsz?"
            if not dialog.confirm(self, msg, ok_text="Kilépés"):
                return
            self.screens["progress"].cancel_event.set()
            self.destroy()
            return
        if dialog.confirm(self, msg, ok_text="Kilépés"):
            self.destroy()

    # --- háttérszálak ---
    def post(self, fn: Callable, *args) -> None:
        """Másik szálból: fn(*args) végrehajtása a GUI-szálon."""
        self._queue.put((fn, args))

    def _poll(self) -> None:
        try:
            while True:
                fn, args = self._queue.get_nowait()
                fn(*args)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(80, self._poll)

    def run_in_thread(self, work: Callable[[], Any], on_done: Callable[[Any], None],
                      on_error: Callable[[Exception], None]) -> threading.Thread:
        def target() -> None:
            try:
                result = work()
            except Exception as exc:  # noqa: BLE001 - a GUI-n jelezzük
                self.post(on_error, exc)
            else:
                self.post(on_done, result)

        t = threading.Thread(target=target, daemon=True)
        t.start()
        return t
