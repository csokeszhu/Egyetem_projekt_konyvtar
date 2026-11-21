"""[6.] Feldolgozás folyamatban – háttérszálon fut, megszakítható."""
from __future__ import annotations

import threading
import time
import tkinter as tk
from tkinter import ttk

from kef.gui import dialog
from kef.gui.base import Screen, fmt_int
from kef.pipeline import Cancelled
from kef.service import run_processing


class ProgressScreen(Screen):
    number = 6
    title = "Feldolgozás folyamatban"

    def __init__(self, app) -> None:
        super().__init__(app)
        self.cancel_event = threading.Event()
        self.started = 0.0
        self.step_lbl = ttk.Label(self.body, text="", style="Value.TLabel")
        self.step_lbl.pack(anchor="w")
        self.pbar = ttk.Progressbar(self.body, mode="determinate", maximum=100)
        self.pbar.pack(fill="x", pady=(8, 4))
        info = ttk.Frame(self.body)
        info.pack(fill="x")
        self.pct_lbl = ttk.Label(info, text="0%", style="Muted.TLabel")
        self.pct_lbl.pack(side="left")
        self.time_lbl = ttk.Label(info, text="Eltelt idő: 0 mp", style="Muted.TLabel")
        self.time_lbl.pack(side="right")

        box = ttk.LabelFrame(self.body, text="Napló", padding=8)
        box.pack(fill="both", expand=True, pady=(12, 0))
        self.log = tk.Text(box, height=16, font=("Consolas", 9), bg="#FBFCFD", relief="flat", state="disabled")
        sb = ttk.Scrollbar(box, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bottom = ttk.Frame(self.body)
        bottom.pack(fill="x", pady=(12, 0))
        self.cancel_btn = ttk.Button(bottom, text="Megszakítás", command=self.cancel)
        self.cancel_btn.pack(side="right")

    def write(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"{time.strftime('%H:%M:%S')}  {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def on_show(self, start: bool = False, **kwargs) -> None:
        if not start:
            return
        self.cancel_event = threading.Event()
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.pbar["value"] = 0
        self.cancel_btn.state(["!disabled"])
        app = self.app
        lr, s, cfg = app.load_result, app.transform_settings, app.split_config
        self.write(f"Forrás: {lr.path.name} – {fmt_int(lr.valid_rows)} érvényes, {fmt_int(lr.error_rows)} hibás sor")
        self.write(f"Szempont: {cfg.szempont}, kimeneti mappa: {cfg.kimeneti_mappa}")
        self.step_lbl.configure(text="Indítás…")
        self.started = time.perf_counter()
        app.busy = True
        self._tick()
        app.run_in_thread(
            lambda: run_processing(app.db, app.user, lr, s, cfg, progress=self._progress_from_thread,
                                   cancel=self.cancel_event),
            self._done, self._failed)

    def _tick(self) -> None:
        if self.app.busy and self.app.current == "progress":
            self.time_lbl.configure(text=f"Eltelt idő: {time.perf_counter() - self.started:.0f} mp")
            self.after(500, self._tick)

    def _progress_from_thread(self, i: int, total: int, msg: str) -> None:
        self.app.post(self._progress, i, total, msg)

    def _progress(self, i: int, total: int, msg: str) -> None:
        pct = int(100 * i / total) if total else 0
        self.pbar["value"] = pct
        self.pct_lbl.configure(text=f"{pct}%  ({i}/{total})")
        self.step_lbl.configure(text=msg)
        self.write(msg)

    def cancel(self) -> None:
        if dialog.confirm(self, "Biztosan megszakítod a feldolgozást?\nA már létrehozott fájlok törlődnek.",
                          ok_text="Megszakítás"):
            self.cancel_event.set()
            self.cancel_btn.state(["disabled"])
            self.write("Megszakítás kérve – az aktuális fájl befejezése után leáll…")

    def _done(self, result) -> None:
        self.app.busy = False
        self.write(f"Kész: {len(result.files)} fájl, {result.seconds:.1f} mp.")
        self.app.last_run = result
        self.app.show("result")

    def _failed(self, exc: Exception) -> None:
        self.app.busy = False
        if isinstance(exc, Cancelled):
            self.write("A feldolgozás megszakítva, a részleges fájlok törölve.")
            dialog.warning(self, "A feldolgozás megszakítva. A részleges fájlokat töröltük.")
        else:
            self.write(f"HIBA: {exc}")
            dialog.error(self, f"A feldolgozás sikertelen:\n{exc}")
        self.app.show("split")
