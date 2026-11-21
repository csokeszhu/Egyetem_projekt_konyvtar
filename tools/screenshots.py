"""A dokumentációhoz szükséges képernyőképek automatikus elkészítése.

Végigvezeti az alkalmazást a teljes folyamaton (bemutató adatbázissal), és minden
ablakról PNG-t ment a docs/screenshots mappába. Egyben füstteszt is: ha bármelyik
lépés hibát dob, a szkript hibakóddal áll le.

    python tools/screenshots.py
"""
from __future__ import annotations

import ctypes
import shutil
import sys
import traceback
from ctypes import wintypes
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from PIL import ImageGrab  # noqa: E402

from kef import config  # noqa: E402
from kef.auth import login  # noqa: E402
from kef.db import Database  # noqa: E402
from kef.gui.app import App  # noqa: E402
from kef.gui.dialog import MessageDialog  # noqa: E402

OUT = config.BASE_DIR / "docs" / "screenshots"
WORK = config.BASE_DIR / "docs" / "_demo"


def window_rect(widget) -> tuple[int, int, int, int]:
    """A teljes ablak (kerettel, címsorral) képernyő-koordinátái DWM-en keresztül."""
    hwnd = int(widget.winfo_toplevel().wm_frame(), 16)
    rect = wintypes.RECT()
    ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(rect), ctypes.sizeof(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def grab(widget, name: str) -> None:
    top = widget.winfo_toplevel()
    if top.state() in ("iconic", "withdrawn"):
        top.deiconify()
    top.lift()
    top.attributes("-topmost", True)
    for _ in range(3):
        widget.update_idletasks()
        widget.update()
        top.after(120)
    left, t, right, bottom = window_rect(widget)
    if right - left < 300:
        raise RuntimeError(f"A(z) {name} ablak nem látható (méret: {right - left}×{bottom - t}).")
    img = ImageGrab.grab(bbox=window_rect(widget), all_screens=True)
    img.save(OUT / f"{name}.png")
    print("  mentve:", name)


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    db = Database(WORK / "kef_demo.db")
    # előzmény a bemutatóhoz: egy korábbi, megszakított futtatás a feldolgozó nevében
    feldolgozo = login(db, "feldolgozo", "Feldolgozo1")
    rid = db.start_run(feldolgozo.id, str(config.SAMPLE_DIR / "kolcsonzesek_2026_10.xlsx"), 88_412, 88_190, 222,
                       "fiok", str(WORK / "output"))
    db.finish_run(rid, "megszakitva", "A felhasználó megszakította; a részleges fájlok törölve.")
    db.save_settings({"kimeneti_mappa": str(WORK / "output")}, feldolgozo.id)

    app = App(db=db)
    app.geometry("1180x760+40+30")
    app.attributes("-topmost", True)
    state = {"failed": False}

    steps = []

    def step(delay: int):
        def deco(fn):
            steps.append((delay, fn))
            return fn
        return deco

    @step(600)
    def s_login():
        scr = app.screens["login"]
        scr.user_var.set("admin")
        scr.pw_var.set("Admin123!")
        grab(app, "01_bejelentkezes")
        scr.do_login()

    @step(500)
    def s_menu():
        grab(app, "02_fomenu")
        app.screens["menu"]._new()
        app.screens["load"].use_sample()
        app.screens["load"].load()

    @step(700)
    def s_loading():
        grab(app, "03a_betoltes_folyamatban")

    def wait_loaded():
        if app.load_result is None:
            app.after(300, wait_loaded)
            return
        app.after(400, s_loaded)

    def s_loaded():
        grab(app, "03_adatbetoltes")
        app.screens["load"].nb.select(1)
        app.after(300, s_errors)

    def s_errors():
        grab(app, "03b_hibas_sorok")
        app.screens["load"].nb.select(0)
        app.show("rules", wizard=True)
        app.after(400, s_rules)

    def s_rules():
        grab(app, "04_formazasi_szabalyok")
        app.screens["rules"].next()
        app.after(600, s_split)

    def s_split():
        grab(app, "05_szetbontas")
        scr = app.screens["split"]
        cfg = scr.read()
        d = MessageDialog(app, f"Biztosan elindítod a feldolgozást?\n\n12 fájl + összesítő jön létre ide:\n"
                               f"{cfg.kimeneti_mappa}", "confirm", ok_text="Indítás")
        d.attributes("-topmost", True)
        app.after(500, lambda: s_dialog(d, cfg))

    def s_dialog(d, cfg):
        grab(d, "10_megerosites")
        d.destroy()
        app.split_config = cfg
        app.show("progress", start=True)
        app.after(9000, s_progress)

    def s_progress():
        grab(app, "06_feldolgozas")
        wait_result()

    def wait_result():
        if app.current != "result":
            app.after(500, wait_result)
            return
        app.after(500, s_result)

    def s_result():
        grab(app, "07_eredmeny")
        app.show("history")
        app.after(500, s_history)

    def s_history():
        grab(app, "08_elozmenyek")
        app.show("users")
        app.after(500, s_users)

    def s_users():
        scr = app.screens["users"]
        scr.table.tree.selection_set(scr.table.tree.get_children()[1])
        app.after(300, s_users2)

    def s_users2():
        grab(app, "09_felhasznalok")
        # megtekintő szerepkörű főmenü (letiltott gombok)
        app.user = login(db, "megtekinto", "Megtekinto1")
        app.show("menu")
        app.after(400, s_viewer)

    def s_viewer():
        grab(app, "02b_fomenu_megtekinto")
        app.after(200, app.destroy)

    def run_steps(i=0):
        if i >= len(steps):
            wait_loaded()
            return
        delay, fn = steps[i]

        def go():
            try:
                fn()
            except Exception:
                traceback.print_exc()
                state["failed"] = True
                app.destroy()
                return
            run_steps(i + 1)
        app.after(delay, go)

    def report_callback_exception(exc, val, tb):
        traceback.print_exception(exc, val, tb)
        state["failed"] = True
        app.destroy()

    app.report_callback_exception = report_callback_exception
    run_steps()
    app.mainloop()
    db.close()
    return 1 if state["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
