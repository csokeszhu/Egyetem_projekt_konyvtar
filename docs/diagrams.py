"""A dokumentáció ábráinak előállítása matplotlibbel (PNG a docs/diagrams mappába).

    python docs/diagrams.py
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch  # noqa: E402

OUT = Path(__file__).resolve().parent / "diagrams"
PRIMARY = "#1F4E78"
LIGHT = "#DCE8F4"
GREY = "#5F6B7A"
BLUE = "#1565C0"
ORANGE = "#C55A11"
plt.rcParams["font.family"] = "Segoe UI"


def box(ax, x, y, w, h, text, fc="white", ec=PRIMARY, fs=9, bold=False, color="black", lw=1.4, ls="-"):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                                fc=fc, ec=ec, lw=lw, ls=ls))
    ax.text(x, y, text, ha="center", va="center", fontsize=fs, color=color,
            fontweight="bold" if bold else "normal", wrap=True)


def arrow(ax, p1, p2, color="black", rad=0.0, label=None, lw=1.2, fs=7.5, ls="-", label_pos=0.5, both=False):
    a = FancyArrowPatch(p1, p2, arrowstyle="<|-|>" if both else "-|>", mutation_scale=11, color=color, lw=lw,
                        connectionstyle=f"arc3,rad={rad}", shrinkA=2, shrinkB=2, linestyle=ls)
    ax.add_patch(a)
    if label:
        mx = p1[0] + (p2[0] - p1[0]) * label_pos
        my = p1[1] + (p2[1] - p1[1]) * label_pos
        # az ív közepének közelítése
        dx, dy = p2[0] - p1[0], p2[1] - p1[1]
        mx += -dy * rad * 0.5
        my += dx * rad * 0.5
        ax.text(mx, my, label, fontsize=fs, color=color, ha="center", va="center",
                bbox=dict(fc="white", ec="none", pad=0.6))


def finish(fig, ax, name, xlim, ylim):
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.savefig(OUT / name, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  ", name)


# ---------------------------------------------------------------------------------------------
def funkciofa():
    tree = [
        ("Belépés és\njogosultság", ["Bejelentkezés", "Kijelentkezés,\nkilépés", "Saját jelszó\nmódosítása",
                                     "Felhasználók\nkezelése (admin)", "Jelszó\nvisszaállítása (admin)"]),
        ("Adatbetöltés", ["Forrásfájl\nkiválasztása", "Beolvasás\n(.xlsx / .csv)", "Validáció\n(oszlop, sor)",
                          "Előnézet és\nhibalista"]),
        ("Formázás", ["Szabályok\nmegtekintése", "Szabályok\nmódosítása (admin)", "Számított\noszlopok",
                      "Késedelmi díj\nszámítása", "Fejlécek és\nformátumok"]),
        ("Szétbontás", ["Szempont és\nidőszak", "Fájlnév-minta,\nkimeneti mappa", "Előnézet",
                        "Fájlok írása,\nösszesítő", "Megszakítás"]),
        ("Előzmények", ["Futtatások\nlistázása", "Szűrés dátum,\nfelhasználó szerint", "Fájlok\nmegtekintése",
                        "Mappa\nmegnyitása"]),
    ]
    fig, ax = plt.subplots(figsize=(15, 8.2))
    root = (7.5, 9.2)
    box(ax, *root, 4.6, 0.9, "Kölcsönzési Export Feldolgozó (KEF)", fc=PRIMARY, ec=PRIMARY, fs=11, bold=True,
        color="white")
    xs = [1.5, 4.5, 7.5, 10.5, 13.5]
    ax.plot([xs[0], xs[-1]], [8.2, 8.2], color=GREY, lw=1.2)
    ax.plot([root[0], root[0]], [8.75, 8.2], color=GREY, lw=1.2)
    for x, (name, leaves) in zip(xs, tree):
        ax.plot([x, x], [8.2, 7.75], color=GREY, lw=1.2)
        box(ax, x, 7.35, 2.5, 0.8, name, fc=LIGHT, fs=10, bold=True)
        lx = x - 1.05
        ax.plot([lx, lx], [6.95, 6.0 - (len(leaves) - 1) * 1.15], color=GREY, lw=1)
        for j, leaf in enumerate(leaves):
            y = 6.0 - j * 1.15
            ax.plot([lx, x - 0.85], [y, y], color=GREY, lw=1)
            box(ax, x + 0.2, y, 2.1, 0.85, leaf, fs=8.5)
    finish(fig, ax, "funkciofa.png", (0, 15), (0.6, 9.8))


# ---------------------------------------------------------------------------------------------
def bejarhatosag():
    fig, ax = plt.subplots(figsize=(15, 8.8))
    H = 0.95
    nodes = {  # kulcs: (x, y, szélesség, felirat)
        "exit": (0.9, 7.6, 1.7, "Program\nbezárása"),
        1: (0.9, 4.6, 2.3, "[1.]\nBejelentkezés"),
        2: (4.6, 4.6, 2.3, "[2.]\nFőmenü"),
        9: (4.6, 7.6, 2.3, "[9.]\nFelhasználók\nés jelszó"),
        8: (8.4, 7.6, 2.3, "[8.]\nFuttatási\nelőzmények"),
        4: (8.4, 4.6, 2.3, "[4.]\nFormázási\nszabályok"),
        7: (14.0, 4.6, 2.3, "[7.]\nEredmény\nösszesítő"),
        3: (4.6, 1.2, 2.3, "[3.]\nAdatbetöltés\nés validáció"),
        5: (8.4, 1.2, 2.3, "[5.]\nSzétbontás\nbeállításai"),
        10: (11.2, 1.2, 2.0, "[10.]\nMegerősítés"),
        6: (14.0, 1.2, 2.3, "[6.]\nFeldolgozás\nfolyamatban"),
    }
    for k, (x, y, w, t) in nodes.items():
        if k == "exit":
            box(ax, x, y, w, H, t, fc="#EEEEEE", ec=GREY, ls="--", fs=9)
        else:
            box(ax, x, y, w, H, t, fc=LIGHT if k == 2 else "white", fs=9.5, bold=k == 2)

    def link(p1, p2, label="", color="black", both=False, rad=0.0, dx=0.0, dy=0.0, pos=0.5, ha="center"):
        arrow(ax, p1, p2, color=color, rad=rad, both=both)
        if label:
            mx = p1[0] + (p2[0] - p1[0]) * pos + dx
            my = p1[1] + (p2[1] - p1[1]) * pos + dy
            ax.text(mx, my, label, fontsize=7.8, color=color, ha=ha, va="center",
                    bbox=dict(fc="white", ec="none", pad=0.8))

    top, bot = 4.6 + H / 2, 4.6 - H / 2
    link((2.05, 4.6), (3.45, 4.6), "OK /\nKijelentkezés", both=True, dy=0.42)
    link((0.9, top), (0.9, 7.6 - H / 2), "Kilépés / [X]")
    link((3.55, top), (1.75, 7.45), "Kilépés / [X]", pos=0.45)
    link((4.6, top), (4.6, 7.6 - H / 2), "Felhasználók és jelszó /\nFőmenü", both=True)
    link((5.6, top), (7.6, 7.6 - H / 2), "Előzmények /\nFőmenü", both=True)
    link((5.75, 4.6), (7.25, 4.6), "Formázási\nszabályok /\nVissza", both=True, dy=-0.55)
    link((4.3, bot), (4.3, 1.2 + H / 2), "Új\nfeldolgozás", color=BLUE, dx=-0.15, ha="right")
    link((4.9, 1.2 + H / 2), (4.9, bot), "Főmenü", dx=0.12, ha="left", pos=0.35)
    link((5.75, 1.55), (7.5, bot), "Tovább /\nVissza", color=BLUE, both=True, dx=0.15, dy=-0.1, ha="left")
    link((8.4, bot), (8.4, 1.2 + H / 2), "Tovább /\nVissza", color=BLUE, both=True, dx=0.15, ha="left")
    link((9.55, 1.2), (10.2, 1.2), "Indítás /\nMégse", color=BLUE, both=True, dy=0.62)
    link((12.2, 1.2), (12.85, 1.2), "Igen /\nMegszakítás", color=BLUE, both=True, dy=0.62)
    link((14.0, 1.2 + H / 2), (14.0, bot), "kész", color=BLUE)
    arrow(ax, (14.0, 1.2 - H / 2), (8.4, 1.2 - H / 2), color="black", rad=-0.18)
    ax.text(11.2, -0.02, "hiba vagy megszakítás → vissza a [5.]-re", fontsize=7.8, ha="center",
            bbox=dict(fc="white", ec="none", pad=0.8))
    arrow(ax, (14.0, top), (5.3, top), color="black", rad=0.3)
    ax.text(9.65, 6.15, "Főmenü", fontsize=7.8, ha="center", bbox=dict(fc="white", ec="none", pad=0.8))
    link((12.85, 4.3), (5.75, 1.0), "Új feldolgozás", color=BLUE, pos=0.22, dy=0.12)

    # jelmagyarázat
    lx, ly = 10.5, 8.35
    ax.add_patch(FancyBboxPatch((lx - 0.15, ly - 1.75), 5.0, 2.05, boxstyle="round,pad=0.02", fc="white",
                                ec=GREY, lw=0.8))
    ax.text(lx, ly, "Jelmagyarázat", fontsize=9, fontweight="bold", va="center")
    ax.plot([lx, lx + 0.6], [ly - 0.42] * 2, color="black", lw=1.4)
    ax.text(lx + 0.75, ly - 0.42, "mindenki (admin, feldolgozó, megtekintő)", fontsize=8, va="center")
    ax.plot([lx, lx + 0.6], [ly - 0.8] * 2, color=BLUE, lw=1.4)
    ax.text(lx + 0.75, ly - 0.8, "csak admin és feldolgozó", fontsize=8, va="center")
    ax.text(lx, ly - 1.18, "[4.] és [9.] szerkeszthető része csak adminnak", fontsize=8, va="center", color=ORANGE)
    ax.text(lx, ly - 1.52, "[10.] mentés / indítás / kilépés előtt mindenhol", fontsize=8, va="center", color=GREY)
    finish(fig, ax, "bejarhatosag.png", (-0.6, 15.4), (-0.6, 8.6))


# ---------------------------------------------------------------------------------------------
def er_diagram():
    fig, ax = plt.subplots(figsize=(15, 9.2))

    def entity(x, y, name, cols, w=3.6, color=PRIMARY):
        row_h = 0.34
        h = row_h * (len(cols) + 1) + 0.1
        top = y
        ax.add_patch(FancyBboxPatch((x, top - h), w, h, boxstyle="round,pad=0.02,rounding_size=0.06", fc="white",
                                    ec=color, lw=1.4))
        ax.add_patch(FancyBboxPatch((x, top - row_h - 0.05), w, row_h + 0.05, boxstyle="round,pad=0.02,rounding_size=0.06",
                                    fc=color, ec=color, lw=1.4))
        ax.text(x + w / 2, top - row_h / 2 - 0.03, name, ha="center", va="center", color="white", fontsize=10,
                fontweight="bold")
        for i, (key, col, typ) in enumerate(cols):
            yy = top - row_h * (i + 1.5) - 0.07
            ax.text(x + 0.12, yy, key, fontsize=7.5, color=ORANGE, va="center", fontweight="bold")
            ax.text(x + 0.62, yy, col, fontsize=8.2, va="center")
            ax.text(x + w - 0.12, yy, typ, fontsize=7.5, va="center", ha="right", color=GREY)
        return (x, top - h, w, h)

    fel = entity(0.3, 8.6, "felhasznalo", [
        ("PK", "id", "INTEGER"), ("UQ", "azonosito", "TEXT"), ("", "nev", "TEXT"), ("", "jelszo_hash", "TEXT"),
        ("", "szerepkor", "TEXT"), ("", "aktiv", "INTEGER"), ("", "letrehozva", "TEXT"),
        ("", "utolso_belepes", "TEXT")])
    fut = entity(5.6, 8.6, "futtatas", [
        ("PK", "id", "INTEGER"), ("FK", "felhasznalo_id", "INTEGER"), ("", "inditas", "TEXT"),
        ("", "befejezes", "TEXT"), ("", "forrasfajl", "TEXT"), ("", "sorok_be", "INTEGER"),
        ("", "ervenyes_sorok", "INTEGER"), ("", "hibas_sorok", "INTEGER"), ("", "szempont", "TEXT"),
        ("", "kimeneti_mappa", "TEXT"), ("", "statusz", "TEXT"), ("", "uzenet", "TEXT")])
    ki = entity(11.0, 8.6, "kimeneti_fajl", [
        ("PK", "id", "INTEGER"), ("FK", "futtatas_id", "INTEGER"), ("", "csoport_kulcs", "TEXT"),
        ("", "csoport_nev", "TEXT"), ("", "fajlnev", "TEXT"), ("", "sorok_szama", "INTEGER"),
        ("", "kesesben", "INTEGER"), ("", "osszeg_dij", "INTEGER")])
    bea = entity(0.3, 4.6, "beallitas", [
        ("PK", "kulcs", "TEXT"), ("", "ertek", "TEXT"), ("FK", "modositotta", "INTEGER"), ("", "modositva", "TEXT")])

    exp_cols = ["kolcsonzes_id", "olvaso_id", "olvaso_nev", "olvaso_szul_ev", "olvaso_tipus", "leltari_szam",
                "isbn", "cim", "szerzo", "kiado", "kiadas_eve", "kategoria", "fiok_kod", "fiok_nev",
                "kolcsonzes_datum", "lejarat_datum", "visszahozas_datum", "hosszabbitasok_szama", "konyvtaros_id",
                "megjegyzes"]
    x0, y0 = 5.6, 3.7
    ax.add_patch(FancyBboxPatch((x0, y0 - 3.45), 9.0, 3.45, boxstyle="round,pad=0.02,rounding_size=0.06",
                                fc="#FAFBFC", ec=GREY, lw=1.2, ls="--"))
    ax.text(x0 + 4.5, y0 - 0.25, "Bemenő export (fájl, nem adatbázis-tábla) – 20 oszlop + 6 számított oszlop",
            ha="center", fontsize=9.5, fontweight="bold", color=GREY)
    for i, c in enumerate(exp_cols):
        ax.text(x0 + 0.2 + (i // 7) * 3.0, y0 - 0.65 - (i % 7) * 0.33, c, fontsize=8)
    ax.text(x0 + 0.2, y0 - 3.15, "+ kolcsonzesi_napok, keses_napok, kesedelmi_dij, statusz, ev_honap, korcsoport",
            fontsize=8, color=BLUE)

    def rel(p1, p2, l1, l2, rad=0.0):
        arrow(ax, p1, p2, color=PRIMARY, rad=rad)
        ax.text(p1[0] + 0.15, p1[1] + 0.15, l1, fontsize=9, color=PRIMARY, fontweight="bold")
        ax.text(p2[0] - 0.35, p2[1] + 0.15, l2, fontsize=9, color=PRIMARY, fontweight="bold")

    rel((fel[0] + fel[2], 8.0), (fut[0], 8.0), "1", "N")
    ax.text((fel[0] + fel[2] + fut[0]) / 2, 8.25, "indítja", fontsize=8, ha="center", color=GREY)
    rel((fut[0] + fut[2], 8.0), (ki[0], 8.0), "1", "N")
    ax.text((fut[0] + fut[2] + ki[0]) / 2, 8.25, "létrehozza", fontsize=8, ha="center", color=GREY)
    arrow(ax, (1.6, fel[1]), (1.6, 4.6), color=PRIMARY)
    ax.text(1.75, fel[1] - 0.25, "1", fontsize=9, color=PRIMARY, fontweight="bold")
    ax.text(1.75, 4.75, "N", fontsize=9, color=PRIMARY, fontweight="bold")
    ax.text(1.85, (fel[1] + 4.6) / 2, "módosítja", fontsize=8, color=GREY)
    arrow(ax, (7.4, fut[1]), (7.4, y0), color=GREY, ls="--")
    ax.text(7.55, (fut[1] + y0) / 2, "forrasfajl → a feldolgozott export", fontsize=8, color=GREY)
    finish(fig, ax, "er_diagram.png", (0, 15), (0, 8.8))


# ---------------------------------------------------------------------------------------------
def folyamat():
    fig, ax = plt.subplots(figsize=(15, 5.6))
    ax.text(0.1, 5.25, "Eredeti munkahelyi megoldás (több eszköz, kézi indítás)", fontsize=11, fontweight="bold",
            color=GREY)
    steps1 = ["SAP lekérdezés,\nkézi export\n(~90 000 sor, 20 oszlop)", "Power Automate\nfolyamat\n(fájl átvétele, indítás)",
              "VBA makró\n(formázás,\núj oszlopok)", "Office Script\n(szétbontás\ncégenként)", "Cégenkénti\nExcel-fájlok"]
    steps2 = ["Könyvtári rendszer\nexportja\n(.xlsx / .csv)", "[3.] Betöltés\nés validáció\n(hibás sorok kiszűrése)",
              "[4.] Formázás\n(6 számított oszlop,\nfejlécek, formátumok)", "[5.]–[6.] Szétbontás\nfiók / kategória /\nkönyv szerint",
              "Fiókonkénti fájlok\n+ összesítő\n+ napló (SQLite)"]
    xs = [1.4, 4.4, 7.4, 10.4, 13.4]
    for i, (x, t) in enumerate(zip(xs, steps1)):
        box(ax, x, 4.25, 2.5, 1.2, t, fc="#F2F2F2", ec=GREY, fs=8.5)
        if i:
            arrow(ax, (xs[i - 1] + 1.27, 4.25), (x - 1.27, 4.25), color=GREY)
    ax.text(0.1, 2.55, "Kölcsönzési Export Feldolgozó (egy alkalmazás, egy kattintásos folyamat)", fontsize=11,
            fontweight="bold", color=PRIMARY)
    for i, (x, t) in enumerate(zip(xs, steps2)):
        box(ax, x, 1.5, 2.5, 1.25, t, fc=LIGHT if 0 < i < 4 else "white", fs=8.5)
        if i:
            arrow(ax, (xs[i - 1] + 1.27, 1.5), (x - 1.27, 1.5), color=PRIMARY)
    finish(fig, ax, "folyamat.png", (0, 15), (0.6, 5.5))


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    funkciofa()
    bejarhatosag()
    er_diagram()
    folyamat()
