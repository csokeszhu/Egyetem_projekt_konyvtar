"""A terv (docs/KEF_terv.docx) összerakása python-docx-szel – egyszerű, alap Word-formázással.

Előtte: python tools/screenshots.py  és  python docs/diagrams.py
Utána:  python docs/build_docs.py   (PDF: python docs/make_pdf.py)
"""
from __future__ import annotations

import sqlite3
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from kef import config  # noqa: E402
from kef.pipeline import TransformSettings, load_file, transform  # noqa: E402

DOCS = ROOT / "docs"
SHOTS = DOCS / "screenshots"
DIAG = DOCS / "diagrams"
DEMO_DB = DOCS / "_demo" / "kef_demo.db"
OUT = DOCS / "KEF_terv.docx"

doc = Document()


# ---------------------------------------------------------------- segédfüggvények
def p(text: str = "", bold: str | None = None):
    par = doc.add_paragraph()
    if bold:
        par.add_run(bold).bold = True
    par.add_run(text)
    return par


def bullets(items: list[str]) -> None:
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


def table(headers: list[str], rows: list[list], widths: list[float] | None = None, size: float = 10) -> None:
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = ""
        t.rows[0].cells[i].paragraphs[0].add_run(h).bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            if widths:
                c.width = Cm(widths[i])
            for par in c.paragraphs:
                for r in par.runs:
                    r.font.size = Pt(size)
    doc.add_paragraph()


def image(path: Path, caption: str, width: float = 16) -> None:
    doc.add_picture(str(path), width=Cm(width))
    p(caption).runs[0].italic = True


def func(name: str, rows: list[tuple[str, str]]) -> None:
    """Egy funkció leírása a mintamegoldás szerinti formában."""
    p(bold=name)
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    for label, text in rows:
        c = t.add_row().cells
        c[0].text = ""
        c[0].paragraphs[0].add_run(label).bold = True
        c[1].text = text
        c[0].width, c[1].width = Cm(3.5), Cm(13)
    doc.add_paragraph()


# ---------------------------------------------------------------- tartalom
def cover() -> None:
    doc.add_heading("Kölcsönzési Export Feldolgozó (KEF)", level=0)
    p("IT modellezés és alkalmazásfejlesztés – gyakorlati feladat")
    p("Készítették: ……………………, ……………………, ……………………, ……………………")
    p(f"{date.today():%Y. %m.}")


def chapter1() -> None:
    doc.add_heading("1. Feladatspecifikáció", level=1)
    p("Egy városi könyvtárhálózatnak van egy központi könyvtára és 11 fiókkönyvtára. A könyvtári rendszerből "
      "havonta kiexportálják az összes kölcsönzést egy nagy Excel-táblázatba (kb. 90 000 sor, 20 oszlop). "
      "A fiókvezetők viszont csak a saját fiókjuk adatait szeretnék látni: ki mit vitt el, mi van még kint, "
      "kinek mennyi késedelmi díja van. Eddig ezt valaki kézzel szűrte és mentette el fiókonként, ami több órás "
      "munka volt, és gyakran becsúszott hiba.")
    p("Az ötlet egy valós munkahelyi feladatból jött: ott egy SAP-ból letöltött, hasonló méretű listát kellett "
      "megformázni (új oszlopokat hozzáadni), majd cégenként külön fájlokba szétbontani. Ott ezt Power Automate, "
      "egy VBA-makró és egy Office Script együtt oldotta meg. Mi ugyanezt a feladatot egy könyvtárra írtuk át, "
      "és egyetlen programmal oldjuk meg.")
    p("A program feladata röviden:")
    bullets([
        "beolvassa az exportot (.xlsx vagy .csv) és kiszűri a hibás sorokat,",
        "hozzáad néhány számított oszlopot (kölcsönzési napok, késés, késedelmi díj, státusz, év-hónap, korcsoport),",
        "fiókonként (vagy kategóriánként, könyvenként) külön, formázott Excel-fájlokba bontja az adatokat,",
        "minden futtatást elment, hogy később vissza lehessen nézni.",
    ])

    doc.add_heading("Megoldási alternatívák", level=2)
    table(["Megoldás", "Előny", "Hátrány"], [
        ["Kézzel Excelben", "nem kell programozni", "lassú, sok hiba, minden hónapban újra"],
        ["Egy VBA-makró", "gyorsan megírható", "nehéz karbantartani, nincs jogosultság és napló"],
        ["Power Automate + Office Script + VBA (a munkahelyi megoldás)", "felhőben ütemezhető",
         "három külön eszköz, licenc kell, nehéz hibát keresni"],
        ["Saját Python-program (ezt választottuk)", "egy helyen minden, gyors, tesztelhető, ingyenes",
         "meg kell írni, Python kell hozzá"],
    ], [5.5, 5, 6])
    image(DIAG / "folyamat.png", "Az eredeti munkahelyi folyamat és a mi programunk összehasonlítása")

    doc.add_heading("Felhasználók", level=2)
    table(["Szerepkör", "Ki?", "Mit tehet?"], [
        ["Adminisztrátor", "a könyvtár informatikusa", "mindent; szabályok és felhasználók kezelése"],
        ["Feldolgozó", "aki a havi exportot feldolgozza", "betöltés, formázás, szétbontás, előzmények"],
        ["Megtekintő", "fiókvezető", "csak az előzményeket és a fájlokat nézheti"],
    ], [3.5, 5, 8])

    doc.add_heading("Szakterületi követelmények", level=2)
    bullets([
        "A kölcsönzési idő 28 nap, legfeljebb 5-ször hosszabbítható.",
        "Késedelmi díj: a lejárat után napi 50 Ft, de kölcsönzésenként legfeljebb 3000 Ft. Diákok és "
        "nyugdíjasok 50% kedvezményt kapnak. Ezeket az admin módosíthatja.",
        "A még kint lévő könyvek késését egy viszonyítási napig (alapból a mai nap) számoljuk.",
        "12 fiók van (F01–F12); ismeretlen fiókkódú sor nem kerülhet egyik fiók fájljába sem.",
        "Olvasótípus csak diák, felnőtt vagy nyugdíjas lehet.",
        "A fájlokban olvasói nevek vannak (személyes adat), ezért minden fiók csak a saját adatait kapja meg.",
    ])

    doc.add_heading("Funkcionális követelmények", level=2)
    bullets([
        "Bejelentkezés azonosítóval és jelszóval, szerepkör szerinti menü.",
        "Export beolvasása (.xlsx, .csv), oszlopok és sorok ellenőrzése, hibás sorok listája.",
        "Előnézet a beolvasott adatokról.",
        "6 számított oszlop hozzáadása; a díjszabályokat az admin állíthatja.",
        "Szétbontás fiók / kategória / könyv szerint, választható időszakkal, mappával és fájlnévvel.",
        "Összesítő fájl és a hibás sorok listája.",
        "A feldolgozás megszakítható.",
        "Futtatások naplózása és visszakeresése.",
        "Felhasználók kezelése (admin), saját jelszó módosítása.",
    ])

    doc.add_heading("Nem funkcionális követelmények", level=2)
    bullets([
        "A 90 000 soros fájl feldolgozása legyen meg 1 percen belül (nálunk kb. 40 mp).",
        "A program ne fagyjon le feldolgozás közben.",
        "Windows 10/11-en fusson, internet nélkül is.",
        "A jelszavakat ne simán tároljuk, hanem hash-elve.",
        "Ha a feldolgozás megszakad, ne maradjon félkész fájl.",
        "Magyar nyelvű felület, mentés és indítás előtt rákérdez.",
    ])


def chapter2() -> None:
    doc.add_heading("2. Funkcionális terv", level=1)
    p("A program fő funkcióit az alábbi fa mutatja. Alatta az egyes funkciók leírása következik; a [szám] az "
      "ablak sorszáma a 3. fejezetben.")
    image(DIAG / "funkciofa.png", "Funkciófa")

    func("Bejelentkezés – [1.]", [
        ("Bemenő adatok", "Felhasználói azonosító (3–30 karakter, betű, szám, pont, kötőjel; kötelező)\n"
                          "Jelszó (8–30 karakter, betű és szám is kell bele; kötelező, pontokkal jelenik meg)"),
        ("Kimenő adatok", "Sikeres belépésnél megnyílik a Főmenü [2.]. Hibás adatnál: „Helytelen felhasználói "
                          "azonosító és/vagy jelszó!”"),
        ("Mikor hívható", "A program ezzel indul, és kijelentkezés után ide jön vissza."),
        ("Leírás", "Beírja az azonosítót és a jelszót, majd a Bejelentkezés gombra kattint (vagy Entert nyom). "
                   "A Kilépés gomb bezárja a programot. Elfelejtett jelszót az admin tud visszaállítani."),
    ])
    func("Kijelentkezés, Kilépés", [
        ("Bemenő adatok", "Kijelentkezés / Kilépés gomb vagy az ablak [X] gombja."),
        ("Kimenő adatok", "Megerősítés után vissza a Bejelentkezéshez, illetve a program bezárul."),
        ("Mikor hívható", "Főmenü [2.]; az [X] minden ablakban."),
        ("Mellékhatás", "A betöltött, de fel nem dolgozott adatok elvesznek."),
    ])
    func("Adatbetöltés – [3.]", [
        ("Bemenő adatok", "A forrásfájl (Tallózás… gomb, vagy a Mintaadat gomb a mellékelt 90 000 soros "
                          "mintához). Csak .xlsx és .csv."),
        ("Kimenő adatok", "Összes / érvényes / hibás sorok száma, fiókok száma, időszak; előnézet az első 100 "
                          "sorról; a hibás sorok listája (sorszám, azonosító, hiba oka). Ha hiányzik egy oszlop: "
                          "„Hiányzó oszlop(ok): …”"),
        ("Mikor hívható", "Főmenü → Új feldolgozás (csak admin és feldolgozó)."),
        ("Leírás", "A program soronként ellenőrzi: kötelező mezők nem üresek, a dátumok jók, a lejárat és a "
                   "visszahozás nem a kölcsönzés előtti, nincs kétszer ugyanaz az azonosító, a fiókkód és az "
                   "olvasótípus létezik. A hibás sorok kimaradnak, de az összesítőbe bekerülnek."),
    ])
    func("Formázási szabályok – [4.]", [
        ("Bemenő adatok", "Díj/nap (alapból 50 Ft), díjplafon (3000 Ft), kedvezmény (50%), viszonyítási "
                          "dátum, melyik számított oszlop kell (6 jelölőnégyzet)."),
        ("Kimenő adatok", "A beállítások az aktuális feldolgozásra érvényesek; az admin elmentheti őket "
                          "alapértelmezettnek."),
        ("Mikor hívható", "A [3.] Tovább gombjával, vagy a Főmenüből (módosítani csak az admin tud)."),
    ])
    func("Formázás (számított oszlopok)", [
        ("Bemenő adatok", "Az érvényes sorok és a formázási szabályok."),
        ("Kimenő adatok", "6 új oszlop: kölcsönzési napok, késés, késedelmi díj, státusz, év-hónap, korcsoport "
                          "(a képletek a 4. fejezetben)."),
        ("Mikor hívható", "A feldolgozás indításakor automatikusan."),
        ("Leírás", "A munkahelyi megoldásban ezt csinálta a VBA-makró."),
    ])
    func("Szétbontás beállítása – [5.]", [
        ("Bemenő adatok", "Szempont (fiók / kategória / könyv), időszak (nem kötelező), kimeneti mappa, "
                          "fájlnév-minta (pl. KEF_{idoszak}_{kulcs}_{nev}), kell-e összesítő."),
        ("Kimenő adatok", "Előnézet: hány fájl lesz, mi lesz a nevük, hány sor kerül bele."),
        ("Mikor hívható", "A [4.] Tovább gombjával."),
    ])
    func("Feldolgozás indítása – [5.] → [10.] → [6.] → [7.]", [
        ("Bemenő adatok", "A betöltött adatok és a beállítások."),
        ("Kimenő adatok", "Rákérdez, hogy biztosan indítja-e [10.], utána a [6.] ablakban látszik a haladás. "
                          "A végén az Eredmény [7.] ablak mutatja a létrejött fájlokat."),
        ("Mellékhatás", "Létrejönnek a fájlok (az azonos nevűek felülíródnak), a futtatás bekerül a naplóba."),
        ("Leírás", "Minden fiók kap egy fájlt egy „Kölcsönzések” és egy „Összesítő” lappal; a késésben lévő "
                   "sorok pirosak. A munkahelyi megoldásban ezt csinálta az Office Script."),
    ])
    func("Megszakítás – [6.]", [
        ("Bemenő adatok", "Megszakítás gomb, megerősítés."),
        ("Kimenő adatok", "Az addig elkészült fájlok törlődnek, vissza a [5.] ablakhoz."),
    ])
    func("Előzmények – [8.]", [
        ("Bemenő adatok", "Szűrés dátum (-tól, -ig) és felhasználó szerint (nem kötelező)."),
        ("Kimenő adatok", "A korábbi futtatások listája, és a kijelölt futtatás fájljai."),
        ("Mikor hívható", "Főmenü → Futtatási előzmények (mindenki)."),
    ])
    func("Jelszó módosítása és felhasználókezelés – [9.]", [
        ("Bemenő adatok", "Régi jelszó, új jelszó kétszer. Adminnak még: új felhasználó adatai, szerepkör, "
                          "aktív-e."),
        ("Kimenő adatok", "„Az új jelszó mentése megtörtént.” vagy hibaüzenet (rossz régi jelszó, nem egyezik "
                          "a kettő, túl rövid)."),
        ("Mikor hívható", "Főmenü → Felhasználók és jelszó. A felhasználókezelő rész csak adminnak látszik."),
    ])


SCREENS = [
    ("01_bejelentkezes", "[1.] Bejelentkezés", "Azonosító és jelszó megadása."),
    ("02_fomenu", "[2.] Főmenü", "Innen érhető el minden; a szerepkörtől függ, melyik gomb aktív."),
    ("03_adatbetoltes", "[3.] Adatbetöltés", "Fájl kiválasztása, betöltés, előnézet és a hibás sorok fül."),
    ("04_formazasi_szabalyok", "[4.] Formázási szabályok", "Díjszabályok és a hozzáadandó oszlopok."),
    ("05_szetbontas", "[5.] Szétbontás beállításai", "Szempont, időszak, mappa, fájlnév; jobb oldalt előnézet."),
    ("06_feldolgozas", "[6.] Feldolgozás folyamatban", "Folyamatjelző és napló, megszakítás gomb."),
    ("07_eredmeny", "[7.] Eredmény", "A létrejött fájlok listája; innen megnyitható a mappa."),
    ("08_elozmenyek", "[8.] Előzmények", "A korábbi futtatások szűrhető listája."),
    ("09_felhasznalok", "[9.] Felhasználók és jelszó", "Jelszócsere, adminnak felhasználókezelés is."),
    ("10_megerosites", "[10.] Megerősítés", "Közös „Biztosan…?” ablak mentés, indítás, kilépés előtt."),
]


def chapter3() -> None:
    doc.add_heading("3. Képernyőtervek", level=1)
    p("A programnak 10 ablaka van. A képek a kész programról készültek.")
    for name, title, text in SCREENS:
        doc.add_heading(title, level=2)
        p(text)
        doc.add_picture(str(SHOTS / f"{name}.png"), width=Cm(9 if name == "10_megerosites" else 15))
    doc.add_heading("Az ablakok bejárhatósága", level=2)
    p("A nyilakon az a gomb van, amivel át lehet menni egyik ablakból a másikba. A kék nyilakat csak az admin "
      "és a feldolgozó használhatja.")
    image(DIAG / "bejarhatosag.png", "Bejárhatósági gráf")


def sample_rows():
    lr = load_file(config.SAMPLE_DIR / "kolcsonzesek_export.csv")
    df = transform(lr.valid, TransformSettings(viszonyitasi_datum=date(2026, 11, 15)))
    picks = [
        df[(df["statusz"] == "visszahozva") & (df["keses_napok"] == 0)].iloc[0],
        df[(df["statusz"] == "visszahozva") & (df["keses_napok"] > 5) & (df["olvaso_tipus"] == "felnőtt")].iloc[0],
        df[(df["statusz"] == "visszahozva") & (df["keses_napok"] > 5) & (df["olvaso_tipus"] == "diák")].iloc[0],
        df[df["statusz"] == "kint"].iloc[0],
        df[df["statusz"] == "késésben"].iloc[-1],
    ]
    return lr, picks


def d(ts) -> str:
    return "" if ts is None or str(ts) == "NaT" else ts.strftime("%Y.%m.%d.")


def chapter4() -> None:
    doc.add_heading("4. Adatmodell", level=1)
    p("Két fajta adatunk van: a bemenő export (egy Excel-fájl, ezt csak olvassuk), és a program saját kis "
      "SQLite-adatbázisa (felhasználók, beállítások, futtatások, létrejött fájlok).")
    image(DIAG / "er_diagram.png", "Az adatmodell")

    doc.add_heading("A bemenő export oszlopai", level=2)
    kotelezo = set(config.MANDATORY_COLUMNS)
    typ = {"olvaso_szul_ev": "egész", "kiadas_eve": "egész", "hosszabbitasok_szama": "egész",
           "kolcsonzes_datum": "dátum", "lejarat_datum": "dátum", "visszahozas_datum": "dátum"}
    table(["Oszlop", "Típus", "Kötelező"],
          [[c, typ.get(c, "szöveg"), "igen" if c in kotelezo else "nem"] for c in config.REQUIRED_COLUMNS],
          [6, 3, 3], size=9)

    doc.add_heading("Számított oszlopok", level=2)
    table(["Oszlop", "Hogyan számoljuk"], [
        ["kolcsonzesi_napok", "visszahozás (vagy viszonyítási nap) − kölcsönzés napja"],
        ["keses_napok", "visszahozás (vagy viszonyítási nap) − lejárat, ha pozitív, különben 0"],
        ["kesedelmi_dij", "késés × 50 Ft, max. 3000 Ft; diáknak, nyugdíjasnak a fele"],
        ["statusz", "visszahozva / késésben (még kint és lejárt) / kint"],
        ["ev_honap", "a kölcsönzés éve és hónapja, pl. 2025-02"],
        ["korcsoport", "0–14, 15–18, 19–35, 36–64, 65+"],
    ], [4.5, 12])

    doc.add_heading("Mintaadatok", level=2)
    lr, picks = sample_rows()
    p("Néhány sor a 90 000 soros mintából, a számított oszlopokkal (viszonyítási nap: 2026.11.15.):")
    table(["Azonosító", "Típus", "Cím", "Fiók", "Kölcsönzés", "Lejárat", "Vissza", "Késés", "Díj", "Státusz"],
          [[r["kolcsonzes_id"], r["olvaso_tipus"], r["cim"], r["fiok_kod"], d(r["kolcsonzes_datum"]),
            d(r["lejarat_datum"]), d(r["visszahozas_datum"]), r["keses_napok"], r["kesedelmi_dij"], r["statusz"]]
           for r in picks], size=8)
    p("Példák kiszűrt hibás sorokra:")
    errs = lr.errors.drop_duplicates("hiba").head(5)
    table(["Sor", "Azonosító", "Hiba"], [[r.sor, r.kolcsonzes_id or "(üres)", r.hiba] for r in errs.itertuples()],
          [2, 3.5, 11], size=9)

    if DEMO_DB.exists():
        con = sqlite3.connect(DEMO_DB)
        con.row_factory = sqlite3.Row
        p("A felhasznalo tábla:")
        table(["id", "azonosito", "nev", "szerepkor", "aktiv"],
              [[u["id"], u["azonosito"], u["nev"], u["szerepkor"], u["aktiv"]]
               for u in con.execute("SELECT * FROM felhasznalo ORDER BY id")], size=9)
        p("A futtatas tábla (néhány oszlop):")
        table(["id", "felhasznalo_id", "inditas", "sorok_be", "hibas_sorok", "szempont", "statusz"],
              [[r["id"], r["felhasznalo_id"], r["inditas"], r["sorok_be"], r["hibas_sorok"], r["szempont"],
                r["statusz"]] for r in con.execute("SELECT * FROM futtatas ORDER BY id")], size=9)
        p("A kimeneti_fajl tábla (első 3 sor):")
        table(["id", "futtatas_id", "csoport", "fajlnev", "sorok", "osszeg_dij"],
              [[f["id"], f["futtatas_id"], f["csoport_kulcs"], f["fajlnev"], f["sorok_szama"], f["osszeg_dij"]]
               for f in con.execute("SELECT * FROM kimeneti_fajl ORDER BY id LIMIT 3")], size=8)
        con.close()


TESTS = [
    ("1", "bejelentkezés", "[1.]", "Bejelentkezés", "feldolgozo / Feldolgozo1",
     "Megnyílik a Főmenü: „Üdvözlünk, Kovács Anna!”"),
    ("2", "bejelentkezés", "[1.]", "Bejelentkezés", "feldolgozo / feldolgozo1 (kisbetűs jelszó)",
     "Hiba: „Helytelen felhasználói azonosító és/vagy jelszó!”"),
    ("3", "adatbetöltés", "[3.]", "Betöltés", "olyan fájl, amiből hiányzik a „cim” oszlop",
     "Hiba: „Hiányzó oszlop(ok): cim”, a Tovább gomb inaktív"),
    ("4", "adatbetöltés", "[3.]", "Betöltés", "4 sor, a 3. sorban üres az olvaso_id",
     "3 érvényes, 1 hibás sor: 4. sor, „Kötelező mező üres: olvaso_id”"),
    ("5", "késedelmi díj", "[4.]→[6.]", "Tovább, Indítás",
     "lejárat 2025.03.01., visszahozás 2025.03.11.; a) felnőtt b) diák",
     "a) 10 nap késés, 500 Ft  b) 10 nap késés, 250 Ft"),
    ("6", "díjplafon", "[4.]→[6.]", "Tovább, Indítás",
     "lejárat 2025.01.01., nincs visszahozva, viszonyítási nap 2025.06.01.",
     "151 nap késés, díj 3000 Ft (plafon), státusz: késésben"),
    ("7", "szétbontás", "[5.]→[7.]", "Feldolgozás indítása", "10 sor: F01-ből 4, F02-ből 3, F03-ból 3",
     "3 fájl + összesítő, a sorok száma 4 + 3 + 3 = 10"),
    ("8", "adatbetöltés", "[3.]", "Betöltés", "2 sor, a 2.-ban fiókkód: F99",
     "1 érvényes, 1 hibás: „Ismeretlen fiókkód: F99”"),
    ("9", "szétbontás", "[5.]", "Feldolgozás indítása", "olyan kimeneti mappa, ami nem hozható létre",
     "Hiba: „A kimeneti mappa nem írható: …”, nem indul el"),
    ("10", "jogosultság", "[2.]", "Új feldolgozás", "megtekinto felhasználóval belépve",
     "Az Új feldolgozás gomb inaktív, nem lehet feldolgozást indítani"),
]


def chapter5() -> None:
    doc.add_heading("5. Tesztelési terv", level=1)
    p("A teszteseteket a feldolgozás fő részeire (betöltés, díjszámítás, szétbontás) és a jogosultságokra "
      "írtuk. Mind a 10-et automatikus tesztként (pytest) is megírtuk, és mind lefut hiba nélkül.")
    table(["№", "Funkció", "Ablak", "Akciógomb", "Bemenő adatok", "Tervezett kimenő adatok"],
          [list(t) for t in TESTS], [0.8, 2.3, 1.8, 2.6, 4.6, 4.9], size=9)


def main() -> None:
    cover()
    chapter1()
    chapter2()
    chapter3()
    chapter4()
    chapter5()
    doc.save(OUT)
    print("Mentve:", OUT)


if __name__ == "__main__":
    main()
