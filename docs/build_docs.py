"""A teljes dokumentáció (KEF_dokumentacio.docx) előállítása python-docx-szel.

Előfeltétel: python tools/screenshots.py (képernyőképek + bemutató adatbázis) és
python docs/diagrams.py (ábrák). Utána:

    python docs/build_docs.py
"""
from __future__ import annotations

import sqlite3
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from kef import config  # noqa: E402
from kef.pipeline import TransformSettings, load_file, transform  # noqa: E402

DOCS = ROOT / "docs"
SHOTS = DOCS / "screenshots"
DIAG = DOCS / "diagrams"
DEMO_DB = DOCS / "_demo" / "kef_demo.db"
OUT = DOCS / "KEF_dokumentacio.docx"

PRIMARY = RGBColor(0x1F, 0x4E, 0x78)
HEAD_FILL = "1F4E78"
LABEL_FILL = "DCE8F4"
ZEBRA_FILL = "F4F7FB"

doc = Document()
_fig = [0]
_tab = [0]


# ============================================================ alap segédfüggvények
def setup() -> None:
    st = doc.styles
    normal = st["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    for name, size, before, after in [("Heading 1", 18, 18, 10), ("Heading 2", 14, 14, 6), ("Heading 3", 12, 10, 4)]:
        h = st[name]
        h.font.name = "Calibri"
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = PRIMARY
        rfonts = h.element.rPr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            h.element.rPr.append(rfonts)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(a), "Calibri")
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            if rfonts.get(qn(a)) is not None:
                del rfonts.attrib[qn(a)]
        h.paragraph_format.space_before = Pt(before)
        h.paragraph_format.space_after = Pt(after)
        h.paragraph_format.keep_with_next = True
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2.2)
        s.top_margin = s.bottom_margin = Cm(2.0)


def shade(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def repeat_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def no_split(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_field(paragraph, instr: str, placeholder: str = "") -> None:
    run = paragraph.add_run()
    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    t = OxmlElement("w:instrText")
    t.set(qn("xml:space"), "preserve")
    t.text = instr
    s = OxmlElement("w:fldChar")
    s.set(qn("w:fldCharType"), "separate")
    run._r.extend([b, t, s])
    paragraph.add_run(placeholder)
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(e)


HEADINGS: list[tuple[int, str]] = []
TOC_MARKER = []
PAGES: dict[str, int] = {}


def h1(text: str) -> None:
    doc.add_heading(text, level=1)
    HEADINGS.append((1, text))


def h2(text: str) -> None:
    doc.add_heading(text, level=2)
    HEADINGS.append((2, text))


def build_toc() -> None:
    """Statikus tartalomjegyzék a TOC_MARKER helyére (oldalszámok a docs/_toc_pages.json-ból).

    A Word saját TOC-mezőjének frissítése automatizálva lefagyhat, ezért a make_pdf.py két
    menetben dolgozik: első export → oldalszámok kiolvasása a PDF-ből → újragenerálás.
    """
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

    marker = TOC_MARKER[0]
    for level, text in HEADINGS:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.tab_stops.add_tab_stop(Cm(16.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        pf.left_indent = Cm(0 if level == 1 else 0.8)
        pf.space_before = Pt(8 if level == 1 else 0)
        pf.space_after = Pt(2)
        r = p.add_run(f"{text}\t{PAGES.get(text, '')}")
        r.bold = level == 1
        r.font.size = Pt(11 if level == 1 else 10.5)
        marker._p.addprevious(p._p)
    marker._p.getparent().remove(marker._p)


def h3(text: str) -> None:
    doc.add_heading(text, level=3)


def para(text: str = "", bold_prefix: str | None = None, italic=False, size=None, align=None, after=None):
    p = doc.add_paragraph()
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if align:
        p.alignment = align
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    return p


def bullets(items: list[str], style="List Bullet") -> None:
    for it in items:
        p = doc.add_paragraph(style=style)
        if "::" in it:
            head, tail = it.split("::", 1)
            p.add_run(head).bold = True
            p.add_run(tail)
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(2)


def figure(path: Path, caption: str, width_cm: float = 16.5) -> None:
    doc.add_picture(str(path), width=Cm(width_cm))
    doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.paragraphs[-1].paragraph_format.keep_with_next = True
    _fig[0] += 1
    p = para(f"{_fig[0]}. ábra – {caption}", italic=True, size=9, align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=10)
    p.runs[0].font.color.rgb = RGBColor(0x5F, 0x6B, 0x7A)


def caption_table(text: str) -> None:
    _tab[0] += 1
    p = para(f"{_tab[0]}. táblázat – {text}", italic=True, size=9, after=3)
    p.paragraph_format.keep_with_next = True
    p.runs[0].font.color.rgb = RGBColor(0x5F, 0x6B, 0x7A)


def _cell_text(cell, text: str, bold=False, size=9.5, color=None) -> None:
    cell.text = ""
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(line)
        r.bold = bold
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color


def table(headers: list[str], rows: list[list], widths_cm: list[float], size=9.5, zebra=True, caption=None):
    if caption:
        caption_table(caption)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        _cell_text(c, h, bold=True, size=size, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, HEAD_FILL)
    repeat_header(t.rows[0])
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            _cell_text(cells[i], v, size=size)
            if zebra and ri % 2 == 1:
                shade(cells[i], ZEBRA_FILL)
        no_split(t.rows[-1])
    for row in t.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def func_spec(title: str, subtitle: str, items: list[tuple[str, str]]) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = PRIMARY
    p.add_run(f"  ({subtitle})").italic = True
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.autofit = False
    for label, text in items:
        cells = t.add_row().cells
        _cell_text(cells[0], label, bold=True, size=9.5)
        shade(cells[0], LABEL_FILL)
        _cell_text(cells[1], text, size=9.5)
    for row in t.rows:
        row.cells[0].width = Cm(3.6)
        row.cells[1].width = Cm(13.0)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def page_break() -> None:
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def footer_page_numbers() -> None:
    for s in doc.sections:
        p = s.footer.paragraphs[0]
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.text = ""
        r = p.add_run(f"{config.APP_SHORT} – {config.APP_NAME}   |   ")
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x8A, 0x94, 0xA3)
        add_field(p, "PAGE", "1")


def new_section(landscape: bool) -> None:
    from docx.enum.section import WD_SECTION

    s = doc.add_section(WD_SECTION.NEW_PAGE)
    if landscape:
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = Cm(29.7), Cm(21)
    else:
        s.orientation = WD_ORIENT.PORTRAIT
        s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2.0)


# ============================================================ adatok a mintaadatból és a bemutató DB-ből
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


def demo_db():
    if not DEMO_DB.exists():
        raise SystemExit("Hiányzik a bemutató adatbázis – előbb futtasd: python tools/screenshots.py")
    con = sqlite3.connect(DEMO_DB)
    con.row_factory = sqlite3.Row
    return con


def fmt(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def d(ts) -> str:
    return "" if ts is None or str(ts) == "NaT" else ts.strftime("%Y.%m.%d.")


# ============================================================ tartalom
def cover() -> None:
    for _ in range(5):
        doc.add_paragraph()
    p = para(config.APP_NAME, size=30, align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=2)
    p.runs[0].bold = True
    p.runs[0].font.color.rgb = PRIMARY
    para(f"({config.APP_SHORT})", size=16, align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=24)
    para("Városi könyvtárhálózat kölcsönzési exportjainak feldolgozása és fiókkönyvtáranként való szétbontása",
         size=13, align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=40)
    para("IT modellezés és alkalmazásfejlesztés – gyakorlati feladat", size=12,
         align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=4)
    para("Tervdokumentáció és megvalósítás", size=12, align=WD_PARAGRAPH_ALIGNMENT.CENTER, after=60)
    t = doc.add_table(rows=4, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate([("Csoport tagjai:", "……………………………………"), ("", "……………………………………"),
                                ("Gyakorlatvezető:", "……………………………………"),
                                ("Beadás:", f"{date.today():%Y. %m.}")]):
        _cell_text(t.rows[i].cells[0], k, bold=True, size=11)
        _cell_text(t.rows[i].cells[1], v, size=11)
        t.rows[i].cells[0].width = Cm(4.5)
        t.rows[i].cells[1].width = Cm(7.5)
    page_break()
    h = doc.add_paragraph()
    r = h.add_run("Tartalomjegyzék")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = PRIMARY
    TOC_MARKER.append(doc.add_paragraph())
    page_break()


def chapter1() -> None:
    h1("1. Feladatspecifikáció")
    h2("1.1. Háttér és a megoldandó probléma")
    para("A feladat ötlete egy valós munkahelyi feladatból származik. Egy vállalatnál havonta egy kb. 90 000 soros, "
         "20 oszlopos listát kellett az SAP-ból exportálni, azt meg kellett formázni (új, számított oszlopok "
         "hozzáadása, fejlécek, formátumok), majd cégenként szétbontani külön Excel-fájlokba, hogy minden "
         "leányvállalat csak a saját adatait kapja meg. A folyamatot végül Power Automate, Office Script és "
         "VBA-makró kombinációjával automatizáltuk. A dokumentációban ugyanezt a problémát egy, a gyakorlati "
         "feladatlistában is szereplő területre – könyvek kölcsönzése, könyvtár – vetítjük át.")
    para("A Városi Könyvtárhálózat (VK) egy központi könyvtárból és 11 fiókkönyvtárból áll (a valós példában ezek "
         "a „cégek”). Az integrált könyvtári rendszer – a valós példában az SAP szerepét tölti be – havonta "
         "exportálja az elmúlt időszak összes kölcsönzését egyetlen nagy táblázatba (kb. 90 000 sor, 20 oszlop). "
         "A fiókvezetőknek a saját fiókjuk forgalmára van szükségük: ki mit kölcsönzött, mi van még kint, kinek "
         "mennyi késedelmi díja keletkezett. Eddig egy munkatárs kézzel szűrte, formázta és mentette el a "
         "fiókonkénti fájlokat, ami havonta több órát vett igénybe, és gyakran előfordultak hibák (rossz szűrés, "
         "kimaradt sorok, eltérő formátumok, hibás díjszámítás).")
    h2("1.2. A rendszer célja")
    para("A Kölcsönzési Export Feldolgozó (KEF) egy asztali alkalmazás, amely a könyvtári rendszer exportját "
         "beolvassa és ellenőrzi, a hibás sorokat kiszűri, a hasznos számított oszlopokkal (kölcsönzési napok, "
         "késés, késedelmi díj, státusz, év-hónap, korcsoport) kiegészíti, egységesen formázza, majd egyetlen "
         "folyamatban fiókkönyvtáranként (vagy kategóriánként, könyvenként) külön Excel-fájlokba bontja. "
         "Minden futtatást naplóz, így később visszakereshető, ki, mikor, milyen forrásból, milyen eredménnyel "
         "dolgozott. A cél a havi több órás kézi munka néhány percre csökkentése és a hibák kiküszöbölése.")
    h2("1.3. Megoldási alternatívák")
    para("A probléma többféleképpen megoldható; az alábbi táblázat a mérlegelt lehetőségeket hasonlítja össze.")
    table(["Alternatíva", "Előnyök", "Hátrányok"], [
        ["Kézi feldolgozás Excelben (szűrés, másolás, mentés)", "Nem kell fejleszteni; bárki el tudja végezni.",
         "Havonta több óra; sok hibalehetőség; nem reprodukálható; 90 000 sornál lassú."],
        ["Egyetlen VBA-makró az Excelben", "Gyorsan elkészíthető, az Excelen belül fut.",
         "Nehezen karbantartható, nincs jogosultságkezelés és naplózás; makróbiztonsági figyelmeztetések; "
         "nagy adatnál lassú."],
        ["Power Automate + Office Script + VBA (az eredeti munkahelyi megoldás)",
         "Felhőben ütemezhető, integrálható e-maillel, SharePointtal.",
         "Három különböző eszköz, három hibaforrás; licencigény; nehéz hibakeresés; offline nem működik."],
        ["Saját asztali alkalmazás Pythonban (választott)",
         "Egy eszköz, egy kattintásos folyamat; validáció, jogosultság, naplózás; gyors (pandas); "
         "automatikusan tesztelhető; ingyenes.",
         "Fejlesztési ráfordítás; Python futtatókörnyezet szükséges a gépen."],
    ], [4.2, 6.2, 6.2], caption="A mérlegelt megoldási alternatívák")
    para("A választott megoldás a negyedik: a KEF egyetlen alkalmazásban fogja össze azt, amit a valós példában három "
         "eszköz végzett. Az alábbi ábra a két folyamatot veti össze.")
    figure(DIAG / "folyamat.png", "Az eredeti (SAP – Power Automate – VBA – Office Script) és a KEF folyamata")

    h2("1.4. Felhasználói körök és jogosultságok")
    table(["Szerepkör", "Ki tölti be?", "Mit tehet?"], [
        ["Adminisztrátor", "a központi könyvtár informatikusa",
         "minden funkció; formázási szabályok módosítása; felhasználók létrehozása, letiltása, jelszó visszaállítása"],
        ["Feldolgozó", "a havi exportot feldolgozó munkatárs",
         "adatbetöltés, formázás, szétbontás, előzmények, saját jelszó módosítása; a szabályokat csak megtekinti"],
        ["Megtekintő", "fiókvezető, igazgatóság",
         "futtatási előzmények és a létrejött fájlok megtekintése, saját jelszó; új feldolgozást nem indíthat"],
    ], [3.2, 4.4, 9.0], caption="Szerepkörök")

    h2("1.5. Szakterületi követelmények")
    bullets([
        "Kölcsönzési idő:: a kölcsönzési idő 28 nap, amely legfeljebb 5 alkalommal hosszabbítható (a lejárat "
        "a könyvtári rendszerben már a hosszabbításokkal együtt szerepel).",
        "Késedelmi díj:: a lejárat utáni minden megkezdett napra 50 Ft, kölcsönzésenként legfeljebb 3 000 Ft "
        "(plafon). Diák és nyugdíjas olvasók 50% kedvezményt kapnak. Az értékeket az adminisztrátor módosíthatja.",
        "Még kint lévő könyvek:: a késést egy viszonyítási dátumig (alapértelmezés: a feldolgozás napja) számoljuk; "
        "ha ekkor a lejárat már elmúlt, a státusz „késésben”, egyébként „kint”.",
        "Fiókkönyvtárak:: a hálózat 12 egységből áll (F01–F12); ismeretlen fiókkódú sor nem kerülhet egyetlen "
        "fiók fájljába sem.",
        "Olvasótípusok:: diák, felnőtt, nyugdíjas – más érték hibának minősül.",
        "Azonosítók:: a kölcsönzés-azonosító egyedi; az ismétlődő sorok közül csak az első érvényes.",
        "Adatvédelem (GDPR):: a fájlok olvasói nevet tartalmaznak, ezért csak a belső hálózaton tárolhatók; a "
        "fiókvezető csak a saját fiókja adatait kapja meg (ezt a szétbontás biztosítja); a program a forrásadatot "
        "nem módosítja és nem küldi tovább.",
    ])

    h2("1.6. Funkcionális követelmények")
    table(["Azonosító", "Követelmény"], [
        ["FK-01", "Bejelentkezés felhasználói azonosítóval és jelszóval; szerepkör szerinti menü."],
        ["FK-02", ".xlsx és .csv formátumú export beolvasása (a CSV elválasztóját a program felismeri)."],
        ["FK-03", "Fájlszintű ellenőrzés: kötelező oszlopok megléte, nem üres fájl, támogatott formátum."],
        ["FK-04", "Soronkénti validáció: kötelező mezők, dátumok, dátumsorrend, duplikált azonosító, ismeretlen "
                  "fiókkód és olvasótípus, hosszabbítások száma; hibalista a forrásfájl sorszámával."],
        ["FK-05", "Előnézet az első 100 érvényes sorról és összesítés (összes / érvényes / hibás sor, fiókok, időszak)."],
        ["FK-06", "Hat számított oszlop hozzáadása a beállított szabályok szerint; az oszlopok egyenként kikapcsolhatók."],
        ["FK-07", "Formázás: olvasható magyar fejlécek, dátumformátum, Ft-formátum, rögzített fejléc, szűrő, "
                  "oszlopszélesség, késésben lévő sorok színezése, összesítő munkalap."],
        ["FK-08", "Szétbontás fiókkönyvtár, kategória vagy könyv (ISBN) szerint, opcionális időszakszűréssel, "
                  "megadható kimeneti mappába és fájlnév-mintával; előnézet a létrejövő fájlokról."],
        ["FK-09", "Összesítő fájl a csoportokról és a hibás sorokról."],
        ["FK-10", "A feldolgozás megszakítható; megszakításkor a részleges fájlok törlődnek."],
        ["FK-11", "Minden futtatás és a létrejött fájlok naplózása; előzmények szűrése dátum és felhasználó szerint."],
        ["FK-12", "Felhasználókezelés (admin): létrehozás, szerepkör, letiltás, jelszó visszaállítása; "
                  "saját jelszó módosítása mindenkinek."],
    ], [2.4, 14.2], caption="Funkcionális követelmények")

    h2("1.7. Nem funkcionális követelmények")
    table(["Azonosító", "Leírás", "Cél-érték", "Elfogadható tartomány"], [
        ["NFK-01", "Teljesítmény: 90 000 soros export teljes feldolgozása (betöltés + formázás + szétbontás)",
         "< 45 mp", "< 60 mp (mért: kb. 37–45 mp)"],
        ["NFK-02", "A felület a feldolgozás alatt sem fagy le (háttérszál, folyamatjelző)", "mindig reagál", "–"],
        ["NFK-03", "Platform: Windows 10/11, Python 3.12, internetkapcsolat nélkül is működik", "offline", "–"],
        ["NFK-04", "Biztonság: jelszavak csak sózott PBKDF2-SHA256 hash formában tárolva", "120 000 iteráció",
         "≥ 100 000"],
        ["NFK-05", "Megbízhatóság: hiba vagy megszakítás esetén nem marad félkész kimenet", "0 félkész fájl", "–"],
        ["NFK-06", "Használhatóság: magyar nyelvű felület, minden mentés/indítás előtt megerősítés",
         "max. 6 lépés", "–"],
        ["NFK-07", "Karbantarthatóság: automatikus tesztek a fő funkciókra", "26 teszt", "≥ 10 teszt"],
    ], [2.0, 7.6, 3.0, 4.0], caption="Nem funkcionális követelmények")

    h2("1.8. Követelményjegyzék-bejegyzések (példák)")
    para("A követelmények pontosításához az előadáson bemutatott „Követelményjegyzék-bejegyzés” formalapot "
         "használtuk; két kitöltött példa:")
    for rec in [
        [("Projekt/rendszer", "KEF"), ("Követelmény AZ", "FK-08"), ("Forrás", "Központi könyvtár, feldolgozó munkatárs"),
         ("Prioritás", "kötelező (magas)"), ("Tulajdonos", "Városi Könyvtárhálózat igazgatósága"),
         ("Funkcionális követelmény", "A validált és formázott kölcsönzési adatokat fiókkönyvtáranként külön, "
                                      "formázott Excel-fájlba kell menteni, egységes fájlnévvel."),
         ("Nem funkcionális követelmény", "90 000 sor esetén legfeljebb 60 mp; megszakításkor nem maradhat félkész fájl."),
         ("Haszon", "A havi kb. 3–4 óra kézi munka néhány percre csökken; nincs kimaradt vagy rossz fiókba került sor."),
         ("Megjegyzések / javasolt megoldás", "pandas groupby + xlsxwriter; a kimeneti fájlnév mintából képezve."),
         ("Kapcsolódó követelmények", "FK-06, FK-07, FK-09, FK-10, NFK-01, NFK-05"),
         ("Megoldás", "Szétbontás funkció, [5.] és [6.] ablak (2.2. fejezet).")],
        [("Projekt/rendszer", "KEF"), ("Követelmény AZ", "FK-06"), ("Forrás", "Fiókvezetők kérése"),
         ("Prioritás", "kötelező"), ("Tulajdonos", "Központi könyvtár"),
         ("Funkcionális követelmény", "A program számolja ki minden kölcsönzés késését és késedelmi díját a "
                                      "szakterületi szabályok szerint (díj/nap, plafon, kedvezmény)."),
         ("Nem funkcionális követelmény", "A díjszabályokat csak adminisztrátor módosíthatja; a változás "
                                          "naplózott (ki, mikor)."),
         ("Haszon", "A fiókvezető azonnal látja a kintlévőségeket, nem kell kézzel számolni."),
         ("Megjegyzések / javasolt megoldás", "Vektorizált számítás pandasban; a beállítások a beallitas táblában."),
         ("Kapcsolódó követelmények", "FK-07, 1.5. szakterületi követelmények"),
         ("Megoldás", "Formázás funkció, [4.] ablak; T05 és T06 teszteset.")],
    ]:
        t = doc.add_table(rows=0, cols=2)
        t.style = "Table Grid"
        for k, v in rec:
            c = t.add_row().cells
            _cell_text(c[0], k, bold=True, size=9.5)
            shade(c[0], LABEL_FILL)
            _cell_text(c[1], v, size=9.5)
            c[0].width, c[1].width = Cm(4.6), Cm(12.0)
        doc.add_paragraph()

    h2("1.9. Fogalomszótár")
    table(["Fogalom", "Jelentés"], [
        ["Export", "A könyvtári rendszerből letöltött, egy időszak összes kölcsönzését tartalmazó táblázat."],
        ["Fiók(könyvtár)", "A hálózat egy egysége (F01–F12); a szétbontás alapértelmezett szempontja."],
        ["Érvényes / hibás sor", "A validáción átment, illetve fennakadt sor; a hibás sor nem kerül a fiókfájlokba."],
        ["Számított oszlop", "A program által a formázás során hozzáadott oszlop (pl. késedelmi díj)."],
        ["Viszonyítási dátum", "Ehhez a naphoz képest számoljuk a még kint lévő könyvek késését."],
        ["Szétbontás", "Az adatok csoportosítása egy szempont szerint, csoportonként külön fájlba."],
        ["Futtatás", "Egy teljes feldolgozás (formázás + szétbontás) a naplóban egy sorral."],
        ["Összesítő fájl", "A csoportok sorszámát, késéseit, díjait és a hibás sorokat tartalmazó fájl."],
    ], [4.0, 12.6], caption="Fogalomszótár")


def chapter2() -> None:
    page_break()
    h1("2. Funkcionális terv")
    h2("2.1. Funkciófa")
    para("A specifikációból adódó fő funkciókat az alábbi fa foglalja össze. A funkciók részletes leírása a "
         "2.2. fejezetben, az ablakok sorszámára ([n.]) hivatkozva található.")
    figure(DIAG / "funkciofa.png", "A KEF funkciófája")
    h2("2.2. A funkciók részletezése")
    specs = [
        ("Bejelentkezés", "a felhasználót bejelentkezteti – [1.] ablak", [
            ("Bemenő adatok", "Felhasználói azonosító (karakteres, 3–30 karakter: angol betű, szám, pont, kötőjel, "
                              "aláhúzás; kis- és nagybetű különbözik; kötelező).\nJelszó (karakteres, 8–30 karakter, "
                              "betűt és számjegyet is tartalmaz, szóköz nélkül; a beírt karakterek helyett pontok "
                              "látszanak; kötelező)."),
            ("Kimenő adatok", "Sikeres belépéskor a Főmenü [2.] nyílik meg a szerepkörnek megfelelő gombokkal; az "
                              "utolsó belépés ideje mentésre kerül.\nÜres mező, nem létező azonosító vagy rossz jelszó "
                              "esetén hibaüzenet [10.]: „Helytelen felhasználói azonosító és/vagy jelszó!” (nem árulja "
                              "el, melyik volt hibás). Letiltott fióknál: „A felhasználói fiók le van tiltva.”"),
            ("Mikor hívható meg", "A program ezzel az ablakkal indul; kijelentkezés után ide tér vissza."),
            ("Mellékhatások", "Hibás próbálkozás után a jelszómező törlődik."),
            ("Részletes leírás", "A felhasználó beírja azonosítóját és jelszavát, majd a Bejelentkezés gombra "
                                 "kattint (vagy Entert nyom a jelszómezőben). A program az adatbázisban tárolt "
                                 "PBKDF2-hash alapján ellenőrzi a jelszót. A Kilépés gomb megerősítés után leállítja "
                                 "a programot. Elfelejtett jelszó esetén az adminisztrátor állítja vissza."),
        ]),
        ("Kijelentkezés és Kilépés", "[2.] ablak, illetve bármely ablak [X] gombja", [
            ("Bemenő adatok", "A Kijelentkezés / Kilépés gomb, illetve az ablak [X] gombja."),
            ("Kimenő adatok", "Megerősítés [10.] („Biztosan kijelentkezel?” / „Biztosan kilépsz a programból?”). "
                              "Igen esetén kijelentkezéskor a Bejelentkezés [1.] jelenik meg, kilépéskor a program "
                              "leáll. Mégse esetén az aktuális ablak marad."),
            ("Mikor hívható meg", "Kijelentkezés: Főmenü [2.]. Kilépés: [1.], [2.] és az [X] minden ablakban."),
            ("Mellékhatások", "A betöltött, még fel nem dolgozott adatok elvesznek. Futó feldolgozás közben a "
                              "kilépés külön figyelmeztet; ilyenkor a részleges fájlok törlődnek."),
            ("Részletes leírás", "Kijelentkezéskor a varázsló állapota (betöltött fájl, beállítások) törlődik, így a "
                                 "következő felhasználó nem látja az előző adatait."),
        ]),
        ("Adatbetöltés", "a könyvtári export beolvasása – [3.] ablak", [
            ("Bemenő adatok", "Forrásfájl elérési útja (Tallózás… gombbal vagy beírva; .xlsx, .xlsm vagy .csv; "
                              "kötelező). A Mintaadat gomb a mellékelt 90 000 soros mintafájlt választja ki."),
            ("Kimenő adatok", "Összesítés: összes sor, érvényes sor, hibás sor, fiókok száma, időszak (első és utolsó "
                              "kölcsönzés). Előnézet fül: az első 100 érvényes sor 10 fő oszlopa. Hibás sorok fül: "
                              "a forrásfájl sorszáma, kölcsönzés-azonosító, hibaok.\nFájlszintű hibánál hibaüzenet "
                              "[10.], pl. „Hiányzó oszlop(ok): cim”, „A fájl nem tartalmaz adatsort.”, „Nem támogatott "
                              "fájlformátum (csak .xlsx és .csv).”"),
            ("Mikor hívható meg", "Főmenü [2.] → Új feldolgozás (csak admin és feldolgozó), illetve az Eredmény "
                                  "összesítő [7.] Új feldolgozás gombjával."),
            ("Mellékhatások", "A beolvasás háttérszálon fut (90 000 sornál kb. 5 mp), közben folyamatjelző látszik. "
                              "Új fájl betöltése törli az előző betöltés eredményét."),
            ("Részletes leírás", "A program minden cellát szövegként olvas be, majd oszloponként ellenőriz (lásd "
                                 "Validáció). A Tovább gomb csak akkor aktív, ha van legalább egy érvényes sor. A "
                                 "hibás sorok nem állítják meg a feldolgozást: kimaradnak a fiókfájlokból, és az "
                                 "összesítő fájl „Hibás sorok” lapjára kerülnek."),
        ]),
        ("Validáció", "a betöltés része – [3.] ablak", [
            ("Bemenő adatok", "A beolvasott export sorai."),
            ("Kimenő adatok", "Érvényes sorok (típusos adatokkal) és hibalista. Egy sorhoz több hibaok is tartozhat, "
                              "pontosvesszővel elválasztva."),
            ("Mikor hívható meg", "Automatikusan, minden betöltéskor."),
            ("Mellékhatások", "–"),
            ("Részletes leírás", "Ellenőrzött szabályok: kötelező mezők (kolcsonzes_id, olvaso_id, olvaso_tipus, "
                                 "leltari_szam, cim, fiok_kod, fiok_nev, kolcsonzes_datum, lejarat_datum) nem üresek; "
                                 "dátumok ÉÉÉÉ-HH-NN vagy ÉÉÉÉ.HH.NN. formátumúak és létező napok; a lejárat és a "
                                 "visszahozás nem korábbi a kölcsönzésnél; a kölcsönzés-azonosító egyedi; a fiókkód "
                                 "F01–F12; az olvasótípus diák/felnőtt/nyugdíjas; a hosszabbítások száma 0–5 egész. "
                                 "Az ISBN-ből a kötőjelek, a szövegekből a felesleges szóközök eltávolításra kerülnek."),
        ]),
        ("Formázási szabályok", "megtekintés és módosítás – [4.] ablak", [
            ("Bemenő adatok", "Díj / késedelmes nap (egész, 0–10 000 Ft; alapért. 50).\nDíjplafon (egész, "
                              "0–100 000 Ft; alapért. 3 000).\nKedvezmény diáknak és nyugdíjasnak (egész, 0–100%; "
                              "alapért. 50).\nViszonyítási dátum (ÉÉÉÉ.HH.NN.; alapért. a mai nap).\nHozzáadandó "
                              "számított oszlopok (6 jelölőnégyzet; alapért. mind)."),
            ("Kimenő adatok", "A varázslóban a beállítások az aktuális futtatásra érvényesek. A Mentés "
                              "alapértelmezettként gomb (csak admin) megerősítés [10.] után a beallitas táblába ment. "
                              "Hibás érték esetén hibaüzenet, pl. „A kedvezmény 0 és 100% között lehet.”"),
            ("Mikor hívható meg", "Varázslóból: [3.] Tovább gomb. Önállóan: Főmenü [2.] → Formázási szabályok "
                                  "(mindenkinek, de csak admin módosíthat)."),
            ("Mellékhatások", "A mentett alapértékek minden későbbi feldolgozásra hatnak."),
            ("Részletes leírás", "Nem admin felhasználónál a díjmezők és jelölőnégyzetek inaktívak; a viszonyítási "
                                 "dátum futtatásonként mindenkinek megadható. A Gyári értékek gomb visszatölti az "
                                 "eredeti beállításokat (mentés nélkül)."),
        ]),
        ("Formázás (számított oszlopok)", "a feldolgozás 1. lépése – [6.] ablakban fut", [
            ("Bemenő adatok", "Érvényes sorok és a formázási szabályok."),
            ("Kimenő adatok", "Kiegészített adatok: kölcsönzési napok, késés (nap), késedelmi díj (Ft), státusz, "
                              "év-hónap, korcsoport (képletek a 4.3. fejezetben). Rendezés fiókkód, kölcsönzési "
                              "dátum és azonosító szerint."),
            ("Mikor hívható meg", "A Feldolgozás indítása után automatikusan."),
            ("Mellékhatások", "A forrásfájl nem módosul."),
            ("Részletes leírás", "A valós megoldásban ezt a lépést végezte a VBA-makró. A számítás vektorizáltan, "
                                 "90 000 soron kb. 1 mp alatt fut."),
        ]),
        ("Szétbontás beállítása", "[5.] ablak", [
            ("Bemenő adatok", "Szempont (rádiógomb: fiókkönyvtár / kategória / könyv (ISBN); alapért. fiók).\n"
                              "Időszak -tól, -ig (ÉÉÉÉ.HH.NN., nem kötelező; a -tól nem lehet későbbi az -ig-nél).\n"
                              "Kimeneti mappa (kötelező, írhatónak kell lennie).\nFájlnév-minta (kötelező; tartalmaznia "
                              "kell a {kulcs} helyőrzőt; tiltott karakterek: < > : \" / \\ | ? *; alapért. "
                              "KEF_{idoszak}_{kulcs}_{nev}).\nÖsszesítő fájl, Hibás sorok lap (jelölőnégyzetek)."),
            ("Kimenő adatok", "Előnézet: a létrejövő fájlok száma, neve és sorszáma. Hibás beállításnál az előnézet "
                              "helyén hibaüzenet."),
            ("Mikor hívható meg", "A [4.] ablak Tovább gombjával (csak varázslóból)."),
            ("Mellékhatások", "Adminnál a mappa és a minta alapértelmezettként mentődik."),
            ("Részletes leírás", "A fájlnévben az ékezetes betűk ékezet nélkülire, a szóközök és tiltott karakterek "
                                 "aláhúzásra cserélődnek (pl. „Szőlőhegyi Fiók” → Szolohegyi_Fiok). 200-nál több "
                                 "létrejövő fájl esetén (pl. könyv szerinti bontás) külön figyelmeztetés jelenik meg."),
        ]),
        ("Feldolgozás indítása és mentés", "[5.] → [10.] → [6.] → [7.]", [
            ("Bemenő adatok", "A betöltött adatok, a formázási szabályok és a szétbontási beállítások."),
            ("Kimenő adatok", "Megerősítés [10.]: „Biztosan elindítod a feldolgozást? N fájl + összesítő jön létre "
                              "ide: …” (ha a mappában már vannak .xlsx fájlok, figyelmeztet a felülírásra). Igen után "
                              "a [6.] ablak mutatja a haladást (folyamatjelző, aktuális fájl, eltelt idő, napló). "
                              "Siker esetén az Eredmény összesítő [7.] nyílik meg."),
            ("Mikor hívható meg", "[5.] Feldolgozás indítása gomb; csak admin és feldolgozó szerepkörrel."),
            ("Mellékhatások", "Új sor a futtatas táblában (statusz: fut → sikeres / hibas / megszakitva), a "
                              "létrejött fájlok a kimeneti_fajl táblába kerülnek. Az azonos nevű fájlok felülíródnak."),
            ("Részletes leírás", "A feldolgozás háttérszálon fut, így a felület nem fagy le. Minden csoport fájlja "
                                 "két munkalapot kap: „Kölcsönzések” (formázott adatok) és „Összesítő” (darabszámok, "
                                 "díjösszeg, első/utolsó kölcsönzés). A végén elkészül az összesítő fájl."),
        ]),
        ("Megszakítás", "[6.] ablak", [
            ("Bemenő adatok", "Megszakítás gomb, megerősítés [10.]."),
            ("Kimenő adatok", "Az aktuális fájl befejezése után a feldolgozás leáll, a már létrehozott fájlok "
                              "törlődnek, figyelmeztetés jelenik meg, majd a [5.] ablak lesz aktív."),
            ("Mikor hívható meg", "Feldolgozás közben."),
            ("Mellékhatások", "A futtatás „megszakitva” státusszal kerül a naplóba."),
            ("Részletes leírás", "Így nem marad olyan félkész kimenet, amelyben csak a fiókok egy része szerepel."),
        ]),
        ("Eredmény megtekintése", "[7.] ablak", [
            ("Bemenő adatok", "Az utolsó sikeres futtatás eredménye."),
            ("Kimenő adatok", "Futtatás száma, fájlok száma, kiírt sorok, késésben lévő tételek, késedelmi díj "
                              "összesen, kihagyott hibás sorok, futásidő; táblázat fájlonként (csoport, név, fájlnév, "
                              "sorok, késésben, díj)."),
            ("Mikor hívható meg", "Sikeres feldolgozás után automatikusan."),
            ("Mellékhatások", "A Mappa megnyitása / Kijelölt fájl megnyitása gomb a Windows Intézőt, illetve az "
                              "Excelt indítja el."),
            ("Részletes leírás", "Innen a Főmenübe vagy egy új feldolgozás [3.] elejére lehet lépni."),
        ]),
        ("Előzmények szűrése", "[8.] ablak", [
            ("Bemenő adatok", "Indítás -tól, -ig (ÉÉÉÉ.HH.NN., nem kötelező), Felhasználó (lista; alapért. "
                              "Mindenki)."),
            ("Kimenő adatok", "Futtatások listája (azonosító, indítás, felhasználó, forrásfájl, sorok, hibás sorok, "
                              "szempont, fájlok száma, státusz, üzenet), a nem sikeresek pirossal; alatta a kijelölt "
                              "futtatás fájljai. Hibás dátumnál: „Érvénytelen dátum…”."),
            ("Mikor hívható meg", "Főmenü [2.] → Futtatási előzmények (minden szerepkör)."),
            ("Mellékhatások", "–"),
            ("Részletes leírás", "A Lista frissítése gomb alkalmazza a szűrőket, a Szűrők törlése visszaállítja az "
                                 "alapértelmezést. A Kimeneti mappa megnyitása a kijelölt futtatás mappáját nyitja meg."),
        ]),
        ("Saját jelszó módosítása", "[9.] ablak felső része", [
            ("Bemenő adatok", "Régi jelszó, új jelszó, új jelszó még egyszer (mind kötelező; az új jelszóra a "
                              "Bejelentkezésnél leírt szabály vonatkozik)."),
            ("Kimenő adatok", "Siker: „Az új jelszó mentése megtörtént.” Hibák: „A régi jelszó hibás.”, „Az új jelszó "
                              "és a megerősítése nem egyezik.”, „Az új jelszó nem egyezhet meg a régivel.”, illetve "
                              "a jelszószabály szövege."),
            ("Mikor hívható meg", "Főmenü [2.] → Felhasználók és jelszó (minden szerepkör)."),
            ("Mellékhatások", "A következő bejelentkezéstől az új jelszó érvényes."),
            ("Részletes leírás", "A jelszó csak hash formában kerül az adatbázisba; a program sehol nem jeleníti meg."),
        ]),
        ("Felhasználók kezelése", "[9.] ablak alsó része – csak adminisztrátor", [
            ("Bemenő adatok", "Új felhasználó: azonosító, név, szerepkör, kezdő jelszó.\nKijelölt felhasználó: "
                              "szerepkör, aktív jelölőnégyzet."),
            ("Kimenő adatok", "A felhasználók listája (azonosító, név, szerepkör, aktív, létrehozva, utolsó belépés). "
                              "Hibák: „Ilyen azonosítójú felhasználó már létezik.”, „Saját magadat nem tilthatod le és "
                              "nem fokozhatod le.”"),
            ("Mikor hívható meg", "[9.] ablak, ha a bejelentkezett felhasználó adminisztrátor (másnak ez a rész nem "
                                  "is látszik)."),
            ("Mellékhatások", "Letiltott felhasználó nem tud belépni; a korábbi futtatásai megmaradnak."),
            ("Részletes leírás", "A Jelszó visszaállítása megerősítés után ideiglenes jelszót állít be, amelyet a "
                                 "program egyszer kiír; a felhasználónak belépés után meg kell változtatnia."),
        ]),
    ]
    for title, sub, items in specs:
        func_spec(title, sub, items)


SCREEN_TEXT = [
    ("01_bejelentkezes", "[1.] Bejelentkezés",
     "A program indításakor jelenik meg. Az azonosító és a jelszó megadása után a Bejelentkezés gomb (vagy Enter) "
     "a Főmenübe visz; hibás adatoknál a [10.] hibaüzenet jelenik meg."),
    ("02_fomenu", "[2.] Főmenü",
     "A program „irányítóközpontja”. A gombok a szerepkörtől függően aktívak; alul a legutóbbi futtatás adatai "
     "láthatók. Az alábbi második kép a megtekintő szerepkörű felhasználó főmenüje (az Új feldolgozás inaktív)."),
    ("03_adatbetoltes", "[3.] Adatbetöltés és validáció",
     "A forrásfájl kiválasztása és beolvasása. Az összesítő sor mutatja az összes, érvényes és hibás sorok számát, "
     "alatta két fül: előnézet és a hibás sorok listája (második kép)."),
    ("04_formazasi_szabalyok", "[4.] Formázási szabályok",
     "A késedelmi díj paraméterei, a viszonyítási dátum és a hozzáadandó számított oszlopok. Csak admin módosíthat "
     "és menthet alapértelmezettet."),
    ("05_szetbontas", "[5.] Szétbontás beállításai",
     "Szempont, időszak, kimeneti mappa és fájlnév-minta megadása; jobb oldalon élő előnézet a létrejövő fájlokról."),
    ("06_feldolgozas", "[6.] Feldolgozás folyamatban",
     "Folyamatjelző, az éppen írt fájl, eltelt idő és napló. A Megszakítás gomb megerősítés után leállítja a "
     "feldolgozást és törli a részleges fájlokat."),
    ("07_eredmeny", "[7.] Eredmény összesítő",
     "A sikeres futtatás összesítése és fájlonkénti táblázata; innen megnyitható a kimeneti mappa vagy egy fájl."),
    ("08_elozmenyek", "[8.] Futtatási előzmények",
     "Az összes korábbi futtatás szűrhető listája; a kijelölt futtatás fájljai alul jelennek meg."),
    ("09_felhasznalok", "[9.] Felhasználók és jelszó",
     "Felül a saját jelszó módosítása (mindenkinek), alul a felhasználókezelés (csak adminisztrátornak látszik)."),
    ("10_megerosites", "[10.] Megerősítés / üzenet",
     "Közös párbeszédablak: megerősítés (Igen / Mégse), információ, figyelmeztetés és hibaüzenet. Minden mentés, "
     "indítás, megszakítás, kijelentkezés és kilépés előtt megjelenik."),
]


def chapter3() -> None:
    page_break()
    h1("3. Felhasználói felület terve")
    para("Az alkalmazás 10 ablakból áll. Az ablakok fejlécében a sorszám ([n.]) és a bejelentkezett felhasználó "
         "látható. Az alábbi képek az elkészült alkalmazásról, a 90 000 soros mintaadat feldolgozása közben "
         "készültek.")
    h2("3.1. Ablaktervek")
    for name, title, text in SCREEN_TEXT:
        h3(title)
        para(text)
        width = 9.5 if name == "10_megerosites" else 15.5
        figure(SHOTS / f"{name}.png", title, width)
        if name == "02_fomenu":
            figure(SHOTS / "02b_fomenu_megtekinto.png", "[2.] Főmenü megtekintő szerepkörrel", 15.5)
        if name == "03_adatbetoltes":
            figure(SHOTS / "03b_hibas_sorok.png", "[3.] Hibás sorok fül a 270 kiszűrt sorral", 15.5)
    h2("3.2. Az ablakok bejárhatósága")
    para("A gráf az ablakok közötti átmeneteket mutatja; a nyilakon az átmenetet kiváltó gomb neve szerepel. A kék "
         "nyilak csak adminisztrátor és feldolgozó szerepkörrel járhatók be (feldolgozás). A [10.] megerősítő "
         "ablak a gráfban csak a feldolgozás indításánál szerepel, de minden mentés, megszakítás, kijelentkezés és "
         "kilépés előtt is megjelenik; Mégse esetén mindig az előző ablak marad aktív.")
    figure(DIAG / "bejarhatosag.png", "Az ablakok bejárhatósági gráfja")


def chapter4() -> None:
    page_break()
    h1("4. Adatmodell")
    h2("4.1. Áttekintés")
    para("Az adatok két helyen élnek: (1) a bemenő export egy fájl, amelyet a program csak olvas; (2) az alkalmazás "
         "saját adatait (felhasználók, beállítások, futtatások, létrejött fájlok) egy helyi SQLite-adatbázis "
         "(data/kef.db) tárolja. A kimenet Excel-fájlokból áll.")
    figure(DIAG / "er_diagram.png", "Az adatmodell (ER-diagram)")

    h2("4.2. A bemenő export szerkezete")
    desc = {
        "kolcsonzes_id": ("szöveg", "igen", "a kölcsönzés egyedi azonosítója", "K20000003"),
        "olvaso_id": ("szöveg", "igen", "olvasójegy-azonosító", "O105470"),
        "olvaso_nev": ("szöveg", "nem", "az olvasó neve", "Rácz Gergely"),
        "olvaso_szul_ev": ("egész", "nem", "születési év (korcsoporthoz)", "1980"),
        "olvaso_tipus": ("szöveg", "igen", "diák / felnőtt / nyugdíjas", "felnőtt"),
        "leltari_szam": ("szöveg", "igen", "a példány leltári száma", "L0035908"),
        "isbn": ("szöveg", "nem", "ISBN-13", "9789630695947"),
        "cim": ("szöveg", "igen", "a könyv címe", "Rejtett csillag III."),
        "szerzo": ("szöveg", "nem", "szerző", "Biró Eszter"),
        "kiado": ("szöveg", "nem", "kiadó", "Tudástár Kiadó"),
        "kiadas_eve": ("egész", "nem", "kiadás éve", "1996"),
        "kategoria": ("szöveg", "nem", "12 kategória egyike", "Gyermekkönyv"),
        "fiok_kod": ("szöveg", "igen", "F01–F12", "F02"),
        "fiok_nev": ("szöveg", "igen", "a fiók neve", "Belvárosi Fiók"),
        "kolcsonzes_datum": ("dátum", "igen", "a kölcsönzés napja", "2024.11.15."),
        "lejarat_datum": ("dátum", "igen", "határidő (hosszabbításokkal)", "2025.03.07."),
        "visszahozas_datum": ("dátum", "nem", "üres, ha még kint van", "2024.11.17."),
        "hosszabbitasok_szama": ("egész", "nem", "0–5, üres = 0", "3"),
        "konyvtaros_id": ("szöveg", "nem", "a kölcsönzést rögzítő könyvtáros", "KT024"),
        "megjegyzes": ("szöveg", "nem", "szabad szöveg", "Sérült borító"),
    }
    table(["Oszlop", "Típus", "Kötelező", "Leírás", "Példa"],
          [[c, *desc[c]] for c in config.REQUIRED_COLUMNS], [3.9, 1.6, 1.7, 5.2, 4.2], size=8.5,
          caption="A bemenő export 20 oszlopa")

    h2("4.3. Számított oszlopok")
    table(["Oszlop", "Szabály / képlet", "Példa"], [
        ["kolcsonzesi_napok", "(visszahozás vagy viszonyítási dátum) − kölcsönzés dátuma, napokban",
         "2025.02.01. → 2025.03.11. = 38"],
        ["keses_napok", "max(0, (visszahozás vagy viszonyítási dátum) − lejárat)", "lejárat 03.01., vissza 03.11. = 10"],
        ["kesedelmi_dij", "min(késés × díj/nap, plafon) × (1 − kedvezmény, ha diák vagy nyugdíjas), egészre kerekítve",
         "10 × 50 = 500 Ft; diáknak 250 Ft"],
        ["statusz", "visszahozva, ha van visszahozás; különben késésben, ha késés > 0; különben kint", "késésben"],
        ["ev_honap", "a kölcsönzés éve és hónapja ÉÉÉÉ-HH alakban", "2025-02"],
        ["korcsoport", "viszonyítási év − születési év: 0–14, 15–18, 19–35, 36–64, 65+ (hiányzó: ismeretlen)",
         "36–64"],
    ], [3.4, 8.6, 4.6], caption="Számított oszlopok")

    h2("4.4. Mintaadatok a bemenő exportból")
    lr, picks = sample_rows()
    para(f"A mellékelt mintafájl ({fmt(lr.total_rows)} sor, ebből {fmt(lr.valid_rows)} érvényes és "
         f"{lr.error_rows} hibás) néhány jellemző sora a számított oszlopokkal (viszonyítási dátum: 2026.11.15.):")
    table(["Azonosító", "Olvasó", "Típus", "Cím", "Fiók", "Kölcsönzés", "Lejárat", "Vissza", "Késés", "Díj",
           "Státusz"],
          [[r["kolcsonzes_id"], r["olvaso_id"], r["olvaso_tipus"], r["cim"], r["fiok_kod"], d(r["kolcsonzes_datum"]),
            d(r["lejarat_datum"]), d(r["visszahozas_datum"]), str(r["keses_napok"]), str(r["kesedelmi_dij"]),
            r["statusz"]] for r in picks],
          [1.7, 1.4, 1.3, 2.2, 0.9, 1.7, 1.7, 1.7, 1.0, 1.0, 2.0], size=7.5,
          caption="Mintasorok a bemenő exportból a számított oszlopokkal")
    errs = lr.errors.drop_duplicates("hiba").head(5)
    table(["Forrás sora", "Kölcsönzés azonosító", "Hiba"],
          [[str(r.sor), r.kolcsonzes_id or "(üres)", r.hiba] for r in errs.itertuples()], [2.6, 4.0, 10.0], size=9,
          caption="Példák a validáció által kiszűrt sorokra")

    h2("4.5. Az alkalmazás adatbázisa mintaadatokkal")
    con = demo_db()
    users = con.execute("SELECT * FROM felhasznalo ORDER BY id").fetchall()
    table(["id", "azonosito", "nev", "jelszo_hash", "szerepkor", "aktiv", "utolso_belepes"],
          [[str(u["id"]), u["azonosito"], u["nev"], u["jelszo_hash"][:26] + "…", u["szerepkor"], str(u["aktiv"]),
            u["utolso_belepes"] or "NULL"] for u in users],
          [0.8, 2.3, 2.7, 4.4, 2.2, 1.1, 3.1], size=8, caption="felhasznalo")
    runs = con.execute("SELECT * FROM futtatas ORDER BY id").fetchall()
    table(["id", "felh. id", "inditas", "forrasfajl", "sorok_be", "ervenyes", "hibas", "szempont", "statusz"],
          [[str(r["id"]), str(r["felhasznalo_id"]), r["inditas"], Path(r["forrasfajl"]).name, str(r["sorok_be"]),
            str(r["ervenyes_sorok"]), str(r["hibas_sorok"]), r["szempont"], r["statusz"]] for r in runs],
          [0.7, 1.2, 2.9, 3.7, 1.4, 1.4, 1.1, 1.7, 2.3], size=8, caption="futtatas (a befejezes, kimeneti_mappa és "
                                                                          "uzenet oszlopok nélkül)")
    files = con.execute("SELECT * FROM kimeneti_fajl ORDER BY id LIMIT 4").fetchall()
    table(["id", "futt. id", "csoport_kulcs", "csoport_nev", "fajlnev", "sorok", "kesesben", "osszeg_dij"],
          [[str(f["id"]), str(f["futtatas_id"]), f["csoport_kulcs"], f["csoport_nev"], f["fajlnev"],
            str(f["sorok_szama"]), str(f["kesesben"]), str(f["osszeg_dij"])] for f in files],
          [0.7, 1.2, 2.1, 2.7, 5.6, 1.2, 1.5, 1.6], size=8, caption="kimeneti_fajl (első 4 sor)")
    sets = con.execute("SELECT * FROM beallitas ORDER BY kulcs").fetchall()
    table(["kulcs", "ertek", "modositotta", "modositva"],
          [[s["kulcs"], s["ertek"] if len(s["ertek"]) < 60 else "…" + s["ertek"][-50:], str(s["modositotta"]),
            s["modositva"]] for s in sets] +
          [["dij_per_nap", "50 (alapértelmezés, amíg az admin nem ment mást)", "NULL", "NULL"]],
          [3.2, 8.0, 2.1, 3.3], size=8, caption="beallitas")
    con.close()

    h2("4.6. A kimeneti fájlok szerkezete")
    bullets([
        "Csoportfájl (pl. KEF_202411-202611_F02_Belvarosi_Fiok.xlsx):: „Kölcsönzések” munkalap a csoport összes "
        "sorával (20 + 6 oszlop, magyar fejlécek, szűrő, rögzített fejléc, késésben lévő sorok színezve) és "
        "„Összesítő” munkalap (kölcsönzések, visszahozva, kint, késésben, díj összesen, olvasók és könyvek száma, "
        "első/utolsó kölcsönzés).",
        "Összesítő fájl (KEF_…_OSSZESITO.xlsx):: csoportonként egy sor (csoport, név, fájl, sorok, késésben, díj) "
        "és a „Hibás sorok” munkalap a forrásfájl sorszámával és a hiba okával.",
    ])


TESTS = [
    ("T01", "bejelentkezés", "[1.]", "Bejelentkezés",
     "helyes azonosító és jelszó\nazonosító: feldolgozo\njelszó: Feldolgozo1 (●●●●●●●●●●●)",
     "Megnyílik a Főmenü [2.]: „Üdvözlünk, Kovács Anna!”, szerepkör: Feldolgozó; az utolsó belépés ideje mentve."),
    ("T02", "bejelentkezés", "[1.]", "Bejelentkezés",
     "a) feldolgozo / feldolgozo1 (kisbetűs jelszó)\nb) Feldolgozo / Feldolgozo1 (nagybetűs azonosító)\n"
     "c) üres azonosító és jelszó\nd) nincsilyen / Valami123",
     "Hibaüzenet [10.]: „Helytelen felhasználói azonosító és/vagy jelszó!”; a jelszómező törlődik, az [1.] "
     "marad aktív."),
    ("T03", "adatbetöltés", "[3.]", "Betöltés",
     "hianyos.xlsx: 2 adatsor, a „cim” oszlop hiányzik",
     "Hibaüzenet [10.]: „Hiányzó oszlop(ok): cim”; nincs előnézet, a Tovább gomb inaktív."),
    ("T04", "validáció", "[3.]", "Betöltés",
     "4 adatsor; a 3. adatsorban (K3) az olvaso_id üres",
     "Összes sor: 4, érvényes: 3, hibás: 1. Hibás sorok fül: sor 4, K3, „Kötelező mező üres: olvaso_id”."),
    ("T05", "formázás – késedelmi díj", "[4.] → [6.]", "Tovább, Indítás",
     "kölcsönzés 2025.02.01., lejárat 2025.03.01., visszahozás 2025.03.11.; díj 50 Ft/nap, plafon 3000 Ft, "
     "kedvezmény 50%\na) felnőtt olvasó\nb) diák olvasó",
     "a) késés 10 nap, díj 500 Ft, kölcsönzési napok 38, státusz: visszahozva\nb) késés 10 nap, díj 250 Ft"),
    ("T06", "formázás – díjplafon", "[4.] → [6.]", "Tovább, Indítás",
     "kölcsönzés 2024.12.01., lejárat 2025.01.01., nincs visszahozva; viszonyítási dátum 2025.06.01.",
     "Késés 151 nap; 151 × 50 = 7550 Ft helyett a díj 3000 Ft (plafon); státusz: késésben, a sor színezve."),
    ("T07", "szétbontás", "[5.] → [10.] → [6.] → [7.]", "Feldolgozás indítása, Indítás",
     "10 érvényes sor: F01 – 4, F02 – 3, F03 – 3 sor; szempont: fiókkönyvtár; minta: KEF_{idoszak}_{kulcs}_{nev}",
     "3 fájl + összesítő: KEF_202502-202502_F01_Kozponti_Konyvtar.xlsx (4 sor), …_F02_Belvarosi_Fiok.xlsx (3), "
     "…_F03_Kertvarosi_Fiok.xlsx (3), KEF_202502-202502_OSSZESITO.xlsx; sorösszeg 10 = érvényes sorok; az A1 "
     "cella „Kölcsönzés azonosító”, az 1. sor rögzítve."),
    ("T08", "validáció", "[3.]", "Betöltés",
     "2 adatsor; a 2. sorban fiok_kod = F99",
     "Érvényes: 1, hibás: 1 – „Ismeretlen fiókkód: F99”; az F99-es sor egyik fiók fájljába sem kerül."),
    ("T09", "szétbontás", "[5.]", "Feldolgozás indítása",
     "kimeneti mappa egy létező fájl „alatt”: …\\ez_egy_fajl.txt\\kimenet",
     "Hibaüzenet [10.]: „A kimeneti mappa nem írható: …”; a feldolgozás nem indul el, az [5.] marad aktív."),
    ("T10", "jogosultság", "[2.]", "Új feldolgozás",
     "bejelentkezés: megtekinto / Megtekinto1, majd feldolgozás indítása",
     "Az Új feldolgozás gomb inaktív; közvetlen hívásnál: „Ehhez a művelethez nincs jogosultságod.”; nem jön "
     "létre futtatás és fájl sem."),
]


def chapter5() -> None:
    new_section(landscape=True)
    h1("5. Tesztelési terv")
    h2("5.1. Tesztelési stratégia")
    para("A tesztelés a feldolgozási láncra (betöltés, validáció, formázás, szétbontás), a jogosultságkezelésre és "
         "a naplózásra koncentrál. A 10 kidolgozott teszteset mindegyike automatizált pytest-tesztként is "
         "megvalósult (tests/test_tesztesetek.py, T01–T10), további 14 egységteszt a határeseteket fedi le "
         "(CSV-beolvasás, duplikált azonosító, megszakítás, kategória szerinti bontás, jelszómódosítás stb.). "
         "A grafikus felületet a tools/screenshots.py füstteszt járja végig a 90 000 soros mintaadattal; ez "
         "készítette a 3. fejezet képernyőképeit is. A gombok bal egérgombbal és – ahol jelölve – Enterrel is "
         "működnek; a teszteseteknél bal egérgombos kattintást feltételezünk.")
    h2("5.2. Tesztesetek")
    table(["№", "Funkció", "Ablak", "Akciógomb", "Bemenő adatok", "Tervezett kimenő adatok", "Eredmény"],
          [[n, f, w, b, i, o, "sikeres (pytest)"] for n, f, w, b, i, o in TESTS],
          [1.1, 2.6, 2.4, 2.8, 6.4, 7.6, 2.2], size=8.5, caption="A 10 kidolgozott teszteset")
    h2("5.3. A tesztek futtatása és eredménye")
    para("Futtatás a projekt gyökérmappájából:  python -m pytest -q tests")
    para("Eredmény a beadott változaton: 26 teszt, mind sikeres (26 passed). A teljes 90 000 soros mintaadat "
         "feldolgozása a mérések szerint 37 mp (parancssorból) és 45 mp (a felületről, háttérszálon) – ez megfelel "
         "az NFK-01 követelménynek.")
    new_section(landscape=False)


def chapter6() -> None:
    h1("6. Megvalósítás")
    h2("6.1. Felhasznált technológiák")
    table(["Komponens", "Technológia", "Szerep"], [
        ["Programnyelv", "Python 3.12", "a teljes alkalmazás"],
        ["Felület", "Tkinter / ttk", "10 ablak, háttérszálas feldolgozás"],
        ["Adatfeldolgozás", "pandas, numpy", "beolvasás, validáció, számított oszlopok, csoportosítás"],
        ["Excel olvasás / írás", "python-calamine, xlsxwriter", "gyors .xlsx olvasás; formázott fájlok írása"],
        ["Adatbázis", "SQLite (sqlite3)", "felhasználók, beállítások, futtatások, fájlok"],
        ["Biztonság", "hashlib PBKDF2-SHA256", "jelszavak tárolása"],
        ["Tesztelés", "pytest", "26 automatikus teszt"],
        ["Dokumentáció", "python-docx, matplotlib", "ez a dokumentum és az ábrák"],
    ], [3.6, 4.8, 8.2], caption="Technológiák")
    h2("6.2. A forráskód felépítése")
    table(["Mappa / fájl", "Tartalom"], [
        ["kef/config.py", "konstansok: oszlopok, fiókok, alapértelmezett szabályok"],
        ["kef/pipeline/loader.py", "beolvasás és validáció (Adatbetöltés, Validáció)"],
        ["kef/pipeline/transform.py", "számított oszlopok (Formázás – a valós példában a VBA-makró)"],
        ["kef/pipeline/splitter.py", "szétbontás és Excel-formázás (a valós példában az Office Script)"],
        ["kef/service.py", "a teljes feldolgozás jogosultság-ellenőrzéssel és naplózással"],
        ["kef/db.py, kef/auth.py", "adatbázis, bejelentkezés, jelszavak, szerepkörök"],
        ["kef/gui/", "a 10 ablak (egy modul ablakonként) és a közös elemek"],
        ["tools/generate_sample.py", "90 000 soros szintetikus mintaadat (a valós SAP-adat helyett)"],
        ["tools/screenshots.py", "GUI-füstteszt és képernyőképek"],
        ["tests/", "pytest tesztek (T01–T10 + egységtesztek)"],
        ["docs/", "ábrák és a dokumentáció generálása"],
    ], [5.2, 11.4], caption="A projekt szerkezete")
    h2("6.3. Telepítés és futtatás")
    bullets([
        "Python 3.12 telepítése, majd a függőségek:: pip install -r requirements.txt",
        "Mintaadat előállítása:: python tools/generate_sample.py",
        "Az alkalmazás indítása:: python -m kef.main (bemutató felhasználók: lásd README.md)",
        "Tesztek:: python -m pytest -q tests",
    ])
    h2("6.4. Továbbfejlesztési lehetőségek")
    bullets([
        "Ütemezett, automatikus futtatás (Windows Feladatütemező) – a Power Automate ütemezésének megfelelője.",
        "A fiókfájlok automatikus e-mailben küldése a fiókvezetőknek vagy feltöltése SharePointra.",
        "Közvetlen kapcsolat a könyvtári rendszer adatbázisával, így az export kézi letöltése is elmaradhat.",
        "Kimutatások (diagramok) a havi forgalomról az összesítő fájlban.",
    ])


def main() -> None:
    import json

    pages_file = DOCS / "_toc_pages.json"
    if "--pages" in sys.argv and pages_file.exists():
        PAGES.update(json.loads(pages_file.read_text(encoding="utf-8")))
    setup()
    cover()
    chapter1()
    chapter2()
    chapter3()
    chapter4()
    chapter5()
    chapter6()
    build_toc()
    footer_page_numbers()
    (DOCS / "_toc_headings.json").write_text(json.dumps([t for _, t in HEADINGS], ensure_ascii=False), "utf-8")
    doc.core_properties.title = f"{config.APP_NAME} – dokumentáció"
    doc.core_properties.author = "KEF projektcsapat"
    doc.save(OUT)
    print("Mentve:", OUT)


if __name__ == "__main__":
    main()
