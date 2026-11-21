"""Közös beállítások és konstansok a Kölcsönzési Export Feldolgozóhoz (KEF)."""
from __future__ import annotations

from pathlib import Path

APP_NAME = "Kölcsönzési Export Feldolgozó"
APP_SHORT = "KEF"
APP_VERSION = "1.0"

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SAMPLE_DIR = DATA_DIR / "sample"
DB_PATH = DATA_DIR / "kef.db"
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

# --- A könyvtári rendszer exportjának oszlopai (technikai mezőnevek, mint egy SAP-exportban) ---
REQUIRED_COLUMNS: list[str] = [
    "kolcsonzes_id",
    "olvaso_id",
    "olvaso_nev",
    "olvaso_szul_ev",
    "olvaso_tipus",
    "leltari_szam",
    "isbn",
    "cim",
    "szerzo",
    "kiado",
    "kiadas_eve",
    "kategoria",
    "fiok_kod",
    "fiok_nev",
    "kolcsonzes_datum",
    "lejarat_datum",
    "visszahozas_datum",
    "hosszabbitasok_szama",
    "konyvtaros_id",
    "megjegyzes",
]

# Ezek nem lehetnek üresek egy érvényes sorban.
MANDATORY_COLUMNS: list[str] = [
    "kolcsonzes_id",
    "olvaso_id",
    "olvaso_tipus",
    "leltari_szam",
    "cim",
    "fiok_kod",
    "fiok_nev",
    "kolcsonzes_datum",
    "lejarat_datum",
]

DATE_COLUMNS: list[str] = ["kolcsonzes_datum", "lejarat_datum", "visszahozas_datum"]

READER_TYPES = ("diák", "felnőtt", "nyugdíjas")
MAX_EXTENSIONS = 5

# --- Számított oszlopok (a „formázás” lépés során adódnak hozzá) ---
COMPUTED_COLUMNS: list[str] = [
    "kolcsonzesi_napok",
    "keses_napok",
    "kesedelmi_dij",
    "statusz",
    "ev_honap",
    "korcsoport",
]

STATUS_RETURNED = "visszahozva"
STATUS_OUT = "kint"
STATUS_LATE = "késésben"

# Olvasható oszlopfejlécek a kimeneti fájlokban.
DISPLAY_NAMES: dict[str, str] = {
    "kolcsonzes_id": "Kölcsönzés azonosító",
    "olvaso_id": "Olvasó azonosító",
    "olvaso_nev": "Olvasó neve",
    "olvaso_szul_ev": "Születési év",
    "olvaso_tipus": "Olvasótípus",
    "leltari_szam": "Leltári szám",
    "isbn": "ISBN",
    "cim": "Cím",
    "szerzo": "Szerző",
    "kiado": "Kiadó",
    "kiadas_eve": "Kiadás éve",
    "kategoria": "Kategória",
    "fiok_kod": "Fiókkód",
    "fiok_nev": "Fiókkönyvtár",
    "kolcsonzes_datum": "Kölcsönzés dátuma",
    "lejarat_datum": "Lejárat dátuma",
    "visszahozas_datum": "Visszahozás dátuma",
    "hosszabbitasok_szama": "Hosszabbítások",
    "konyvtaros_id": "Könyvtáros",
    "megjegyzes": "Megjegyzés",
    "kolcsonzesi_napok": "Kölcsönzési napok",
    "keses_napok": "Késés (nap)",
    "kesedelmi_dij": "Késedelmi díj (Ft)",
    "statusz": "Státusz",
    "ev_honap": "Év-hónap",
    "korcsoport": "Korcsoport",
}

# --- Fiókkönyvtárak (a valós példában ezek voltak a „cégek”) ---
BRANCHES: dict[str, str] = {
    "F01": "Központi Könyvtár",
    "F02": "Belvárosi Fiók",
    "F03": "Kertvárosi Fiók",
    "F04": "Újtelepi Fiók",
    "F05": "Hegyaljai Fiók",
    "F06": "Egyetemvárosi Fiók",
    "F07": "Vasútállomási Fiók",
    "F08": "Ligetsori Fiók",
    "F09": "Malomkerti Fiók",
    "F10": "Patakparti Fiók",
    "F11": "Szőlőhegyi Fiók",
    "F12": "Gyermekkönyvtár",
}

# --- Szétbontási szempontok: kulcs -> (csoportosító oszlop, név oszlop, felirat) ---
SPLIT_KEYS: dict[str, tuple[str, str | None, str]] = {
    "fiok": ("fiok_kod", "fiok_nev", "Fiókkönyvtár szerint"),
    "kategoria": ("kategoria", None, "Kategória szerint"),
    "konyv": ("isbn", "cim", "Könyv (ISBN) szerint"),
}
DEFAULT_FILENAME_PATTERN = "KEF_{idoszak}_{kulcs}_{nev}"
FILENAME_PLACEHOLDERS = ("{idoszak}", "{kulcs}", "{nev}", "{datum}")
MAX_OUTPUT_FILES_WARNING = 200

# --- Alapértelmezett formázási szabályok (a beallitas táblában felülírhatók) ---
DEFAULT_SETTINGS: dict[str, str] = {
    "dij_per_nap": "50",          # Ft / késedelmes nap
    "dij_plafon": "3000",         # Ft / kölcsönzés
    "kedvezmeny_szazalek": "50",  # diák és nyugdíjas olvasóknak
    "szamitott_oszlopok": ",".join(COMPUTED_COLUMNS),
    "kimeneti_mappa": str(DEFAULT_OUTPUT_DIR),
    "fajlnev_minta": DEFAULT_FILENAME_PATTERN,
}

DATE_DISPLAY_FORMAT = "%Y.%m.%d."
