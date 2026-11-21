"""Szintetikus kölcsönzési export előállítása (a valós SAP-export helyett).

Használat:
    python tools/generate_sample.py                 # 90 000 sor -> data/sample/
    python tools/generate_sample.py --rows 5000 --seed 1

A kimenet .xlsx és .csv (pontosvesszővel elválasztva) formában is elkészül.
A szándékosan hibás sorok (kb. 0,3%) a validáció bemutatását szolgálják.
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from kef import config  # noqa: E402

EXPORT_DATE = date(2026, 11, 15)
PERIOD_DAYS = 730

LAST_NAMES = [
    "Nagy", "Kovács", "Tóth", "Szabó", "Horváth", "Varga", "Kiss", "Molnár", "Németh", "Farkas",
    "Balogh", "Papp", "Takács", "Juhász", "Lakatos", "Mészáros", "Oláh", "Simon", "Rácz", "Fekete",
    "Szilágyi", "Török", "Fehér", "Balázs", "Gál", "Kis", "Szűcs", "Kocsis", "Orsós", "Pintér",
    "Fodor", "Szalai", "Sipos", "Magyar", "Lukács", "Gulyás", "Biró", "Király", "Katona", "László",
]
FIRST_NAMES = [
    "Anna", "Bence", "Csilla", "Dániel", "Eszter", "Ferenc", "Gábor", "Hanna", "István", "Judit",
    "Katalin", "László", "Márton", "Nóra", "Olivér", "Péter", "Réka", "Sándor", "Tamás", "Zsófia",
    "Ádám", "Borbála", "Csaba", "Dóra", "Emese", "Flóra", "Gergely", "Hajnalka", "Imre", "Janka",
    "Kristóf", "Lilla", "Máté", "Noémi", "Orsolya", "Panna", "Richárd", "Szilvia", "Tünde", "Viktor",
    "Zoltán", "Ágnes", "Balázs", "Cecília", "Dénes", "Erika", "Gyula", "Ilona", "Jenő", "Kinga",
]
ADJECTIVES = [
    "csendes", "elveszett", "titkos", "utolsó", "hosszú", "fehér", "sötét", "kék", "régi", "távoli",
    "rejtett", "hideg", "aranyló", "szomorú", "vidám", "ismeretlen", "álmodó", "örök", "néma", "vad",
    "tavaszi", "őszi", "téli", "nyári", "üres", "zöld", "vörös", "furcsa", "bátor", "kis",
]
NOUNS = [
    "folyó", "levél", "kert", "város", "erdő", "ház", "híd", "tenger", "hegy", "út",
    "ablak", "kapu", "torony", "sziget", "vonat", "lámpás", "tükör", "óra", "térkép", "kulcs",
    "madár", "róka", "farkas", "hajó", "harang", "könyvtár", "csillag", "szél", "eső", "nyár",
    "dal", "álom", "emlék", "titok", "nap", "hold", "kő", "fa", "virág", "tó",
]
PUBLISHERS = [
    "Lapozó Kiadó", "Betűvető Kiadó", "Fókusz Könyvek", "Hársfa Kiadó", "Mérföldkő Kiadó",
    "Papírhajó Kiadó", "Tudástár Kiadó", "Kék Bagoly Könyvek", "Iránytű Kiadó", "Gondolat-jel Kiadó",
]
CATEGORIES = [
    "Szépirodalom", "Krimi", "Gyermekkönyv", "Ifjúsági", "Tudományos", "Történelem",
    "Életrajz", "Szakkönyv", "Sci-fi és fantasy", "Útikönyv", "Nyelvkönyv", "Művészet",
]
CATEGORY_WEIGHTS = np.array([22, 14, 12, 9, 6, 7, 5, 7, 8, 3, 3, 4], dtype=float)
NOTES = ["Sérült borító", "Hosszabbítás telefonon", "Előjegyzett példány", "Olvasói kérésre félretéve"]

COLUMNS = config.REQUIRED_COLUMNS


def _article(word: str) -> str:
    return "Az" if word[0].lower() in "aáeéiíoóöőuúüű" else "A"


def _isbn13(rng: np.random.Generator) -> str:
    digits = [9, 7, 8, 9, 6, 3] + list(rng.integers(0, 10, size=6))
    total = sum(d * (1 if i % 2 == 0 else 3) for i, d in enumerate(digits))
    check = (10 - total % 10) % 10
    return "".join(str(d) for d in digits) + str(check)


def _make_books(rng: np.random.Generator, n_books: int) -> pd.DataFrame:
    titles: set[str] = set()
    rows = []
    while len(rows) < n_books:
        adj = ADJECTIVES[rng.integers(len(ADJECTIVES))]
        noun = NOUNS[rng.integers(len(NOUNS))]
        form = rng.integers(4)
        if form == 0:
            title = f"{_article(adj)} {adj} {noun}"
        elif form == 1:
            noun2 = NOUNS[rng.integers(len(NOUNS))]
            title = f"{noun.capitalize()} és {noun2}"
        elif form == 2:
            title = f"{_article(noun)} {noun} {['titka', 'meséje', 'dala', 'útja', 'árnyéka'][rng.integers(5)]}"
        else:
            title = f"{adj.capitalize()} {noun}"
        if title in titles:
            title = f"{title} {['II.', 'III.', 'IV.'][rng.integers(3)]}"
            if title in titles:
                continue
        titles.add(title)
        author = f"{LAST_NAMES[rng.integers(len(LAST_NAMES))]} {FIRST_NAMES[rng.integers(len(FIRST_NAMES))]}"
        rows.append(
            {
                "isbn": _isbn13(rng),
                "cim": title,
                "szerzo": author,
                "kiado": PUBLISHERS[rng.integers(len(PUBLISHERS))],
                "kiadas_eve": int(rng.integers(1965, 2026)),
                "kategoria": CATEGORIES[rng.choice(len(CATEGORIES), p=CATEGORY_WEIGHTS / CATEGORY_WEIGHTS.sum())],
                "popularitas": float(rng.pareto(1.6) + 1),
            }
        )
    return pd.DataFrame(rows)


def _make_readers(rng: np.random.Generator, n_readers: int, branch_codes: list[str]) -> pd.DataFrame:
    birth = rng.integers(1940, 2019, size=n_readers)
    age = EXPORT_DATE.year - birth
    types = np.where(age <= 18, "diák", np.where(age >= 65, "nyugdíjas", "felnőtt"))
    # 19–25 évesek egy része egyetemista -> diák
    uni = (age >= 19) & (age <= 25) & (rng.random(n_readers) < 0.6)
    types = np.where(uni, "diák", types)
    names = [
        f"{LAST_NAMES[rng.integers(len(LAST_NAMES))]} {FIRST_NAMES[rng.integers(len(FIRST_NAMES))]}"
        for _ in range(n_readers)
    ]
    weights = np.array([18, 12, 10, 8, 7, 11, 6, 7, 6, 5, 4, 6], dtype=float)[: len(branch_codes)]
    home = rng.choice(branch_codes, size=n_readers, p=weights / weights.sum())
    return pd.DataFrame(
        {
            "olvaso_id": [f"O{100000 + i}" for i in range(n_readers)],
            "olvaso_nev": names,
            "olvaso_szul_ev": birth,
            "olvaso_tipus": types,
            "home": home,
        }
    )


def generate(rows: int = 90_000, seed: int = 42, error_rate: float = 0.003) -> pd.DataFrame:
    """Visszaad egy `rows` soros, a könyvtári rendszer exportját utánzó DataFrame-et."""
    rng = np.random.default_rng(seed)
    branch_codes = list(config.BRANCHES)
    n_books = max(50, min(3000, rows // 30))
    n_readers = max(100, min(8000, rows // 11))
    books = _make_books(rng, n_books)
    readers = _make_readers(rng, n_readers, branch_codes)

    r_idx = rng.integers(0, n_readers, size=rows)
    pop = books["popularitas"].to_numpy()
    b_idx = rng.choice(n_books, size=rows, p=pop / pop.sum())
    rd = readers.iloc[r_idx].reset_index(drop=True)
    bk = books.iloc[b_idx].reset_index(drop=True)

    # 80% az olvasó saját fiókjában kölcsönöz
    other = rng.choice(branch_codes, size=rows)
    branch = np.where(rng.random(rows) < 0.8, rd["home"].to_numpy(), other)

    start = EXPORT_DATE - timedelta(days=PERIOD_DAYS)
    offs = np.sort(rng.integers(0, PERIOD_DAYS, size=rows))
    loan = pd.to_datetime(start) + pd.to_timedelta(offs, unit="D")
    ext = rng.choice([0, 1, 2, 3], size=rows, p=[0.62, 0.25, 0.1, 0.03])
    due = loan + pd.to_timedelta(28 * (1 + ext), unit="D")

    # visszahozás: a legtöbben határidőn belül, ~12% késik, ~1% sosem hozza vissza
    keep_days = np.minimum(rng.gamma(2.2, 8.0, size=rows).astype(int) + 1, 28 * (1 + ext))
    late = rng.random(rows) < 0.12
    late_days = rng.gamma(1.4, 9.0, size=rows).astype(int) + 1
    ret_days = np.where(late, 28 * (1 + ext) + late_days, keep_days)
    ret = loan + pd.to_timedelta(ret_days, unit="D")
    export_ts = pd.Timestamp(EXPORT_DATE)
    lost = rng.random(rows) < 0.01
    ret = pd.Series(ret).where((ret <= export_ts) & ~lost, pd.NaT)

    librarians = {code: [f"KT{int(code[1:]) * 10 + k:03d}" for k in range(1, 5)] for code in branch_codes}
    lib = [librarians[b][rng.integers(4)] for b in branch]

    notes = np.full(rows, "", dtype=object)
    note_mask = rng.random(rows) < 0.04
    notes[note_mask] = rng.choice(NOTES, size=note_mask.sum())
    notes[lost & ret.isna().to_numpy()] = "Elveszett – pótlás folyamatban"

    # leltári szám: könyv + fiók egyedi példánya
    leltari = [f"L{(bi * 13 + int(b[1:])) % 9_999_999:07d}" for bi, b in zip(b_idx, branch)]

    df = pd.DataFrame(
        {
            "kolcsonzes_id": [f"K{20_000_000 + i}" for i in range(rows)],
            "olvaso_id": rd["olvaso_id"],
            "olvaso_nev": rd["olvaso_nev"],
            "olvaso_szul_ev": rd["olvaso_szul_ev"],
            "olvaso_tipus": rd["olvaso_tipus"],
            "leltari_szam": leltari,
            "isbn": bk["isbn"],
            "cim": bk["cim"],
            "szerzo": bk["szerzo"],
            "kiado": bk["kiado"],
            "kiadas_eve": bk["kiadas_eve"],
            "kategoria": bk["kategoria"],
            "fiok_kod": branch,
            "fiok_nev": [config.BRANCHES[b] for b in branch],
            "kolcsonzes_datum": loan,
            "lejarat_datum": due,
            "visszahozas_datum": ret,
            "hosszabbitasok_szama": ext,
            "konyvtaros_id": lib,
            "megjegyzes": notes,
        }
    )
    _inject_errors(df, rng, error_rate)
    return df[COLUMNS]


def _inject_errors(df: pd.DataFrame, rng: np.random.Generator, error_rate: float) -> None:
    """Szándékos adathibák, mint egy valós exportban (a validáció ezeket szűri ki)."""
    n = len(df)
    n_err = int(n * error_rate)
    if n_err < 5:
        return
    idx = rng.choice(np.arange(1, n), size=n_err, replace=False)
    parts = np.array_split(idx, 5)
    df["kolcsonzes_datum"] = df["kolcsonzes_datum"].astype(object)
    df.loc[parts[0], "olvaso_id"] = ""                                   # üres kötelező mező
    df.loc[parts[1], "kolcsonzes_datum"] = "2025.13.45."                 # érvénytelen dátum
    df.loc[parts[2], "kolcsonzes_id"] = df.loc[parts[2] - 1, "kolcsonzes_id"].to_numpy()  # duplikált azonosító
    df.loc[parts[3], "fiok_kod"] = "F99"                                 # ismeretlen fiók
    loan = pd.to_datetime(df.loc[parts[4], "kolcsonzes_datum"])
    df.loc[parts[4], "visszahozas_datum"] = loan - pd.Timedelta(days=3)  # visszahozás a kölcsönzés előtt


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--rows", type=int, default=90_000)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--out", type=Path, default=config.SAMPLE_DIR)
    parser.add_argument("--no-xlsx", action="store_true", help="csak CSV készüljön")
    args = parser.parse_args()

    args.out.mkdir(parents=True, exist_ok=True)
    df = generate(args.rows, args.seed)
    csv_path = args.out / "kolcsonzesek_export.csv"
    df.to_csv(csv_path, sep=";", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    print(f"CSV:  {csv_path}  ({len(df):,} sor, {len(df.columns)} oszlop)")
    if not args.no_xlsx:
        xlsx_path = args.out / "kolcsonzesek_export.xlsx"
        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
            df.to_excel(xw, index=False, sheet_name="EXPORT")
        print(f"XLSX: {xlsx_path}")


if __name__ == "__main__":
    main()
