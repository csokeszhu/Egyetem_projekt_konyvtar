"""1. lépés – a könyvtári rendszer exportjának beolvasása és validálása."""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from kef import config


class ValidationError(Exception):
    """A teljes fájlt érintő hiba (pl. hiányzó oszlop, rossz formátum)."""


@dataclass
class LoadResult:
    path: Path
    total_rows: int
    valid: pd.DataFrame
    errors: pd.DataFrame  # oszlopok: sor, kolcsonzes_id, hiba
    extra_columns: list[str] = field(default_factory=list)

    @property
    def valid_rows(self) -> int:
        return len(self.valid)

    @property
    def error_rows(self) -> int:
        return len(self.errors)


SUPPORTED_SUFFIXES = (".xlsx", ".xlsm", ".csv")


def _read_raw(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as fh:
            sample = fh.read(4096)
        try:
            sep = csv.Sniffer().sniff(sample, delimiters=";,\t").delimiter
        except csv.Error:
            sep = ";"
        return pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    try:
        return pd.read_excel(path, dtype=str, engine="calamine", keep_default_na=False)
    except ImportError:  # pragma: no cover - ha nincs telepítve a python-calamine
        return pd.read_excel(path, dtype=str, engine="openpyxl", keep_default_na=False)


def parse_dates(series: pd.Series) -> pd.Series:
    """Elfogadja: 2025-03-01, 2025-03-01 00:00:00, 2025.03.01., 2025.03.01 — minden más NaT."""
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r"\.$", "", regex=True).str.replace(".", "-", regex=False)
    s = s.str.slice(0, 10)
    return pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")


def load_file(path: str | Path, known_branches: dict[str, str] | None = None) -> LoadResult:
    """Beolvassa és soronként validálja az exportot. Fájlszintű hibánál ValidationError-t dob."""
    path = Path(path)
    if not path.exists():
        raise ValidationError(f"A fájl nem található: {path}")
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ValidationError("Nem támogatott fájlformátum (csak .xlsx és .csv).")
    known_branches = config.BRANCHES if known_branches is None else known_branches

    try:
        raw = _read_raw(path)
    except Exception as exc:  # sérült / nem olvasható fájl
        raise ValidationError(f"A fájl nem olvasható: {exc}") from exc

    raw.columns = [str(c).strip().lower() for c in raw.columns]
    missing = [c for c in config.REQUIRED_COLUMNS if c not in raw.columns]
    if missing:
        raise ValidationError("Hiányzó oszlop(ok): " + ", ".join(missing))
    extra = [c for c in raw.columns if c not in config.REQUIRED_COLUMNS]
    if raw.empty:
        raise ValidationError("A fájl nem tartalmaz adatsort.")

    df = raw[config.REQUIRED_COLUMNS].copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    problems: dict[int, list[str]] = {}

    def flag(mask: pd.Series, message: str | pd.Series) -> None:
        for i in mask[mask].index:
            msg = message if isinstance(message, str) else message.loc[i]
            problems.setdefault(i, []).append(msg)

    # kötelező mezők
    for col in config.MANDATORY_COLUMNS:
        flag(df[col] == "", f"Kötelező mező üres: {col}")

    # dátumok
    dates = {c: parse_dates(df[c]) for c in config.DATE_COLUMNS}
    for col in config.DATE_COLUMNS:
        flag((df[col] != "") & dates[col].isna(), f"Érvénytelen dátum: {col}")
    loan, due, ret = dates["kolcsonzes_datum"], dates["lejarat_datum"], dates["visszahozas_datum"]
    flag(due.notna() & loan.notna() & (due < loan), "A lejárat korábbi, mint a kölcsönzés")
    flag(ret.notna() & loan.notna() & (ret < loan), "A visszahozás korábbi, mint a kölcsönzés")

    # duplikált kölcsönzés-azonosító (az első előfordulás marad érvényes)
    dup = (df["kolcsonzes_id"] != "") & df["kolcsonzes_id"].duplicated(keep="first")
    flag(dup, "Duplikált kölcsönzés-azonosító")

    # fiókkód
    unknown = (df["fiok_kod"] != "") & ~df["fiok_kod"].isin(known_branches)
    flag(unknown, "Ismeretlen fiókkód: " + df["fiok_kod"])

    # olvasótípus
    bad_type = (df["olvaso_tipus"] != "") & ~df["olvaso_tipus"].isin(config.READER_TYPES)
    flag(bad_type, "Ismeretlen olvasótípus: " + df["olvaso_tipus"])

    # hosszabbítások száma
    ext = pd.to_numeric(df["hosszabbitasok_szama"].replace("", "0"), errors="coerce")
    flag(ext.isna() | (ext < 0) | (ext > config.MAX_EXTENSIONS) | (ext % 1 != 0),
         f"Hosszabbítások száma 0–{config.MAX_EXTENSIONS} egész lehet")

    error_idx = sorted(problems)
    errors = pd.DataFrame(
        {
            "sor": [i + 2 for i in error_idx],  # Excel-sorszám (1. sor a fejléc)
            "kolcsonzes_id": df.loc[error_idx, "kolcsonzes_id"].tolist(),
            "hiba": ["; ".join(problems[i]) for i in error_idx],
        }
    )

    valid = df.drop(index=error_idx).copy()
    for col in config.DATE_COLUMNS:
        valid[col] = dates[col].drop(index=error_idx)
    valid["hosszabbitasok_szama"] = ext.drop(index=error_idx).astype(int)
    valid["olvaso_szul_ev"] = pd.to_numeric(valid["olvaso_szul_ev"], errors="coerce").astype("Int64")
    valid["kiadas_eve"] = pd.to_numeric(valid["kiadas_eve"], errors="coerce").astype("Int64")
    valid["isbn"] = valid["isbn"].map(lambda s: re.sub(r"[^0-9Xx]", "", s))
    valid = valid.reset_index(drop=True)

    return LoadResult(path=path, total_rows=len(df), valid=valid, errors=errors, extra_columns=extra)
