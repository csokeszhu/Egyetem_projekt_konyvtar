"""3. lépés – szétbontás külön Excel-fájlokba (a valós megoldásban: Power Automate + Office Script)."""
from __future__ import annotations

import os
import re
import tempfile
import threading
import unicodedata
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

import pandas as pd
import xlsxwriter

from kef import config


class OutputError(Exception):
    """A kimenet nem hozható létre (pl. nem írható mappa, hibás fájlnév-minta)."""


class Cancelled(Exception):
    """A felhasználó megszakította a feldolgozást."""


@dataclass
class SplitConfig:
    szempont: str = "fiok"
    kimeneti_mappa: Path = config.DEFAULT_OUTPUT_DIR
    fajlnev_minta: str = config.DEFAULT_FILENAME_PATTERN
    idoszak_tol: date | None = None
    idoszak_ig: date | None = None
    hibas_sorok_fajl: bool = True
    osszesito_fajl: bool = True

    def validate(self) -> None:
        if self.szempont not in config.SPLIT_KEYS:
            raise OutputError(f"Ismeretlen szétbontási szempont: {self.szempont}")
        if "{kulcs}" not in self.fajlnev_minta:
            raise OutputError("A fájlnév-mintának tartalmaznia kell a {kulcs} helyőrzőt.")
        if re.search(r'[<>:"/\\|?*]', self.fajlnev_minta.replace("{", "").replace("}", "")):
            raise OutputError('A fájlnév-minta nem tartalmazhatja: < > : " / \\ | ? *')
        if self.idoszak_tol and self.idoszak_ig and self.idoszak_tol > self.idoszak_ig:
            raise OutputError("Érvénytelen időszak: a kezdő dátum későbbi a záró dátumnál.")


@dataclass
class OutputFile:
    csoport_kulcs: str
    csoport_nev: str
    fajlnev: str
    sorok_szama: int
    osszeg_dij: int
    kesesben: int


ProgressCb = Callable[[int, int, str], None]


def safe_name(text: str, max_len: int = 60) -> str:
    """Ékezetmentes, fájlnévben biztonságos szöveg."""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("_")
    return text[:max_len].rstrip("_") or "ures"


def ensure_writable(folder: Path) -> None:
    try:
        folder.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=folder, prefix=".kef_")
        os.close(fd)
        os.remove(tmp)
    except OSError as exc:
        raise OutputError(f"A kimeneti mappa nem írható: {folder}") from exc


def filter_period(df: pd.DataFrame, tol: date | None, ig: date | None) -> pd.DataFrame:
    mask = pd.Series(True, index=df.index)
    if tol:
        mask &= df["kolcsonzes_datum"] >= pd.Timestamp(tol)
    if ig:
        mask &= df["kolcsonzes_datum"] <= pd.Timestamp(ig)
    return df[mask]


def period_label(df: pd.DataFrame, cfg: SplitConfig) -> str:
    start = pd.Timestamp(cfg.idoszak_tol) if cfg.idoszak_tol else df["kolcsonzes_datum"].min()
    end = pd.Timestamp(cfg.idoszak_ig) if cfg.idoszak_ig else df["kolcsonzes_datum"].max()
    if pd.isna(start) or pd.isna(end):
        return "ures"
    return f"{start:%Y%m}-{end:%Y%m}"


def group_counts(df: pd.DataFrame, cfg: SplitConfig) -> pd.Series:
    """Előnézethez: csoportonkénti sorszám (a szétbontás előtt)."""
    key_col, _, _ = config.SPLIT_KEYS[cfg.szempont]
    return filter_period(df, cfg.idoszak_tol, cfg.idoszak_ig)[key_col].value_counts().sort_index()


def build_filename(cfg: SplitConfig, idoszak: str, kulcs: str, nev: str) -> str:
    name = cfg.fajlnev_minta.format(
        idoszak=idoszak, kulcs=safe_name(kulcs, 30), nev=safe_name(nev, 40) if nev else "",
        datum=date.today().strftime("%Y%m%d"),
    )
    name = re.sub(r"_{2,}", "_", name).strip("_")
    return name + ".xlsx"


def _python_columns(df: pd.DataFrame) -> list[list]:
    """Oszloponként natív Python értékek (gyors xlsxwriter-íráshoz); hiányzó érték -> None."""
    cols = []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            cols.append(s.astype(object).where(s.notna(), None).tolist())
        elif pd.api.types.is_integer_dtype(s) or pd.api.types.is_float_dtype(s):
            cols.append(s.astype(object).where(s.notna(), None).tolist())
        else:
            cols.append(s.astype(object).where(s.notna() & (s.astype(str) != ""), None).tolist())
    return cols


def _write_sheet(wb: xlsxwriter.Workbook, df: pd.DataFrame, sheet: str) -> None:
    """Egy munkalap formázott kiírása: fejléc, szűrő, rögzítés, szélesség, számformátumok."""
    ws = wb.add_worksheet(sheet)
    headers = [config.DISPLAY_NAMES.get(c, c) for c in df.columns]
    head = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78",
                          "text_wrap": True, "valign": "vcenter", "border": 1})
    money = wb.add_format({"num_format": '#,##0 "Ft"'})
    late_fmt = wb.add_format({"bg_color": "#F8CBAD"})
    for i, col in enumerate(df.columns):
        sample = df[col].head(500).astype(str)
        width = 13 if col in config.DATE_COLUMNS else min(max(len(headers[i]), int(sample.str.len().max() or 0)) + 2, 42)
        ws.set_column(i, i, width, money if col == "kesedelmi_dij" else None)
    ws.set_row(0, 30)
    ws.write_row(0, 0, headers, head)
    for r, row in enumerate(zip(*_python_columns(df)), start=1):
        ws.write_row(r, 0, row)
    ws.freeze_panes(1, 0)
    n = len(df)
    ws.autofilter(0, 0, max(n, 1), len(headers) - 1)
    if "statusz" in df.columns and n:
        letter = _col_letter(df.columns.get_loc("statusz"))
        ws.conditional_format(1, 0, n, len(headers) - 1, {
            "type": "formula", "criteria": f'=${letter}2="{config.STATUS_LATE}"', "format": late_fmt})


def _write_table(wb: xlsxwriter.Workbook, df: pd.DataFrame, sheet: str, widths: list[int]) -> None:
    """Kis táblázat (összesítő, hibalista) egyszerű kiírása."""
    ws = wb.add_worksheet(sheet)
    head = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1})
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.write_row(0, 0, list(df.columns), head)
    for r, row in enumerate(zip(*_python_columns(df)), start=1):
        ws.write_row(r, 0, row)
    ws.freeze_panes(1, 0)


def _col_letter(idx: int) -> str:
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = [("Kölcsönzések száma", len(df))]
    if "statusz" in df.columns:
        rows += [
            ("Visszahozva", int((df["statusz"] == config.STATUS_RETURNED).sum())),
            ("Kint (határidőn belül)", int((df["statusz"] == config.STATUS_OUT).sum())),
            ("Késésben", int((df["statusz"] == config.STATUS_LATE).sum())),
        ]
    if "kesedelmi_dij" in df.columns:
        rows.append(("Késedelmi díj összesen (Ft)", int(df["kesedelmi_dij"].sum())))
    rows += [
        ("Különböző olvasók", int(df["olvaso_id"].nunique())),
        ("Különböző könyvek (ISBN)", int(df["isbn"].nunique())),
        ("Első kölcsönzés", df["kolcsonzes_datum"].min().strftime(config.DATE_DISPLAY_FORMAT) if len(df) else "-"),
        ("Utolsó kölcsönzés", df["kolcsonzes_datum"].max().strftime(config.DATE_DISPLAY_FORMAT) if len(df) else "-"),
    ]
    return pd.DataFrame(rows, columns=["Mutató", "Érték"])


def _workbook(path: Path) -> xlsxwriter.Workbook:
    return xlsxwriter.Workbook(str(path), {"constant_memory": True, "default_date_format": "yyyy.mm.dd."})


def split(
    df: pd.DataFrame,
    cfg: SplitConfig,
    errors: pd.DataFrame | None = None,
    progress: ProgressCb | None = None,
    cancel: threading.Event | None = None,
) -> list[OutputFile]:
    """Csoportonként egy-egy formázott .xlsx fájlt ír a kimeneti mappába.

    Megszakításkor a már létrehozott fájlokat törli és Cancelled kivételt dob.
    """
    cfg.validate()
    out_dir = Path(cfg.kimeneti_mappa)
    ensure_writable(out_dir)

    data = filter_period(df, cfg.idoszak_tol, cfg.idoszak_ig)
    if data.empty:
        raise OutputError("A megadott időszakban nincs kölcsönzés.")
    key_col, name_col, _ = config.SPLIT_KEYS[cfg.szempont]
    idoszak = period_label(data, cfg)
    groups = list(data.groupby(key_col, sort=True))
    total = len(groups) + int(cfg.osszesito_fajl)
    created: list[Path] = []
    results: list[OutputFile] = []

    try:
        for i, (key, part) in enumerate(groups, start=1):
            if cancel is not None and cancel.is_set():
                raise Cancelled()
            nev = str(part[name_col].iloc[0]) if name_col else ""
            fname = build_filename(cfg, idoszak, str(key), nev)
            path = out_dir / fname
            if progress:
                progress(i - 1, total, f"{fname} írása ({len(part):,} sor)…".replace(",", " "))
            created.append(path)
            with _workbook(path) as wb:
                _write_sheet(wb, part, "Kölcsönzések")
                _write_table(wb, _summary(part), "Összesítő", [32, 18])
            results.append(OutputFile(
                csoport_kulcs=str(key), csoport_nev=nev, fajlnev=fname, sorok_szama=len(part),
                osszeg_dij=int(part["kesedelmi_dij"].sum()) if "kesedelmi_dij" in part else 0,
                kesesben=int((part["statusz"] == config.STATUS_LATE).sum()) if "statusz" in part else 0,
            ))

        if cfg.osszesito_fajl:
            if cancel is not None and cancel.is_set():
                raise Cancelled()
            fname = build_filename(cfg, idoszak, "OSSZESITO", "")
            if progress:
                progress(total - 1, total, f"{fname} írása…")
            overview = pd.DataFrame([{
                "Csoport": r.csoport_kulcs, "Név": r.csoport_nev, "Fájl": r.fajlnev,
                "Sorok": r.sorok_szama, "Késésben": r.kesesben, "Késedelmi díj (Ft)": r.osszeg_dij,
            } for r in results])
            path = out_dir / fname
            created.append(path)
            with _workbook(path) as wb:
                _write_table(wb, overview, "Összesítő", [26, 26, 55, 12, 12, 18])
                if errors is not None and cfg.hibas_sorok_fajl and len(errors):
                    err = errors.rename(columns={"sor": "Sor (forrásfájl)", "kolcsonzes_id": "Kölcsönzés azonosító",
                                                 "hiba": "Hiba"})
                    _write_table(wb, err, "Hibás sorok", [16, 22, 60])
        if progress:
            progress(total, total, "Kész.")
    except Cancelled:
        for p in created:
            p.unlink(missing_ok=True)
        raise
    except OSError as exc:
        raise OutputError(f"Hiba a fájl írásakor: {exc}") from exc
    return results
