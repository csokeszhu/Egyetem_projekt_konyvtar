from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from kef import config  # noqa: E402
from kef.db import Database  # noqa: E402

BASE_ROW = {
    "kolcsonzes_id": "K1",
    "olvaso_id": "O100001",
    "olvaso_nev": "Kiss Anna",
    "olvaso_szul_ev": "1985",
    "olvaso_tipus": "felnőtt",
    "leltari_szam": "L0000001",
    "isbn": "978-963-00-0000-1",
    "cim": "A csendes folyó",
    "szerzo": "Nagy Péter",
    "kiado": "Lapozó Kiadó",
    "kiadas_eve": "2015",
    "kategoria": "Szépirodalom",
    "fiok_kod": "F01",
    "fiok_nev": "Központi Könyvtár",
    "kolcsonzes_datum": "2025-02-01",
    "lejarat_datum": "2025-03-01",
    "visszahozas_datum": "2025-02-20",
    "hosszabbitasok_szama": "0",
    "konyvtaros_id": "KT011",
    "megjegyzes": "",
}


def make_rows(*overrides: dict) -> pd.DataFrame:
    rows = []
    for i, o in enumerate(overrides, start=1):
        row = dict(BASE_ROW, kolcsonzes_id=f"K{i}")
        row.update(o)
        rows.append(row)
    return pd.DataFrame(rows, columns=config.REQUIRED_COLUMNS)


def write_xlsx(df: pd.DataFrame, path: Path) -> Path:
    df.to_excel(path, index=False, sheet_name="EXPORT")
    return path


@pytest.fixture
def db(tmp_path):
    database = Database(tmp_path / "test.db")
    yield database
    database.close()
