"""További egységtesztek a betöltéshez, formázáshoz, szétbontáshoz és a naplózáshoz."""
from __future__ import annotations

import sys
import threading
from datetime import date
from pathlib import Path

import pytest

from conftest import make_rows, write_xlsx
from kef.auth import AuthError, change_password, create_user, login
from kef.pipeline import Cancelled, SplitConfig, TransformSettings, ValidationError, load_file, split, transform
from kef.pipeline.splitter import OutputError, build_filename, safe_name
from kef.service import run_processing

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))


def test_csv_betoltes_pontosvesszovel(tmp_path):
    p = tmp_path / "export.csv"
    make_rows({}, {"kolcsonzes_datum": "2025.02.03."}).to_csv(p, sep=";", index=False, encoding="utf-8-sig")
    lr = load_file(p)
    assert lr.valid_rows == 2
    assert lr.valid["kolcsonzes_datum"].iloc[1] == __import__("pandas").Timestamp(2025, 2, 3)


def test_nem_tamogatott_formatum(tmp_path):
    p = tmp_path / "export.txt"
    p.write_text("x")
    with pytest.raises(ValidationError, match="Nem támogatott"):
        load_file(p)


def test_ures_fajl(tmp_path):
    with pytest.raises(ValidationError, match="nem tartalmaz adatsort"):
        load_file(write_xlsx(make_rows({}).iloc[0:0], tmp_path / "ures.xlsx"))


def test_duplikalt_azonosito_es_rossz_datumok(tmp_path):
    df = make_rows({}, {}, {"kolcsonzes_datum": "2025.13.45."}, {"visszahozas_datum": "2025-01-01"})
    df.loc[1, "kolcsonzes_id"] = "K1"
    lr = load_file(write_xlsx(df, tmp_path / "e.xlsx"))
    hibak = dict(zip(lr.errors["sor"], lr.errors["hiba"]))
    assert hibak[3] == "Duplikált kölcsönzés-azonosító"
    assert hibak[4] == "Érvénytelen dátum: kolcsonzes_datum"
    assert hibak[5] == "A visszahozás korábbi, mint a kölcsönzés"
    assert lr.valid_rows == 1


def test_szamitott_oszlopok_kikapcsolhatok(tmp_path):
    lr = load_file(write_xlsx(make_rows({}), tmp_path / "e.xlsx"))
    out = transform(lr.valid, TransformSettings(szamitott_oszlopok=["statusz"]))
    assert "statusz" in out.columns and "kesedelmi_dij" not in out.columns


def test_ervenytelen_beallitas():
    with pytest.raises(ValueError):
        TransformSettings(kedvezmeny_szazalek=150).validate()


def test_fajlnev_ekezetmentes():
    assert safe_name("Szőlőhegyi Fiók") == "Szolohegyi_Fiok"
    cfg = SplitConfig()
    assert build_filename(cfg, "202501-202512", "Sci-fi és fantasy", "") == "KEF_202501-202512_Sci-fi_es_fantasy.xlsx"


def test_fajlnev_minta_kulcs_nelkul():
    with pytest.raises(OutputError, match="kulcs"):
        SplitConfig(fajlnev_minta="KEF_{idoszak}").validate()


def test_kategoria_szerinti_szetbontas_es_idoszak(tmp_path):
    df = make_rows({"kategoria": "Krimi"}, {"kategoria": "Krimi", "kolcsonzes_datum": "2025-04-01",
                                            "lejarat_datum": "2025-04-29", "visszahozas_datum": ""},
                   {"kategoria": "Életrajz"})
    lr = load_file(write_xlsx(df, tmp_path / "e.xlsx"))
    data = transform(lr.valid, TransformSettings(viszonyitasi_datum=date(2025, 6, 1)))
    cfg = SplitConfig(szempont="kategoria", kimeneti_mappa=tmp_path / "ki", idoszak_ig=date(2025, 3, 1))
    files = split(data, cfg)
    assert {f.csoport_kulcs: f.sorok_szama for f in files} == {"Krimi": 1, "Életrajz": 1}


def test_megszakitas_torli_a_fajlokat(tmp_path):
    lr = load_file(write_xlsx(make_rows({}, {"fiok_kod": "F02", "fiok_nev": "B"}), tmp_path / "e.xlsx"))
    data = transform(lr.valid, TransformSettings())
    cancel = threading.Event()

    def progress(i, total, msg):
        if i == 1:
            cancel.set()

    with pytest.raises(Cancelled):
        split(data, SplitConfig(kimeneti_mappa=tmp_path / "ki"), progress=progress, cancel=cancel)
    assert list((tmp_path / "ki").glob("*.xlsx")) == []


def test_futtatas_naplozasa(db, tmp_path):
    lr = load_file(write_xlsx(make_rows({}, {"fiok_kod": "F02", "fiok_nev": "B"}), tmp_path / "e.xlsx"))
    user = login(db, "feldolgozo", "Feldolgozo1")
    res = run_processing(db, user, lr, TransformSettings(), SplitConfig(kimeneti_mappa=tmp_path / "ki"))
    runs = db.list_runs()
    assert len(runs) == 1 and runs[0]["statusz"] == "sikeres" and runs[0]["fajlok"] == 2
    assert [f["csoport_kulcs"] for f in db.list_output_files(res.run_id)] == ["F01", "F02"]


def test_jelszo_modositas_es_felhasznalo_letrehozas(db):
    admin = login(db, "admin", "Admin123!")
    create_user(db, admin, "uj.kollega", "Új Kolléga", "feldolgozo", "Kezdo1234")
    u = login(db, "uj.kollega", "Kezdo1234")
    with pytest.raises(AuthError, match="nem egyezik"):
        change_password(db, u, "Kezdo1234", "Masik5678", "Masik5679")
    with pytest.raises(AuthError, match="8–30"):
        change_password(db, u, "Kezdo1234", "rovid1", "rovid1")
    change_password(db, u, "Kezdo1234", "Masik5678", "Masik5678")
    assert login(db, "uj.kollega", "Masik5678").id == u.id


def test_generator_kis_mintaval():
    from generate_sample import generate

    df = generate(rows=2000, seed=7)
    assert df.shape == (2000, 20)
    assert (df["fiok_kod"] == "F99").sum() > 0       # szándékos hibák is vannak
