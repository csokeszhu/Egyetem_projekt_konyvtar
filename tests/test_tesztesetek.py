"""A dokumentáció 5. fejezetében szereplő 10 teszteset automatizált változata (T01–T10)."""
from __future__ import annotations

from datetime import date

import openpyxl
import pytest

from conftest import make_rows, write_xlsx
from kef.auth import LOGIN_ERROR, AuthError, PermissionDenied, login
from kef.pipeline import OutputError, SplitConfig, TransformSettings, ValidationError, load_file, split, transform
from kef.service import run_processing


def test_T01_bejelentkezes_helyes_adatokkal(db):
    user = login(db, "feldolgozo", "Feldolgozo1")
    assert user.nev == "Kovács Anna"
    assert user.szerepkor == "feldolgozo"
    assert db.get_user(user.id)["utolso_belepes"] is not None


@pytest.mark.parametrize("azon, jelszo", [("feldolgozo", "feldolgozo1"), ("Feldolgozo", "Feldolgozo1"),
                                          ("", ""), ("nincsilyen", "Valami123")])
def test_T02_bejelentkezes_hibas_adatokkal(db, azon, jelszo):
    with pytest.raises(AuthError, match=LOGIN_ERROR):
        login(db, azon, jelszo)


def test_T03_hianyzo_oszlop(tmp_path):
    df = make_rows({}, {}).drop(columns=["cim"])
    path = write_xlsx(df, tmp_path / "hianyos.xlsx")
    with pytest.raises(ValidationError, match="Hiányzó oszlop\\(ok\\): cim"):
        load_file(path)


def test_T04_ures_kotelezo_mezo(tmp_path):
    df = make_rows({}, {}, {"olvaso_id": ""}, {})
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    assert (lr.total_rows, lr.valid_rows, lr.error_rows) == (4, 3, 1)
    err = lr.errors.iloc[0]
    assert err["sor"] == 4            # 3. adatsor = 4. Excel-sor
    assert err["kolcsonzes_id"] == "K3"
    assert err["hiba"] == "Kötelező mező üres: olvaso_id"


def test_T05_keses_es_dij_szamitasa(tmp_path):
    df = make_rows({"lejarat_datum": "2025-03-01", "visszahozas_datum": "2025-03-11"},
                   {"lejarat_datum": "2025-03-01", "visszahozas_datum": "2025-03-11", "olvaso_tipus": "diák"})
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    out = transform(lr.valid, TransformSettings(dij_per_nap=50, dij_plafon=3000, kedvezmeny_szazalek=50,
                                                viszonyitasi_datum=date(2025, 6, 1)))
    felnott, diak = out.iloc[0], out.iloc[1]
    assert felnott["keses_napok"] == 10 and felnott["kesedelmi_dij"] == 500
    assert diak["keses_napok"] == 10 and diak["kesedelmi_dij"] == 250
    assert felnott["statusz"] == "visszahozva"
    assert felnott["kolcsonzesi_napok"] == 38


def test_T06_dijplafon_es_kesesben_statusz(tmp_path):
    df = make_rows({"kolcsonzes_datum": "2024-12-01", "lejarat_datum": "2025-01-01", "visszahozas_datum": ""})
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    out = transform(lr.valid, TransformSettings(viszonyitasi_datum=date(2025, 6, 1)))
    row = out.iloc[0]
    assert row["keses_napok"] == 151
    assert row["kesedelmi_dij"] == 3000      # 151 × 50 = 7550 Ft, de a plafon 3000 Ft
    assert row["statusz"] == "késésben"


def test_T07_szetbontas_harom_fiokra(tmp_path):
    f = [("F01", "Központi Könyvtár")] * 4 + [("F02", "Belvárosi Fiók")] * 3 + [("F03", "Kertvárosi Fiók")] * 3
    df = make_rows(*[{"fiok_kod": k, "fiok_nev": n} for k, n in f])
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    data = transform(lr.valid, TransformSettings(viszonyitasi_datum=date(2025, 6, 1)))
    out_dir = tmp_path / "ki"
    files = split(data, SplitConfig(kimeneti_mappa=out_dir), errors=lr.errors)

    assert [x.csoport_kulcs for x in files] == ["F01", "F02", "F03"]
    assert [x.sorok_szama for x in files] == [4, 3, 3]
    written = sorted(p.name for p in out_dir.glob("*.xlsx"))
    assert written == ["KEF_202502-202502_F01_Kozponti_Konyvtar.xlsx", "KEF_202502-202502_F02_Belvarosi_Fiok.xlsx",
                       "KEF_202502-202502_F03_Kertvarosi_Fiok.xlsx", "KEF_202502-202502_OSSZESITO.xlsx"]
    ws = openpyxl.load_workbook(out_dir / written[0])["Kölcsönzések"]
    assert ws.max_row == 1 + 4 and ws["A1"].value == "Kölcsönzés azonosító" and ws.freeze_panes == "A2"
    assert sum(x.sorok_szama for x in files) == lr.valid_rows == 10


def test_T08_ismeretlen_fiokkod(tmp_path):
    df = make_rows({}, {"fiok_kod": "F99", "fiok_nev": "Ismeretlen"})
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    assert lr.valid_rows == 1
    assert lr.errors.iloc[0]["hiba"] == "Ismeretlen fiókkód: F99"


def test_T09_nem_irhato_kimeneti_mappa(tmp_path):
    df = make_rows({})
    lr = load_file(write_xlsx(df, tmp_path / "export.xlsx"))
    data = transform(lr.valid, TransformSettings())
    blocker = tmp_path / "ez_egy_fajl.txt"
    blocker.write_text("nem mappa")
    with pytest.raises(OutputError, match="A kimeneti mappa nem írható"):
        split(data, SplitConfig(kimeneti_mappa=blocker / "kimenet"))


def test_T10_megtekinto_nem_indithat_feldolgozast(db, tmp_path):
    lr = load_file(write_xlsx(make_rows({}), tmp_path / "export.xlsx"))
    viewer = login(db, "megtekinto", "Megtekinto1")
    assert not viewer.can("feldolgozas")
    with pytest.raises(PermissionDenied):
        run_processing(db, viewer, lr, TransformSettings(), SplitConfig(kimeneti_mappa=tmp_path / "ki"))
    assert db.list_runs() == []                    # nem jött létre futtatás
    assert not (tmp_path / "ki").exists()          # és fájl sem
