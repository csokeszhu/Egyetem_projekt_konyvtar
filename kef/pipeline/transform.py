"""2. lépés – „formázás”: a valós megoldásban ezt a VBA-makró végezte (új oszlopok, tisztítás)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import numpy as np
import pandas as pd

from kef import config


@dataclass
class TransformSettings:
    dij_per_nap: int = 50
    dij_plafon: int = 3000
    kedvezmeny_szazalek: int = 50
    viszonyitasi_datum: date = field(default_factory=date.today)
    szamitott_oszlopok: list[str] = field(default_factory=lambda: list(config.COMPUTED_COLUMNS))

    def validate(self) -> None:
        if not (0 <= self.dij_per_nap <= 10_000):
            raise ValueError("A késedelmi díj/nap 0 és 10 000 Ft között lehet.")
        if not (0 <= self.dij_plafon <= 100_000):
            raise ValueError("A díjplafon 0 és 100 000 Ft között lehet.")
        if not (0 <= self.kedvezmeny_szazalek <= 100):
            raise ValueError("A kedvezmény 0 és 100% között lehet.")
        unknown = [c for c in self.szamitott_oszlopok if c not in config.COMPUTED_COLUMNS]
        if unknown:
            raise ValueError("Ismeretlen számított oszlop: " + ", ".join(unknown))

    @classmethod
    def from_store(cls, values: dict[str, str], ref_date: date | None = None) -> "TransformSettings":
        cols = [c for c in values.get("szamitott_oszlopok", "").split(",") if c]
        return cls(
            dij_per_nap=int(values.get("dij_per_nap", config.DEFAULT_SETTINGS["dij_per_nap"])),
            dij_plafon=int(values.get("dij_plafon", config.DEFAULT_SETTINGS["dij_plafon"])),
            kedvezmeny_szazalek=int(values.get("kedvezmeny_szazalek", config.DEFAULT_SETTINGS["kedvezmeny_szazalek"])),
            viszonyitasi_datum=ref_date or date.today(),
            szamitott_oszlopok=cols,
        )


AGE_BINS = [-1, 14, 18, 35, 64, 200]
AGE_LABELS = ["0–14", "15–18", "19–35", "36–64", "65+"]


def transform(df: pd.DataFrame, settings: TransformSettings) -> pd.DataFrame:
    """Hozzáadja a számított oszlopokat. A bemenetet nem módosítja."""
    settings.validate()
    out = df.copy()
    ref = pd.Timestamp(settings.viszonyitasi_datum)

    loan = out["kolcsonzes_datum"]
    due = out["lejarat_datum"]
    ret = out["visszahozas_datum"]
    end = ret.fillna(ref)

    kolcs_napok = (end - loan).dt.days.clip(lower=0)
    keses = (end - due).dt.days.clip(lower=0)
    discount = np.where(out["olvaso_tipus"].isin(["diák", "nyugdíjas"]), settings.kedvezmeny_szazalek / 100, 0.0)
    dij = np.minimum(keses * settings.dij_per_nap, settings.dij_plafon) * (1 - discount)
    statusz = np.where(ret.notna(), config.STATUS_RETURNED,
                       np.where(keses > 0, config.STATUS_LATE, config.STATUS_OUT))
    age = ref.year - out["olvaso_szul_ev"].astype("Float64")
    korcsoport = pd.cut(age.astype(float), bins=AGE_BINS, labels=AGE_LABELS).astype(object).where(age.notna(), "ismeretlen")

    computed = {
        "kolcsonzesi_napok": kolcs_napok.astype(int),
        "keses_napok": keses.astype(int),
        "kesedelmi_dij": np.round(dij).astype(int),
        "statusz": statusz,
        "ev_honap": loan.dt.strftime("%Y-%m"),
        "korcsoport": korcsoport,
    }
    for name in config.COMPUTED_COLUMNS:
        if name in settings.szamitott_oszlopok:
            out[name] = computed[name]

    # tisztítás: egységes nagybetűs kezdés a címben, felesleges szóközök nélkül
    for col in ("olvaso_nev", "cim", "szerzo", "kiado", "megjegyzes"):
        out[col] = out[col].str.replace(r"\s+", " ", regex=True).str.strip()

    return out.sort_values(["fiok_kod", "kolcsonzes_datum", "kolcsonzes_id"]).reset_index(drop=True)
