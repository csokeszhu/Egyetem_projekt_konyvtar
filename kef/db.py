"""SQLite adatbázis: felhasználók, futtatások, kimeneti fájlok, beállítások."""
from __future__ import annotations

import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from kef import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS felhasznalo (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    azonosito      TEXT    NOT NULL UNIQUE,
    nev            TEXT    NOT NULL,
    jelszo_hash    TEXT    NOT NULL,
    szerepkor      TEXT    NOT NULL CHECK (szerepkor IN ('admin', 'feldolgozo', 'megtekinto')),
    aktiv          INTEGER NOT NULL DEFAULT 1,
    letrehozva     TEXT    NOT NULL,
    utolso_belepes TEXT
);
CREATE TABLE IF NOT EXISTS futtatas (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    felhasznalo_id  INTEGER NOT NULL REFERENCES felhasznalo(id),
    inditas         TEXT    NOT NULL,
    befejezes       TEXT,
    forrasfajl      TEXT    NOT NULL,
    sorok_be        INTEGER NOT NULL,
    ervenyes_sorok  INTEGER NOT NULL,
    hibas_sorok     INTEGER NOT NULL,
    szempont        TEXT    NOT NULL,
    kimeneti_mappa  TEXT    NOT NULL,
    statusz         TEXT    NOT NULL CHECK (statusz IN ('fut', 'sikeres', 'hibas', 'megszakitva')),
    uzenet          TEXT
);
CREATE TABLE IF NOT EXISTS kimeneti_fajl (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    futtatas_id   INTEGER NOT NULL REFERENCES futtatas(id) ON DELETE CASCADE,
    csoport_kulcs TEXT    NOT NULL,
    csoport_nev   TEXT,
    fajlnev       TEXT    NOT NULL,
    sorok_szama   INTEGER NOT NULL,
    kesesben      INTEGER NOT NULL DEFAULT 0,
    osszeg_dij    INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS beallitas (
    kulcs       TEXT PRIMARY KEY,
    ertek       TEXT NOT NULL,
    modositotta INTEGER REFERENCES felhasznalo(id),
    modositva   TEXT
);
"""

# Helyi teszt-/bemutató felhasználók (első indításkor jönnek létre, lásd README).
SEED_USERS = [
    ("admin", "Rendszergazda", "Admin123!", "admin"),
    ("feldolgozo", "Kovács Anna", "Feldolgozo1", "feldolgozo"),
    ("megtekinto", "Tóth Péter", "Megtekinto1", "megtekinto"),
]


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Database:
    def __init__(self, path: str | Path = config.DB_PATH, seed: bool = True) -> None:
        self.path = Path(path)
        if str(path) != ":memory:":
            self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(str(path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        if seed:
            self._seed()

    # --- alacsony szintű segédek ---
    def execute(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Cursor:
        with self._lock:
            cur = self.conn.execute(sql, tuple(params))
            self.conn.commit()
            return cur

    def query(self, sql: str, params: Iterable[Any] = ()) -> list[sqlite3.Row]:
        with self._lock:
            return self.conn.execute(sql, tuple(params)).fetchall()

    def close(self) -> None:
        self.conn.close()

    def _seed(self) -> None:
        from kef.auth import hash_password  # körkörös import elkerülése

        if self.query("SELECT 1 FROM felhasznalo LIMIT 1"):
            return
        for azon, nev, jelszo, szerep in SEED_USERS:
            self.execute(
                "INSERT INTO felhasznalo (azonosito, nev, jelszo_hash, szerepkor, letrehozva) VALUES (?,?,?,?,?)",
                (azon, nev, hash_password(jelszo), szerep, now()),
            )

    # --- felhasználók ---
    def get_user_by_login(self, azonosito: str) -> sqlite3.Row | None:
        rows = self.query("SELECT * FROM felhasznalo WHERE azonosito = ?", (azonosito,))
        return rows[0] if rows else None

    def get_user(self, user_id: int) -> sqlite3.Row | None:
        rows = self.query("SELECT * FROM felhasznalo WHERE id = ?", (user_id,))
        return rows[0] if rows else None

    def list_users(self) -> list[sqlite3.Row]:
        return self.query("SELECT * FROM felhasznalo ORDER BY azonosito")

    # --- beállítások ---
    def get_settings(self) -> dict[str, str]:
        values = dict(config.DEFAULT_SETTINGS)
        for row in self.query("SELECT kulcs, ertek FROM beallitas"):
            values[row["kulcs"]] = row["ertek"]
        return values

    def save_settings(self, values: dict[str, str], user_id: int) -> None:
        ts = now()
        for k, v in values.items():
            self.execute(
                "INSERT INTO beallitas (kulcs, ertek, modositotta, modositva) VALUES (?,?,?,?) "
                "ON CONFLICT(kulcs) DO UPDATE SET ertek=excluded.ertek, modositotta=excluded.modositotta, "
                "modositva=excluded.modositva",
                (k, str(v), user_id, ts),
            )

    # --- futtatások ---
    def start_run(self, user_id: int, forrasfajl: str, sorok_be: int, ervenyes: int, hibas: int,
                  szempont: str, mappa: str) -> int:
        cur = self.execute(
            "INSERT INTO futtatas (felhasznalo_id, inditas, forrasfajl, sorok_be, ervenyes_sorok, hibas_sorok, "
            "szempont, kimeneti_mappa, statusz) VALUES (?,?,?,?,?,?,?,?, 'fut')",
            (user_id, now(), forrasfajl, sorok_be, ervenyes, hibas, szempont, mappa),
        )
        return int(cur.lastrowid)

    def finish_run(self, run_id: int, statusz: str, uzenet: str = "") -> None:
        self.execute("UPDATE futtatas SET befejezes = ?, statusz = ?, uzenet = ? WHERE id = ?",
                     (now(), statusz, uzenet, run_id))

    def add_output_files(self, run_id: int, files: Iterable[Any]) -> None:
        for f in files:
            self.execute(
                "INSERT INTO kimeneti_fajl (futtatas_id, csoport_kulcs, csoport_nev, fajlnev, sorok_szama, "
                "kesesben, osszeg_dij) VALUES (?,?,?,?,?,?,?)",
                (run_id, f.csoport_kulcs, f.csoport_nev, f.fajlnev, f.sorok_szama, f.kesesben, f.osszeg_dij),
            )

    def list_runs(self, tol: str | None = None, ig: str | None = None, user_id: int | None = None) -> list[sqlite3.Row]:
        sql = ("SELECT f.*, u.azonosito AS felhasznalo, "
               "(SELECT COUNT(*) FROM kimeneti_fajl k WHERE k.futtatas_id = f.id) AS fajlok "
               "FROM futtatas f JOIN felhasznalo u ON u.id = f.felhasznalo_id WHERE 1=1")
        params: list[Any] = []
        if tol:
            sql += " AND f.inditas >= ?"
            params.append(tol)
        if ig:
            sql += " AND f.inditas <= ?"
            params.append(ig + " 23:59:59")
        if user_id:
            sql += " AND f.felhasznalo_id = ?"
            params.append(user_id)
        return self.query(sql + " ORDER BY f.id DESC", params)

    def list_output_files(self, run_id: int) -> list[sqlite3.Row]:
        return self.query("SELECT * FROM kimeneti_fajl WHERE futtatas_id = ? ORDER BY csoport_kulcs", (run_id,))
