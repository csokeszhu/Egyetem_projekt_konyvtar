"""A teljes feldolgozás összefogása (jogosultság-ellenőrzés + naplózás az adatbázisba)."""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass

from kef.auth import User, require
from kef.db import Database
from kef.pipeline import Cancelled, LoadResult, OutputFile, SplitConfig, TransformSettings, split, transform
from kef.pipeline.splitter import ProgressCb


@dataclass
class RunResult:
    run_id: int
    files: list[OutputFile]
    seconds: float

    @property
    def total_rows(self) -> int:
        return sum(f.sorok_szama for f in self.files)

    @property
    def total_fee(self) -> int:
        return sum(f.osszeg_dij for f in self.files)


def run_processing(
    db: Database,
    user: User,
    load: LoadResult,
    settings: TransformSettings,
    cfg: SplitConfig,
    progress: ProgressCb | None = None,
    cancel: threading.Event | None = None,
) -> RunResult:
    """Formázás + szétbontás; a futtatást és a létrejött fájlokat naplózza."""
    require(user, "feldolgozas")
    started = time.perf_counter()
    run_id = db.start_run(user.id, str(load.path), load.total_rows, load.valid_rows, load.error_rows,
                          cfg.szempont, str(cfg.kimeneti_mappa))
    try:
        if progress:
            progress(0, 1, f"Formázás: {len(settings.szamitott_oszlopok)} számított oszlop hozzáadása…")
        df = transform(load.valid, settings)
        files = split(df, cfg, errors=load.errors, progress=progress, cancel=cancel)
    except Cancelled:
        db.finish_run(run_id, "megszakitva", "A felhasználó megszakította; a részleges fájlok törölve.")
        raise
    except Exception as exc:
        db.finish_run(run_id, "hibas", str(exc))
        raise
    db.add_output_files(run_id, files)
    seconds = time.perf_counter() - started
    rows = f"{sum(f.sorok_szama for f in files):,}".replace(",", " ")
    db.finish_run(run_id, "sikeres", f"{len(files)} fájl, {rows} sor, {seconds:.1f} mp")
    return RunResult(run_id=run_id, files=files, seconds=seconds)
