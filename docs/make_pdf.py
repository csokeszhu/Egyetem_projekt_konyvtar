"""A terv elkészítése Word- és PDF-formában:  python docs/make_pdf.py  (Microsoft Word kell hozzá)."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

DOCS = Path(__file__).resolve().parent

subprocess.run([sys.executable, str(DOCS / "build_docs.py")], check=True)
subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File",
                str(DOCS / "export_pdf.ps1")], check=True, timeout=180, stdin=subprocess.DEVNULL)
