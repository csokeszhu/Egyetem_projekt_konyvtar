"""A dokumentáció teljes előállítása: docx → PDF, oldalszámos tartalomjegyzékkel (két menetben).

    python docs/make_pdf.py

1. build_docs.py (tartalomjegyzék oldalszámok nélkül) → export_pdf.ps1 (Word)
2. a fejezetcímek oldalszámának kiolvasása a PDF-ből (PyMuPDF)
3. build_docs.py --pages (oldalszámokkal) → export_pdf.ps1
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pymupdf

DOCS = Path(__file__).resolve().parent
PDF = DOCS / "KEF_dokumentacio.pdf"


def run(*args: str, timeout: int = 300) -> None:
    subprocess.run(args, check=True, timeout=timeout, stdin=subprocess.DEVNULL)


def build(with_pages: bool) -> None:
    run(sys.executable, str(DOCS / "build_docs.py"), *(["--pages"] if with_pages else []))


def export() -> None:
    run("powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(DOCS / "export_pdf.ps1"), timeout=180)


def find_pages() -> dict[str, int]:
    headings = json.loads((DOCS / "_toc_headings.json").read_text(encoding="utf-8"))
    doc = pymupdf.open(PDF)
    pages: dict[str, int] = {}
    start = 2  # a címlap és a tartalomjegyzék utáni oldaltól keresünk (0-tól számozva)
    for h in headings:
        for i in range(start, doc.page_count):
            if doc[i].search_for(h):
                pages[h] = i + 1
                start = i
                break
    missing = [h for h in headings if h not in pages]
    if missing:
        print("Figyelem, nem található a PDF-ben:", missing)
    return pages


def main() -> None:
    build(False)
    export()
    pages = find_pages()
    (DOCS / "_toc_pages.json").write_text(json.dumps(pages, ensure_ascii=False, indent=1), encoding="utf-8")
    build(True)
    export()
    if find_pages() != pages:
        print("Figyelem: a második menetben eltolódtak az oldalszámok.")
    print("Kész:", PDF)


if __name__ == "__main__":
    main()
