"""Show who the files in this folder say they were made by.

Usage:
    python tools\\check_metadata.py

Prints the author fields of every spreadsheet and PDF, the size of every screenshot,
and checks that the Windows user name of whoever ran this appears in none of them.
The user name itself is never printed.
"""

from __future__ import annotations

import getpass
import sys
import zipfile
from pathlib import Path

import pdfplumber
from openpyxl import load_workbook
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
EXPECTED = "AZG Engineering"
SKIP = {".venv", ".git", ".pytest_tmp", "_work", "__pycache__"}


def files(pattern: str) -> list[Path]:
    return sorted(p for p in ROOT.rglob(pattern) if not SKIP.intersection(p.relative_to(ROOT).parts))


def contains_user_name(path: Path, needle: bytes) -> bool:
    data = path.read_bytes()
    if needle in data.lower():
        return True
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            return any(needle in archive.read(name).lower() for name in archive.namelist())
    return False


def main() -> int:
    problems = 0
    needle = getpass.getuser().lower().encode()
    checked = []

    for path in files("*.xlsx"):
        properties = load_workbook(path).properties
        macro_parts = [n for n in zipfile.ZipFile(path).namelist() if "vba" in n.lower()]
        good = properties.creator == EXPECTED and properties.lastModifiedBy == EXPECTED and not macro_parts
        problems += not good
        checked.append(path)
        print(f"{'OK ' if good else 'BAD'} {path.relative_to(ROOT)}: author={properties.creator!r} "
              f"last modified by={properties.lastModifiedBy!r} macros={'yes' if macro_parts else 'none'}")

    for path in files("*.pdf"):
        with pdfplumber.open(path) as pdf:
            author = pdf.metadata.get("Author")
        good = author == EXPECTED
        problems += not good
        checked.append(path)
        print(f"{'OK ' if good else 'BAD'} {path.relative_to(ROOT)}: author={author!r}")

    for path in files("*.png"):
        with Image.open(path) as image:
            size, extra = image.size, sorted(image.info)
        good = size == (1600, 1200) and not extra
        problems += not good
        checked.append(path)
        print(f"{'OK ' if good else 'BAD'} {path.relative_to(ROOT)}: {size[0]}x{size[1]} embedded text fields={extra or 'none'}")

    leaks = [p for p in checked if contains_user_name(p, needle)]
    problems += len(leaks)
    print(f"Windows user name found in: {[str(p.relative_to(ROOT)) for p in leaks] or 'no file'} "
          f"({len(checked)} files checked)")
    print("RESULT:", "all good" if not problems else f"{problems} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
