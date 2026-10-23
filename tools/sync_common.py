#!/usr/bin/env python3
"""Копирует общий модуль docx_common.py из gost-report в gost-business.

Каждый навык ставится отдельной папкой, поэтому общий код лежит копией в обоих.
Править модуль нужно в gost-report, затем запустить:

    python tools/sync_common.py          # обновить копию
    python tools/sync_common.py --check  # только проверить, что копии совпадают
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "skills" / "gost-report" / "scripts" / "docx_common.py"
COPIES = [ROOT / "skills" / "gost-business" / "scripts" / "docx_common.py"]
COMMAND = "python tools/sync_common.py"


def stale() -> list[Path]:
    """Копии, которые отличаются от модуля в gost-report или отсутствуют."""
    data = SOURCE.read_bytes()
    return [p for p in COPIES if not p.exists() or p.read_bytes() != data]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Обновить копии docx_common.py в навыках")
    parser.add_argument("--check", action="store_true", help="только проверить, ничего не менять")
    args = parser.parse_args(argv)
    outdated = stale()
    if args.check:
        for p in outdated:
            print(f"Отличается от gost-report: {p.relative_to(ROOT)}. Обновить: {COMMAND}")
        return 1 if outdated else 0
    for p in outdated:
        shutil.copyfile(SOURCE, p)
        print(f"Обновлено: {p.relative_to(ROOT)}")
    if not outdated:
        print("Копии уже совпадают.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
