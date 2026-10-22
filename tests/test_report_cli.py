"""Запуск check_docx.py gost-report как скрипта, а не через импорт."""

import subprocess
import sys
from pathlib import Path

from docbuilder import good_document

SCRIPT = Path(__file__).resolve().parent.parent / "skills" / "gost-report" / "scripts" / "check_docx.py"


def test_runs_from_another_folder(tmp_path):
    """Скрипт находит общую часть docx_common.py рядом с собой, а не в текущей папке."""
    path = tmp_path / "good.docx"
    good_document().save(path)
    work = tmp_path / "other"
    work.mkdir()
    result = subprocess.run([sys.executable, str(SCRIPT), str(path)], cwd=work,
                            capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Нарушений не найдено" in result.stdout
