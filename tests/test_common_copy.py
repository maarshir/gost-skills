"""Копия docx_common.py в gost-business и запуск навыка отдельной папкой."""

import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document

from test_check_business import Letter
from test_fix_business_docx import broken

ROOT = Path(__file__).resolve().parent.parent
SYNC = ROOT / "tools" / "sync_common.py"
SOURCE = ROOT / "skills" / "gost-report" / "scripts" / "docx_common.py"
COPY = ROOT / "skills" / "gost-business" / "scripts" / "docx_common.py"


def run(args, cwd):
    return subprocess.run([sys.executable, *map(str, args)], cwd=cwd,
                          capture_output=True, text=True, encoding="utf-8")


def test_copies_are_identical():
    assert COPY.read_bytes() == SOURCE.read_bytes(), (
        "docx_common.py в gost-business отличается от gost-report. "
        "Править модуль в gost-report, затем: python tools/sync_common.py")


def test_sync_check_passes():
    result = run([SYNC, "--check"], ROOT)
    assert result.returncode == 0, result.stdout + result.stderr


def test_sync_check_names_command(tmp_path):
    """В копии репозитория портим копию модуля: --check называет команду, без флага чинит."""
    repo = tmp_path / "repo"
    for part in ("tools", "skills/gost-report/scripts", "skills/gost-business/scripts"):
        (repo / part).mkdir(parents=True)
    shutil.copy(SYNC, repo / "tools")
    shutil.copy(SOURCE, repo / "skills/gost-report/scripts")
    (repo / "skills/gost-business/scripts/docx_common.py").write_text("# старая копия\n", encoding="utf-8")
    result = run([repo / "tools" / "sync_common.py", "--check"], repo)
    assert result.returncode == 1
    assert "python tools/sync_common.py" in result.stdout
    assert run([repo / "tools" / "sync_common.py"], repo).returncode == 0
    assert (repo / "skills/gost-business/scripts/docx_common.py").read_bytes() == SOURCE.read_bytes()


def test_business_runs_without_other_skills(tmp_path):
    """Только папка gost-business: проверка и --fix работают, копия проходит проверку."""
    skill = tmp_path / "skills" / "gost-business"
    shutil.copytree(ROOT / "skills" / "gost-business", skill,
                    ignore=shutil.ignore_patterns("__pycache__"))
    script = skill / "scripts" / "check_docx.py"
    path = Path(broken().save(tmp_path / "letter.docx"))
    result = run([script, path, "--fix"], tmp_path)
    assert "Traceback" not in result.stderr, result.stderr
    assert "ИСПРАВЛЕНО" in result.stdout
    copy = tmp_path / "letter.fixed.docx"
    Document(copy)
    again = run([script, copy], tmp_path)
    assert again.returncode == 0, again.stdout + again.stderr
    assert "состав и порядок реквизитов" in again.stdout


def test_business_good_letter_from_own_folder(tmp_path):
    skill = tmp_path / "gost-business"
    shutil.copytree(ROOT / "skills" / "gost-business", skill,
                    ignore=shutil.ignore_patterns("__pycache__"))
    path = Path(Letter().save(tmp_path / "letter.docx"))
    result = run([skill / "scripts" / "check_docx.py", path], tmp_path)
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Нарушений не найдено" in result.stdout
