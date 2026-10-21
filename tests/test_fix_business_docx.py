"""Флаг --fix в gost-business: копия письма с полями и стилем «Обычный» по профилю."""

import json
from pathlib import Path

from docx.shared import Mm, Pt

from test_check_business import Letter, cb


def codes(report):
    return {i.code for i in report.issues}


def check(path, config=None):
    return cb.Checker(Path(path), cb.load_config(config)).run()


def broken():
    letter = Letter()
    s = letter.doc.sections[0]
    s.left_margin, s.right_margin = Mm(15), Mm(5)
    s.bottom_margin = Mm(15)
    normal = letter.doc.styles["Normal"]
    normal.font.name = "Courier New"
    normal.font.size = Pt(11)
    return letter


def test_fix_writes_copy_that_passes(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "letter.docx"))
    assert codes(check(path)) >= {"margin", "font", "size-small"}
    before = path.read_bytes()
    assert cb.main([str(path), "--fix"]) == 0
    out = capsys.readouterr().out
    assert path.read_bytes() == before
    copy = tmp_path / "letter.fixed.docx"
    assert copy.exists()
    assert "ИСПРАВЛЕНО" in out and "ОСТАЛОСЬ ИСПРАВИТЬ ВРУЧНУЮ" in out
    assert "Поле левое 15 мм → 20 мм" in out
    assert "Поле правое 5 мм → 10 мм" in out
    assert "Шрифт стиля «Обычный»: Courier New → Times New Roman" in out
    assert "Размер шрифта стиля «Обычный»: 11 пт → 12 пт" in out
    assert cb.CLEAN in out
    report = check(copy)
    assert not codes(report) & {"margin", "font", "size-small", "size-large"}, \
        [(i.code, i.message) for i in report.issues]


def test_fix_keeps_allowed_font(tmp_path, capsys):
    letter = Letter()
    letter.doc.styles["Normal"].font.name = "Arial"
    path = Path(letter.save(tmp_path / "letter.docx"))
    assert cb.main([str(path), "--fix"]) == 0
    out = capsys.readouterr().out
    assert "ничего: однозначных ошибок для исправления нет" in out


def test_fix_lowers_large_size(tmp_path, capsys):
    letter = Letter()
    letter.doc.styles["Normal"].font.size = Pt(16)
    path = Path(letter.save(tmp_path / "letter.docx"))
    assert "size-large" in codes(check(path))
    cb.main([str(path), "--fix"])
    assert "Размер шрифта стиля «Обычный»: 16 пт → 14 пт" in capsys.readouterr().out
    assert "size-large" not in codes(check(tmp_path / "letter.fixed.docx"))


def test_fix_leaves_indent_and_dates_for_manual(tmp_path, capsys):
    letter = broken()
    letter.para("Просим").paragraph_format.first_line_indent = None
    letter.para("05.03.2026").runs[0].text = "5.3.2026 № 12-04/37"
    path = Path(letter.save(tmp_path / "letter.docx"))
    assert cb.main([str(path), "--fix"]) == 1
    capsys.readouterr()
    report = check(tmp_path / "letter.fixed.docx")
    left = codes(report)
    assert "indent" in left
    assert "date-format" in left
    assert not left & {"margin", "font", "size-small"}


def test_fix_does_not_touch_title_dot(tmp_path):
    letter = Letter()
    letter.para("О поставке").runs[0].text = "О поставке бумаги."
    path = Path(letter.save(tmp_path / "letter.docx"))
    cb.main([str(path), "--fix"])
    assert "title-dot" in codes(check(tmp_path / "letter.fixed.docx"))


def test_fix_does_not_overwrite_without_force(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "letter.docx"))
    copy = tmp_path / "letter.fixed.docx"
    copy.write_bytes(b"old")
    assert cb.main([str(path), "--fix"]) == 2
    assert "--force" in capsys.readouterr().out
    assert copy.read_bytes() == b"old"
    assert cb.main([str(path), "--fix", "--force"]) == 0
    assert copy.read_bytes() != b"old"


def test_fix_json_lists_fixed(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "letter.docx"))
    cb.main([str(path), "--fix", "--json"])
    data = json.loads(capsys.readouterr().out)
    assert data["file"] == "letter.fixed.docx"
    assert any("Поле левое" in item for item in data["fixed"])


def test_without_fix_nothing_changes(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "letter.docx"))
    assert cb.main([str(path)]) == 1
    out = capsys.readouterr().out
    assert "ИСПРАВЛЕНО" not in out
    assert not (tmp_path / "letter.fixed.docx").exists()
    cb.main([str(path), "--json"])
    assert "fixed" not in json.loads(capsys.readouterr().out)


def test_fix_keeps_wide_margin(tmp_path, capsys):
    letter = Letter()
    s = letter.doc.sections[0]
    s.left_margin, s.top_margin = Mm(30), Mm(15)
    path = Path(letter.save(tmp_path / "letter.docx"))
    cb.main([str(path), "--fix"])
    out = capsys.readouterr().out
    assert "Поле верхнее 15 мм → 20 мм" in out
    assert "Поле левое" not in out
    copy = cb.docx.Document(str(tmp_path / "letter.fixed.docx")).sections[0]
    assert round(copy.left_margin / cb.EMU_PER_MM) == 30
