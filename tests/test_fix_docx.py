"""Флаг --fix: исправленная копия рядом с исходным файлом."""

import json
from pathlib import Path

from docx.oxml.ns import qn
from docx.shared import Mm, Pt

import check_docx
from docbuilder import good_document


def codes(report):
    return {i.code for i in report.issues}


def check(path, config=None):
    return check_docx.Checker(Path(path), check_docx.load_config(config)).run()


def broken():
    b = good_document()
    section = b.doc.sections[0]
    section.left_margin, section.right_margin = Mm(20), Mm(10)
    section.top_margin = Mm(25)
    normal = b.doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    return b


def test_fix_writes_copy_that_passes(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "work.docx"))
    before = path.read_bytes()
    assert check_docx.main([str(path), "--fix"]) == 0
    out = capsys.readouterr().out
    assert path.read_bytes() == before
    copy = tmp_path / "work.fixed.docx"
    assert copy.exists()
    assert "ИСПРАВЛЕНО" in out and "ОСТАЛОСЬ ИСПРАВИТЬ ВРУЧНУЮ" in out
    assert "Поле левое 20 мм → 30 мм" in out
    assert "Шрифт стиля «Обычный»: Calibri → Times New Roman" in out
    assert "Размер шрифта стиля «Обычный»: 11 пт → 12 пт" in out
    report = check(copy)
    assert not codes(report) & {"margin", "font", "size-small"}, [(i.code, i.message) for i in report.issues]


def test_fix_uses_profile_size(tmp_path):
    cfg = tmp_path / "m.toml"
    cfg.write_text("[text]\nsize_pt = 14\n", encoding="utf-8")
    b = good_document()
    b.doc.styles["Normal"].font.size = Pt(12)
    path = Path(b.save(tmp_path / "work.docx"))
    assert check_docx.main([str(path), "--fix", "--config", str(cfg)]) == 0
    assert "size" not in codes(check(tmp_path / "work.fixed.docx", cfg))


def test_fix_replaces_theme_font(tmp_path):
    b = good_document()
    fonts = b.doc.styles["Normal"].element.rPr.rFonts
    fonts.set(qn("w:asciiTheme"), "minorHAnsi")
    fonts.set(qn("w:hAnsiTheme"), "minorHAnsi")
    path = Path(b.save(tmp_path / "work.docx"))
    assert "font" in codes(check(path))
    check_docx.main([str(path), "--fix"])
    assert "font" not in codes(check(tmp_path / "work.fixed.docx"))


def test_fix_leaves_direct_formatting_for_manual(tmp_path, capsys):
    b = good_document()
    p = [x for x in b.doc.paragraphs if len(x.text) > 60][-1]
    for run in p.runs:
        run.font.name = "Arial"
    path = Path(b.save(tmp_path / "work.docx"))
    assert check_docx.main([str(path), "--fix"]) == 1
    out = capsys.readouterr().out
    assert "ничего: однозначных ошибок для исправления нет" in out
    assert "Шрифт Arial вместо Times New Roman" in out


def test_fix_does_not_overwrite_without_force(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "work.docx"))
    copy = tmp_path / "work.fixed.docx"
    copy.write_bytes(b"old")
    assert check_docx.main([str(path), "--fix"]) == 2
    assert copy.read_bytes() == b"old"
    assert "--force" in capsys.readouterr().out
    assert check_docx.main([str(path), "--fix", "--force"]) == 0
    assert copy.read_bytes() != b"old"


def test_fix_json_lists_fixed(tmp_path, capsys):
    path = Path(broken().save(tmp_path / "work.docx"))
    check_docx.main([str(path), "--fix", "--json"])
    data = json.loads(capsys.readouterr().out)
    assert data["file"] == "work.fixed.docx"
    assert any("Поле правое" in item for item in data["fixed"])
    assert data["errors"] == 0


def test_without_fix_no_copy(tmp_path):
    path = Path(broken().save(tmp_path / "work.docx"))
    assert check_docx.main([str(path)]) == 1
    assert not (tmp_path / "work.fixed.docx").exists()
