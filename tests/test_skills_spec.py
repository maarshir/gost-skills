"""Все навыки соответствуют спецификации agentskills.io/specification.

Проверяем то, из-за чего навык не загрузится или загрузится не тогда:
поля и их длину, имя папки, размер SKILL.md и ссылки на файлы внутри навыка.
"""

import json
import re
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parent.parent
SKILLS = sorted(p.parent for p in (ROOT / "skills").glob("*/SKILL.md"))
ALLOWED = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")


def split(skill: Path) -> tuple[dict, str]:
    text = (skill / "SKILL.md").read_text(encoding="utf-8")
    assert text.startswith("---\n"), "SKILL.md должен начинаться с YAML между ---"
    _, head, body = text.split("---\n", 2)
    return yaml.safe_load(head), body


def test_there_are_skills():
    assert SKILLS


@pytest.mark.parametrize("skill", SKILLS, ids=lambda p: p.name)
def test_frontmatter(skill):
    meta, _ = split(skill)
    assert set(meta) <= ALLOWED, f"лишние поля: {set(meta) - ALLOWED}"
    name = meta["name"]
    assert 1 <= len(name) <= 64 and NAME_RE.match(name), name
    assert name == skill.name, "имя должно совпадать с папкой"
    desc = meta["description"]
    assert isinstance(desc, str) and 1 <= len(desc) <= 1024
    if "compatibility" in meta:
        assert 1 <= len(meta["compatibility"]) <= 500
    if "metadata" in meta:
        assert all(isinstance(k, str) and isinstance(v, str) for k, v in meta["metadata"].items())


@pytest.mark.parametrize("skill", SKILLS, ids=lambda p: p.name)
def test_body_is_short(skill):
    _, body = split(skill)
    assert len(body.splitlines()) < 500
    assert len(body) < 20_000  # около 5000 токенов, как советует спецификация


@pytest.mark.parametrize("skill", SKILLS, ids=lambda p: p.name)
def test_referenced_files_exist(skill):
    for md in skill.rglob("*.md"):
        text = md.read_text(encoding="utf-8")
        links = re.findall(r"\]\(([^)#]+)\)", text)
        paths = re.findall(r"`((?:scripts|assets|references)/[^`\s]+)`", text)
        for link in links:
            if link.startswith(("http://", "https://", "mailto:")):
                continue
            assert (md.parent / link).exists(), f"{md.name}: нет файла {link}"
        for path in paths:
            assert (skill / path).exists(), f"{md.name}: нет файла {path}"


def test_marketplace_lists_every_skill():
    data = json.loads((ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
    listed = {Path(s).name for plugin in data["plugins"] for s in plugin["skills"]}
    assert listed == {s.name for s in SKILLS}


def test_every_issue_code_is_documented():
    """Каждый код из отчёта скрипта объяснён в formatting.md и в fix-in-word.md."""
    skill = ROOT / "skills" / "gost-report"
    source = (skill / "scripts" / "check_docx.py").read_text(encoding="utf-8")
    codes = set()
    for code in re.findall(r"(?:ERROR|WARNING),\s*f?\"([a-z_{}\-]+)\"", source):
        if "{kind}" in code:
            codes |= {code.replace("{kind}", k) for k in ("figure", "table")}
        elif "{code}" in code:
            codes |= {code.replace("{code}", k) for k in ("figure", "table")}
        else:
            codes.add(code)
    assert len(codes) > 30
    formatting = (skill / "references" / "formatting.md").read_text(encoding="utf-8")
    fixes = (skill / "references" / "fix-in-word.md").read_text(encoding="utf-8")
    for code in codes:
        assert f"`{code}`" in formatting, f"{code} не описан в formatting.md"
        prefix = code.split("-")[0]
        assert (f"`{code}`" in fixes or f"`{prefix}-*`" in fixes or
                f"`*-{code.split('-', 1)[-1]}`" in fixes), f"{code} без способа исправить в fix-in-word.md"


def test_every_bibliography_code_is_documented():
    """Каждый код проверки списка литературы объяснён в rules.md."""
    skill = ROOT / "skills" / "gost-bibliography"
    source = (skill / "scripts" / "check_bibliography.py").read_text(encoding="utf-8")
    block = source.split("RULES = {", 1)[1].split("\n}\n", 1)[0]
    codes = set(re.findall(r'^    "([a-z\-]+)":', block, re.M))
    assert len(codes) > 30
    rules = (skill / "references" / "rules.md").read_text(encoding="utf-8")
    for code in codes:
        assert f"`{code}`" in rules, f"{code} не описан в rules.md"
    documented = set(re.findall(r"^\| `([a-z\-]+)` \|", rules, re.M))
    assert documented <= codes, f"в rules.md есть коды, которых нет в скрипте: {documented - codes}"
