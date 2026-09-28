#!/usr/bin/env python3
"""Проверка оформления .docx по ГОСТ 7.32-2017 или по своей методичке.

Скрипт ничего не меняет в файле. Он читает разметку документа и печатает,
что не так и где: номер абзаца, начало его текста и раздел, в котором он стоит.

    python check_docx.py работа.docx
    python check_docx.py работа.docx --config методичка.toml
    python check_docx.py работа.docx --json

Код выхода: 0 если ошибок нет (замечания допустимы), 1 если есть ошибки,
2 если файл не удалось прочитать.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib

import docx
from lxml import etree
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

if hasattr(sys.stdout, "reconfigure"):
    # Консоль Windows по умолчанию не в UTF-8, а отчёт должен напечататься всегда.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(__file__).resolve().parent
DEFAULT_PROFILE = HERE.parent / "assets" / "gost-7.32.toml"

ERROR, WARNING = "ошибка", "замечание"
TWIPS_PER_CM = 566.93
EMU_PER_MM = 36000

# Буквы для приложений: кириллица без Ё, З, Й, О, Ч, Ъ, Ы, Ь (ГОСТ 7.32, 6.17)
APPENDIX_LETTERS = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"

# Названия, которые часто пишут вместо принятых в ГОСТ 7.32
ALIASES = {
    "ОГЛАВЛЕНИЕ": "СОДЕРЖАНИЕ",
    "СПИСОК ЛИТЕРАТУРЫ": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "БИБЛИОГРАФИЧЕСКИЙ СПИСОК": "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
}
BIBLIOGRAPHY = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"

APPENDIX_RE = re.compile(r"^ПРИЛОЖЕНИЕ\s+([А-ЯЁA-Z])(?=\s|$|\()")
SECTION_RE = re.compile(r"^(\d+(?:\.\d+)*)(\.?)\s+(\S.*)$")
FIGURE_RE = re.compile(r"^(Рисунок|Рис\.?)\s*((?:[А-ЯЁA-Z]\.)?\d+(?:\.\d+)?)\s*(.*)$")
TABLE_RE = re.compile(r"^(Таблица|Табл\.?)\s*((?:[А-ЯЁA-Z]\.)?\d+(?:\.\d+)?)\s*(.*)$")
NUMBER_TOKEN = r"(?:[А-ЯЁA-Z]\.)?\d+(?:\.\d+)?"
FIG_REF_RE = re.compile(
    r"\b(?:рисун\w*|рис\.)\s*(" + NUMBER_TOKEN + r"(?:\s*(?:,|и|–|—|-)\s*" + NUMBER_TOKEN + r")*)",
    re.IGNORECASE,
)
TAB_REF_RE = re.compile(
    r"\b(?:таблиц\w*|табл\.)\s*(" + NUMBER_TOKEN + r"(?:\s*(?:,|и|–|—|-)\s*" + NUMBER_TOKEN + r")*)",
    re.IGNORECASE,
)
CITE_RE = re.compile(r"\[([^\[\]]{1,80})\]")
CAPTION_SEPARATORS = "–—-.:"


# ---------------------------------------------------------------- настройки

def deep_merge(base: dict, extra: dict) -> dict:
    out = dict(base)
    for key, value in extra.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = deep_merge(out[key], value)
        else:
            out[key] = value
    return out


def load_config(path: Path | None = None) -> dict:
    with open(DEFAULT_PROFILE, "rb") as f:
        config = tomllib.load(f)
    if path is not None:
        with open(path, "rb") as f:
            config = deep_merge(config, tomllib.load(f))
    return config


# ---------------------------------------------------------------- отчёт

@dataclass
class Place:
    paragraph: int | None = None
    text: str = ""
    section: str = ""

    def describe(self) -> str:
        parts = []
        if self.paragraph is not None:
            parts.append(f"абзац {self.paragraph}")
        if self.text:
            parts.append(f"«{self.text}»")
        if self.section:
            parts.append(f"(раздел «{self.section}»)")
        return " ".join(parts)


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    rule: str = ""
    places: list[Place] = field(default_factory=list)


class Report:
    def __init__(self) -> None:
        self._issues: dict[tuple, Issue] = {}
        self.notes: list[str] = []

    def add(self, severity, code, message, rule="", place: Place | None = None):
        key = (severity, code, message)
        issue = self._issues.get(key)
        if issue is None:
            issue = self._issues[key] = Issue(severity, code, message, rule)
        if place is not None:
            issue.places.append(place)

    @property
    def issues(self) -> list[Issue]:
        order = {ERROR: 0, WARNING: 1}
        return sorted(self._issues.values(), key=lambda i: order[i.severity])

    def count(self, severity: str) -> int:
        return sum(1 for i in self._issues.values() if i.severity == severity)


# ---------------------------------------------------------------- разметка

def is_on(el) -> bool | None:
    """Значение флага вроде <w:b/>: нет элемента → None, есть → True, val=0 → False."""
    if el is None:
        return None
    val = el.get(qn("w:val"))
    return val is None or val.lower() not in ("0", "false", "off")


class Styles:
    """Итоговые свойства абзаца и фрагмента с учётом стилей и значений по умолчанию."""

    def __init__(self, document) -> None:
        root = document.styles.element
        self.by_id = {s.get(qn("w:styleId")): s for s in root.findall(qn("w:style"))}
        self.default_para = next(
            (s for s in root.findall(qn("w:style"))
             if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true", "on")),
            None,
        )
        defaults = root.find(qn("w:docDefaults"))
        self.ppr_default = defaults.find(qn("w:pPrDefault") + "/" + qn("w:pPr")) if defaults is not None else None
        self.rpr_default = defaults.find(qn("w:rPrDefault") + "/" + qn("w:rPr")) if defaults is not None else None
        self.theme = self._theme_fonts(document)

    @staticmethod
    def _theme_fonts(document) -> dict:
        fonts = {}
        for rel in document.part.rels.values():
            if rel.reltype.endswith("/theme"):
                root = etree.fromstring(rel.target_part.blob)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                for kind in ("major", "minor"):
                    latin = root.find(f".//a:{kind}Font/a:latin", ns)
                    if latin is not None:
                        fonts[kind] = latin.get("typeface")
        return fonts

    def _chain(self, style_id: str | None):
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            style = self.by_id.get(style_id)
            if style is None:
                return
            yield style
            based = style.find(qn("w:basedOn"))
            style_id = based.get(qn("w:val")) if based is not None else None

    def para_styles(self, p) -> list:
        ppr = p.find(qn("w:pPr"))
        ps = ppr.find(qn("w:pStyle")) if ppr is not None else None
        if ps is not None:
            return list(self._chain(ps.get(qn("w:val"))))
        if self.default_para is not None:
            return list(self._chain(self.default_para.get(qn("w:styleId"))))
        return []

    def style_name(self, p) -> str:
        chain = self.para_styles(p)
        if not chain:
            return ""
        name = chain[0].find(qn("w:name"))
        return name.get(qn("w:val")) if name is not None else ""

    def para_prop(self, p, getter):
        sources = [p.find(qn("w:pPr"))]
        sources += [s.find(qn("w:pPr")) for s in self.para_styles(p)]
        sources.append(self.ppr_default)
        for src in sources:
            if src is not None:
                value = getter(src)
                if value is not None:
                    return value
        return None

    def run_prop(self, p, r, getter):
        sources = [r.find(qn("w:rPr"))]
        rpr = r.find(qn("w:rPr"))
        rs = rpr.find(qn("w:rStyle")) if rpr is not None else None
        if rs is not None:
            sources += [s.find(qn("w:rPr")) for s in self._chain(rs.get(qn("w:val")))]
        sources += [s.find(qn("w:rPr")) for s in self.para_styles(p)]
        sources.append(self.rpr_default)
        for src in sources:
            if src is not None:
                value = getter(src)
                if value is not None:
                    return value
        return None

    # --- отдельные свойства

    def font(self, p, r) -> str | None:
        def get(rpr):
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None:
                return None
            for attr in ("w:hAnsiTheme", "w:hAnsi", "w:asciiTheme", "w:ascii"):
                val = fonts.get(qn(attr))
                if val:
                    if attr.endswith("Theme"):
                        return self.theme.get("major" if val.startswith("major") else "minor")
                    return val
            return None
        return self.run_prop(p, r, get)

    def size(self, p, r) -> float:
        def get(rpr):
            sz = rpr.find(qn("w:sz"))
            return int(sz.get(qn("w:val"))) / 2 if sz is not None else None
        return self.run_prop(p, r, get) or 10.0  # в Word без указания размера 10 пт

    def bold(self, p, r) -> bool:
        return bool(self.run_prop(p, r, lambda rpr: is_on(rpr.find(qn("w:b")))))

    def caps(self, p, r) -> bool:
        return bool(self.run_prop(p, r, lambda rpr: is_on(rpr.find(qn("w:caps")))))

    def color(self, p, r) -> str | None:
        def get(rpr):
            c = rpr.find(qn("w:color"))
            return c.get(qn("w:val")) if c is not None else None
        return self.run_prop(p, r, get)

    def alignment(self, p) -> str:
        def get(ppr):
            jc = ppr.find(qn("w:jc"))
            return jc.get(qn("w:val")) if jc is not None else None
        value = self.para_prop(p, get) or "left"
        return {"start": "left", "end": "right", "distribute": "both"}.get(value, value)

    def first_line_cm(self, p) -> float | None:
        def get(ppr):
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                return None
            if ind.get(qn("w:firstLineChars")) or ind.get(qn("w:hangingChars")):
                return "chars"
            if ind.get(qn("w:hanging")) is not None:
                return -int(ind.get(qn("w:hanging"))) / TWIPS_PER_CM
            if ind.get(qn("w:firstLine")) is not None:
                return int(ind.get(qn("w:firstLine"))) / TWIPS_PER_CM
            return None
        value = self.para_prop(p, get)
        if value == "chars":
            return None
        return value or 0.0

    def line_spacing(self, p) -> tuple[str, float]:
        """('multiple', 1.5) для обычного интервала или ('exact', 18.0) в пунктах."""
        def get(ppr):
            sp = ppr.find(qn("w:spacing"))
            if sp is None or sp.get(qn("w:line")) is None:
                return None
            line = int(sp.get(qn("w:line")))
            rule = sp.get(qn("w:lineRule")) or "auto"
            if rule == "auto":
                return ("multiple", round(line / 240, 2))
            return (rule, line / 20)
        return self.para_prop(p, get) or ("multiple", 1.0)

    def page_break_before(self, p) -> bool:
        return bool(self.para_prop(p, lambda ppr: is_on(ppr.find(qn("w:pageBreakBefore")))))

    def outline_level(self, p) -> int | None:
        def get(ppr):
            lvl = ppr.find(qn("w:outlineLvl"))
            return int(lvl.get(qn("w:val"))) if lvl is not None else None
        level = self.para_prop(p, get)
        return level if level is not None and level < 9 else None


def para_text(p) -> str:
    parts = []
    for node in p.iter(qn("w:t"), qn("w:tab")):
        parts.append(node.text or "" if node.tag == qn("w:t") else "\t")
    return "".join(parts)


def text_runs(p) -> list:
    return [r for r in p.iter(qn("w:r")) if "".join(t.text or "" for t in r.findall(qn("w:t"))).strip()]


def has_drawing(el) -> bool:
    return any(True for _ in el.iter(qn("w:drawing"), qn("w:pict"), qn("w:object")))


def has_formula(el) -> bool:
    return any(True for _ in el.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))


def has_page_break(p) -> bool:
    return any(br.get(qn("w:type")) == "page" for br in p.iter(qn("w:br")))


def has_section_break(p) -> bool:
    ppr = p.find(qn("w:pPr"))
    if ppr is None or ppr.find(qn("w:sectPr")) is None:
        return False
    kind = ppr.find(qn("w:sectPr")).find(qn("w:type"))
    return kind is None or kind.get(qn("w:val")) != "continuous"


def short(text: str, limit: int = 40) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def normalize(text: str) -> str:
    return " ".join(text.split()).upper().rstrip(".")


# ---------------------------------------------------------------- проверка

@dataclass
class Block:
    kind: str            # "p" или "tbl"
    el: object
    index: int | None    # номер абзаца для отчёта (таблицы не нумеруются)
    text: str = ""
    role: str = "body"   # title, toc, structural, appendix, heading, figure, table_title, bib, body, empty
    section: str = ""
    match: object = None


class Checker:
    def __init__(self, path: Path, config: dict) -> None:
        self.path = path
        self.config = config
        self.doc = docx.Document(str(path))
        self.styles = Styles(self.doc)
        self.report = Report()
        self.blocks = self._read_blocks()

    # --- чтение и разметка ролей

    def _read_blocks(self) -> list[Block]:
        blocks, n = [], 0
        for el in self.doc.element.body:
            if el.tag == qn("w:p"):
                n += 1
                blocks.append(Block("p", el, n, para_text(el)))
            elif el.tag == qn("w:tbl"):
                blocks.append(Block("tbl", el, None))
        return blocks

    def structural_name(self, text: str) -> str | None:
        norm = normalize(text)
        names = [normalize(s) for s in self.config["headings"]["structural"]]
        if norm in names:
            return norm
        if norm in ALIASES:
            return norm
        return None

    def is_heading_style(self, p) -> bool:
        name = self.styles.style_name(p).lower()
        if name.startswith(("heading", "заголовок")) or name == "title":
            return True
        return self.styles.outline_level(p) is not None

    def is_all_bold(self, p) -> bool:
        runs = text_runs(p)
        return bool(runs) and all(self.styles.bold(p, r) for r in runs)

    def assign_roles(self) -> None:
        # Всё до первого структурного элемента считается титульным листом и не проверяется.
        # Если структурных элементов нет вовсе, проверяем текст с самого начала.
        started = not any(
            b.kind == "p" and (self.structural_name(b.text.strip()) or APPENDIX_RE.match(normalize(b.text)))
            for b in self.blocks
        )
        if started:
            self.report.notes.append("Не найдено ни одного структурного элемента («ВВЕДЕНИЕ», «ЗАКЛЮЧЕНИЕ» и т. п.), "
                                     "поэтому титульный лист не отделён и проверялся как обычный текст.")
        region = ""          # текущий структурный элемент
        section = ""
        for b in self.blocks:
            if b.kind == "tbl":
                b.role = "table" if started else "title"
                b.section = section
                continue
            text = b.text.strip()
            name = self.structural_name(text)
            appendix = APPENDIX_RE.match(normalize(text)) if text else None
            if name or appendix:
                started = True
                b.role = "structural" if name else "appendix"
                b.match = name or appendix
                region = ALIASES.get(name, name) if name else "ПРИЛОЖЕНИЕ"
                section = short(text, 50)
                b.section = section
                continue
            if not started:
                b.role = "title"
                continue
            b.section = section
            if not text:
                b.role = "empty" if not has_drawing(b.el) else "drawing"
                continue
            if region == "СОДЕРЖАНИЕ" or self.styles.style_name(b.el).lower().startswith(("toc", "оглавление")):
                b.role = "toc"
                continue
            fig, tab = FIGURE_RE.match(text), TABLE_RE.match(text)
            if fig and (not fig.group(3) or fig.group(3)[0] in CAPTION_SEPARATORS):
                b.role, b.match = "figure", fig
                continue
            if tab and (not tab.group(3) or tab.group(3)[0] in CAPTION_SEPARATORS):
                b.role, b.match = "table_title", tab
                continue
            if region == BIBLIOGRAPHY:
                b.role = "bib"
                continue
            sec = SECTION_RE.match(text)
            looks_like_heading = self.is_heading_style(b.el) or (
                sec is not None and len(text) <= 200 and self.is_all_bold(b.el)
            )
            if looks_like_heading:
                b.role, b.match = "heading", sec
                section = short(text, 50)
                b.section = section
                continue
            if has_drawing(b.el) and len(text) < 3:
                b.role = "drawing"
                continue
            b.role = "formula" if has_formula(b.el) else "body"

    # --- вспомогательное

    def place(self, b: Block) -> Place:
        return Place(b.index, short(b.text), b.section if b.role not in ("structural", "appendix", "heading") else "")

    def starts_new_page(self, i: int) -> bool:
        b = self.blocks[i]
        if self.styles.page_break_before(b.el):
            return True
        first = next(b.el.iter(qn("w:r")), None)
        if first is not None and has_page_break(first) and not para_text(first).strip():
            return True
        j = i - 1
        while j >= 0:
            prev = self.blocks[j]
            if prev.kind == "tbl":
                return False
            if has_page_break(prev.el) or has_section_break(prev.el):
                return True
            if prev.text.strip() or has_drawing(prev.el):
                return False
            j -= 1
        return True  # начало документа

    def prev_block(self, i: int, skip_empty: bool = True):
        j = i - 1
        while j >= 0:
            b = self.blocks[j]
            if not (skip_empty and b.kind == "p" and b.role == "empty"):
                return b
            j -= 1
        return None

    def next_block(self, i: int, skip_empty: bool = True):
        for b in self.blocks[i + 1:]:
            if not (skip_empty and b.kind == "p" and b.role == "empty"):
                return b
        return None

    # --- проверки

    def run(self) -> Report:
        self.assign_roles()
        self.check_page()
        self.check_page_numbers()
        self.check_structure()
        self.check_headings()
        self.check_text()
        self.check_captions()
        self.check_references()
        self.check_sources()
        return self.report

    def check_page(self) -> None:
        cfg = self.config["page"]
        tol = cfg.get("tolerance_mm", 1)
        for n, section in enumerate(self.doc.sections, start=1):
            if section.page_width is None or section.page_height is None:
                continue
            w, h = section.page_width / EMU_PER_MM, section.page_height / EMU_PER_MM
            where = Place(section=f"часть документа {n}") if len(self.doc.sections) > 1 else None
            landscape = w > h
            if landscape:
                w, h = h, w
            if abs(w - cfg["width_mm"]) > tol or abs(h - cfg["height_mm"]) > tol:
                self.report.add(ERROR, "page-size",
                                f"Размер листа {w:.0f}×{h:.0f} мм, нужен А4 {cfg['width_mm']}×{cfg['height_mm']} мм",
                                "6.1.1", where)
            if landscape:
                self.report.notes.append(f"Часть документа {n} с альбомной ориентацией: поля в ней не проверялись.")
                continue
            margins = {
                "левое": (section.left_margin, cfg["margin_left_mm"]),
                "правое": (section.right_margin, cfg["margin_right_mm"]),
                "верхнее": (section.top_margin, cfg["margin_top_mm"]),
                "нижнее": (section.bottom_margin, cfg["margin_bottom_mm"]),
            }
            for label, (value, need) in margins.items():
                if value is None:
                    continue
                mm = value / EMU_PER_MM
                if abs(mm - need) > tol:
                    self.report.add(ERROR, "margin", f"Поле {label} {mm:.0f} мм, нужно {need} мм", "6.1.1", where)

    def _parts_with_page_field(self, reltype: str) -> list:
        found = []
        for rel in self.doc.part.rels.values():
            if rel.reltype != reltype:
                continue
            el = rel.target_part.element
            instr = " ".join((t.text or "") for t in el.iter(qn("w:instrText")))
            instr += " ".join(f.get(qn("w:instr")) or "" for f in el.iter(qn("w:fldSimple")))
            if re.search(r"\bPAGE\b", instr):
                found.append(el)
        return found

    def check_page_numbers(self) -> None:
        if not self.config["page_numbers"].get("check", True):
            return
        footers = self._parts_with_page_field(RT.FOOTER)
        headers = self._parts_with_page_field(RT.HEADER)
        if not footers and headers:
            self.report.add(ERROR, "page-number-top",
                            "Номер страницы стоит в верхнем колонтитуле, а нужен в центре нижней части", "6.3.1")
            return
        if not footers:
            self.report.add(ERROR, "page-number-missing", "Номера страниц не найдены", "6.3.1")
            return
        for footer in footers:
            for p in footer.iter(qn("w:p")):
                instr = " ".join((t.text or "") for t in p.iter(qn("w:instrText")))
                instr += " ".join(f.get(qn("w:instr")) or "" for f in p.iter(qn("w:fldSimple")))
                if re.search(r"\bPAGE\b", instr) and self.styles.alignment(p) != "center":
                    self.report.add(ERROR, "page-number-align",
                                    "Номер страницы не по центру нижней части страницы", "6.3.1")
        first = self.doc.sections[0]
        title_page_numbered = True
        if first.different_first_page_header_footer:
            fp = first.first_page_footer
            title_page_numbered = (not fp.is_linked_to_previous) and any(
                fp._element is el for el in footers
            )
        if title_page_numbered:
            self.report.add(WARNING, "page-number-title",
                            "На титульном листе будет номер страницы: включите «Особый колонтитул для первой "
                            "страницы» или отдельный раздел для титульного листа", "6.3.1")

    def check_structure(self) -> None:
        cfg = self.config["headings"]
        order = [normalize(s) for s in cfg["structural"]]
        seen: list[str] = []
        appendix_letters = []
        for i, b in enumerate(self.blocks):
            if b.role == "structural":
                name = b.match
                if name in ALIASES and name not in order:
                    self.report.add(WARNING, "alias",
                                    f"«{short(b.text)}»: в ГОСТ 7.32 этот элемент называется «{ALIASES[name]}». "
                                    "Если методичка требует своё название, добавьте его в structural в настройках",
                                    "6.2.1", self.place(b))
                    name = ALIASES[name]
                seen.append(name)
                self._check_centered_title(i, b)
            elif b.role == "appendix":
                appendix_letters.append((b.match.group(1), b))
                self._check_centered_title(i, b)
        for need in cfg.get("required", []):
            if normalize(need) not in seen:
                self.report.add(ERROR, "missing", f"Нет структурного элемента «{normalize(need)}»", "4")
        positions = [order.index(n) for n in seen if n in order]
        if positions != sorted(positions):
            self.report.add(WARNING, "order",
                            "Структурные элементы идут не в том порядке: " + ", ".join(seen), "4")
        expected = iter(APPENDIX_LETTERS)
        for letter, b in appendix_letters:
            letter = letter.upper()
            if letter not in APPENDIX_LETTERS:
                self.report.add(ERROR, "appendix-letter",
                                f"Приложение {letter}: буквы Ё, З, Й, О, Ч, Ъ, Ы, Ь и латиница для приложений "
                                "не используются", "6.17.4", self.place(b))
                continue
            want = next(expected, None)
            if letter != want:
                self.report.add(WARNING, "appendix-order",
                                f"Приложение {letter} идёт вместо ожидаемого {want}", "6.17.4", self.place(b))
                expected = iter(APPENDIX_LETTERS[APPENDIX_LETTERS.index(letter) + 1:])

    def _check_centered_title(self, i: int, b: Block) -> None:
        cfg = self.config["headings"]
        text = b.text.strip()
        runs = text_runs(b.el)
        upper = text.isupper() or (runs and all(self.styles.caps(b.el, r) for r in runs))
        head = text.split("(")[0].strip()
        if not upper and not head.isupper():
            self.report.add(ERROR, "title-case", "Заголовок структурного элемента не прописными буквами",
                            "6.2.1", self.place(b))
        if self.styles.alignment(b.el) != "center":
            self.report.add(ERROR, "title-align", "Заголовок структурного элемента не по центру", "6.2.1",
                            self.place(b))
        if text.endswith("."):
            self.report.add(ERROR, "title-dot", "Точка в конце заголовка", "6.2.1", self.place(b))
        if cfg.get("bold", True) and b.role == "structural" and not self.is_all_bold(b.el):
            self.report.add(ERROR, "title-bold", "Заголовок структурного элемента не полужирный", "6.1.1",
                            self.place(b))
        if cfg.get("new_page", True) and not self.starts_new_page(i):
            self.report.add(ERROR, "new-page", "Структурный элемент или раздел не с новой страницы", "6.2.1",
                            self.place(b))

    def check_headings(self) -> None:
        cfg = self.config["headings"]
        indent = self.config["text"]["indent_cm"]
        tol = self.config["text"]["indent_tolerance_cm"]
        last: list[int] = []
        auto_numbered = False
        for i, b in enumerate(self.blocks):
            if b.role != "heading":
                continue
            text = b.text.strip()
            place = self.place(b)
            if text.endswith("."):
                self.report.add(ERROR, "heading-dot", "Точка в конце заголовка раздела", "6.2.3", place)
            if cfg.get("bold", True) and not self.is_all_bold(b.el):
                self.report.add(ERROR, "heading-bold", "Заголовок раздела не полужирный", "6.2.3", place)
            align = self.styles.alignment(b.el)
            first = self.styles.first_line_cm(b.el)
            if align == "center":
                self.report.add(ERROR, "heading-center",
                                "Заголовок раздела по центру, а его печатают с абзацного отступа", "6.2.3", place)
            elif first is not None and abs(first - indent) > tol and self._numbering(b.el) is None:
                self.report.add(ERROR, "heading-indent",
                                f"Заголовок раздела без абзацного отступа {indent:g} см".replace(".", ","),
                                "6.2.3", place)
            m = b.match
            if m is None:
                if self._numbering(b.el) is not None:
                    auto_numbered = True
                continue
            number, dot, rest = m.groups()
            if dot:
                self.report.add(ERROR, "heading-number-dot", f"Точка после номера «{number}.»", "6.4.1", place)
            if rest[:1].islower():
                self.report.add(ERROR, "heading-lower", "Заголовок раздела начинается со строчной буквы",
                                "6.2.3", place)
            parts = [int(x) for x in number.split(".")]
            if len(parts) == 1 and cfg.get("new_page", True) and not self.starts_new_page(i):
                self.report.add(ERROR, "new-page", "Структурный элемент или раздел не с новой страницы", "6.2.1",
                                place)
            expected = self._next_number(last, len(parts))
            if expected is not None and parts != expected:
                self.report.add(WARNING, "heading-sequence",
                                f"Номер {number} идёт после {'.'.join(map(str, last)) or 'начала'}, "
                                f"ожидался {'.'.join(map(str, expected))}", "6.4", place)
            last = parts
        if auto_numbered:
            self.report.notes.append("У части заголовков номера расставлены автоматической нумерацией Word: "
                                     "их последовательность и точку после номера скрипт не видит.")

    @staticmethod
    def _next_number(last: list[int], depth: int) -> list[int] | None:
        if not last:
            return [1] * depth if depth == 1 else None
        if depth <= len(last):
            return last[: depth - 1] + [last[depth - 1] + 1]
        if depth == len(last) + 1:
            return last + [1]
        return None

    def _numbering(self, p):
        return self.styles.para_prop(p, lambda ppr: ppr.find(qn("w:numPr")))

    def check_text(self) -> None:
        cfg = self.config["text"]
        font = cfg.get("font", "")
        size_min = cfg.get("size_min_pt", 0)
        size_exact = cfg.get("size_pt", 0)
        empty_run: list[Block] = []
        for b in self.blocks:
            if b.kind != "p" or b.role in ("title", "toc"):
                continue
            if b.role == "empty":
                empty_run.append(b)
                continue
            if len(empty_run) >= 3:
                self.report.add(WARNING, "empty-lines",
                                "Несколько пустых строк подряд: для новой страницы нужен разрыв страницы "
                                "(Ctrl+Enter), для отступа интервал абзаца", "",
                                Place(empty_run[0].index, "", empty_run[0].section))
            empty_run = []
            if b.role in ("drawing",):
                continue
            place = self.place(b)
            runs = text_runs(b.el)
            fonts, small, wrong_size, colored = set(), set(), set(), False
            for r in runs:
                name = self.styles.font(b.el, r)
                if font and name and name != font:
                    fonts.add(name)
                size = self.styles.size(b.el, r)
                if size < size_min:
                    small.add(size)
                elif size_exact and abs(size - size_exact) > 0.01:
                    wrong_size.add(size)
                color = (self.styles.color(b.el, r) or "auto").lower()
                if cfg.get("black_only", True) and color not in ("auto", "000000"):
                    colored = True
            for name in sorted(fonts):
                self.report.add(ERROR, "font", f"Шрифт {name} вместо {font}", "6.1.1", place)
            for size in sorted(small):
                self.report.add(ERROR, "size-small", f"Шрифт {size:g} пт, меньше {size_min} пт", "6.1.1", place)
            for size in sorted(wrong_size):
                self.report.add(ERROR, "size", f"Шрифт {size:g} пт, по настройкам нужен {size_exact} пт",
                                "методичка", place)
            if colored:
                self.report.add(ERROR, "color", "Цветной текст, нужен чёрный", "6.1.1", place)
            if b.role not in ("body", "bib", "formula"):
                continue
            rule, value = self.styles.line_spacing(b.el)
            need = cfg.get("line_spacing", 0)
            if need and rule == "multiple" and abs(value - need) > 0.01:
                self.report.add(ERROR, "spacing", f"Междустрочный интервал {value:g} вместо {need:g}".replace(".", ","),
                                "6.1.1", place)
            elif need and rule != "multiple":
                self.report.add(WARNING, "spacing-exact",
                                f"Междустрочный интервал задан точно ({value:g} пт), нужен множитель {need:g}"
                                .replace(".", ","), "6.1.1", place)
            if b.role == "formula":
                continue
            if b.role == "body" and self.is_all_bold(b.el) and len(b.text) < 200:
                self.report.add(WARNING, "bold-body",
                                "Полужирный абзац не распознан как заголовок. Если это заголовок, дайте ему номер "
                                "и стиль «Заголовок»; в основном тексте полужирный не используют", "6.1.1", place)
            if self._numbering(b.el) is None:
                first = self.styles.first_line_cm(b.el)
                need_indent = cfg.get("indent_cm", 0)
                if first is not None and need_indent and abs(first - need_indent) > cfg["indent_tolerance_cm"]:
                    self.report.add(ERROR, "indent",
                                    f"Абзацный отступ {first:.2f} см вместо {need_indent:g} см".replace(".", ","),
                                    "6.1.1", place)
            align = cfg.get("alignment", "")
            if align and b.role == "body":
                got = self.styles.alignment(b.el)
                want = {"justify": "both"}.get(align, align)
                if got != want:
                    names = {"both": "по ширине", "left": "по левому краю", "center": "по центру",
                             "right": "по правому краю"}
                    self.report.add(ERROR, "align",
                                    f"Выравнивание {names.get(got, got)}, нужно {names.get(want, want)}",
                                    "методичка", place)

    def _caption_rest(self, b: Block, kind: str) -> None:
        word, number, rest = b.match.groups()
        place = self.place(b)
        full = "Рисунок" if kind == "figure" else "Таблица"
        if word != full:
            self.report.add(ERROR, f"{kind}-word", f"Слово «{full}» в подписи пишется полностью",
                            "6.5.7" if kind == "figure" else "6.6.3", place)
        dashes = self.config["captions"]["dashes"]
        if rest:
            if rest[0] == "-":
                self.report.add(ERROR, f"{kind}-dash", "В подписи дефис вместо тире: «– Название»",
                                "6.5.7" if kind == "figure" else "6.6.3", place)
            elif rest[0] not in dashes:
                self.report.add(ERROR, f"{kind}-dash", f"После номера «{rest[0]}», а нужно тире: «{full} 1 – Название»",
                                "6.5.7" if kind == "figure" else "6.6.3", place)
            title = rest[1:].strip()
            if title[:1].islower():
                self.report.add(ERROR, f"{kind}-lower", "Название в подписи со строчной буквы",
                                "6.5.7" if kind == "figure" else "6.6.3", place)
            if title.endswith("."):
                self.report.add(ERROR, f"{kind}-dot", "Точка в конце подписи",
                                "6.5.7" if kind == "figure" else "6.6.3", place)

    def check_captions(self) -> None:
        figures, tables = [], []
        for i, b in enumerate(self.blocks):
            if b.role == "figure":
                figures.append(b)
                self._caption_rest(b, "figure")
                if self.styles.alignment(b.el) != "center":
                    self.report.add(ERROR, "figure-align", "Подпись рисунка не по центру", "6.5.7", self.place(b))
                prev, nxt = self.prev_block(i), self.next_block(i)
                prev_pic = prev is not None and prev.kind == "p" and has_drawing(prev.el)
                next_pic = nxt is not None and nxt.kind == "p" and has_drawing(nxt.el) and not has_drawing(b.el)
                if not prev_pic and not has_drawing(b.el):
                    if next_pic:
                        self.report.add(ERROR, "figure-above", "Подпись над рисунком, а её ставят под ним", "6.5.7",
                                        self.place(b))
                    else:
                        self.report.add(WARNING, "figure-missing",
                                        "Рядом с подписью не найден рисунок (если он плавающий, проверьте глазами)",
                                        "6.5.7", self.place(b))
            elif b.role == "table_title":
                tables.append(b)
                self._caption_rest(b, "table")
                if self.styles.alignment(b.el) in ("center", "right"):
                    self.report.add(ERROR, "table-align", "Название таблицы ставят слева, не по центру", "6.6.3",
                                    self.place(b))
                first = self.styles.first_line_cm(b.el)
                if first is not None and abs(first) > self.config["text"]["indent_tolerance_cm"]:
                    self.report.add(ERROR, "table-indent", "Название таблицы с абзацным отступом, нужно без него",
                                    "6.6.3", self.place(b))
                nxt = self.next_block(i, skip_empty=False)
                if nxt is None or nxt.kind != "tbl":
                    self.report.add(ERROR, "table-missing", "Сразу после названия таблицы нет самой таблицы",
                                    "6.6.3", self.place(b))
            elif b.kind == "tbl" and b.role == "table":
                prev = self.prev_block(i, skip_empty=False)
                if prev is None or prev.kind != "p" or (
                    prev.role != "table_title" and not prev.text.strip().lower().startswith("продолжение таблицы")
                ):
                    where = prev if prev is not None and prev.kind == "p" else None
                    self.report.add(WARNING, "table-untitled",
                                    "Перед таблицей нет строки «Таблица N – Название»", "6.6.3",
                                    Place(where.index if where else None, "после абзаца" if where else "", b.section))
        self._check_numbering(figures, "Рисунок", "6.5")
        self._check_numbering(tables, "Таблица", "6.6")

    def _check_numbering(self, items: list[Block], word: str, rule: str) -> None:
        main = [b for b in items if not re.match(r"^[А-ЯЁA-Z]\.", b.match.group(2))]
        if not main:
            return
        numbers = [b.match.group(2) for b in main]
        per_section = all("." in n for n in numbers)
        mixed = not per_section and any("." in n for n in numbers)
        if mixed:
            self.report.add(WARNING, "numbering-mixed",
                            f"{word}: часть номеров сквозные, часть по разделам; выберите одну схему", rule)
            return
        counters: dict[str, int] = {}
        for b, n in zip(main, numbers):
            prefix, _, num = n.rpartition(".") if per_section else ("", "", n)
            want = counters.get(prefix, 0) + 1
            if int(num) != want:
                expected = f"{prefix}.{want}" if prefix else str(want)
                self.report.add(WARNING, "numbering",
                                f"{word} {n}: ожидался номер {expected}", rule, self.place(b))
            counters[prefix] = int(num)

    @staticmethod
    def _numbers(text: str) -> list[str]:
        out = []
        tokens = re.findall(NUMBER_TOKEN + r"|–|—|-", text)
        i = 0
        while i < len(tokens):
            tok = tokens[i]
            if tok in "–—-":
                i += 1
                continue
            if i + 2 < len(tokens) and tokens[i + 1] in "–—-" and tok.isdigit() and tokens[i + 2].isdigit():
                out += [str(k) for k in range(int(tok), int(tokens[i + 2]) + 1)]
                i += 3
                continue
            out.append(tok)
            i += 1
        return out

    def check_references(self) -> None:
        if not self.config["references"].get("figures_and_tables", True):
            return
        for kind, code, regex, word, rule in (
            ("figure", "figure", FIG_REF_RE, "рисунок", "6.5.1"),
            ("table_title", "table", TAB_REF_RE, "таблицу", "6.6"),
        ):
            first_ref: dict[str, int] = {}
            for i, b in enumerate(self.blocks):
                if b.kind == "p" and b.role in ("body", "formula", "bib"):
                    for m in regex.finditer(b.text):
                        for n in self._numbers(m.group(1)):
                            first_ref.setdefault(n, i)
            for i, b in enumerate(self.blocks):
                if b.role != kind:
                    continue
                n = b.match.group(2)
                if n not in first_ref:
                    self.report.add(WARNING, f"{code}-noref", f"В тексте нет ссылки на {word} {n}", rule,
                                    self.place(b))
                elif first_ref[n] > i:
                    self.report.add(WARNING, f"{code}-late",
                                    f"Первая ссылка на {word} {n} стоит после него: помещают после первого упоминания",
                                    rule, self.place(b))

    def check_sources(self) -> None:
        if not self.config["references"].get("sources", True):
            return
        entries = [b for b in self.blocks if b.role == "bib"]
        if not entries:
            return
        total = len(entries)
        for pos, b in enumerate(entries, start=1):
            m = re.match(r"^(\d+)\s*[.)]?\s", b.text.strip())
            if m and int(m.group(1)) != pos:
                self.report.add(WARNING, "source-number", f"Источник с номером {m.group(1)} стоит {pos}-м", "6.16",
                                self.place(b))
        cited: list[int] = []
        for b in self.blocks:
            if b.kind != "p" or b.role not in ("body", "formula", "figure", "table_title"):
                continue
            for m in CITE_RE.finditer(b.text):
                inner = re.split(r",\s*с\.|,\s*p\.|;\s*с\.", m.group(1))[0]
                if not re.fullmatch(r"[\d\s,–—\-]+", inner.strip()):
                    continue
                for n in self._numbers(inner):
                    k = int(n)
                    if k > total or k < 1:
                        self.report.add(ERROR, "source-missing",
                                        f"Ссылка [{k}], а в списке источников их {total}", "6.16", self.place(b))
                    cited.append(k)
        if not cited:
            self.report.add(WARNING, "sources-uncited",
                            "В тексте нет ссылок на источники вида [1]", "6.16")
            return
        for k in range(1, total + 1):
            if k not in cited:
                self.report.add(WARNING, "source-uncited", f"На источник {k} нет ссылки в тексте", "6.16",
                                self.place(entries[k - 1]))
        order = []
        for k in cited:
            if k not in order and k <= total:
                order.append(k)
        if order != sorted(order):
            self.report.add(WARNING, "source-order",
                            "Источники пронумерованы не в порядке первых ссылок в тексте: первые ссылки идут "
                            + ", ".join(map(str, order[:10])) + ("…" if len(order) > 10 else ""), "6.16")


# ---------------------------------------------------------------- вывод

def render_text(path: Path, config: dict, report: Report, limit: int = 5) -> str:
    lines = [f"Проверка: {path.name}, профиль «{config.get('name', '')}»",
             f"Ошибок: {report.count(ERROR)}, замечаний: {report.count(WARNING)}", ""]
    for severity, title in ((ERROR, "ОШИБКИ"), (WARNING, "ЗАМЕЧАНИЯ")):
        items = [i for i in report.issues if i.severity == severity]
        if not items:
            continue
        lines.append(title)
        for n, issue in enumerate(items, start=1):
            rule = f" (п. {issue.rule})" if issue.rule and issue.rule[0].isdigit() else (
                f" ({issue.rule})" if issue.rule else "")
            count = f", мест: {len(issue.places)}" if len(issue.places) > 1 else ""
            lines.append(f"{n}. {issue.message}{rule}{count}")
            for place in issue.places[:limit]:
                text = place.describe()
                if text:
                    lines.append(f"   {text}")
            if len(issue.places) > limit:
                lines.append(f"   и ещё {len(issue.places) - limit}")
        lines.append("")
    if report.notes:
        lines.append("НЕ ПРОВЕРЕНО")
        lines += [f"- {note}" for note in report.notes]
        lines.append("")
    if not report.issues:
        lines.append("Нарушений не найдено. Проверьте глазами то, что скрипт не видит: смысл заголовков, "
                     "содержание списка литературы, качество рисунков.")
    return "\n".join(lines).rstrip() + "\n"


def render_json(path: Path, config: dict, report: Report) -> str:
    data = {
        "file": path.name,
        "profile": config.get("name", ""),
        "errors": report.count(ERROR),
        "warnings": report.count(WARNING),
        "issues": [
            {"severity": i.severity, "code": i.code, "message": i.message, "rule": i.rule,
             "places": [p.__dict__ for p in i.places]}
            for i in report.issues
        ],
        "notes": report.notes,
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка оформления .docx по ГОСТ 7.32-2017")
    parser.add_argument("file", type=Path, help="файл .docx")
    parser.add_argument("--config", type=Path, help="настройки методички (.toml), дополняют профиль ГОСТ")
    parser.add_argument("--json", action="store_true", help="вывод в JSON")
    parser.add_argument("--limit", type=int, default=5, help="сколько мест показывать на одно нарушение")
    args = parser.parse_args(argv)

    if args.file.suffix.lower() != ".docx":
        print(f"Нужен файл .docx, а передан {args.file.name}. Файл .doc сначала сохраните как .docx.")
        return 2
    try:
        config = load_config(args.config)
        checker = Checker(args.file, config)
    except FileNotFoundError as e:
        print(f"Файл не найден: {e.filename or e}")
        return 2
    except tomllib.TOMLDecodeError as e:
        print(f"Ошибка в файле настроек: {e}")
        return 2
    except Exception as e:  # повреждённый архив, не Word и т. п.
        print(f"Не удалось прочитать {args.file.name}: {e}")
        return 2
    report = checker.run()
    if args.json:
        print(render_json(args.file, config, report))
    else:
        print(render_text(args.file, config, report, args.limit), end="")
    return 1 if report.count(ERROR) else 0


if __name__ == "__main__":
    sys.exit(main())
