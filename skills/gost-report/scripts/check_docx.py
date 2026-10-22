#!/usr/bin/env python3
"""Проверка оформления .docx по ГОСТ 7.32-2017 или по своей методичке.

Скрипт читает разметку документа и печатает, что не так и где: номер абзаца,
начало его текста и раздел, в котором он стоит. Исходный файл он не меняет.
С флагом --fix однозначные ошибки (поля, шрифт и размер стиля «Обычный»,
тире и точка в подписях, точка после номера заголовка) исправляются в копии
«имя.fixed.docx», а отчёт показывает, что осталось.

    python check_docx.py работа.docx
    python check_docx.py работа.docx --config методичка.toml
    python check_docx.py работа.docx --json
    python check_docx.py работа.docx --fix

Код выхода: 0 если ошибок нет (замечания допустимы), 1 если есть ошибки,
2 если файл не удалось прочитать.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = Path(__file__).resolve().parent
# Общая часть лежит рядом: путь от файла, чтобы скрипт запускался из любой папки
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import docx_common  # noqa: E402
# Часть имён здесь не используется: их берёт gost-business из этого модуля
from docx_common import (  # noqa: E402,F401
    EMU_PER_MM, ERROR, TWIPS_PER_CM, WARNING, Issue, Place, Report, Styles, deep_merge, docx,
    fixed_path, has_drawing, has_formula, has_page_break, has_section_break, is_on, normalize,
    para_text, qn, render_json, short, text_runs, tomllib,
)

DEFAULT_PROFILE = HERE.parent / "assets" / "gost-7.32.toml"

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

def load_config(path: Path | None = None) -> dict:
    with open(DEFAULT_PROFILE, "rb") as f:
        config = tomllib.load(f)
    if path is not None:
        with open(path, "rb") as f:
            config = deep_merge(config, tomllib.load(f))
    return config


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
    def __init__(self, path: Path, config: dict, document=None) -> None:
        self.path = path
        self.config = config
        self.doc = document if document is not None else docx.Document(str(path))
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


# ---------------------------------------------------------------- исправление

class Fixer(docx_common.Fixer):
    """Исправляет в копии только то, что однозначно следует из профиля.

    Поля страницы и шрифт и размер стиля «Обычный» правит общая часть.
    Здесь добавлены подписи и заголовки: в подписях рисунков и таблиц ставится
    тире после номера и убирается точка в конце, в заголовках разделов
    убирается точка после номера. Если нужное место разбито на несколько
    фрагментов, абзац не трогается и остаётся в отчёте как ошибка.
    """

    def run(self) -> list[str]:
        self.fix_margins()
        self.fix_normal_style()
        self.fix_captions_and_headings()
        return self.fixed

    def fix_captions_and_headings(self) -> None:
        checker = Checker(self.path, self.config, self.doc)
        checker.assign_roles()
        dashes = self.config["captions"].get("dashes") or []
        for b in checker.blocks:
            if b.kind != "p" or b.match is None:
                continue
            before = para_text(b.el)
            lead = len(before) - len(before.lstrip())
            if b.role == "heading":
                self._fix_heading(b, lead)
            elif b.role in ("figure", "table_title"):
                self._fix_caption(b, lead, dashes)
            after = para_text(b.el)
            if after != before:
                kind = "Заголовок" if b.role == "heading" else "Подпись"
                self.fixed.append(f"{kind} «{short(before.strip(), 60)}» → «{short(after.strip(), 60)}» "
                                  f"(абзац {b.index})")

    def _fix_heading(self, b: Block, lead: int) -> None:
        m = b.match
        if not m.group(2):
            return
        # номер, точка и первая буква названия должны быть в одном фрагменте
        self._replace_in_one_run(b.el, lead + m.start(1), lead + m.start(3) + 1,
                                 m.group(1) + " " + m.group(3)[0])

    def _fix_caption(self, b: Block, lead: int, dashes: list[str]) -> None:
        m = b.match
        rest = m.group(3)
        title = rest[1:].strip()
        if not rest or not title:
            return
        # сначала точка в конце: её место дальше по тексту и не сдвигается правкой тире
        if title.endswith(".") and not title.endswith(".."):
            end = lead + m.end(3)
            self._replace_in_one_run(b.el, end - 1, end, "")
        if dashes and rest[0] not in dashes:
            title_start = m.start(3) + len(rest) - len(rest[1:].lstrip())
            # конец номера, знак и первая буква названия должны быть в одном фрагменте
            self._replace_in_one_run(b.el, lead + m.end(2), lead + title_start + 1,
                                     f" {dashes[0]} " + title[0])


# ---------------------------------------------------------------- вывод

CLEAN = ("Нарушений не найдено. Проверьте глазами то, что скрипт не видит: смысл заголовков, "
         "содержание списка литературы, качество рисунков.")


def render_text(path: Path, config: dict, report: Report, limit: int = 5,
                fixed: list[str] | None = None) -> str:
    return docx_common.render_text(path, config, report, limit, fixed, clean=CLEAN)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка оформления .docx по ГОСТ 7.32-2017")
    parser.add_argument("file", type=Path, help="файл .docx")
    parser.add_argument("--config", type=Path, help="настройки методички (.toml), дополняют профиль ГОСТ")
    parser.add_argument("--json", action="store_true", help="вывод в JSON")
    parser.add_argument("--limit", type=int, default=5, help="сколько мест показывать на одно нарушение")
    parser.add_argument("--fix", action="store_true",
                        help="исправить однозначное в копии имя.fixed.docx и показать, что осталось")
    parser.add_argument("--force", action="store_true", help="с --fix: перезаписать готовую копию")
    args = parser.parse_args(argv)

    if args.file.suffix.lower() != ".docx":
        print(f"Нужен файл .docx, а передан {args.file.name}. Файл .doc сначала сохраните как .docx.")
        return 2
    target, fixed = args.file, None
    if args.fix:
        target = fixed_path(args.file)
        if target.exists() and not args.force:
            print(f"Копия {target.name} уже есть. Чтобы перезаписать её, добавьте --force.")
            return 2
    try:
        config = load_config(args.config)
        if args.fix:
            fixer = Fixer(args.file, config)
            fixed = fixer.run()
            fixer.save(target)
        checker = Checker(target, config)
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
        print(render_json(target, config, report, fixed))
    else:
        print(render_text(target, config, report, args.limit, fixed), end="")
    return 1 if report.count(ERROR) else 0


if __name__ == "__main__":
    sys.exit(main())
