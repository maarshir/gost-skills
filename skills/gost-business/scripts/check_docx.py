#!/usr/bin/env python3
"""Проверка деловых документов .docx по ГОСТ Р 7.0.97-2016 или по инструкции организации.

Скрипт ничего не меняет в файле. Он проверяет лист, шрифт, абзацный отступ
и даты и печатает, что не так и где: номер абзаца и начало его текста.

    python check_docx.py письмо.docx
    python check_docx.py письмо.docx --config инструкция.toml
    python check_docx.py письмо.docx --json

Чтение разметки Word и формат отчёта общие с навыком gost-report, поэтому
папка gost-report должна лежать рядом (так ставятся все навыки репозитория).

Код выхода: 0 если ошибок нет (замечания допустимы), 1 если есть ошибки,
2 если файл не удалось прочитать.
"""

from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_PROFILE = HERE.parent / "assets" / "gost-r-7.0.97.toml"
REPORT_SCRIPT = HERE.parent.parent / "gost-report" / "scripts" / "check_docx.py"


def _load_report_module():
    """Модуль проверки gost-report под своим именем: у обоих скриптов имя check_docx."""
    if not REPORT_SCRIPT.exists():
        print("Не найден навык gost-report рядом с gost-business: скрипт берёт из него чтение .docx. "
              "Поставьте все навыки репозитория maarshir/gost-skills.")
        sys.exit(2)
    spec = importlib.util.spec_from_file_location("gost_report_check_docx", REPORT_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


base = _load_report_module()
docx, qn, tomllib = base.docx, base.qn, base.tomllib
ERROR, WARNING = base.ERROR, base.WARNING
Place, Report, Styles = base.Place, base.Report, base.Styles
EMU_PER_MM = base.EMU_PER_MM

MONTHS = ("января", "февраля", "марта", "апреля", "мая", "июня",
          "июля", "августа", "сентября", "октября", "ноября", "декабря")

# Правильная цифровая дата и всё, что на неё похоже, но записано иначе
DATE_OK_RE = re.compile(r"(?<![\d.])(\d{2})\.(\d{2})\.(\d{4})(?![\d.])")
DATE_LOOSE_RE = re.compile(r"(?<![\d./\-])(\d{1,2})([./\-])(\d{1,2})\2(\d{2}|\d{4})(?![\d./\-])")
DATE_ISO_RE = re.compile(r"(?<![\d.\-])(\d{4})-(\d{2})-(\d{2})(?![\d\-])")
DATE_WORDS_RE = re.compile(
    r"(?<!\d)(\d{1,2})\s+(" + "|".join(MONTHS) + r")\s+(\d{4})(\s*(?:г\.|года\b|г\b))?",
    re.IGNORECASE,
)
PHONE_RE = re.compile(r"(?:\+7|8)[\s(\-]*\d{3}[\s)\-]*\d{3}[\s\-]?\d{2}[\s\-]?\d{2}|тел\.|@")


def load_config(path: Path | None = None) -> dict:
    with open(DEFAULT_PROFILE, "rb") as f:
        config = tomllib.load(f)
    if path is not None:
        with open(path, "rb") as f:
            config = base.deep_merge(config, tomllib.load(f))
    return config


def real_date(day: int, month: int, year: int) -> bool:
    try:
        dt.date(year, month, day)
    except ValueError:
        return False
    return True


class Checker:
    def __init__(self, path: Path, config: dict) -> None:
        self.path = path
        self.config = config
        self.doc = docx.Document(str(path))
        self.styles = Styles(self.doc)
        self.report = Report()
        # Абзацы тела документа по порядку; таблицы пропускаем: в них ГОСТ допускает мельче.
        self.paras = [el for el in self.doc.element.body if el.tag == qn("w:p")]
        self.texts = [base.para_text(p) for p in self.paras]
        self.executor = self._executor_paragraphs()

    def place(self, i: int) -> Place:
        return Place(i + 1, base.short(self.texts[i]))

    def run(self) -> Report:
        self.check_page()
        self.check_text()
        self.check_dates()
        return self.report

    # --- лист

    def check_page(self) -> None:
        cfg = self.config["page"]
        tol = cfg.get("tolerance_mm", 1)
        sizes = [tuple(s) for s in cfg["sizes_mm"]]
        names = " или ".join(f"{w}×{h} мм" for w, h in sizes)
        rule = cfg.get("rule", "")
        for n, section in enumerate(self.doc.sections, start=1):
            where = Place(section=f"часть документа {n}") if len(self.doc.sections) > 1 else None
            if section.page_width is not None and section.page_height is not None:
                w, h = sorted((section.page_width / EMU_PER_MM, section.page_height / EMU_PER_MM))
                if not any(abs(w - sw) <= tol and abs(h - sh) <= tol for sw, sh in sizes):
                    self.report.add(ERROR, "page-size", f"Размер листа {w:.0f}×{h:.0f} мм, нужен {names}",
                                    rule, where)
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
                if mm < need - tol:
                    self.report.add(ERROR, "margin", f"Поле {label} {mm:.0f} мм, нужно не меньше {need} мм",
                                    rule, where)

    # --- шрифт и отступ

    def _executor_paragraphs(self) -> set[int]:
        """Отметка об исполнителе: короткие абзацы в самом конце, среди них телефон или почта."""
        tail = []
        for i in range(len(self.paras) - 1, -1, -1):
            text = self.texts[i].strip()
            if not text:
                if tail:
                    break
                continue
            if len(text) > 80 or len(tail) == 3:
                break
            tail.append(i)
        if any(PHONE_RE.search(self.texts[i]) for i in tail):
            return set(tail)
        return set()

    def check_text(self) -> None:
        cfg = self.config["text"]
        rule = cfg.get("rule", "")
        fonts = cfg.get("fonts", [])
        size_min, size_max = cfg.get("size_min_pt", 0), cfg.get("size_max_pt", 0)
        need_indent = cfg.get("indent_cm", 0)
        for i, p in enumerate(self.paras):
            text = self.texts[i].strip()
            if not text:
                continue
            place = self.place(i)
            wrong_fonts, small, large = set(), set(), set()
            for r in base.text_runs(p):
                name = self.styles.font(p, r)
                if fonts and name and name not in fonts:
                    wrong_fonts.add(name)
                size = self.styles.size(p, r)
                if size_min and size < size_min - 0.01 and i not in self.executor:
                    small.add(size)
                elif size_max and size > size_max + 0.01:
                    large.add(size)
            allowed = ", ".join(fonts)
            for name in sorted(wrong_fonts):
                self.report.add(ERROR, "font", f"Шрифт {name}, нужен один из: {allowed}", rule, place)
            for size in sorted(small):
                self.report.add(ERROR, "size-small", f"Шрифт {size:g} пт, меньше {size_min:g} пт", rule, place)
            for size in sorted(large):
                self.report.add(WARNING, "size-large", f"Шрифт {size:g} пт, больше {size_max:g} пт", rule, place)
            # Отступ проверяем только у абзацев текста: реквизиты короткие и печатаются без отступа.
            if not need_indent or len(text) < cfg.get("body_min_chars", 100):
                continue
            if self.styles.alignment(p) in ("center", "right"):
                continue
            ppr = p.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:numPr")) is not None:
                continue  # пункты списка с номерами Word: отступ задаёт список
            first = self.styles.first_line_cm(p)
            if first is not None and abs(first - need_indent) > cfg.get("indent_tolerance_cm", 0.05):
                self.report.add(ERROR, "indent",
                                f"Абзацный отступ {first:.2f} см вместо {need_indent:g} см".replace(".", ","),
                                rule, place)

    # --- даты

    def check_dates(self) -> None:
        cfg = self.config["date"]
        rule = cfg.get("rule", "")
        words_allowed = cfg.get("words_allowed", True)
        for i, raw in enumerate(self.texts):
            text = raw.replace(" ", " ")
            if not text.strip():
                continue
            place = self.place(i)
            for m in DATE_OK_RE.finditer(text):
                if not real_date(int(m.group(1)), int(m.group(2)), int(m.group(3))):
                    self.report.add(ERROR, "date-invalid", f"Такой даты нет: {m.group(0)}", rule, place)
            for m in DATE_LOOSE_RE.finditer(text):
                if DATE_OK_RE.fullmatch(m.group(0)):
                    continue
                day, month, year = int(m.group(1)), int(m.group(3)), m.group(4)
                if not (1 <= day <= 31 and 1 <= month <= 12):
                    continue  # номер или дробь, а не дата
                if len(year) == 2:
                    if len(m.group(1)) < 2 or len(m.group(3)) < 2:
                        continue  # «1.2.10» скорее номер пункта, чем дата
                    year = "20" + year
                good = f"{day:02d}.{month:02d}.{year}"
                self.report.add(ERROR, "date-format", f"Дата {m.group(0)} записана не так, нужно {good}",
                                rule, place)
            for m in DATE_ISO_RE.finditer(text):
                year, month, day = m.groups()
                if real_date(int(day), int(month), int(year)):
                    self.report.add(ERROR, "date-format",
                                    f"Дата {m.group(0)} записана не так, нужно {day}.{month}.{year}", rule, place)
            for m in DATE_WORDS_RE.finditer(text):
                day, month, year, tail = m.group(1), m.group(2).lower(), m.group(3), (m.group(4) or "").strip()
                num = MONTHS.index(month) + 1
                if not real_date(int(day), num, int(year)):
                    self.report.add(ERROR, "date-invalid", f"Такой даты нет: {m.group(0).strip()}", rule, place)
                    continue
                if not words_allowed:
                    self.report.add(ERROR, "date-words",
                                    f"Дата словами «{m.group(0).strip()}», по настройкам нужна цифрами: "
                                    f"{int(day):02d}.{num:02d}.{year}", rule, place)
                elif tail != "г.":
                    self.report.add(WARNING, "date-year",
                                    f"После года нужно «г.»: «{int(day)} {month} {year} г.»", rule, place)


CLEAN = ("Нарушений не найдено. Проверьте глазами то, что скрипт не видит: состав и порядок реквизитов, "
         "адресат, подпись, отметки о приложении и об исполнителе.")


def render_text(path: Path, config: dict, report: Report, limit: int = 5) -> str:
    """Отчёт как у gost-report, только итог для чистого документа про деловые реквизиты."""
    text = base.render_text(path, config, report, limit)
    if report.issues:
        return text
    last = text.rstrip("\n").rsplit("\n", 1)
    return (last[0] + "\n" if len(last) > 1 else "") + CLEAN + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка деловых документов .docx по ГОСТ Р 7.0.97-2016")
    parser.add_argument("file", type=Path, help="файл .docx")
    parser.add_argument("--config", type=Path, help="настройки инструкции организации (.toml), дополняют ГОСТ")
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
        print(base.render_json(args.file, config, report))
    else:
        print(render_text(args.file, config, report, args.limit), end="")
    return 1 if report.count(ERROR) else 0


if __name__ == "__main__":
    sys.exit(main())
