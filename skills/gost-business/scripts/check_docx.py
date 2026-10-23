#!/usr/bin/env python3
"""Проверка деловых документов .docx по ГОСТ Р 7.0.97-2016 или по инструкции организации.

Исходный файл скрипт не меняет. Он проверяет лист, шрифт, абзацный отступ,
даты, падеж адресата, расшифровку подписи, точку после заголовка к тексту,
отметки о приложении и об исполнителе и печатает, что не так и где:
номер абзаца и начало его текста.

    python check_docx.py письмо.docx
    python check_docx.py письмо.docx --config инструкция.toml
    python check_docx.py письмо.docx --json
    python check_docx.py письмо.docx --fix

С --fix рядом появляется копия письмо.fixed.docx: в ней поля страницы и шрифт
и размер стиля «Обычный» поставлены по профилю, отчёт показывает, что исправлено
и что осталось исправить вручную.

Чтение разметки Word и формат отчёта лежат рядом в docx_common.py: это копия
модуля из gost-report, поэтому навык работает и без других папок репозитория.

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


def _load_common():
    """Общая часть из docx_common.py рядом, по пути от файла и под своим именем.

    Папку скрипта в sys.path не добавляем: там лежит check_docx.py, и при общем
    запуске тестов он закрыл бы одноимённый скрипт gost-report.
    """
    spec = importlib.util.spec_from_file_location("gost_business_docx_common", HERE / "docx_common.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


base = _load_common()

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
# Первое слово должности в именительном падеже. По нему узнаём адресат, записанный
# не в дательном падеже, и строку подписи. Список намеренно короткий: незнакомую
# должность скрипт пропустит, а не отметит зря.
POSITIONS = {
    "директор", "генеральный", "исполнительный", "технический", "коммерческий", "финансовый",
    "первый", "заместитель", "начальник", "руководитель", "председатель", "президент",
    "вице-президент", "глава", "ректор", "проректор", "декан", "министр", "управляющий",
    "заведующий", "главный", "старший", "ведущий", "менеджер", "специалист", "инженер",
    "бухгалтер", "секретарь", "инспектор", "советник", "помощник", "губернатор", "мэр",
}
GRIF_RE = re.compile(r"^\s*(УТВЕРЖДА|УТВЕРЖДЕН|СОГЛАСОВАН)")
# «Петров П.П.» или «Петров П.» в конце строки
NAME_AFTER_RE = re.compile(r"([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\s+([А-ЯЁ])\.(?:\s*([А-ЯЁ])\.)?$")
TITLE_RE = re.compile(r"^(О|Об|Обо)\s+\S")
# Сокращения, после которых точка в конце заголовка уместна: «за 2025 г.»
ABBR_END_RE = re.compile(r"(?:\b(?:г|гг|др|пр|руб|коп)|т\.\s?[дп])\.$")
PHONE_RE = re.compile(r"(?:\+7|8)[\s(\-]*\d{3}[\s)\-]*\d{3}[\s\-]?\d{2}[\s\-]?\d{2}|тел\.|@")
# Сам номер телефона: не меньше пяти цифр подряд, можно через пробелы, скобки и дефисы
PHONE_DIGITS_RE = re.compile(r"(?:\d[\s()\-]*){5,}")
# Исполнитель: «А.П. Сидорова», «Сидорова А.П.» или фамилия, имя и отчество полностью
WORD = r"[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?"
EXECUTOR_NAME_RE = re.compile(
    rf"[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.\s?)?{WORD}|{WORD}\s+[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.)?"
    rf"|(?:{WORD}\s+){{2}}{WORD}"
)
# Отметка о приложении: «Приложение: отчёт на 5 л. в 1 экз.», для нескольких «Приложения:» и пункты
ATTACH_RE = re.compile(r"^\s*Приложени[ея]\s*:\s*(.*)$", re.S)
ITEM_RE = re.compile(r"^\s*\d+\.\s")
SHEETS_RE = re.compile(r"\d+\s*л\.")
COPIES_RE = re.compile(r"\d+\s*экз\.")


def load_config(path: Path | None = None) -> dict:
    with open(DEFAULT_PROFILE, "rb") as f:
        config = tomllib.load(f)
    if path is not None:
        with open(path, "rb") as f:
            config = base.deep_merge(config, tomllib.load(f))
    return config


def first_word(text: str) -> str:
    m = re.match(r"\s*([А-ЯЁа-яё\-]+)", text)
    return m.group(1) if m else ""


def dative(word: str) -> str:
    """Дательный падеж первого слова должности: «Директор» → «Директору»."""
    low = word.lower()
    if low.endswith(("ый", "ой")):
        out = word[:-2] + "ому"
    elif low.endswith("ий"):
        out = word[:-2] + ("ому" if low[-3:-2] in "кгх" else "ему")
    elif low.endswith("ь"):
        out = word[:-1] + "ю"
    elif low.endswith("а"):
        out = word[:-1] + "е"
    else:
        out = word + "у"
    return out


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
        # Первый абзац текста делит документ на шапку (адресат, заголовок) и конец (подпись).
        min_chars = config["text"].get("body_min_chars", 100)
        self.body_start = next((i for i, t in enumerate(self.texts) if len(t.strip()) >= min_chars), None)

    def place(self, i: int) -> Place:
        return Place(i + 1, base.short(self.texts[i]))

    def run(self) -> Report:
        self.check_page()
        self.check_text()
        self.check_dates()
        self.check_addressee()
        self.check_title()
        self.check_signature()
        self.check_attachment()
        self.check_executor()
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

    # --- реквизиты по тексту

    def head(self) -> list[int]:
        """Непустые абзацы шапки, до первого абзаца текста."""
        end = self.body_start if self.body_start is not None else len(self.paras)
        return [i for i in range(end) if self.texts[i].strip()]

    def left_indent_cm(self, p) -> float:
        def get(ppr):
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                return None
            val = ind.get(qn("w:left")) or ind.get(qn("w:start"))
            return int(val) / base.TWIPS_PER_CM if val is not None else None
        return self.styles.para_prop(p, get) or 0.0

    def shifted_right(self, i: int) -> bool:
        """Адресат стоит справа: выравнивание по правому краю, большой отступ слева или табуляция."""
        p = self.paras[i]
        return (self.styles.alignment(p) == "right" or self.left_indent_cm(p) >= 5
                or self.texts[i].startswith("\t"))

    def check_addressee(self) -> None:
        rule = self.config["addressee"].get("rule", "")
        grif_left = 0
        for i in self.head():
            text = self.texts[i].strip()
            if GRIF_RE.match(text):
                grif_left = 4  # должность в грифе утверждения пишется в именительном
                continue
            if grif_left:
                grif_left -= 1
                continue
            word = first_word(text)
            if word.lower() not in POSITIONS or not self.shifted_right(i):
                continue
            self.report.add(ERROR, "addressee-case",
                            f"Адресат в именительном падеже: «{word}», нужно «{dative(word)}»",
                            rule, self.place(i))

    def check_title(self) -> None:
        rule = self.config["title"].get("rule", "")
        for i in self.head():
            text = self.texts[i].strip()
            if TITLE_RE.match(text) and text.endswith(".") and not ABBR_END_RE.search(text):
                self.report.add(ERROR, "title-dot", "Точка в конце заголовка к тексту", rule, self.place(i))

    def check_signature(self) -> None:
        cfg = self.config["signature"]
        if self.body_start is None or not cfg.get("initials_first", True):
            return
        rule = cfg.get("rule", "")
        prev = ""
        for i in range(self.body_start + 1, len(self.paras)):
            raw = self.texts[i].strip()
            if not raw:
                continue
            if i in self.executor:
                break  # в отметке об исполнителе фамилия может стоять первой
            m = NAME_AFTER_RE.search(raw)
            if m and len(raw) < 100:
                before = raw[:m.start()]
                in_line = (first_word(raw).lower() in POSITIONS
                           and re.search(r"\t|\s{2,}", before) is not None)
                alone = not before.strip() and first_word(prev).lower() in POSITIONS
                if in_line or alone:
                    surname, initials = m.group(1), m.group(2) + "." + (m.group(3) + "." if m.group(3) else "")
                    self.report.add(ERROR, "signature-initials",
                                    f"Расшифровка подписи «{m.group(0)}», нужно «{initials} {surname}»",
                                    rule, self.place(i))
            prev = raw

    def attachment_items(self, i: int, rest: str) -> list[tuple[int, str]]:
        """Пункты отметки о приложении: та же строка или пункты с номерами под ней."""
        items = [(i, rest)] if rest else []
        if rest and not ITEM_RE.match(rest):
            return items
        for j in range(i + 1, len(self.paras)):
            text = self.texts[j].strip()
            ppr = self.paras[j].find(qn("w:pPr"))
            numbered = ppr is not None and ppr.find(qn("w:numPr")) is not None
            if not text or not (ITEM_RE.match(text) or numbered):
                break
            items.append((j, text))
        return items

    def check_attachment(self) -> None:
        cfg = self.config.get("attachment", {})
        if not cfg.get("check", True):
            return
        rule = cfg.get("rule", "")
        for i, raw in enumerate(self.texts):
            m = ATTACH_RE.match(raw.replace(" ", " "))
            if not m:
                continue
            for j, text in self.attachment_items(i, m.group(1).strip()):
                low = text.lower()
                # Сброшюрованное приложение листами не считают (5.20)
                if not SHEETS_RE.search(text) and "сброшюр" not in low:
                    self.report.add(ERROR, "attachment-sheets", "В отметке о приложении нет числа листов",
                                    rule, self.place(j))
                # «...и приложение к нему, всего на 5 л.» по ГОСТу пишут без экземпляров
                if not COPIES_RE.search(text) and "всего" not in low:
                    self.report.add(WARNING, "attachment-copies", "В отметке о приложении нет числа экземпляров",
                                    rule, self.place(j))

    def check_executor(self) -> None:
        """Отметку ищем в конце документа по телефону или почте; если её нет, ничего не отмечаем."""
        cfg = self.config.get("executor", {})
        if not cfg.get("check", True) or not self.executor:
            return
        rule = cfg.get("rule", "")
        lines = sorted(self.executor)
        texts = [self.texts[i].replace(" ", " ") for i in lines]
        place = self.place(lines[0])
        if not any(EXECUTOR_NAME_RE.search(t) for t in texts):
            self.report.add(ERROR, "executor-name",
                            "В отметке об исполнителе нет инициалов или имени и отчества", rule, place)
        if not any(PHONE_DIGITS_RE.search(t) for t in texts):
            self.report.add(ERROR, "executor-phone", "В отметке об исполнителе нет номера телефона", rule, place)


class Fixer(base.Fixer):
    """Исправляет в копии узкие поля страницы и шрифт и размер стиля «Обычный».

    Подписи и заголовки разделов в деловых документах устроены иначе, чем в отчёте,
    поэтому правка подписей из gost-report здесь не вызывается. Абзацный отступ,
    даты и реквизиты остаются в отчёте: по ним скрипт не может решить однозначно.
    """

    def run(self) -> list[str]:
        cfg = self.config["text"]
        self.fix_margins(at_least=True)  # раздел 4 задаёт поля «не менее»
        self.fix_normal_style(fonts=cfg.get("fonts", []), size_min=cfg.get("size_min_pt", 0),
                              size_max=cfg.get("size_max_pt", 0), size_exact=0)
        return self.fixed


CLEAN = ("Нарушений не найдено. Проверьте глазами то, что скрипт не видит: состав и порядок реквизитов, "
         "место отметки об исполнителе на последнем листе.")


def render_text(path: Path, config: dict, report: Report, limit: int = 5,
                fixed: list[str] | None = None) -> str:
    """Отчёт как у gost-report, только итог для чистого документа про деловые реквизиты."""
    return base.render_text(path, config, report, limit, fixed, clean=CLEAN)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка деловых документов .docx по ГОСТ Р 7.0.97-2016")
    parser.add_argument("file", type=Path, help="файл .docx")
    parser.add_argument("--config", type=Path, help="настройки инструкции организации (.toml), дополняют ГОСТ")
    parser.add_argument("--json", action="store_true", help="вывод в JSON")
    parser.add_argument("--limit", type=int, default=5, help="сколько мест показывать на одно нарушение")
    parser.add_argument("--fix", action="store_true",
                        help="исправить поля и стиль «Обычный» в копии имя.fixed.docx и показать, что осталось")
    parser.add_argument("--force", action="store_true", help="с --fix: перезаписать готовую копию")
    args = parser.parse_args(argv)

    if args.file.suffix.lower() != ".docx":
        print(f"Нужен файл .docx, а передан {args.file.name}. Файл .doc сначала сохраните как .docx.")
        return 2
    target, fixed = args.file, None
    if args.fix:
        target = base.fixed_path(args.file)
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
        print(base.render_json(target, config, report, fixed))
    else:
        print(render_text(target, config, report, args.limit, fixed), end="")
    return 1 if report.count(ERROR) else 0


if __name__ == "__main__":
    sys.exit(main())
