#!/usr/bin/env python3
"""Проверка списка литературы по ГОСТ Р 7.0.100-2018 или по своей методичке.

Вход: текстовый файл (одна запись в строке), .docx курсовой (скрипт сам найдёт
раздел со списком) или «-» для чтения из stdin. Файл не меняется.

    python check_bibliography.py список.txt
    python check_bibliography.py работа.docx --config методичка.toml
    python check_bibliography.py список.txt --json

Код выхода: 0 если ошибок нет (замечания допустимы), 1 если есть ошибки,
2 если файл не удалось прочитать или в .docx не нашёлся список.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib

if hasattr(sys.stdout, "reconfigure"):
    # Консоль Windows по умолчанию не в UTF-8, а отчёт должен напечататься всегда.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(__file__).resolve().parent
DEFAULT_PROFILE = HERE.parent / "assets" / "gost-7.0.100.toml"

ERROR, WARNING = "ошибка", "замечание"
DASHES = "–—"

# Заголовки раздела со списком в .docx (сравниваются в верхнем регистре)
HEADINGS = {
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ЛИТЕРАТУРЫ", "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ",
    "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "СПИСОК ИСТОЧНИКОВ", "ЛИТЕРАТУРА",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ И ЛИТЕРАТУРЫ",
}

NUMBER_RE = re.compile(r"^\s*(?:\[(\d{1,3})\]|(\d{1,3})[.)])\s+")
URL_RE = re.compile(r"(?:https?://|www\.)\S+?(?=[.,;)]?(?:\s|$))", re.IGNORECASE)
ACCESS_RE = re.compile(r"\(\s*дата\s+обращения\s*:?\s*([^)]*)\)", re.IGNORECASE)
ISBN_RE = re.compile(r"ISBN[\s\dXx\-–]+")
YEAR_RE = re.compile(r"(?<!\d)(1[5-9]\d{2}|20\d{2})(?!\d)")
PERSON_RE = re.compile(r"[А-ЯЁA-Z]\.\s?(?:[А-ЯЁA-Z]\.\s?)?[А-ЯЁA-Z][а-яёa-z\-]+")
LEGAL_RE = re.compile(
    r"федеральн\w* закон|кодекс|указ президента|постановлени|распоряжени|конституци|"
    r"\bГОСТ\b|национальный стандарт|межгосударственный стандарт|СанПиН|\bСП \d",
    re.IGNORECASE,
)
PUBLISHER_RE = re.compile(r"[А-ЯЁA-Z][^:/]{0,40}\s:\s[^,:/]{1,80},\s*(?:1[5-9]\d{2}|20\d{2})")
CONTENT_RE = re.compile(r"текст\s*:\s*([а-яё]+)", re.IGNORECASE)

# Код: (серьёзность, текст, пункт стандарта)
RULES = {
    "hyphen-dash": (ERROR, "Между областями дефис « - », нужно тире «. – »", "4.6.2"),
    "dash-without-dot": (ERROR, "Тире без точки перед ним, нужно «. – »", "4.6.2"),
    "no-area-dash": (WARNING, "Области разделены точкой, по ГОСТу «. – »", "4.6.4"),
    "end-dot": (ERROR, "Запись не заканчивается точкой", "4.6.1"),
    "double-punct": (ERROR, "Два знака подряд («..», «,,», «,.»)", "4.6.1"),
    "space-colon": (ERROR, "У двоеточия нужен пробел с обеих сторон: «Москва : Наука»", "4.6.5"),
    "space-slash": (ERROR, "У косой черты нужен пробел с обеих сторон: «Заглавие / И. О. Фамилия», « // »", "4.6.5"),
    "space-before": (ERROR, "Пробел перед точкой или запятой", "4.6.5"),
    "heading-comma": (ERROR, "После фамилии в начале записи нужна запятая: «Петров, П. П.»", "ГОСТ 7.80-2000"),
    "initials-space": (ERROR, "Инициалы пишутся через пробел: «П. П. Петров»", "ГОСТ 7.80-2000"),
    "et-al": (ERROR, "«и др.» пишется в квадратных скобках: «[и др.]»", "5.2.6.8"),
    "authors-many": (WARNING, "Пять и больше авторов: после косой черты первые три и «[и др.]»", "5.2.6.8"),
    "heading-four": (WARNING, "Четыре и больше авторов: запись начинается с заглавия", "5.2.6.8"),
    "year-missing": (ERROR, "Не найден год издания или публикации", "5.5"),
    "year-future": (ERROR, "Год ещё не наступил", "5.5"),
    "place-abbr": (ERROR, "Место издания сокращено, методичка требует полностью: «Москва», «Санкт-Петербург»", "настройки"),
    "volume-missing": (WARNING, "У книги не указан объём: «– 256 с.»", "5.6"),
    "pages-missing": (WARNING, "У статьи не указаны страницы: «– С. 15–20.»", "7"),
    "pages-abbr": (ERROR, "«стр.» не сокращение по ГОСТу: «с.» для объёма, «С.» для страниц", "ГОСТ Р 7.0.12"),
    "pages-hyphen": (ERROR, "Диапазон страниц через дефис, нужно тире: «С. 15–20»", "7"),
    "pages-space": (ERROR, "После «С.» нужен пробел: «С. 15»", "4.6.5"),
    "pages-case": (ERROR, "Страницы статьи «С. 15–20» с большой буквы, объём книги «256 с.» с маленькой", "5.6, 7"),
    "number-sign": (ERROR, "После «№» нужен пробел: «№ 3»", "4.6.5"),
    "stale": (WARNING, "Источник старше, чем разрешает методичка", "настройки"),
    "url-prefix": (ERROR, "Адрес без «URL:»", "5.8"),
    "url-access-date": (ERROR, "У адреса нет даты обращения: «(дата обращения: 15.09.2026)»", "5.8"),
    "url-date-format": (ERROR, "Дата обращения не в виде ДД.ММ.ГГГГ или такой даты нет", "5.8"),
    "url-date-future": (ERROR, "Дата обращения ещё не наступила", "5.8"),
    "old-electronic": (WARNING, "«[Электронный ресурс]» из старого ГОСТ 7.1-2003, сейчас в конце «– Текст : электронный.»", "5.10"),
    "old-access": (WARNING, "«Режим доступа:» из старого ГОСТ 7.1-2003, сейчас «URL:»", "5.8"),
    "content-type-format": (ERROR, "Вид содержания пишется так: «Текст : электронный» или «Текст : непосредственный»", "5.10"),
    "content-type-missing": (ERROR, "Методичка требует в конце «– Текст : непосредственный.» или «электронный»", "5.10"),
    "content-type-mismatch": (WARNING, "Есть адрес в сети, а указано «непосредственный»", "5.10"),
    "numbering": (ERROR, "Номер записи не совпадает с её местом в списке", "ГОСТ 7.32, 6.16"),
    "duplicate": (ERROR, "Такая запись или такой адрес уже есть в списке", "ГОСТ 7.32, 6.16"),
    "order": (ERROR, "Нарушен порядок списка", "настройки"),
}


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

def short(text: str, limit: int = 45) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


@dataclass
class Place:
    entry: int
    text: str = ""
    detail: str = ""

    def describe(self) -> str:
        out = f"запись {self.entry} «{short(self.text)}»"
        return f"{out}: {self.detail}" if self.detail else out


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    rule: str
    places: list[Place] = field(default_factory=list)


@dataclass
class Report:
    issues: list[Issue] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    entries: int = 0

    def add(self, code: str, place: Place, severity: str | None = None) -> None:
        default, message, rule = RULES[code]
        severity = severity or default
        for issue in self.issues:
            if issue.code == code and issue.severity == severity:
                issue.places.append(place)
                return
        self.issues.append(Issue(severity, code, message, rule, [place]))

    def count(self, severity: str) -> int:
        return sum(1 for i in self.issues if i.severity == severity)

    def codes(self) -> set[str]:
        return {i.code for i in self.issues}


# ---------------------------------------------------------------- разбор записи

@dataclass
class Entry:
    index: int          # место в списке, с 1
    number: int | None  # номер, написанный в тексте
    text: str           # без номера


def split_number(line: str) -> tuple[int | None, str]:
    m = NUMBER_RE.match(line)
    if not m:
        return None, line.strip()
    return int(m.group(1) or m.group(2)), line[m.end():].strip()


def make_entries(lines: list[str]) -> list[Entry]:
    out = []
    for line in lines:
        if not line.strip():
            continue
        number, text = split_number(line)
        out.append(Entry(len(out) + 1, number, text))
    return out


def mask(text: str) -> str:
    """Закрывает адреса, дату обращения и «URL:» буквами той же длины.

    В них двоеточия, косые черты и точки не подчиняются правилам знаков,
    а позиции остальных символов не меняются.
    """
    def fill(m: re.Match) -> str:
        return "x" * len(m.group(0))

    text = URL_RE.sub(fill, text)
    text = ACCESS_RE.sub(fill, text)
    text = re.sub(r"URL\s*:", fill, text)
    text = re.sub(r"Режим доступа\s*:", fill, text, flags=re.IGNORECASE)
    return text


def is_legal(text: str) -> bool:
    return bool(LEGAL_RE.search(text))


def years(text: str) -> list[int]:
    clean = ISBN_RE.sub(" ", mask(text))
    clean = re.sub(r"ГОСТ(?:\s?Р)?\s?[\d.]+[–—\-]\d{4}", " ", clean)
    return [int(y) for y in YEAR_RE.findall(clean)]


def responsibility(text: str) -> str:
    """Сведения об ответственности: от первой « / » до конца первой области."""
    if " / " not in text:
        return ""
    rest = text.split(" / ", 1)[1]
    return re.split(r"\.\s[–—]\s| // ", rest, maxsplit=1)[0]


def check_entry(entry: Entry, config: dict, today: dt.date, report: Report) -> None:
    text = entry.text
    masked = mask(text)
    place = lambda detail="": Place(entry.index, text, detail)  # noqa: E731
    legal = is_legal(text)
    has_url = bool(URL_RE.search(text))
    is_part = " // " in text or "//" in masked

    # Знаки между областями
    separator = config.get("areas", {}).get("separator", "dash")
    if separator != "any":
        if re.search(r"\s-\s", masked):
            report.add("hyphen-dash", place())
        for m in re.finditer(rf"(\S)\s[{DASHES}]\s", masked):
            if m.group(1) in ".":
                continue
            severity = ERROR if (m.group(1).isdigit() or m.group(1) in ")]»") else WARNING
            report.add("dash-without-dot", place(f"«…{short(text[max(0, m.start() - 12):m.end() + 6], 30)}…»"),
                       severity)
        if separator == "dash" and not re.search(rf"\.\s[{DASHES}]\s", masked) and len(text) > 40:
            report.add("no-area-dash", place())
    if not text.rstrip().endswith("."):
        report.add("end-dot", place())
    if re.search(r"(?<!\.)\.\.(?!\.)|,,|,\.|;;", masked):
        report.add("double-punct", place())

    # Пробелы вокруг предписанных знаков
    if re.search(r"\S:|:(?=\S)", masked):
        report.add("space-colon", place())
    if re.search(r"[^\s/]/|/(?=[^\s/])", masked):
        report.add("space-slash", place())
    if re.search(r"\s[.,](?=\s|$)", masked):
        report.add("space-before", place())

    # Авторы
    if re.match(r"^[А-ЯЁA-Z][а-яёa-z\-]+\s+[А-ЯЁA-Z]\.\s?(?:[А-ЯЁA-Z]\.)?\s", text):
        report.add("heading-comma", place())
    if re.search(r"(?<![А-Яа-яЁёA-Za-z])[А-ЯЁA-Z]\.[А-ЯЁA-Z]\.", masked):
        report.add("initials-space", place())
    if re.search(r"(?<!\[)\b(и др\.|et al\.)", text):
        report.add("et-al", place())
    persons = PERSON_RE.findall(responsibility(text))
    if len(persons) >= 5 and "[и др.]" not in text and "[et al.]" not in text:
        report.add("authors-many", place(f"авторов {len(persons)}"))
    starts_with_author = re.match(r"^[А-ЯЁA-Z][а-яёa-z\-]+,\s[А-ЯЁA-Z]\.", text)
    if starts_with_author and (len(persons) >= 4 or "[и др.]" in text):
        report.add("heading-four", place())

    # Год, место, объём, страницы
    found = years(text)
    publication = [y for y in found if y <= today.year]
    if not found:
        report.add("year-missing", place(), WARNING if has_url else ERROR)
    elif max(found) > today.year:
        report.add("year-future", place(str(max(found))))
    if not config.get("place", {}).get("allow_abbreviations", True):
        m = re.search(r"(?<![А-Яа-яЁё])(М|СПб|Л)\.\s?:", masked)
        if m:
            report.add("place-abbr", place(f"«{m.group(1)}.»"))
    if not is_part and not legal and PUBLISHER_RE.search(masked) and not re.search(r"\d+\s*[сСpP]\.", masked):
        report.add("volume-missing", place())
    if is_part and not has_url and not re.search(r"\b[СсPp]\.\s?\d", masked):
        report.add("pages-missing", place())
    if re.search(r"(?<![А-Яа-яЁё])[Сс]тр\.|страниц", text):
        report.add("pages-abbr", place())
    if re.search(r"\b[СсPp]\.\s?\d+\s?-\s?\d+", masked):
        report.add("pages-hyphen", place())
    if re.search(r"(?<![А-Яа-яЁёA-Za-z])[СсP]\.\d", masked):
        report.add("pages-space", place())
    if (is_part and re.search(rf"[{DASHES}]\sс\.\s?\d", masked)) or re.search(r"\d\s?С\.(?!\s?\d)", masked):
        report.add("pages-case", place())
    if re.search(r"№\d|\bN\s?\d", masked):
        report.add("number-sign", place())
    max_age = int(config.get("freshness", {}).get("max_age_years", 0) or 0)
    if max_age and publication and not legal and max(publication) < today.year - max_age:
        report.add("stale", place(f"{max(publication)} год"))

    # Электронные ресурсы
    if "[электронный ресурс]" in text.lower():
        report.add("old-electronic", place())
    old_access = bool(re.search(r"режим доступа", text, re.IGNORECASE))
    if old_access:
        report.add("old-access", place())
    for m in URL_RE.finditer(text):
        before = text[: m.start()].rstrip()
        if not re.search(r"URL\s*:$", before) and not old_access:
            report.add("url-prefix", place(short(m.group(0), 40)))
            break
    if has_url:
        access = ACCESS_RE.search(text)
        if not access:
            report.add("url-access-date", place())
        else:
            raw = access.group(1).strip()
            try:
                if not re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", raw):
                    raise ValueError
                date = dt.datetime.strptime(raw, "%d.%m.%Y").date()
            except ValueError:
                report.add("url-date-format", place(f"«{raw}»"))
            else:
                if date > today:
                    report.add("url-date-future", place(raw))

    # Вид содержания
    content = CONTENT_RE.search(text)
    if content:
        if not re.search(r"Текст : (непосредственный|электронный)\b", text):
            report.add("content-type-format", place(f"«{short(content.group(0), 30)}»"))
        if has_url and content.group(1).lower() == "непосредственный":
            report.add("content-type-mismatch", place())
    elif config.get("content_type", {}).get("required", False) and not legal:
        report.add("content-type-missing", place())


# ---------------------------------------------------------------- весь список

def sort_key(text: str) -> tuple[int, str]:
    letters = re.sub(r"^[^0-9A-Za-zА-Яа-яЁё]+", "", text).lower().replace("ё", "е")
    group = 1 if re.match(r"[a-z]", letters) else 0  # сначала кириллица, потом латиница
    return group, letters


def check_list(entries: list[Entry], config: dict, report: Report) -> None:
    numbered = [e for e in entries if e.number is not None]
    if numbered:
        for e in entries:
            if e.number != e.index:
                got = "нет номера" if e.number is None else f"номер {e.number}"
                report.add("numbering", Place(e.index, e.text, f"{got}, ждём {e.index}"))

    seen_text: dict[str, int] = {}
    seen_url: dict[str, int] = {}
    for e in entries:
        norm = re.sub(r"[\s–—\-]+", " ", e.text.lower()).strip(" .")
        urls = [u.lower().rstrip("/") for u in URL_RE.findall(e.text)]
        first = seen_text.get(norm) or next((seen_url[u] for u in urls if u in seen_url), None)
        if first:
            report.add("duplicate", Place(e.index, e.text, f"повторяет запись {first}"))
        seen_text.setdefault(norm, e.index)
        for u in urls:
            seen_url.setdefault(u, e.index)

    order = config.get("order", {})
    if order.get("mode", "any") != "alphabet":
        return
    rest = entries
    if order.get("legal_first", False):
        seen_ordinary = False
        for e in entries:
            if is_legal(e.text):
                if seen_ordinary:
                    report.add("order", Place(e.index, e.text, "закон или стандарт после обычных источников"))
            else:
                seen_ordinary = True
        rest = [e for e in entries if not is_legal(e.text)]
    for prev, cur in zip(rest, rest[1:]):
        if sort_key(cur.text) < sort_key(prev.text):
            report.add("order", Place(cur.index, cur.text, f"по алфавиту должна стоять раньше записи {prev.index}"))


def check_entries(lines: list[str], config: dict | None = None, today: dt.date | None = None) -> Report:
    config = config if config is not None else load_config()
    today = today or dt.date.today()
    entries = make_entries(lines)
    report = Report(entries=len(entries))
    for entry in entries:
        check_entry(entry, config, today, report)
    check_list(entries, config, report)
    return report


# ---------------------------------------------------------------- чтение файлов

def has_numbering(p) -> bool:
    """Автоматический номер Word: у самого абзаца или в его стиле (например, «Нумерованный список»)."""
    if p._p.pPr is not None and p._p.pPr.numPr is not None:
        return True
    style = p.style
    while style is not None:
        ppr = style.element.pPr
        if ppr is not None and ppr.numPr is not None:
            return True
        style = style.base_style
    return False


def read_docx(path: Path) -> tuple[list[str], list[str]]:
    """Абзацы раздела со списком литературы и заметки о том, что не проверено."""
    import docx  # только для .docx

    paragraphs = [p for p in docx.Document(str(path)).paragraphs]
    start = None
    for i, p in enumerate(paragraphs):
        if " ".join(p.text.split()).upper().rstrip(".:") in HEADINGS:
            start = i + 1  # берём последний такой заголовок: первый обычно в содержании
    if start is None:
        raise LookupError("В документе не найден заголовок «Список использованных источников» "
                          "или похожий. Скопируйте список в текстовый файл, по записи в строке.")
    lines, auto_numbered = [], False
    for p in paragraphs[start:]:
        text = p.text.strip()
        if text.upper().startswith("ПРИЛОЖЕНИЕ"):
            break
        if not text:
            continue
        lines.append(text)
        if has_numbering(p):
            auto_numbered = True
    notes = []
    if auto_numbered:
        notes.append("Нумерация списка автоматическая (номера ставит Word), поэтому номера не проверялись.")
    return lines, notes


def read_source(name: str) -> tuple[list[str], list[str]]:
    if name == "-":
        return sys.stdin.read().splitlines(), []
    path = Path(name)
    if path.suffix.lower() == ".docx":
        return read_docx(path)
    if path.suffix.lower() == ".doc":
        raise ValueError("Файл .doc сначала сохраните как .docx или скопируйте список в текстовый файл.")
    return path.read_text(encoding="utf-8-sig").splitlines(), []


# ---------------------------------------------------------------- вывод

def render_text(name: str, config: dict, report: Report, limit: int = 5) -> str:
    lines = [f"Проверка списка: {name}, профиль «{config.get('name', '')}»",
             f"Записей: {report.entries}. Ошибок: {report.count(ERROR)}, замечаний: {report.count(WARNING)}", ""]
    for severity, title in ((ERROR, "ОШИБКИ"), (WARNING, "ЗАМЕЧАНИЯ")):
        items = [i for i in report.issues if i.severity == severity]
        if not items:
            continue
        lines.append(title)
        for n, issue in enumerate(items, start=1):
            rule = f" (п. {issue.rule})" if issue.rule[0].isdigit() else f" ({issue.rule})"
            count = f", записей: {len(issue.places)}" if len(issue.places) > 1 else ""
            lines.append(f"{n}. {issue.message}{rule}{count}")
            for place in issue.places[:limit]:
                lines.append(f"   {place.describe()}")
            if len(issue.places) > limit:
                lines.append(f"   и ещё {len(issue.places) - limit}")
        lines.append("")
    if report.notes:
        lines.append("НЕ ПРОВЕРЕНО")
        lines += [f"- {note}" for note in report.notes]
        lines.append("")
    if not report.issues:
        lines.append("Нарушений не найдено. Проверьте глазами то, что скрипт не видит: "
                     "верны ли сами данные (авторы, годы, страницы) и есть ли на каждый источник ссылка в тексте.")
    return "\n".join(lines).rstrip() + "\n"


def render_json(name: str, config: dict, report: Report) -> str:
    return json.dumps({
        "file": name,
        "profile": config.get("name", ""),
        "entries": report.entries,
        "errors": report.count(ERROR),
        "warnings": report.count(WARNING),
        "issues": [{
            "severity": i.severity, "code": i.code, "message": i.message, "rule": i.rule,
            "places": [{"entry": p.entry, "text": p.text, "detail": p.detail} for p in i.places],
        } for i in report.issues],
        "notes": report.notes,
    }, ensure_ascii=False, indent=2)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка списка литературы по ГОСТ Р 7.0.100-2018")
    parser.add_argument("file", help="текстовый файл (запись в строке), .docx или «-» для stdin")
    parser.add_argument("--config", type=Path, help="настройки методички (.toml), дополняют профиль ГОСТ")
    parser.add_argument("--json", action="store_true", help="вывод в JSON")
    parser.add_argument("--limit", type=int, default=5, help="сколько записей показывать на одно нарушение")
    parser.add_argument("--today", type=dt.date.fromisoformat, help="сегодняшняя дата ГГГГ-ММ-ДД (для проверки будущих дат)")
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
        lines, notes = read_source(args.file)
    except FileNotFoundError as e:
        print(f"Файл не найден: {e.filename or e}")
        return 2
    except tomllib.TOMLDecodeError as e:
        print(f"Ошибка в файле настроек: {e}")
        return 2
    except (LookupError, ValueError, UnicodeDecodeError) as e:
        print(str(e))
        return 2
    except Exception as e:  # повреждённый .docx и т. п.
        print(f"Не удалось прочитать {args.file}: {e}")
        return 2
    report = check_entries(lines, config, args.today)
    report.notes += notes
    if not report.entries:
        print("Список пустой: не найдено ни одной записи.")
        return 2
    name = "stdin" if args.file == "-" else Path(args.file).name
    print(render_json(name, config, report) if args.json else render_text(name, config, report, args.limit), end="\n" if args.json else "")
    return 1 if report.count(ERROR) else 0


if __name__ == "__main__":
    sys.exit(main())
