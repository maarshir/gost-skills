"""Общая часть скриптов проверки .docx: чтение разметки Word, отчёт и исправление в копии.

Этим модулем пользуются check_docx.py навыков gost-report и gost-business.
Здесь только то, что не зависит от стандарта: итоговые свойства абзаца с учётом
стилей, список нарушений с местами, вывод текстом и в JSON, поля страницы
и стиль «Обычный» в копии файла. Сами правила лежат в check_docx.py каждого навыка.
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib

import docx
from lxml import etree
from docx.oxml.ns import qn

if hasattr(sys.stdout, "reconfigure"):
    # Консоль Windows по умолчанию не в UTF-8, а отчёт должен напечататься всегда.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ERROR, WARNING = "ошибка", "замечание"
TWIPS_PER_CM = 566.93
EMU_PER_MM = 36000

CLEAN = "Нарушений не найдено."


# ---------------------------------------------------------------- настройки

def deep_merge(base: dict, extra: dict) -> dict:
    out = dict(base)
    for key, value in extra.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = deep_merge(out[key], value)
        else:
            out[key] = value
    return out


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


# ---------------------------------------------------------------- исправление

def fixed_path(path: Path) -> Path:
    return path.with_name(path.stem + ".fixed.docx")


class Fixer:
    """Исправляет в копии только то, что однозначно следует из профиля.

    Поля страницы ставятся по профилю, у стиля «Обычный» шрифт и размер.
    Шрифт и размер, заданные прямо у текста или в других стилях, не трогаются:
    там скрипт не может знать, задумано это или нет.

    Текст абзаца _replace_in_one_run меняет только внутри одного фрагмента (w:t),
    поэтому оформление остаётся прежним. Навык дополняет класс своими правками
    и своим run.
    """

    def __init__(self, path: Path, config: dict) -> None:
        self.path = path
        self.config = config
        self.doc = docx.Document(str(path))
        self.fixed: list[str] = []

    def run(self) -> list[str]:
        self.fix_margins()
        self.fix_normal_style()
        return self.fixed

    @staticmethod
    def _replace_in_one_run(p, start: int, end: int, new: str) -> bool:
        """Заменяет символы start:end текста абзаца, если они лежат в одном w:t."""
        pos = 0
        for node in p.iter(qn("w:t"), qn("w:tab")):
            text = (node.text or "") if node.tag == qn("w:t") else "\t"
            if pos <= start and end <= pos + len(text):
                if node.tag != qn("w:t"):
                    return False
                a, b = start - pos, end - pos
                node.text = text[:a] + new + text[b:]
                if node.text != node.text.strip():
                    node.set(qn("xml:space"), "preserve")
                return True
            if pos + len(text) > start:
                return False  # место начинается в одном фрагменте, а кончается в другом
            pos += len(text)
        return False

    def fix_margins(self, at_least: bool = False) -> None:
        """Поля по профилю. С at_least значения профиля наименьшие: широкое поле не трогается."""
        cfg = self.config["page"]
        tol = cfg.get("tolerance_mm", 1)
        sides = (("left_margin", "левое", "margin_left_mm"), ("right_margin", "правое", "margin_right_mm"),
                 ("top_margin", "верхнее", "margin_top_mm"), ("bottom_margin", "нижнее", "margin_bottom_mm"))
        many = len(self.doc.sections) > 1
        for n, section in enumerate(self.doc.sections, start=1):
            if section.page_width is None or section.page_height is None:
                continue
            if section.page_width > section.page_height:
                continue  # альбомные части проверка тоже пропускает
            for attr, label, key in sides:
                value = getattr(section, attr)
                if value is None:
                    continue
                mm = value / EMU_PER_MM
                wrong = mm < cfg[key] - tol if at_least else abs(mm - cfg[key]) > tol
                if wrong:
                    setattr(section, attr, round(cfg[key] * EMU_PER_MM))
                    where = f" (часть документа {n})" if many else ""
                    self.fixed.append(f"Поле {label} {mm:.0f} мм → {cfg[key]} мм{where}")

    def _normal_style(self):
        root = self.doc.styles.element
        for s in root.findall(qn("w:style")):
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true", "on"):
                return s
        return None

    def fix_normal_style(self, fonts: list[str] | None = None, size_min: float | None = None,
                         size_max: float = 0, size_exact: float | None = None) -> None:
        """Шрифт и размер стиля «Обычный».

        По умолчанию значения берутся из профиля [text]: один шрифт font, размер size_pt
        или не меньше size_min_pt. Другой навык с другим профилем передаёт их сам:
        fonts допускает несколько шрифтов (неподходящий меняется на первый),
        size_max ограничивает размер сверху.
        """
        style = self._normal_style()
        if style is None:
            return
        cfg = self.config["text"]
        if fonts is None:
            fonts = [cfg["font"]] if cfg.get("font") else []
        if size_min is None:
            size_min = cfg.get("size_min_pt", 0)
        if size_exact is None:
            size_exact = cfg.get("size_pt", 0)
        styles = Styles(self.doc)

        def resolved(getter):
            for src in (style.find(qn("w:rPr")), styles.rpr_default):
                if src is not None:
                    value = getter(src)
                    if value is not None:
                        return value
            return None

        def font_of(rpr):
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None:
                return None
            for attr in ("w:hAnsiTheme", "w:hAnsi", "w:asciiTheme", "w:ascii"):
                val = fonts.get(qn(attr))
                if val:
                    if attr.endswith("Theme"):
                        return styles.theme.get("major" if val.startswith("major") else "minor")
                    return val
            return None

        def size_of(rpr):
            sz = rpr.find(qn("w:sz"))
            return int(sz.get(qn("w:val"))) / 2 if sz is not None else None

        name = style.find(qn("w:name"))
        label = "«Обычный»" if name is None or name.get(qn("w:val")) in ("Normal", "Обычный") else (
            f"«{name.get(qn('w:val'))}»")
        rpr = style.find(qn("w:rPr"))

        need_font = fonts[0] if fonts else ""
        current = resolved(font_of)
        if fonts and current not in fonts:
            if rpr is None:
                rpr = style.get_or_add_rPr() if hasattr(style, "get_or_add_rPr") else etree.SubElement(
                    style, qn("w:rPr"))
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = etree.SubElement(rpr, qn("w:rFonts"))
            # Тема перекрывает явное имя шрифта, поэтому ссылки на тему убираются
            for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme"):
                rfonts.attrib.pop(qn(attr), None)
            for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                rfonts.set(qn(attr), need_font)
            self._order_rpr(rpr)
            self.fixed.append(f"Шрифт стиля {label}: {current or 'не задан'} → {need_font}")

        size = resolved(size_of) or 10.0
        target = None
        if size_exact and abs(size - size_exact) > 0.01:
            target = size_exact
        elif not size_exact and size < size_min:
            target = size_min
        elif not size_exact and size_max and size > size_max + 0.01:
            target = size_max
        if target:
            if rpr is None:
                rpr = style.get_or_add_rPr() if hasattr(style, "get_or_add_rPr") else etree.SubElement(
                    style, qn("w:rPr"))
            half = str(round(target * 2))
            for tag in ("w:sz", "w:szCs"):
                el = rpr.find(qn(tag))
                if el is None:
                    el = etree.SubElement(rpr, qn(tag))
                el.set(qn("w:val"), half)
            self._order_rpr(rpr)
            self.fixed.append(f"Размер шрифта стиля {label}: {size:g} пт → {target:g} пт")

    @staticmethod
    def _order_rpr(rpr) -> None:
        """Word строго следит за порядком элементов в w:rPr: sz и szCs ставятся на своё место."""
        order = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
                 "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
                 "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
                 "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
                 "specVanish", "oMath"]
        rank = {qn(f"w:{t}"): i for i, t in enumerate(order)}
        children = list(rpr)
        children.sort(key=lambda el: rank.get(el.tag, len(order)))
        for el in children:
            rpr.remove(el)
            rpr.append(el)

    def save(self, out: Path) -> None:
        self.doc.save(str(out))


# ---------------------------------------------------------------- вывод

def render_text(path: Path, config: dict, report: Report, limit: int = 5,
                fixed: list[str] | None = None, clean: str = CLEAN) -> str:
    """Отчёт для человека. clean печатается, когда нарушений нет."""
    lines = [f"Проверка: {path.name}, профиль «{config.get('name', '')}»"]
    if fixed is not None:
        lines.append("ИСПРАВЛЕНО")
        lines += [f"- {item}" for item in fixed] or ["- ничего: однозначных ошибок для исправления нет"]
        lines += ["", "ОСТАЛОСЬ ИСПРАВИТЬ ВРУЧНУЮ"]
    lines += [f"Ошибок: {report.count(ERROR)}, замечаний: {report.count(WARNING)}", ""]
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
        lines.append(clean)
    return "\n".join(lines).rstrip() + "\n"


def render_json(path: Path, config: dict, report: Report, fixed: list[str] | None = None) -> str:
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
    if fixed is not None:
        data["fixed"] = fixed
    return json.dumps(data, ensure_ascii=False, indent=2)
