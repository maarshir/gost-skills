"""Собирает учебную курсовую с типичными ошибками, чтобы попробовать проверку.

    python examples/make_example.py            # создаст example.docx
    python skills/gost-report/scripts/check_docx.py example.docx

Текст работы выдуман, ошибки внесены нарочно: поле, шрифт, заголовок,
подпись рисунка, ссылка на несуществующий источник.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))

from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.shared import Mm  # noqa: E402

from docbuilder import good_document  # noqa: E402


def build(path: str = "example.docx") -> str:
    b = good_document()
    b.doc.sections[0].left_margin = Mm(20)
    for p in b.doc.paragraphs:
        if p.text == "Метод А оказался точнее на выбранном наборе вопросов.":
            for run in p.runs:
                run.font.name = "Arial"
        if p.text == "ЗАКЛЮЧЕНИЕ":
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.page_break_before = False
        if p.text == "1.1 Сравнение":
            p.runs[0].text = "1.1. Сравнение"
        if p.text.startswith("Рисунок 1"):
            p.runs[0].text = "Рис. 1 - Схема поиска."
        if p.text == "Сравнение показало, что метод А подходит лучше.":
            p.runs[0].text = "Сравнение показало, что метод А подходит лучше [3]."
    return b.save(path)


if __name__ == "__main__":
    print("Создан", build(sys.argv[1] if len(sys.argv) > 1 else "example.docx"))
