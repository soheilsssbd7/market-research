#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the Pasargad F-161 Word report from the adjacent Markdown source.

Run from any working directory with: python _f161/build.py
Requires python-docx (see requirements.txt). All inputs/outputs stay in _f161/.
"""
from pathlib import Path
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "F161_REPORT_FINAL.md"
LEDGER = HERE / "F161_EVIDENCE_LEDGER.md"
OUTPUT = HERE / "F161_REPORT_FINAL.docx"


def set_rtl(paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
    paragraph.alignment = alignment
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.append(OxmlElement("w:bidi"))
    for run in paragraph.runs:
        rpr = run._r.get_or_add_rPr()
        if rpr.find(qn("w:rtl")) is None:
            rpr.append(OxmlElement("w:rtl"))
        run.font.name = "Arial"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Arial")


def add_inline(paragraph, text):
    """Render Markdown bold markers while preserving citations/URLs as text."""
    for part in re.split(r"(\*\*.*?\*\*)", text):
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**")
        run = paragraph.add_run(part[2:-2] if is_bold else part)
        run.bold = is_bold
        run.font.name = "Arial"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Arial")


def split_table_row(line):
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def add_markdown_table(document, rows):
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = document.add_table(rows=0, cols=width)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for column, cell in enumerate(cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            value = values[column] if column < len(values) else ""
            paragraph = cell.paragraphs[0]
            add_inline(paragraph, value)
            set_rtl(paragraph)
            if row_index == 0:
                for run in paragraph.runs:
                    run.bold = True
            tcpr = cell._tc.get_or_add_tcPr()
            if tcpr.find(qn("w:bidiVisual")) is None:
                tcpr.append(OxmlElement("w:bidiVisual"))
    document.add_paragraph()


def append_evidence_register(document):
    """Append the source register as readable evidence cards, not a wide table."""
    if not LEDGER.is_file():
        raise FileNotFoundError(f"Evidence ledger missing: {LEDGER}")
    document.add_page_break()
    heading = document.add_paragraph(style="Heading 1")
    add_inline(heading, "پیوست الف ـ دفتر شواهد تفصیلی")
    set_rtl(heading)
    intro = document.add_paragraph()
    add_inline(intro, "این پیوست، ردیف‌های دفتر شواهد را برای ممیزی در همان سند عرضه می‌کند. ردیف‌ها معادل منابع مستقل نیستند؛ بازنشرها، نقل‌قول‌ها و مطالعات محدود باید با خوشه و حدود تعمیم خود خوانده شوند.")
    set_rtl(intro)

    records = []
    for line in LEDGER.read_text(encoding="utf-8").splitlines():
        candidate = line.strip()
        if not candidate.startswith("|"):
            continue
        row = split_table_row(candidate)
        if row and re.fullmatch(r"پ\d+", row[0]):
            records.append(row)

    for row in records:
        row += [""] * max(0, 6 - len(row))
        source_id, title, date, url, evidence, status = row[:6]
        paragraph = document.add_paragraph(style="Heading 2")
        add_inline(paragraph, f"{source_id} ـ {title}")
        set_rtl(paragraph)
        for label, value in (("تاریخ انتشار", date), ("نشانی", url),
                             ("یافته و محدودیت", evidence), ("جایگاه شمارشی", status)):
            p = document.add_paragraph()
            add_inline(p, f"{label}: {value}")
            set_rtl(p)
    print(f"Evidence-register appendix records: {len(records)}")


def build():
    if not SOURCE.is_file():
        raise FileNotFoundError(f"Markdown source missing: {SOURCE}")
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    for name, size in (("Normal", 11), ("Title", 22), ("Heading 1", 17),
                       ("Heading 2", 14), ("Heading 3", 12)):
        style = document.styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        style.paragraph_format.space_after = Pt(6)

    index = 0
    while index < len(lines):
        text = lines[index].strip()
        if not text or text in ("---", "***"):
            index += 1
            continue
        if text.startswith("|"):
            rows = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                row = split_table_row(lines[index])
                # Omit Markdown table separator rows.
                if not row or not all(re.fullmatch(r":?-{3,}:?", c.replace(" ", "")) for c in row):
                    rows.append(row)
                index += 1
            add_markdown_table(document, rows)
            continue
        if text.startswith("#"):
            match = re.match(r"^(#+)\s+(.*)$", text)
            level, title = len(match.group(1)), match.group(2)
            style = "Title" if level == 1 else f"Heading {min(level - 1, 3)}"
            paragraph = document.add_paragraph(style=style)
            add_inline(paragraph, title)
            set_rtl(paragraph)
        elif re.match(r"^[-*]\s+", text):
            paragraph = document.add_paragraph(style="List Bullet")
            add_inline(paragraph, re.sub(r"^[-*]\s+", "", text))
            set_rtl(paragraph)
        elif re.match(r"^\d+[.)]\s+", text):
            paragraph = document.add_paragraph(style="List Number")
            add_inline(paragraph, re.sub(r"^\d+[.)]\s+", "", text))
            set_rtl(paragraph)
        else:
            paragraph = document.add_paragraph()
            add_inline(paragraph, text)
            set_rtl(paragraph)
        index += 1

    append_evidence_register(document)
    document.core_properties.title = "گزارش نهایی دامنه‌مند پژوهش بازار بانک پاسارگاد"
    document.core_properties.subject = "ارزیابی ارتباطات و شواهد عمومی؛ تاریخ قطع ۲۱ اکتبر ۲۰۲۶"
    document.save(OUTPUT)
    # Match the established report-build convention: count whitespace-separated
    # tokens in Word paragraphs and table cells, not Markdown source syntax.
    built = Document(OUTPUT)
    word_count = sum(len(paragraph.text.split()) for paragraph in built.paragraphs)
    for table in built.tables:
        for row in table.rows:
            for cell in row.cells:
                word_count += len(cell.text.split())
    print(f"Saved: {OUTPUT.relative_to(HERE.parent)} ({OUTPUT.stat().st_size:,} bytes)")
    print(f"DOCX whitespace word count (paragraphs + table cells): {word_count:,}")
    print(f"Paragraphs: {len(built.paragraphs):,}; tables: {len(built.tables):,}")
    print(f"Markdown whitespace tokens (includes Markdown syntax): {len(SOURCE.read_text(encoding='utf-8').split()):,}")


if __name__ == "__main__":
    build()
