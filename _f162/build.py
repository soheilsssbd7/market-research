#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the F-161 Persian report and source-register appendix as a styled DOCX.

Run: python _f162/build.py
The document explicitly requests B Nazanin. Rendering requires that font to be
installed on the machine opening the DOCX; no font files are bundled.
"""
from pathlib import Path
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "F162_REPORT_FINAL.md"
LEDGER = HERE / "F162_EVIDENCE_LEDGER.md"
OUTPUT = HERE / "F162_REPORT_FINAL.docx"
FONT = "B Nazanin"
NAVY = "17365D"
TEAL = "167D8D"
PALE = "EAF3F5"
PALE_BLUE = "EAF0F7"
STRIPE = "F5F8FA"
INK = "263746"
MUTED = "5C6B78"
LINK = "087E8B"


def set_run_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None:
        run.italic = italic
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{attr}"), FONT)
    rtl = rpr.find(qn("w:rtl"))
    if rtl is None:
        rtl = OxmlElement("w:rtl")
        rpr.append(rtl)
    return run


def set_rtl(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY):
    paragraph.alignment = alignment
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.append(OxmlElement("w:bidi"))
    if ppr.find(qn("w:widowControl")) is None:
        ppr.append(OxmlElement("w:widowControl"))
    for run in paragraph.runs:
        set_run_font(run)


def add_hyperlink(paragraph, text, url):
    rel = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{attr}"), FONT)
    rpr.append(fonts)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), LINK)
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    rtl = OxmlElement("w:rtl")
    rpr.append(rtl)
    run.append(rpr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_inline(paragraph, text, base_size=None):
    """Render bold Markdown and links while preserving readable URL text."""
    token_re = re.compile(r"(\*\*.*?\*\*|\[[^\]]+\]\(https?://[^)]+\)|https?://[^\s]+)")
    for part in token_re.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, size=base_size, bold=True, color=NAVY)
        elif part.startswith("[") and "](" in part:
            match = re.fullmatch(r"\[([^\]]+)\]\((https?://[^)]+)\)", part)
            if match:
                add_hyperlink(paragraph, match.group(1), match.group(2))
            else:
                set_run_font(paragraph.add_run(part), size=base_size)
        elif part.startswith("http"):
            url = part.rstrip(".,;:،؛)]}")
            suffix = part[len(url):]
            add_hyperlink(paragraph, url, url)
            if suffix:
                set_run_font(paragraph.add_run(suffix), size=base_size)
        else:
            set_run_font(paragraph.add_run(part), size=base_size)


def set_cell_shading(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    margins = tcpr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tcpr.append(margins)
    for m, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            margins.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def repeat_table_header(row):
    trpr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    trpr.append(tbl_header)


def set_table_rtl(table):
    tblpr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tblpr.append(bidi)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True


def split_table_row(line):
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def add_markdown_table(document, rows):
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = document.add_table(rows=0, cols=width)
    table.style = "Table Grid"
    set_table_rtl(table)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        if row_index == 0:
            repeat_table_header(table.rows[-1])
        for column, cell in enumerate(cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            value = values[column] if column < len(values) else ""
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.05
            add_inline(paragraph, value, base_size=9.6)
            set_rtl(paragraph, WD_ALIGN_PARAGRAPH.RIGHT)
            if row_index == 0:
                set_cell_shading(cell, NAVY)
                for run in paragraph.runs:
                    set_run_font(run, size=10.5, bold=True, color="FFFFFF")
            elif row_index % 2 == 0:
                set_cell_shading(cell, STRIPE)
    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)
    spacer.paragraph_format.space_before = Pt(0)


def add_callout(document, text, label=None):
    p = document.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.08)
    p.paragraph_format.right_indent = Inches(0.08)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_together = True
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), PALE)
    ppr.append(shd)
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), TEAL)
    borders.append(left)
    ppr.append(borders)
    if label:
        add_inline(p, f"{label} ـ ", base_size=11.5)
        if p.runs:
            p.runs[-1].bold = True
    add_inline(p, text, base_size=11.5)
    set_rtl(p)
    return p


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, size=10, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t")
    result.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instr, separate, result, end):
        run._r.append(node)


def add_header_footer(section):
    section.different_first_page_header_footer = True
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.paragraph_format.space_after = Pt(4)
    add_inline(header, "پروندهٔ پژوهشی اف۱۶۲  |  بانک صادرات ایران", base_size=10)
    set_rtl(header, WD_ALIGN_PARAGRAPH.RIGHT)
    ppr = header._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "5")
    bottom.set(qn("w:color"), TEAL)
    p_bdr.append(bottom)
    ppr.append(p_bdr)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(3)
    add_inline(footer, "گزارش تحلیلی دامنه‌مند   •   ", base_size=9.5)
    add_page_field(footer)
    set_rtl(footer, WD_ALIGN_PARAGRAPH.CENTER)


def configure_styles(document):
    normal = document.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(12.5)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal._element.rPr.rFonts.set(qn("w:cs"), FONT)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    style_specs = {
        "Title": (29, NAVY, True),
        "Subtitle": (17, TEAL, False),
        "Heading 1": (19, NAVY, True),
        "Heading 2": (15.5, TEAL, True),
        "Heading 3": (13.5, NAVY, True),
    }
    for name, (size, color, bold) in style_specs.items():
        style = document.styles[name]
        style.font.name = FONT
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style._element.rPr.rFonts.set(qn("w:cs"), FONT)
        style.paragraph_format.space_before = Pt(12 if name.startswith("Heading") else 0)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True

    for name in ("List Bullet", "List Number"):
        style = document.styles[name]
        style.font.name = FONT
        style.font.size = Pt(12.2)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.2


def add_cover(document, title, metadata):
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(42)
    p.paragraph_format.space_after = Pt(18)
    add_inline(p, "گزارش پژوهشی  |  F–162", base_size=14)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        set_run_font(run, size=14, bold=True, color=TEAL)

    p = document.add_paragraph(style="Title")
    p.paragraph_format.space_before = Pt(35)
    p.paragraph_format.space_after = Pt(10)
    add_inline(p, title, base_size=29)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)

    subtitle = metadata.get("subtitle", "ارزیابی شواهد عمومی، تجربهٔ مشتری و ارتباطات سازمانی")
    p = document.add_paragraph(style="Subtitle")
    p.paragraph_format.space_after = Pt(35)
    add_inline(p, subtitle, base_size=17)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)

    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    add_inline(p, "بانک صادرات ایران", base_size=18)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        set_run_font(run, size=18, bold=True, color=NAVY)

    cutoff = metadata.get("cutoff", "۲۱ اکتبر ۲۰۲۶، UTC")
    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(22)
    add_inline(p, f"تاریخ قطع شواهد: {cutoff}", base_size=13)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        set_run_font(run, size=13, color=MUTED)

    t = document.add_table(rows=1, cols=1)
    set_table_rtl(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, PALE_BLUE)
    set_cell_margins(cell, top=180, start=220, bottom=180, end=220)
    p = cell.paragraphs[0]
    add_inline(p, "وضعیت معیارها", base_size=13)
    set_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
    for run in p.runs:
        set_run_font(run, size=13, bold=True, color=NAVY)
    p = cell.add_paragraph()
    add_inline(p, "دفتر شامل منابع اولیه، گزارش‌های رسانه‌ای، بازنشرها و مطالعات محدود است. شمار ردیف‌ها معادل منابع مستقل نیست و آستانهٔ عددی منابع مستقل ادعا نمی‌شود.", base_size=12.5)
    set_rtl(p)

    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(30)
    add_inline(p, "گزارش تحلیلی دامنه‌مند  •  نسخهٔ Word همراه با پیوست دفتر شواهد", base_size=11)
    set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)
    for run in p.runs:
        set_run_font(run, size=11, color=MUTED)
    document.add_page_break()


def add_contents(document, headings):
    p = document.add_paragraph(style="Heading 1")
    add_inline(p, "فهرست بخش‌ها", base_size=19)
    set_rtl(p)
    for heading in headings:
        p = document.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.05
        add_inline(p, heading, base_size=11.5)
        set_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    add_inline(p, "برای پیمایش سریع در Word از پنل Navigation نیز می‌توان استفاده کرد؛ عنوان‌های فصل‌ها دارای سبک Heading هستند.", base_size=10.5)
    set_rtl(p)
    document.add_page_break()


def parse_ledger_records():
    records = []
    for line in LEDGER.read_text(encoding="utf-8").splitlines():
        candidate = line.strip()
        if not candidate.startswith("|"):
            continue
        row = split_table_row(candidate)
        if row and re.fullmatch(r"پ\d+", row[0]):
            records.append(row)
    return records


def append_evidence_register(document):
    records = parse_ledger_records()
    document.add_page_break()
    heading = document.add_paragraph(style="Heading 1")
    add_inline(heading, "پیوست الف ـ دفتر شواهد تفصیلی", base_size=19)
    set_rtl(heading)
    intro = document.add_paragraph()
    add_inline(intro, f"این پیوست، {len(records)} ردیف دفتر را به‌شکل کارت‌های قابل‌خواندن عرضه می‌کند. ردیف‌ها معادل منابع مستقل نیستند؛ بازنشرها، نقل‌قول‌ها و مطالعات محدود باید با خوشه و حدود تعمیم خود خوانده شوند.")
    set_rtl(intro)

    for row in records:
        row += [""] * max(0, 6 - len(row))
        source_id, title, date, url, evidence, status = row[:6]
        paragraph = document.add_paragraph(style="Heading 2")
        paragraph.paragraph_format.space_before = Pt(12)
        add_inline(paragraph, f"{source_id} ـ {title}", base_size=15.5)
        set_rtl(paragraph)
        for label, value in (("تاریخ انتشار", date), ("نشانی", url),
                             ("یافته و محدودیت", evidence), ("جایگاه شمارشی", status)):
            p = document.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.12
            add_inline(p, f"{label}: {value}", base_size=10.8)
            set_rtl(p)
    print(f"Evidence-register appendix records: {len(records)}")


def build():
    if not SOURCE.is_file() or not LEDGER.is_file():
        raise FileNotFoundError("The Markdown report and evidence ledger must both be present.")
    lines = SOURCE.read_text(encoding="utf-8").splitlines()
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.78)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.82)
    section.right_margin = Inches(0.82)
    section.header_distance = Inches(0.34)
    section.footer_distance = Inches(0.34)
    configure_styles(document)
    add_header_footer(section)

    title_match = next((re.match(r"^#\s+(.*)$", line.strip()) for line in lines
                        if re.match(r"^#\s+", line.strip())), None)
    title = title_match.group(1) if title_match else "گزارش پژوهشی بانک صادرات ایران"
    metadata = {}
    for line in lines[:12]:
        stripped = line.strip()
        if stripped.startswith("**پروندهٔ"):
            metadata["subtitle"] = re.sub(r"\*", "", stripped).strip()
        elif stripped.startswith("**تاریخ قطع شواهد:"):
            metadata["cutoff"] = re.sub(r"\*", "", stripped).replace("تاریخ قطع شواهد:", "").strip()

    # Cover and contents are presentation elements, separate from the analytical body.
    add_cover(document, title, metadata)
    top_headings = []
    started = False
    for line in lines:
        txt = line.strip()
        if txt.startswith("## "):
            started = True
            if not txt.startswith("## ۲۸."):
                top_headings.append(txt[3:])
            else:
                break
    add_contents(document, top_headings)

    # Main body starts at the first analytical section. The long introductory
    # method note is retained as a shaded editorial callout immediately below.
    start_idx = next((i for i, line in enumerate(lines)
                      if line.strip().startswith("## ۱. خلاصهٔ اجرایی")), 0)
    method_note = next((re.sub(r"^>\s*", "", line.strip()) for line in lines
                        if line.strip().startswith("> **یادداشت روش:**")), None)
    if method_note:
        add_callout(document, method_note.replace("**", ""), "یادداشت روش")

    index = start_idx
    while index < len(lines):
        text = lines[index].strip()
        if text.startswith("## ۲۸. منابع مورد استفاده در این گزارش"):
            break
        if not text or text in ("---", "***") or text.startswith("> **یادداشت روش:**"):
            index += 1
            continue
        if text.startswith("|"):
            rows = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                row = split_table_row(lines[index])
                if not row or not all(re.fullmatch(r":?-{3,}:?", c.replace(" ", "")) for c in row):
                    rows.append(row)
                index += 1
            add_markdown_table(document, rows)
            continue
        if text.startswith("#"):
            match = re.match(r"^(#+)\s+(.*)$", text)
            level, heading = len(match.group(1)), match.group(2)
            style = f"Heading {min(level - 1, 3)}"
            p = document.add_paragraph(style=style)
            add_inline(p, heading)
            set_rtl(p)
        elif re.match(r"^[-*]\s+", text):
            p = document.add_paragraph(style="List Bullet")
            add_inline(p, re.sub(r"^[-*]\s+", "", text))
            set_rtl(p)
        elif re.match(r"^\d+[.)]\s+", text):
            p = document.add_paragraph(style="List Number")
            add_inline(p, re.sub(r"^\d+[.)]\s+", "", text))
            set_rtl(p)
        else:
            p = document.add_paragraph()
            p.paragraph_format.first_line_indent = Inches(0.16)
            add_inline(p, text)
            set_rtl(p)
        index += 1

    append_evidence_register(document)
    document.core_properties.title = title
    document.core_properties.subject = "ارزیابی شواهد عمومی و تجربهٔ مشتری بانک صادرات ایران؛ قطع شواهد ۲۱ اکتبر ۲۰۲۶"
    document.core_properties.author = "پروندهٔ پژوهشی اف۱۶۲"
    document.core_properties.keywords = "بانک صادرات ایران، روابط عمومی، ارتباطات بحران، تجربه مشتری، دفتر شواهد"
    settings = document.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")
    document.save(OUTPUT)

    built = Document(OUTPUT)
    word_count = sum(len(paragraph.text.split()) for paragraph in built.paragraphs)
    for table in built.tables:
        for row in table.rows:
            for cell in row.cells:
                word_count += len(cell.text.split())
    print(f"Saved: {OUTPUT.relative_to(HERE.parent)} ({OUTPUT.stat().st_size:,} bytes)")
    print(f"DOCX whitespace word count (includes cover, contents, appendix): {word_count:,}")
    print(f"Paragraphs: {len(built.paragraphs):,}; tables: {len(built.tables):,}")
    print(f"Requested font: {FONT}; install it on the opening computer for exact rendering.")


if __name__ == "__main__":
    build()
