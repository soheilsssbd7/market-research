# -*- coding: utf-8 -*-
"""سازندهٔ سند F-142 — صنعتی مینو خرمدره (غصینو) — رکوردِ بی‌صدا"""
import sys, os, re, importlib.util

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FA = str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')

_G = re.compile('[«»]')

def _clean(s):
    s = _G.sub('', s)
    s = re.sub(' +', ' ', s)
    return s

def fa(s):
    s = _clean(s)
    return s.translate(FA).replace('%', '٪')

TMAP = {}

_TRE = re.compile(r'(?<![A-Za-z0-9])T(\d{1,2})(?!\d)')

def ren(s):
    return _TRE.sub(lambda m: TMAP.get('T' + m.group(1), m.group(0)), s)

BASE = os.path.dirname(os.path.abspath(__file__))

def load_parts():
    out = []
    for i in range(1, 41):
        p = os.path.join(BASE, 'p%02d.py' % i)
        if not os.path.exists(p):
            continue
        spec = importlib.util.spec_from_file_location('p%02d' % i, p)
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        out.extend(m.P)
        out.append('@PAGEBREAK|')
    return out

from tables import TABLES as TABLES

TMAP = {k: str(i + 1) for i, k in enumerate(sorted(TABLES, key=lambda x: int(x[1:])))}

_BIDI_SUCC = ('w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing',
              'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
              'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')

def _set_rtl(p):
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is not None:
        return
    bidi = OxmlElement('w:bidi')
    pPr.insert_element_before(bidi, *_BIDI_SUCC)

def _auto_widths(head, rows):
    n = len(head)
    lens = []
    for i in range(n):
        vals = [len(str(r[i])) for r in rows] or [1]
        lens.append(max(8, min(60, sum(vals) / len(vals))))
    tot = sum(lens)
    return [max(1.4, round(16.0 * x / tot, 2)) for x in lens]


TABLE_NO = [0]

def add_table(doc, key):
    TABLE_NO[0] += 1
    tup = TABLES[key]
    head, rows = tup[0], tup[1]
    widths = tup[2] if len(tup) > 2 else _auto_widths(head, rows)
    t = doc.add_table(rows=1, cols=len(head))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    bv = OxmlElement('w:bidiVisual')
    tblPr.insert_element_before(bv, 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc',
                                'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd', 'w:tblLayout',
                                'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription')
    hdr = t.rows[0].cells
    for i, h in enumerate(head):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        _set_rtl(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(fa(h))
        r.bold = True
        r.font.size = Pt(9)
        r.font.name = 'Tahoma'
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            _set_rtl(p)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(fa(v))
            r.font.size = Pt(8.5)
            r.font.name = 'Tahoma'
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph()


def main():
    parts = load_parts()
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Tahoma'
    st.font.size = Pt(11)
    st._element.rPr.rFonts.set(qn('w:cs'), 'Tahoma')
    sec = doc.sections[0]
    sec.page_height = Cm(29.7)
    sec.page_width = Cm(21.0)
    sec.left_margin = Cm(2.4)
    sec.right_margin = Cm(2.4)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)

    for raw in parts:
        if '|' in raw:
            kind, _, txt = raw.partition('|')
        else:
            kind, txt = '@P', raw
        if kind == '@PAGEBREAK':
            doc.add_page_break()
            continue
        if kind == '@T':
            add_table(doc, txt)
            continue
        txt = ren(txt)
        p = doc.add_paragraph()
        _set_rtl(p)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.35
        if kind == '@TITLE':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(13)
        elif kind == '@BIG':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(26)
        elif kind == '@SUB':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(14)
        elif kind == '@META':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(fa(txt)); r.font.size = Pt(10.5)
        elif kind == '@STATUS':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0xA6, 0x1B, 0x1B)
        elif kind == '@NOTE':
            r = p.add_run(fa(txt)); r.italic = True; r.font.size = Pt(10)
        elif kind == '@U':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(txt); r.font.size = Pt(9)
        elif kind == '@H1':
            p.paragraph_format.space_before = Pt(16)
            p.paragraph_format.space_after = Pt(6)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(15)
            r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
        elif kind == '@H2':
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(fa(txt)); r.bold = True; r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0x2E, 0x4E, 0x7E)
        elif kind == '@QC':
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(fa(txt.replace('{N}', str(TABLE_NO[0])))); r.italic = True; r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
        else:
            segs = txt.split('**')
            for i, seg in enumerate(segs):
                if not seg:
                    continue
                r = p.add_run(fa(seg))
                r.font.size = Pt(11)
                r.bold = (i % 2 == 1)

    out = os.path.join(os.path.dirname(BASE), 'F-142_صنعتی_مینو_خرمدره_غصینو_گزارش_تحقیقات_بازار_آرون.docx')
    doc.save(out)

    d2 = Document(out)
    words = 0
    for p in d2.paragraphs:
        words += len(p.text.split())
    for t in d2.tables:
        for row in t.rows:
            for c in row.cells:
                words += len(c.text.split())
    print('SAVED', out)
    print('WORDS', words)
    print('PARAS', len(d2.paragraphs), 'TABLES', len(d2.tables))
    h1 = sum(1 for p in d2.paragraphs if p.runs and p.runs[0].bold and p.runs[0].font.size and p.runs[0].font.size.pt >= 15)
    h2 = sum(1 for p in d2.paragraphs if p.runs and p.runs[0].bold and p.runs[0].font.size and 11.5 <= p.runs[0].font.size.pt < 15)
    print('H1', h1, 'H2', h2)
    txt = '\n'.join(p.text for p in d2.paragraphs)
    print('GUL', txt.count('«') + txt.count('»'))
    import re as _re
    lat = set(_re.findall(r'[A-Za-z]+', txt))
    print('LATIN', sorted(lat)[:20])
    print('DBLSPACE', txt.count('  '))


if __name__ == '__main__':
    main()
