# -*- coding: utf-8 -*-
"""Post-build structural checks for the F-067 DOCX. Reports blockers; does not certify facts."""
from collections import Counter
from pathlib import Path
import re
from docx import Document

DOCX = Path(__file__).resolve().parent.parent / "F-067_وسبحان_سرمایه_گذاری_سبحان_گزارش_تحقیقات_بازار_آرون.docx"

def main():
    doc = Document(DOCX)
    words = sum(len(p.text.split()) for p in doc.paragraphs)
    words += sum(len(cell.text.split()) for table in doc.tables for row in table.rows for cell in row.cells)
    paragraphs = [re.sub(r"\s+", " ", p.text.strip()) for p in doc.paragraphs if p.text.strip()]
    paragraph_dups = sum(n - 1 for n in Counter(paragraphs).values() if n > 1)
    text = "\n".join(paragraphs + [c.text for t in doc.tables for r in t.rows for c in r.cells])
    latin = sorted(set(re.findall(r"[A-Za-z]+", text)))
    cjk = re.findall(r"[\u4e00-\u9fff]", text)
    caps = [p.text for p in doc.paragraphs if re.match(r"جدول\s+\d+\s+—", p.text)]
    cap_numbers = [int(m.group(1)) for s in caps if (m := re.match(r"جدول\s+(\d+)", s))]
    body = list(doc._element.body)
    adjacency_issues = 0
    for i, element in enumerate(body):
        if element.tag.endswith("}tbl"):
            nxt = body[i + 1] if i + 1 < len(body) else None
            nxt_text = "".join(nxt.itertext()) if nxt is not None else ""
            if nxt is None or not nxt.tag.endswith("}p") or not nxt_text.startswith("جدول "):
                adjacency_issues += 1
    signatures = [tuple(tuple(cell.text for cell in row.cells) for row in t.rows) for t in doc.tables]
    unique_tables = len(set(signatures))
    h1 = sum(1 for p in doc.paragraphs if p.runs and p.runs[0].bold and p.runs[0].font.size and p.runs[0].font.size.pt >= 15)
    h2 = sum(1 for p in doc.paragraphs if p.runs and p.runs[0].bold and p.runs[0].font.size and 11.5 <= p.runs[0].font.size.pt < 15)
    print("DOCX", DOCX)
    print("WORDS", words, "GAP_TO_30000", max(0, 30000 - words))
    print("PARAGRAPHS", len(doc.paragraphs), "EXACT_DUPLICATE_PARAGRAPHS", paragraph_dups)
    print("TABLES", len(doc.tables), "UNIQUE_TABLE_CONTENTS", unique_tables)
    print("CAPTIONS", len(caps), "SEQUENTIAL", cap_numbers == list(range(1, len(cap_numbers) + 1)), "CAPTION_ADJACENCY_ISSUES", adjacency_issues)
    print("H1", h1, "H2", h2)
    print("LATIN", latin, "CJK", len(cjk))
    print("GUILLEMET", text.count("«"), text.count("»"), "PARENTHESES", text.count("("), text.count(")"))
    appendix_count = len(re.findall(r"(?m)^پیوست\s+[۰-۹0-9]+", "\n".join(paragraphs)))
    print("APPENDIX_HEADINGS", appendix_count, "REQUIRED_10", appendix_count == 10)

if __name__ == "__main__":
    main()
