#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Structural QA for B-120; unmet completion gates remain explicit."""
from pathlib import Path
import re
from zipfile import ZipFile
from urllib.parse import urlsplit
from docx import Document

ROOT = Path(__file__).resolve().parent
REPORT = ROOT / "B120_REPORT_FINAL.md"
LEDGER = ROOT / "B120_EVIDENCE_LEDGER.md"
DOCX = ROOT / "B120_REPORT_FINAL.docx"
report = REPORT.read_text(encoding="utf-8")
ledger = LEDGER.read_text(encoding="utf-8")
words = len(report.split())
prose_lines = [line for line in report.splitlines() if line.strip() and not line.lstrip().startswith(("#", "|", ">", "---", "```"))]
prose_words = sum(len(line.split()) for line in prose_lines)
areas = [line for line in report.splitlines() if re.match(r"^# [۰-۹]{1,2}\. ", line)]
ledger_ids = re.findall(r"^\| (پ\d+) \|", ledger, re.M)
report_refs = set(re.findall(r"پ\d+", report))
unknown = sorted(report_refs - set(ledger_ids), key=lambda x: int(x[1:]))
ledger_urls = []
for row in ledger.splitlines():
    if re.match(r"^\| پ\d+ \|", row):
        ledger_urls.extend(re.findall(r"https?://[^\s|,]+", row))
unique_raw_urls = len(set(ledger_urls))
def canonical_url(url):
    parts = urlsplit(url)
    path = parts.path.rstrip("/") or "/"
    return (parts.scheme.lower(), parts.netloc.lower(), path, parts.query)
canonical_unique_urls = len({canonical_url(url) for url in ledger_urls})
source_dupes = len(ledger_ids) - len(set(ledger_ids))
nonempty = [x.strip() for x in report.splitlines() if x.strip() and not re.fullmatch(r"\|?\s*:?-{3,}:?(\s*\|\s*:?-{3,}:?)*\s*\|?", x.strip())]
exact_dupes = len(nonempty) - len(set(nonempty))

def latin_tokens(text):
    text = re.sub(r"https?://[^\s|]+", " ", text)
    text = re.sub(r"`[^`]*`", " ", text)  # preserve literal handles/identifiers, not prose
    return sorted(set(re.findall(r"(?<![A-Za-z])[A-Za-z]{2,}(?![A-Za-z])", text)))

allow_report = {"UTC", "B", "GRI", "CSR", "ATM", "POS", "CRM", "f", "EVIDENCE", "LEDGER", "REPORT", "FINAL", "md", "QA", "PDF", "Salehnia", "Eshaghi", "Saki", "Android", "URL", "WCAG", "iOS", "paywall", "schema", "API", "Foundation", "NIST", "SP", "Board", "Global", "Initiative", "Reporting", "Standards", "Sustainability", "ISO", "OECD", "Basel", "Committee", "Supervision", "BIS", "Banking", "on", "IntelliNews", "International", "Iran", "ZarinPal", "bne"}
allow_ledger = allow_report | {"HTTP", "AMOS", "SmartPLS", "APA", "MLA", "VANCOUVER", "COI", "ISI", "DOI", "DOR", "PDF", "Salehnia", "Eshaghi", "Saki"}
latin_report = sorted(set(latin_tokens(report)) - allow_report)
latin_ledger = sorted(set(latin_tokens(ledger)) - allow_ledger)

print(f"REPORT_WORDS={words}")
print(f"ANALYTICAL_PROSE_WORDS={prose_words}")
print(f"AUDIT_AREAS={len(areas)}")
print(f"RAW_LEDGER_ROWS={len(ledger_ids)}")
print(f"UNIQUE_LEDGER_IDS={len(set(ledger_ids))}")
print(f"REPEATED_LEDGER_IDS={source_dupes}")
print(f"URL_OCCURRENCES={len(ledger_urls)}")
print(f"UNIQUE_RAW_URLS={unique_raw_urls}")
print(f"CANONICAL_UNIQUE_URLS={canonical_unique_urls}")
print("URL_COUNTS_ARE_NOT_INDEPENDENCE_COUNTS=True")
print(f"SOURCE_AUDIT_FILE_PRESENT={(ROOT / 'B120_SOURCE_AUDIT.md').is_file()}")
print(f"UNKNOWN_CITATIONS={unknown}")
print(f"EXACT_DUPLICATE_NONEMPTY_LINES={exact_dupes}")
print(f"UNRECOGNIZED_LATIN_TOKENS_REPORT={latin_report}")
print(f"UNRECOGNIZED_LATIN_TOKENS_LEDGER={latin_ledger}")

if not DOCX.is_file():
    print("DOCX_MISSING=True")
    docx_ok = False
else:
    with ZipFile(DOCX) as zf:
        bad = zf.testzip()
        assert bad is None, f"Corrupt ZIP member: {bad}"
    doc = Document(DOCX)
    text = "\n".join(p.text for p in doc.paragraphs)
    all_text = text + "\n" + "\n".join(cell.text for table in doc.tables for row in table.rows for cell in row.cells)
    tables = len(doc.tables)
    docx_title = "پروندهٔ ممیزی ارتباطات عمومی بانک ملت" in text
    print(f"DOCX_PARAGRAPHS={len(doc.paragraphs)}")
    print(f"DOCX_TABLES={tables}")
    print(f"DOCX_HAS_REPORT_TITLE={docx_title}")
    docx_ok = docx_title

print("GATE_WORDS_30000=" + ("PASS" if prose_words >= 30000 else "NOT_MET"))
print("GATE_RAW_WHITESPACE_COUNT_30000=" + ("PASS" if words >= 30000 else "NOT_MET"))
print("GATE_12_AREAS=" + ("PASS" if len(areas) == 12 else "FAIL"))
print("GATE_CITATIONS=" + ("PASS" if not unknown else "FAIL"))
print("GATE_INDEPENDENT_SOURCES=NOT_ASSESSED; RAW_ROWS_ARE_NOT_INDEPENDENT_SOURCES")
print("GATE_DOCX=" + ("PASS" if docx_ok else "FAIL"))
if len(areas) != 12 or unknown or not docx_ok:
    raise SystemExit(1)
