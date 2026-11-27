---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins and justification (กั้นหน้า–กั้นหลัง, thaiDistribute / THAI_JUSTIFY,
  A4 3.0/2.0 cm margins, TH Sarabun New 16pt, w:bidi="th-TH", Thai word-break
  ZWSP, and visual QA via MS Word or LibreOffice). Use when creating, fixing,
  or polishing Thai Word documents whose right margin looks ragged or uneven.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) files with even left and right margins (กั้นหน้า–กั้นหลัง) using native Word Thai Distributed alignment (`w:jc="thaiDistribute"`), Complex Script language tagging (`w:bidi="th-TH"`), and punctuation-safe Zero-Width Spaces (`U+200B`).

## Setup

Install Python dependencies from [requirements.txt](./requirements.txt):

```bash
pip install -r requirements.txt
```

## Workflow

Follow these steps in order until all checks pass:

1. **Configure page geometry and styles** using `configure_thai_a4_margins()` in [thai_word_justify.py](./scripts/thai_word_justify.py):
   - Page size: A4 (`21.0 × 29.7 cm`).
   - Official Thai margins: Top `2.5 cm`, Bottom `2.0 cm`, Left (กั้นหน้า) `3.0 cm`, Right (กั้นหลัง) `2.0 cm` → printable width `16.0 cm`.
   - Font: `TH Sarabun New` (`16pt` body, `18–22pt` headings, `13.5–14pt` tables, `12pt` footer), setting `w:ascii`, `w:hAnsi`, `w:cs`, `w:szCs`, and `<w:lang w:val="en-US" w:bidi="th-TH"/>` on every run via `apply_thai_run_style()`.
2. **Apply alignment and Thai word breaks**:
   - Multi-line Thai body prose, callout boxes, and multi-line bullets: use `set_thai_justify()` (`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` → `thaiDistribute`) with `insert_thai_zwsp_clean()`.
   - Headings, short metadata lines, and narrow table cells: keep `WD_ALIGN_PARAGRAPH.LEFT` so narrow columns stay compact without stretching characters.
   - Tables: set fixed layout (`w:tblLayout="fixed"`) summing to `16.0 cm`, repeat header rows (`w:tblHeader`), and prevent mid-row splits (`w:cantSplit`). See [ooxml-thai-guide.md](./references/ooxml-thai-guide.md) for XML details.
3. **Run programmatic audit** (must exit `0` before visual inspection):
   ```bash
   python scripts/thai_word_justify.py check output.docx
   ```
4. **Run visual QA** via Microsoft Word (Windows) or LibreOffice fallback with [render_docx_to_png.py](./scripts/render_docx_to_png.py):
   ```bash
   python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150
   ```
   Inspect every generated `page_*.png` to verify flush left/right margins, clean Thai wrapping, exact table widths, and page numbers (`หน้า X / Y`).

## CLI Usage

Fix an existing `.docx` in one pass, then verify it:

```bash
python scripts/thai_word_justify.py fix input.docx -o output.docx --footer "ชื่อเอกสาร"
python scripts/thai_word_justify.py check output.docx
```

For building a new document from scratch in Python, follow [quick_example.py](./examples/quick_example.py).
