---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins, heading hierarchy, and justification (กั้นหน้า–กั้นหลัง,
  thaiDistribute / THAI_JUSTIFY, A4 3.0/2.0 cm margins, TH Sarabun New 16pt,
  Title 18–20pt centered, H1 16–18pt, H2/H3 16pt indented, w:bidi="th-TH",
  Thai word-break ZWSP, and visual QA via MS Word or LibreOffice). Use when
  creating, fixing, or polishing Thai Word documents whose right margin looks
  ragged or uneven.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) files with even left and right margins (กั้นหน้า–กั้นหลัง), standard Thai heading hierarchy, native Word Thai Distributed alignment (`w:jc="thaiDistribute"`), Complex Script language tagging (`w:bidi="th-TH"`), and punctuation-safe Zero-Width Spaces (`U+200B`).

## Setup

Install Python dependencies from [requirements.txt](./requirements.txt):

```bash
pip install -r requirements.txt
```

## Thai Typography Hierarchy (`TH Sarabun New`)

| ระดับ (Level) | ขนาด (Size) | น้ำหนัก (Style) | การจัดวางและระยะเยื้อง (Alignment & Indent) |
| :--- | :--- | :--- | :--- |
| **ชื่อเรื่อง / ชื่อบท** (หัวข้อใหญ่กลางหน้ากระดาษ) | **18–20 pt** *(ค่าเริ่มต้น 20 pt)* | ตัวหนา (**Bold**) | จัดกึ่งกลางหน้ากระดาษ (`CENTER`) |
| **หัวข้อหลัก ระดับที่ 1** (เช่น `1. ความเป็นมา`) | **16–18 pt** *(ค่าเริ่มต้น 18 pt)* | ตัวหนา (**Bold**) | พิมพ์ชิดขอบซ้าย (`LEFT`, indent `0 cm`) |
| **หัวข้อรอง ระดับที่ 2** (เช่น `1.1 วัตถุประสงค์`) | **16 pt** | ตัวหนา (**Bold**) | พิมพ์ชิดซ้ายหรือเยื้องเข้า 1 ระดับ (`LEFT`, default indent `0.75 cm`) |
| **หัวข้อย่อย ระดับที่ 3** (เช่น `1.1.1 ด้านงบประมาณ`) | **16 pt** | ตัวหนา หรือ ตัวเอน (**Bold / Italic**) | เยื้องลึกเข้าไปจากหัวข้อรอง (`LEFT`, default indent `1.5 cm`) |
| **เนื้อหาปกติ** (Body prose) | **16 pt** | ตัวธรรมดา (**Regular**) | กั้นหน้า–กั้นหลัง (`THAI_JUSTIFY` / `thaiDistribute`), ย่อหน้าแรก `1.25 cm` |

## Thai Layout Presets

| Preset | Margins (L / R / T / B cm) | Printable Width | Body Font & Indent | Page Numbers |
| :--- | :--- | :--- | :--- | :--- |
| `proposal` *(default)* | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, indent `1.25 cm` | Footer center (`หน้า X / Y`) |
| `official` (หนังสือราชการ) | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, indent `1.25 cm` | Header right (supports `--thai-numerals`) |
| `academic` (รายงาน/วิทยานิพนธ์) | `3.81 / 2.54 / 2.54 / 2.54` | `14.65 cm` | `TH Sarabun New 16pt`, indent `1.25 cm` | Header right |

## Workflow

Follow these steps in order until all checks pass:

1. **Configure page geometry and styles** using `configure_thai_a4_margins()` in [thai_word_justify.py](./scripts/thai_word_justify.py):
   - Page size: A4 (`21.0 × 29.7 cm`) with the target preset margins (`proposal`, `official`, or `academic`).
   - Font: `TH Sarabun New` following the **Thai Typography Hierarchy** table above, setting `w:ascii`, `w:hAnsi`, `w:cs`, `w:szCs`, and `<w:lang w:val="en-US" w:bidi="th-TH"/>` on every run via `apply_thai_run_style()`.
2. **Apply heading hierarchy, alignment, and Thai word breaks**:
   - Headings: use `add_thai_heading(doc, text, level="title"|"h1"|"h2"|"h3")` with `keep_with_next = True`.
   - Multi-line Thai body prose, callout boxes, and multi-line bullets: use `set_thai_justify()` (`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` → `thaiDistribute`) with `insert_thai_zwsp_clean()`.
   - Narrow table cells: keep `WD_ALIGN_PARAGRAPH.LEFT` so narrow columns stay compact without stretching characters.
   - Tables: set fixed layout (`w:tblLayout="fixed"`) matching the printable width (`16.0 cm` or `14.65 cm`), repeat header rows (`w:tblHeader`), and prevent mid-row splits (`w:cantSplit`). See [ooxml-thai-guide.md](./references/ooxml-thai-guide.md) for XML details and the inline `python-docx` recipe.
3. **Run programmatic audit** (verifies A4 geometry, margins, `THAI_JUSTIFY`, `w:bidi="th-TH"`, ZWS, and table width overflow — must exit `0`):
   ```bash
   python scripts/thai_word_justify.py check output.docx
   ```
4. **Run visual QA** via Microsoft Word (Windows) or LibreOffice fallback with [render_docx_to_png.py](./scripts/render_docx_to_png.py):
   ```bash
   python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150
   ```
   Inspect every generated `page_*.png` to verify flush left/right margins, heading hierarchy, clean Thai wrapping, exact table widths, and page numbers.

## CLI Usage

Fix an existing `.docx` using a preset and optional Thai numerals (`๐–๙`), then verify it:

```bash
python scripts/thai_word_justify.py fix input.docx -o output.docx --preset proposal --footer "ชื่อเอกสาร"
python scripts/thai_word_justify.py fix input.docx -o official.docx --preset official --thai-numerals
python scripts/thai_word_justify.py check output.docx
```

For building a new document from scratch in Python, follow [quick_example.py](./examples/quick_example.py).
