---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins, แบบที่ 1 hierarchical numbering, and justification (กั้นหน้า–กั้นหลัง,
  thaiDistribute / THAI_JUSTIFY, A4 3.0/2.0 cm margins, TH Sarabun New 16pt,
  1. / 1.1 / 1.1.1 / 1.1.2.1 / 1) cascading indents, w:bidi="th-TH", Thai
  word-break ZWSP, and visual QA via MS Word or LibreOffice). Use when creating,
  fixing, or polishing Thai Word documents whose right margin or numbering looks uneven.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) files with even left and right margins (กั้นหน้า–กั้นหลัง), standard Thai **แบบที่ 1** heading numbering & cascading indentation, native Word Thai Distributed alignment (`w:jc="thaiDistribute"`), Complex Script language tagging (`w:bidi="th-TH"`), and punctuation-safe Zero-Width Spaces (`U+200B`).

## Setup

Install Python dependencies from [requirements.txt](./requirements.txt):

```bash
pip install -r requirements.txt
```

## ระบบเลขหัวข้อและการเยื้องมาตรฐานไทย (แบบที่ 1)

| ระดับ (Level) | รูปแบบตัวเลข (Numbering) | ขนาด / น้ำหนัก | ระยะเยื้อง (`left_indent`) | กฎสำคัญ |
| :--- | :--- | :--- | :--- | :--- |
| `title` (ชื่อเรื่อง / ชื่อบท) | เช่น `แบบที่ 1` / `บทที่ 1` | `18–20 pt` *(ค่าเริ่มต้น 20 pt)* **Bold** | กึ่งกลาง (`CENTER`) | ไม่มีเลขข้อปิดท้าย |
| `h1` (หัวข้อใหญ่) | `1.  หัวข้อใหญ่`, `2.  หัวข้อใหญ่` | `16–18 pt` *(ค่าเริ่มต้น 18 pt)* **Bold** | `0.0 cm` (ชิดขอบซ้าย) | **มีจุดหลังตัวเลข** + เว้น **2 เคาะ** |
| ย่อหน้าใต้ `1.` | ย่อหน้าแรก `1.0 cm` (5 เคาะ) | `16 pt` Regular | `left=0 cm`, `first_line=1.0 cm` | บรรทัดแรกตรงแนว `1.1`, บรรทัดถัดไปชิดซ้าย (`THAI_JUSTIFY`) |
| `h2` (หัวข้อย่อยระดับที่ 1) | `1.1  หัวข้อย่อยระดับที่ 1`, `1.2` | `16 pt` **Bold** | `1.0 cm` (5 เคาะ) | **ไม่มีจุดปิดท้าย** + เว้น **2 เคาะ** |
| ย่อหน้าใต้ `1.1`+ | ย่อหน้าแรก `2.0 cm` (10 เคาะ) | `16 pt` Regular | `left=0 cm`, `first_line=2.0 cm` | บรรทัดแรกเยื้องลึกจาก `1.1`, บรรทัดถัดไปชิดซ้าย (`THAI_JUSTIFY`) |
| `h3` (หัวข้อย่อยระดับที่ 2) | `1.1.1  หัวข้อย่อยระดับที่ 2` | `16 pt` Regular *(หรือ Bold/Italic)* | `1.8 cm` (ตรงอักษรแรกของ `1.1`) | **ไม่มีจุดปิดท้าย** + เว้น **2 เคาะ** |
| `h4` (หัวข้อย่อยระดับที่ 3) | `1.1.2.1  หัวข้อย่อยระดับที่ 3` | `16 pt` Regular | `2.8 cm` (ตรงอักษรแรกของ `1.1.1`) | **ไม่มีจุดปิดท้าย** + เว้น **2 เคาะ** |
| `h5` (หัวข้อย่อยระดับที่ 4) | `1)  หัวข้อย่อยระดับที่ 4`, `2)` | `16 pt` Regular | `4.0 cm` (ตรงอักษรแรกของ `1.1.2.1`) | **ใช้วงเล็บปิด `)` ไม่มีจุด** + เว้น **2 เคาะ** |

```text
                       แบบที่ 1 (CENTER, Bold)
1.  หัวข้อใหญ่ (LEFT 0.0 cm, Bold)
     ย่อหน้า 5 เคาะ (first_line 1.0 cm) ....................................
     1.1  หัวข้อย่อยระดับที่ 1 (1.0 cm, Bold)
          ย่อหน้า 10 เคาะ (first_line 2.0 cm) ..............................
          1.1.1  หัวข้อย่อยระดับที่ 2 (1.8 cm)
          1.1.2  หัวข้อย่อยระดับที่ 2 (1.8 cm)
               1.1.2.1  หัวข้อย่อยระดับที่ 3 (2.8 cm)
                    1)  หัวข้อย่อยระดับที่ 4 (4.0 cm)
                    2)  หัวข้อย่อยระดับที่ 4 (4.0 cm)
               1.1.2.2  หัวข้อย่อยระดับที่ 3 (2.8 cm)
          1.1.3  หัวข้อย่อยระดับที่ 2 (1.8 cm)
     1.2  หัวข้อย่อยระดับที่ 1 (1.0 cm, Bold)
2.  หัวข้อใหญ่ (LEFT 0.0 cm, Bold)
```

## Thai Layout Presets

| Preset | Margins (L / R / T / B cm) | Printable Width | Body Font & Context Indent | Page Numbers |
| :--- | :--- | :--- | :--- | :--- |
| `proposal` *(default)* | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, `1.0 / 2.0 cm` | Footer center (`หน้า X / Y`) |
| `official` (หนังสือราชการ) | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, `1.0 / 2.0 cm` | Header right (supports `--thai-numerals`) |
| `academic` (รายงาน/วิทยานิพนธ์) | `3.81 / 2.54 / 2.54 / 2.54` | `14.65 cm` | `TH Sarabun New 16pt`, `1.0 / 2.0 cm` | Header right |

## Workflow

1. **Configure page geometry and styles** using `configure_thai_a4_margins()` in [thai_word_justify.py](./scripts/thai_word_justify.py) (`21.0 × 29.7 cm`, preset margins, `TH Sarabun New`, `w:cs`, `w:szCs`, `<w:lang w:val="en-US" w:bidi="th-TH"/>`).
2. **Apply แบบที่ 1 heading hierarchy, alignment, and Thai word breaks**:
   - Headings: use `add_thai_heading(doc, text, level="title"|"h1"|"h2"|"h3"|"h4"|"h5")` (normalizes numbering via `normalize_thai_heading_number()`, strips trailing dots on `1.1` / `1.1.1` / `1.1.2.1`, enforces 2 spaces and cascading indents).
   - Body prose: use `add_thai_body(doc, text, under_level="h1"|"h2")` for `THAI_JUSTIFY` + `insert_thai_zwsp_clean()` + `1.0 cm` / `2.0 cm` first-line indent.
   - Tables: keep cell paragraphs `LEFT`, fixed layout (`w:tblLayout="fixed"`) matching printable width (`16.0 cm` or `14.65 cm`), `w:tblHeader`, and `w:cantSplit`. See [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).
3. **Run programmatic audit** (must exit `0`):
   ```bash
   python scripts/thai_word_justify.py check output.docx
   ```
4. **Run visual QA** via MS Word / LibreOffice with [render_docx_to_png.py](./scripts/render_docx_to_png.py) and inspect every `page_*.png`:
   ```bash
   python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150
   ```

## CLI Usage

```bash
python scripts/thai_word_justify.py fix input.docx -o output.docx --preset proposal --footer "ชื่อเอกสาร"
python scripts/thai_word_justify.py fix input.docx -o official.docx --preset official --thai-numerals
python scripts/thai_word_justify.py check output.docx
```

See [quick_example.py](./examples/quick_example.py) for a complete runnable **แบบที่ 1** document example.
