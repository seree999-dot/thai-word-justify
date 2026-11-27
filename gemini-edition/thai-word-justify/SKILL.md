---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins, heading hierarchy, and justification (กั้นหน้า–กั้นหลัง,
  thaiDistribute / THAI_JUSTIFY, A4 3.0/2.0 cm margins, TH Sarabun New 16pt,
  Title 18–20pt centered, H1 16–18pt, H2/H3 16pt indented, w:bidi="th-TH",
  and Thai word-break ZWSP). Use when creating, fixing, or polishing Thai Word
  documents whose right margin looks ragged or uneven.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) documents so left and right margins (กั้นหน้า–กั้นหลัง) and Thai heading hierarchy align cleanly without stretched spaces or broken syllables. This skill contains pure instructions and OOXML specifications with no external scripts required.

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
| `official` (หนังสือราชการ) | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, indent `1.25 cm` | Header right (Thai numerals `๐–๙` optional) |
| `academic` (รายงาน/วิทยานิพนธ์) | `3.81 / 2.54 / 2.54 / 2.54` | `14.65 cm` | `TH Sarabun New 16pt`, indent `1.25 cm` | Header right |

## Core Specification

1. **Page Geometry (A4 Thai Standards)**:
   - Paper size: `21.0 × 29.7 cm` (A4).
   - Default `proposal` / `official` margins: Top `2.5 cm`, Bottom `2.0 cm`, Left / กั้นหน้า `3.0 cm`, Right / กั้นหลัง `2.0 cm` → printable width **`16.0 cm`** (`9072 dxa`).
   - `academic` margins: Left `3.81 cm`, Right/Top/Bottom `2.54 cm` → printable width **`14.65 cm`** (`8307 dxa`).
   - Header / Footer distance: `1.25 cm`.

2. **Typography & Complex Script Language Tagging**:
   - Font family: `TH Sarabun New` (or `TH SarabunPSK`).
   - Follow the **Thai Typography Hierarchy** table above (`Title 18–20pt`, `H1 16–18pt`, `H2 16pt`, `H3 16pt`, `Body 16pt`).
   - Set all four font slots (`w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`) plus Complex Script sizing (`w:szCs` in half-points, e.g., `32` for `16pt`, `36` for `18pt`, `40` for `20pt`), Complex Script bold/italic (`w:bCs`, `w:iCs`), and **`<w:lang w:val="en-US" w:bidi="th-TH"/>`** on every run and on the `Normal` style.

3. **Alignment Strategy (กั้นหน้า–กั้นหลัง)**:
   - **Title / Chapter headings**: `WD_ALIGN_PARAGRAPH.CENTER`, Bold, `18–20 pt`, `keep_with_next = True`.
   - **H1 / H2 / H3 headings**: `WD_ALIGN_PARAGRAPH.LEFT`, `keep_with_next = True`, indented progressively (`0 cm` → `0.75 cm` → `1.5 cm`).
   - **Body prose, callout boxes, and multi-line bullets**: Use Thai Distributed (`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` / `<w:jc w:val="thaiDistribute"/>`) with `line_spacing = 1.15`. Add a `1.25 cm` first-line indent on body prose paragraphs and hanging indent (`left_indent = 1.0 cm`, `first_line_indent = -0.4 cm`) on bullet items.
   - **Short metadata lines and narrow table cells**: Use `WD_ALIGN_PARAGRAPH.LEFT` (`<w:jc w:val="left"/>`) so narrow columns stay readable without stretching characters.
   - Avoid standard `JUSTIFY` (`w:val="both"`, which creates huge gaps) and `DISTRIBUTE` (`w:val="distribute"`, which stretches the last line).

4. **Thai Word-Break Cues (`U+200B` / ZWSP)**:
   - When generating or editing Thai text in `.docx`, insert Zero-Width Space (`\u200b`) cleanly between Thai word boundaries so Word wraps long clauses cleanly toward the right margin.
   - If `pythainlp` is already available in the active Python environment, tokenize Thai character runs (`[\u0E00-\u0E7F]+`) using `newmm`, strip any existing `\u200b` first so processing is idempotent, and remove any `\u200b` right before `ๆ` or `ฯ`. Leave English words, numbers, and URLs untouched.
   - If `pythainlp` is not available, rely on native MS Word Thai justification via `<w:lang w:val="en-US" w:bidi="th-TH"/>` + `<w:jc w:val="thaiDistribute"/>`.

5. **Tables & Dynamic Page Numbers**:
   - Set `<w:tblLayout w:type="fixed"/>` and total `<w:tblW>` equal to the preset printable width (`16.0 cm` / `9072 dxa`, or `14.65 cm` / `8307 dxa`) so tables align flush with both margins without overflowing.
   - Add `<w:tblHeader/>` on row 0 to repeat table headers across page breaks, and `<w:cantSplit/>` on every row.
   - Add a centered footer (`proposal`) or right-aligned header (`official` / `academic`) with `PAGE` / `NUMPAGES` Word fields.

For complete OpenXML tag snippets, troubleshooting patterns, and a copy-ready inline `python-docx` helper, read [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).

## Execution Modes

- **Mode A — Python / `python-docx` available in session**: Build or update the `.docx` inline using the copy-ready helper and OOXML specs in [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).
- **Mode B — Pure text/chat mode without code execution**: Draft the Thai content structured cleanly for Word import (marking Title `20pt` centered, H1 `18pt`, H2 `16pt` indented, H3 `16pt` deep-indented, and Body `16pt` indented) and provide the user with exact Word setup values: A4, preset margins (e.g., Left `3.0 cm` / Right `2.0 cm` / Top `2.5 cm` / Bottom `2.0 cm`), font `TH Sarabun New`, paragraph alignment **Thai Distributed** (กระจายแบบไทย), and Proofing Language **Thai**.

## Self-Audit Checklist

Verify all 6 items before delivering any `.docx` or formatting output:

1. Paper is A4 (`21.0 × 29.7 cm`) with Left `3.0 cm` (or `3.81 cm` academic) and Right `2.0 cm` (or `2.54 cm`).
2. `TH Sarabun New` is set on `w:cs` with matching `w:szCs` and `<w:lang w:val="en-US" w:bidi="th-TH"/>`.
3. Heading hierarchy matches: Title `18–20pt` bold centered, H1 `16–18pt` bold left, H2 `16pt` bold left/1-level indent, H3 `16pt` bold/italic deep-indented, Body `16pt` regular `THAI_JUSTIFY`.
4. No Thai text contains duplicate `\u200b\u200b` or `\u200b` immediately before `ๆ` / `ฯ`.
5. Every table uses fixed layout summing to the printable width (`<= 16.0 cm` or `<= 14.65 cm`) with `w:tblHeader` and `w:cantSplit`.
6. Page numbers match the selected preset (`หน้า X / Y` footer or right-aligned header).
