---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins and justification (กั้นหน้า–กั้นหลัง, thaiDistribute / THAI_JUSTIFY,
  A4 3.0/2.0 cm margins, TH Sarabun New 16pt, w:bidi="th-TH", and Thai word-break
  ZWSP). Use when creating, fixing, or polishing Thai Word documents whose right
  margin looks ragged or uneven.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) documents so left and right margins (กั้นหน้า–กั้นหลัง) align cleanly without stretched spaces or broken syllables. This skill contains pure instructions and OOXML specifications with no external scripts required.

## Core Specification

1. **Page Geometry (A4 Official Thai Standard)**:
   - Paper size: `21.0 × 29.7 cm` (A4).
   - Margins: Top `2.5 cm`, Bottom `2.0 cm`, Left / กั้นหน้า `3.0 cm`, Right / กั้นหลัง `2.0 cm`.
   - Net printable content width: **`16.0 cm`** (`9072 dxa`).
   - Header / Footer distance: `1.25 cm`.

2. **Typography & Complex Script Language Tagging**:
   - Font family: `TH Sarabun New` (or `TH SarabunPSK`).
   - Size hierarchy: Title `22pt` Bold, H1/H2 `18pt` Bold, H3 `16pt` Bold, Body prose `16pt` Regular, Callouts `14.5pt`, Tables `13.5–14pt`, Footer `12pt` Italic.
   - Set all four font slots (`w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`) plus Complex Script sizing (`w:szCs` in half-points, e.g., `32` for `16pt`), Complex Script bold/italic (`w:bCs`, `w:iCs`), and **`<w:lang w:val="en-US" w:bidi="th-TH"/>`** on every run and on the `Normal` style.

3. **Alignment Strategy (กั้นหน้า–กั้นหลัง)**:
   - **Body prose, callout boxes, and multi-line bullets**: Use Thai Distributed (`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` / `<w:jc w:val="thaiDistribute"/>`) with `line_spacing = 1.15`. Add a `1.25 cm` first-line indent on body prose paragraphs and hanging indent (`left_indent = 1.0 cm`, `first_line_indent = -0.4 cm`) on bullet items.
   - **Headings, short metadata lines, and narrow table cells**: Use `WD_ALIGN_PARAGRAPH.LEFT` (`<w:jc w:val="left"/>`) so narrow columns stay readable without stretching characters. Keep headings attached to their following paragraph (`keep_with_next = True`).
   - Avoid standard `JUSTIFY` (`w:val="both"`, which creates huge gaps) and `DISTRIBUTE` (`w:val="distribute"`, which stretches the last line).

4. **Thai Word-Break Cues (`U+200B` / ZWSP)**:
   - When generating or editing Thai text in `.docx`, insert Zero-Width Space (`\u200b`) cleanly between Thai word boundaries so Word wraps long clauses cleanly toward the right margin.
   - If `pythainlp` is already available in the active Python environment, tokenize Thai character runs (`[\u0E00-\u0E7F]+`) using `newmm`, strip any existing `\u200b` first so processing is idempotent, and remove any `\u200b` right before `ๆ` or `ฯ`. Leave English words, numbers, and URLs untouched.
   - If `pythainlp` is not available, rely on native MS Word Thai justification via `<w:lang w:val="en-US" w:bidi="th-TH"/>` + `<w:jc w:val="thaiDistribute"/>` (and place `\u200b` between Thai word boundaries in long clauses when helpful).

5. **Tables & Dynamic Page Numbers**:
   - Set `<w:tblLayout w:type="fixed"/>` and total `<w:tblW>` equal to `16.0 cm` (`9072 dxa`) so tables align flush with both margins.
   - Add `<w:tblHeader/>` on row 0 to repeat table headers across page breaks, and `<w:cantSplit/>` on every row.
   - Add a centered footer with `PAGE` / `NUMPAGES` Word fields (`หน้า X / Y`).

For complete OpenXML tag snippets and troubleshooting patterns, read [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).

## Execution Modes

- **Mode A — Python / `python-docx` available in session**: Build or update the `.docx` inline following the Core Specification and OOXML reference in [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).
- **Mode B — Pure text/chat mode without code execution**: Draft the Thai content structured cleanly for Word import (marking headings, indented paragraphs, callouts, and tables) and provide the user with exact Word setup values: A4, margins Left `3.0 cm` / Right `2.0 cm` / Top `2.5 cm` / Bottom `2.0 cm`, font `TH Sarabun New 16pt`, paragraph alignment **Thai Distributed** (กระจายแบบไทย), and Proofing Language **Thai**.

## Self-Audit Checklist

Verify all 6 items before delivering any `.docx` or formatting output:

1. Paper is A4 (`21.0 × 29.7 cm`) with Left `3.0 cm` (กั้นหน้า) and Right `2.0 cm` (กั้นหลัง).
2. `TH Sarabun New` is set on `w:cs` with matching `w:szCs` and `<w:lang w:val="en-US" w:bidi="th-TH"/>`.
3. Body prose, callouts, and multi-line bullets use `THAI_JUSTIFY` (`thaiDistribute`), while headings and table cells use `LEFT`.
4. No Thai text contains duplicate `\u200b\u200b` or `\u200b` immediately before `ๆ` / `ฯ`.
5. Every table uses fixed layout summing to `16.0 cm` with `w:tblHeader` and `w:cantSplit`.
6. Footer displays dynamic page numbers (`หน้า X / Y`).
