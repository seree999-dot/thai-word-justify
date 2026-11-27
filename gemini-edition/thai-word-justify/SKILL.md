---
name: thai-word-justify
description: >-
  Format and audit Microsoft Word (.docx) documents for publication-grade Thai
  margins, แบบที่ 1 hierarchical numbering, and justification (กั้นหน้า–กั้นหลัง,
  thaiDistribute / THAI_JUSTIFY, A4 3.0/2.0 cm margins, TH Sarabun New 16pt,
  1. / 1.1 / 1.1.1 / 1.1.2.1 / 1) cascading indents, w:bidi="th-TH", Thai
  word-break ZWSP), then upload to Google Drive and deliver a downloadable
  .docx file. Use when creating, fixing, or polishing Thai Word documents.
---

# Thai Word Justify

Format Thai Microsoft Word (`.docx`) documents so left and right margins (กั้นหน้า–กั้นหลัง) and Thai **แบบที่ 1** heading numbering & cascading indentation align cleanly without stretched spaces or broken syllables. This skill contains pure instructions and OOXML specifications with no external scripts required.

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
| `official` (หนังสือราชการ) | `3.0 / 2.0 / 2.5 / 2.0` | `16.0 cm` | `TH Sarabun New 16pt`, `1.0 / 2.0 cm` | Header right (Thai numerals `๐–๙` optional) |
| `academic` (รายงาน/วิทยานิพนธ์) | `3.81 / 2.54 / 2.54 / 2.54` | `14.65 cm` | `TH Sarabun New 16pt`, `1.0 / 2.0 cm` | Header right |

## Core Specification

1. **Page Geometry (A4 Thai Standards)**:
   - Paper size: `21.0 × 29.7 cm` (A4).
   - Default `proposal` / `official` margins: Top `2.5 cm`, Bottom `2.0 cm`, Left / กั้นหน้า `3.0 cm`, Right / กั้นหลัง `2.0 cm` → printable width **`16.0 cm`** (`9072 dxa`).
   - `academic` margins: Left `3.81 cm`, Right/Top/Bottom `2.54 cm` → printable width **`14.65 cm`** (`8307 dxa`).
   - Header / Footer distance: `1.25 cm`.

2. **Typography & Complex Script Language Tagging**:
   - Font family: `TH Sarabun New` (or `TH SarabunPSK`).
   - Follow the **แบบที่ 1** table above (`Title 18–20pt`, `H1 16–18pt`, `H2–H5 & Body 16pt`).
   - Set all four font slots (`w:ascii`, `w:hAnsi`, `w:cs`, `w:eastAsia`) plus Complex Script sizing (`w:szCs` in half-points, e.g., `32` for `16pt`, `36` for `18pt`, `40` for `20pt`), Complex Script bold/italic (`w:bCs`, `w:iCs`), and **`<w:lang w:val="en-US" w:bidi="th-TH"/>`** on every run and on the `Normal` style.

3. **Numbering, Alignment & Indentation Strategy**:
   - **Title / Chapter headings**: `WD_ALIGN_PARAGRAPH.CENTER`, Bold, `18–20 pt`, `keep_with_next = True`.
   - **Numbering punctuation & spacing**: Level 1 uses a dot (`1.  `), multi-level headings (`1.1  `, `1.1.1  `, `1.1.2.1  `) omit any trailing dot, and level-4 sub-items use a closing parenthesis (`1)  `). Place exactly **2 spaces** between the number and heading title.
   - **Cascading indents**: `h1` at `0.0 cm`, `h2` at `1.0 cm`, `h3` at `1.8 cm`, `h4` at `2.8 cm`, `h5` at `4.0 cm`.
   - **Body prose**: Thai Distributed (`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` / `<w:jc w:val="thaiDistribute"/>`), `line_spacing = 1.15`, `left_indent = 0`, `first_line_indent = 1.0 cm` under `1.` (5 spaces) or `2.0 cm` under `1.1`+ (10 spaces).
   - **Narrow table cells**: Use `WD_ALIGN_PARAGRAPH.LEFT` (`<w:jc w:val="left"/>`).

4. **Thai Word-Break Cues (`U+200B` / ZWSP)**:
   - Insert `\u200b` cleanly between Thai words (using `pythainlp` `newmm` if pre-installed, stripping existing `\u200b` first, and never before `ๆ` or `ฯ`). Leave English, numbers, and URLs untouched.

5. **Tables & Dynamic Page Numbers**:
   - Set `<w:tblLayout w:type="fixed"/>` and total `<w:tblW>` equal to the preset printable width (`16.0 cm` / `9072 dxa`, or `14.65 cm` / `8307 dxa`), with `<w:tblHeader/>` and `<w:cantSplit/>`.

For complete OpenXML tag snippets and the copy-ready inline `python-docx` helper, read [ooxml-thai-guide.md](./references/ooxml-thai-guide.md).

## Execution Modes

- **Mode A — Python / `python-docx` available in session**: Always run the inline `python-docx` builder from [ooxml-thai-guide.md](./references/ooxml-thai-guide.md) to generate the actual `.docx` file on disk (never stop at pasting raw code).
- **Mode B — Pure text/chat mode without any file or code execution**: Clearly state that code/file generation is unavailable in this chat session, provide the Thai content structured cleanly with **แบบที่ 1** numbering (`1.  `, `1.1  `, `1.1.1  `, `1.1.2.1  `, `1)  `), and give exact Word setup values: A4, Left `3.0 cm` / Right `2.0 cm`, `TH Sarabun New`, Thai Distributed, and Proofing Language Thai.

## Final Step: Upload to Google Drive & Deliver Downloadable `.docx` (ขั้นตอนสุดท้าย — ส่งไฟล์ไปยัง Google Drive และให้ดาวน์โหลดได้ทันที)

Whenever code execution or file output tools are available:

1. **Generate & verify**: Save the `.docx` file to the session output/workspace directory and confirm the file exists with a non-zero byte size (`> 0` bytes).
2. **Upload to Google Drive**: Upload the `.docx` file directly to the user's **Google Drive** (into the requested folder or My Drive root), preserving the native `.docx` file format so `thaiDistribute`, `w:bidi="th-TH"`, and `TH Sarabun New` remain intact. Keep default private access unless the user explicitly asks to share.
3. **Present immediately at the end of your reply**: Display BOTH the clickable **Google Drive link** and the **immediately downloadable `.docx` file** right in your final message, including filename, file size in bytes/KB, and a 1-line summary confirming `แบบที่ 1` numbering, `TH Sarabun New`, `3.0 / 2.0 cm` margins, and `thaiDistribute`.
4. **Fallback if Google Drive is unavailable**: If Google Drive integration is not connected or authorized in the current chat session, still deliver the downloadable `.docx` file immediately and clearly note that Google Drive can be connected for automatic upload.

## Self-Audit Checklist

1. Paper is A4 (`21.0 × 29.7 cm`) with Left `3.0 cm` (or `3.81 cm` academic) and Right `2.0 cm` (or `2.54 cm`).
2. `TH Sarabun New` is set on `w:cs` with matching `w:szCs` and `<w:lang w:val="en-US" w:bidi="th-TH"/>`.
3. Heading numbering follows **แบบที่ 1** (`1.  ` with dot; `1.1  `, `1.1.1  `, `1.1.2.1  ` without trailing dot; `1)  ` with parenthesis; 2 spaces after number; cascading indents `0.0 → 1.0 → 1.8 → 2.8 → 4.0 cm`).
4. Body prose uses `THAI_JUSTIFY` (`1.0 cm` indent under `1.`, `2.0 cm` indent under `1.1`+).
5. No duplicate `\u200b\u200b`, no `\u200b` before `ๆ` / `ฯ`, and every table fits within printable width (`<= 16.0 cm`).
6. Page numbers match the selected preset.
7. The final Microsoft Word (`.docx`) file is generated, non-empty, uploaded to Google Drive (when connected), and presented with both the Google Drive link and downloadable `.docx` file/link.
