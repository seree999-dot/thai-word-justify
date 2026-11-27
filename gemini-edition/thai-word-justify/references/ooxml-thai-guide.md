# OOXML Reference for Thai Microsoft Word Formatting

Reference details for OpenXML tags required for publication-grade Thai margins and justification in `.docx` documents.

## 1. Alignment (`w:jc`)

| Word Behavior | OOXML Tag | Use Case |
| :--- | :--- | :--- |
| Thai Distributed (`THAI_JUSTIFY`) | `<w:jc w:val="thaiDistribute"/>` | Multi-line Thai prose, callout boxes, multi-line bullets |
| Left (`LEFT`) | `<w:jc w:val="left"/>` | Headings, short metadata lines, narrow table cells |
| Standard Justify (`JUSTIFY`) | `<w:jc w:val="both"/>` | Avoid for Thai — stretches visible spaces awkwardly |
| Distribute (`DISTRIBUTE`) | `<w:jc w:val="distribute"/>` | Avoid — stretches the final short line of a paragraph |

## 2. Complex Script Fonts & Language (`w:rPr`)

Microsoft Word relies on Complex Script (`cs`) font properties and `w:bidi="th-TH"` to engage its Thai shaping and justification engine:

```xml
<w:rPr>
  <w:rFonts w:ascii="TH Sarabun New" w:hAnsi="TH Sarabun New" w:cs="TH Sarabun New" w:eastAsia="TH Sarabun New"/>
  <w:bCs/>
  <w:sz w:val="32"/>
  <w:szCs w:val="32"/>
  <w:lang w:val="en-US" w:bidi="th-TH"/>
</w:rPr>
```

- Font sizes are in **half-points** (`16pt` = `32`, `14pt` = `28`, `18pt` = `36`).
- Omitting `w:szCs` causes Thai text to render at Word's default 10–11pt while Latin text renders at 16pt.

## 3. A4 Thai Page Geometry (`1 cm = 567 dxa`)

- A4 width: `21.0 cm`, height: `29.7 cm`.
- Official Thai margins: Left (กั้นหน้า) `3.0 cm`, Right (กั้นหลัง) `2.0 cm`, Top `2.5 cm`, Bottom `2.0 cm`.
- Net printable width: `21.0 - 3.0 - 2.0 = 16.0 cm` (`9072 dxa`).

## 4. Tables Matching Printable Width

Set `<w:tblW w:w="9072" w:type="dxa"/>` and `<w:tblLayout w:type="fixed"/>` so tables align flush with both margins. Add `<w:tblHeader/>` on row 0 to repeat column headers across page breaks, and `<w:cantSplit/>` on every row to avoid mid-row page splits.

## 5. Common Failure Modes

- **Ragged right margin despite `thaiDistribute`**: Missing `w:bidi="th-TH"` or long Thai runs without `U+200B` word-break cues.
- **Double `U+200B` after re-processing**: Always strip existing `\u200b` before inserting word breaks.
- **Orphaned closing parentheses or `ๆ`**: Never insert `\u200b` before `ๆ`, `ฯ`, or punctuation.
- **Stretched words inside narrow table columns**: Keep table cell paragraphs `LEFT`-aligned rather than `thaiDistribute`.
