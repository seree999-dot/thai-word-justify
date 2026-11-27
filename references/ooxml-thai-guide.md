# OOXML Reference for Thai Microsoft Word Formatting

Reference details for OpenXML tags required for publication-grade Thai margins, **แบบที่ 1** heading numbering & cascading indentation, and justification in `.docx` documents.

## 1. Alignment (`w:jc`)

| Word Behavior | OOXML Tag | Use Case |
| :--- | :--- | :--- |
| Thai Distributed (`THAI_JUSTIFY`) | `<w:jc w:val="thaiDistribute"/>` | Multi-line Thai prose, callout boxes, multi-line bullets |
| Center (`CENTER`) | `<w:jc w:val="center"/>` | Title / Chapter headings (ชื่อเรื่อง / ชื่อบท `18–20pt`) |
| Left (`LEFT`) | `<w:jc w:val="left"/>` | Numbered headings (`1.`, `1.1`, `1.1.1`, `1.1.2.1`, `1)`), metadata, narrow table cells |
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

- Font sizes are in **half-points** (`16pt` = `32`, `18pt` = `36`, `20pt` = `40`, `14pt` = `28`).
- Omitting `w:szCs` causes Thai text to render at Word's default 10–11pt while Latin text renders at 16pt.

## 3. A4 Thai Page Geometry (`1 cm = 567 dxa`)

- A4 width: `21.0 cm`, height: `29.7 cm`.
- Official/Proposal Thai margins: Left (กั้นหน้า) `3.0 cm`, Right (กั้นหลัง) `2.0 cm`, Top `2.5 cm`, Bottom `2.0 cm` → printable width `16.0 cm` (`9072 dxa`).
- Academic/Thesis margins: Left `3.81 cm`, Right `2.54 cm`, Top `2.54 cm`, Bottom `2.54 cm` → printable width `14.65 cm` (`8307 dxa`).

## 4. Tables Matching Printable Width

Set `<w:tblW w:w="9072" w:type="dxa"/>` (or `8307` for academic) and `<w:tblLayout w:type="fixed"/>` so tables align flush with both margins. Add `<w:tblHeader/>` on row 0 to repeat column headers across page breaks, and `<w:cantSplit/>` on every row to avoid mid-row page splits.

## 5. Common Failure Modes

- **Trailing dot on multi-level numbers (`1.1.`, `1.1.1.`)**: Only level 1 (`1.`) has a trailing dot; remove trailing dots on `1.1`, `1.1.1`, `1.1.2.1`, use `1)` for level 4, and keep **2 spaces** after the number.
- **Ragged right margin despite `thaiDistribute`**: Missing `w:bidi="th-TH"` or long Thai runs without `U+200B` word-break cues.
- **Double `U+200B` after re-processing**: Always strip existing `\u200b` before inserting word breaks.
- **Orphaned closing parentheses or `ๆ`**: Never insert `\u200b` before `ๆ`, `ฯ`, or punctuation.

## 6. Inline `python-docx` Recipe (Zero External Scripts)

Copy-ready helper supporting all **แบบที่ 1** heading levels (`title`, `h1`–`h5`) and context-aware `body` indents (`1.0 cm` under `1.`, `2.0 cm` under `1.1`+):

```python
import re
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def setup_thai_docx(left=3.0, right=2.0, top=2.5, bottom=2.0, font="TH Sarabun New", size=16):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.left_margin, s.right_margin = Cm(left), Cm(right)
    s.top_margin, s.bottom_margin = Cm(top), Cm(bottom)
    n = doc.styles["Normal"]
    n.font.name, n.font.size = font, Pt(size)
    n.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    n.paragraph_format.line_spacing = 1.15
    return doc

def _style_run(r, size=16, bold=False, italic=False, font="TH Sarabun New"):
    r.font.name, r.font.size, r.bold, r.italic = font, Pt(size), bold, italic
    rPr = r._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts")) or OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), font)
    if rf.getparent() is None:
        rPr.insert(0, rf)
    for t in ("w:sz", "w:szCs"):
        el = OxmlElement(t)
        el.set(qn("w:val"), str(int(size * 2)))
        rPr.append(el)
    if bold:
        rPr.append(OxmlElement("w:bCs"))
    if italic:
        rPr.append(OxmlElement("w:iCs"))
    lg = OxmlElement("w:lang")
    lg.set(qn("w:val"), "en-US")
    lg.set(qn("w:bidi"), "th-TH")
    rPr.append(lg)

def norm_num(text):
    t = text.lstrip()
    t = re.sub(r"^(\d+(?:\.\d+){1,3})\.?\s+", r"\1  ", t)
    t = re.sub(r"^(\d+\))\s+", r"\1  ", t)
    return re.sub(r"^(\d+)\.?\s+", r"\1.  ", t)

def add_thai_heading(doc, text, level="h1", bold=None, italic=False, font="TH Sarabun New"):
    specs = {"title": (20, True, "CENTER", 0.0), "h1": (18, True, "LEFT", 0.0), "h2": (16, True, "LEFT", 1.0), "h3": (16, False, "LEFT", 1.8), "h4": (16, False, "LEFT", 2.8), "h5": (16, False, "LEFT", 4.0)}
    sz, b, align, ind = specs[level]
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "CENTER" else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.left_indent = Cm(ind)
    txt = text if level == "title" else norm_num(text)
    _style_run(p.add_run(txt), size=sz, bold=(b if bold is None else bold), italic=italic, font=font)
    return p

def add_thai_body(doc, text, under_level="h1", size=16, font="TH Sarabun New"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.first_line_indent = Cm(1.0 if under_level == "h1" else 2.0)
    _style_run(p.add_run(text), size=size, font=font)
    return p
```
