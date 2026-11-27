#!/usr/bin/env python3
"""
thai_word_justify.py — Library and CLI for Thai Microsoft Word (.docx)
margins (กั้นหน้า-กั้นหลัง), thaiDistribute justification, w:bidi="th-TH",
TH Sarabun New complex-script styling, Thai heading hierarchy, presets,
page numbering, fixing, and auditing.
"""

import argparse
import re
import sys
from pathlib import Path
from pythainlp import word_tokenize
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ZWS = "\u200b"
_THAI_RUN_RE = re.compile(r"([\u0E00-\u0E7F]+)")
_ARABIC_TO_THAI = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")

TYPOGRAPHY_HIERARCHY: dict[str, dict] = {
    "title": {
        "label": "ชื่อเรื่อง / ชื่อบท",
        "size_pt": 20.0,
        "min_pt": 18.0,
        "max_pt": 20.0,
        "bold": True,
        "italic": False,
        "align": "CENTER",
        "indent_cm": 0.0,
        "space_before_pt": 0.0,
        "space_after_pt": 8.0,
    },
    "h1": {
        "label": "หัวข้อหลัก (ระดับที่ 1)",
        "size_pt": 18.0,
        "min_pt": 16.0,
        "max_pt": 18.0,
        "bold": True,
        "italic": False,
        "align": "LEFT",
        "indent_cm": 0.0,
        "space_before_pt": 12.0,
        "space_after_pt": 4.0,
    },
    "h2": {
        "label": "หัวข้อรอง (ระดับที่ 2)",
        "size_pt": 16.0,
        "min_pt": 16.0,
        "max_pt": 16.0,
        "bold": True,
        "italic": False,
        "align": "LEFT",
        "indent_cm": 0.75,
        "space_before_pt": 8.0,
        "space_after_pt": 3.0,
    },
    "h3": {
        "label": "หัวข้อย่อย (ระดับที่ 3)",
        "size_pt": 16.0,
        "min_pt": 16.0,
        "max_pt": 16.0,
        "bold": True,
        "italic": False,
        "align": "LEFT",
        "indent_cm": 1.5,
        "space_before_pt": 6.0,
        "space_after_pt": 2.0,
    },
    "body": {
        "label": "เนื้อหาปกติ",
        "size_pt": 16.0,
        "min_pt": 16.0,
        "max_pt": 16.0,
        "bold": False,
        "italic": False,
        "align": "THAI_JUSTIFY",
        "first_line_indent_cm": 1.25,
    },
}

PRESETS: dict[str, dict] = {
    "proposal": {
        "top": 2.5,
        "bottom": 2.0,
        "left": 3.0,
        "right": 2.0,
        "body_size_pt": 16.0,
        "indent": 1.25,
        "page_num_pos": "footer_center",
    },
    "official": {
        "top": 2.5,
        "bottom": 2.0,
        "left": 3.0,
        "right": 2.0,
        "body_size_pt": 16.0,
        "indent": 1.25,
        "page_num_pos": "header_right",
    },
    "academic": {
        "top": 2.54,
        "bottom": 2.54,
        "left": 3.81,
        "right": 2.54,
        "body_size_pt": 16.0,
        "indent": 1.25,
        "page_num_pos": "header_right",
    },
}


def to_thai_numerals(text: str) -> str:
    """Convert ASCII digits 0-9 to Thai numerals ๐-๙."""
    return text.translate(_ARABIC_TO_THAI) if text else text


def insert_thai_zwsp_clean(text: str, engine: str = "newmm") -> str:
    """
    Insert Zero-Width Space (U+200B) cleanly between Thai words.
    Idempotent: strips existing ZWS before tokenizing, leaves English/numbers/URLs
    intact, and prevents ZWS before 'ๆ'/'ฯ' or duplicate ZWS.
    """
    if not text:
        return text
    cleaned = text.replace(ZWS, "")
    parts = _THAI_RUN_RE.split(cleaned)
    out = []
    for part in parts:
        if _THAI_RUN_RE.fullmatch(part):
            tokens = [t for t in word_tokenize(part, engine=engine) if t]
            joined = ZWS.join(tokens)
            joined = re.sub(r"\u200b+([ๆฯ])", r"\1", joined)
            joined = re.sub(r"\u200b{2,}", ZWS, joined)
            out.append(joined)
        else:
            out.append(part)
    return "".join(out)


def hex_to_rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.lstrip("#"))


def apply_thai_run_style(
    run,
    size_pt: float = 16,
    bold: bool | None = False,
    italic: bool | None = False,
    color_hex: str | None = "#202124",
    font_name: str = "TH Sarabun New",
):
    """Apply TH Sarabun New fonts, Complex Script sizing, and w:bidi='th-TH'."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color_hex is not None:
        run.font.color.rgb = hex_to_rgb(color_hex)

    rPr = run._r.get_or_add_rPr()

    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    half_pts = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), half_pts)

    if bold is True and rPr.find(qn("w:bCs")) is None:
        rPr.append(OxmlElement("w:bCs"))
    if italic is True and rPr.find(qn("w:iCs")) is None:
        rPr.append(OxmlElement("w:iCs"))

    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")


def set_thai_justify(paragraph):
    """Set paragraph alignment to Thai Distributed (thaiDistribute / THAI_JUSTIFY)."""
    paragraph.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY


def add_thai_heading(
    doc,
    text: str,
    level: str = "h1",
    size_pt: float | None = None,
    indent_cm: float | None = None,
    italic: bool | None = None,
    bold: bool | None = None,
    color_hex: str = "#202124",
    font_name: str = "TH Sarabun New",
):
    """
    Add a Thai heading enforcing standard typography hierarchy:
    - title: 18-20 pt (default 20), bold, CENTER
    - h1:    16-18 pt (default 18), bold, LEFT (indent 0 cm)
    - h2:    16 pt, bold, LEFT or 1-level indent (default 0.75 cm)
    - h3:    16 pt, bold or italic (default bold=True), deeper indent (default 1.5 cm)
    """
    lvl = level.lower()
    if lvl not in ("title", "h1", "h2", "h3"):
        raise ValueError(f"Unsupported heading level '{level}'. Expected one of: title, h1, h2, h3.")
    cfg = TYPOGRAPHY_HIERARCHY[lvl]

    eff_size = cfg["size_pt"] if size_pt is None else float(size_pt)
    eff_indent = cfg["indent_cm"] if indent_cm is None else float(indent_cm)
    eff_italic = cfg["italic"] if italic is None else bool(italic)
    eff_bold = (not eff_italic if lvl == "h3" and italic is True and bold is None else cfg["bold"]) if bold is None else bool(bold)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if cfg["align"] == "CENTER" else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(cfg["space_before_pt"])
    p.paragraph_format.space_after = Pt(cfg["space_after_pt"])
    if eff_indent > 0:
        p.paragraph_format.left_indent = Cm(eff_indent)

    run = p.add_run(insert_thai_zwsp_clean(text))
    apply_thai_run_style(
        run,
        size_pt=eff_size,
        bold=eff_bold,
        italic=eff_italic,
        color_hex=color_hex,
        font_name=font_name,
    )
    return p


def configure_thai_a4_margins(
    doc,
    top: float = 2.5,
    bottom: float = 2.0,
    left: float = 3.0,
    right: float = 2.0,
    font_name: str = "TH Sarabun New",
    body_size_pt: float = 16,
) -> float:
    """Configure A4 paper and standard Thai margins (returns printable width in cm)."""
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)

    normal = doc.styles["Normal"]
    normal.font.name = font_name
    normal.font.size = Pt(body_size_pt)
    normal.font.color.rgb = hex_to_rgb("#202124")
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    normal.paragraph_format.line_spacing = 1.15

    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    half_pts = str(int(round(body_size_pt * 2)))
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), half_pts)

    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")

    return round(21.0 - left - right, 2)


def _append_word_field(
    paragraph,
    field_name: str,
    size_pt: float,
    color_hex: str,
    font_name: str,
    italic: bool = False,
):
    run_begin = paragraph.add_run()
    apply_thai_run_style(run_begin, size_pt=size_pt, italic=italic, color_hex=color_hex, font_name=font_name)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(fld_begin)

    run_instr = paragraph.add_run()
    apply_thai_run_style(run_instr, size_pt=size_pt, italic=italic, color_hex=color_hex, font_name=font_name)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    apply_thai_run_style(run_sep, size_pt=size_pt, italic=italic, color_hex=color_hex, font_name=font_name)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(fld_sep)

    run_val = paragraph.add_run("1")
    apply_thai_run_style(run_val, size_pt=size_pt, italic=italic, color_hex=color_hex, font_name=font_name)

    run_end = paragraph.add_run()
    apply_thai_run_style(run_end, size_pt=size_pt, italic=italic, color_hex=color_hex, font_name=font_name)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end._r.append(fld_end)


def add_thai_page_number_footer(
    section,
    footer_text: str | None = None,
    font_name: str = "TH Sarabun New",
    size_pt: float = 12,
    color_hex: str = "#5f6368",
):
    """Add a centered footer with a subtle top border and dynamic PAGE / NUMPAGES."""
    footer = section.footer
    p = footer.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(0)

    pPr = p._p.get_or_add_pPr()
    for existing in pPr.findall(qn("w:pBdr")):
        pPr.remove(existing)
    pPr.append(
        parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="4" w:color="DADCE0"/>'
            f'</w:pBdr>'
        )
    )

    prefix_str = f"{footer_text}  |  หน้า " if footer_text else "หน้า "
    r_prefix = p.add_run(insert_thai_zwsp_clean(prefix_str))
    apply_thai_run_style(r_prefix, size_pt=size_pt, italic=True, color_hex=color_hex, font_name=font_name)

    _append_word_field(p, "PAGE", size_pt=size_pt, color_hex=color_hex, font_name=font_name, italic=True)

    r_mid = p.add_run(" / ")
    apply_thai_run_style(r_mid, size_pt=size_pt, italic=True, color_hex=color_hex, font_name=font_name)

    _append_word_field(p, "NUMPAGES", size_pt=size_pt, color_hex=color_hex, font_name=font_name, italic=True)


def add_thai_page_number_header(
    section,
    font_name: str = "TH Sarabun New",
    size_pt: float = 14,
    color_hex: str = "#202124",
    thai_numerals: bool = False,
):
    """Add a right-aligned dynamic PAGE field in the header (official/academic style)."""
    header = section.header
    p = header.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    field_expr = r"PAGE \* thai1" if thai_numerals else "PAGE"
    _append_word_field(p, field_expr, size_pt=size_pt, color_hex=color_hex, font_name=font_name)


def _run_has_only_text(run) -> bool:
    r = run._r
    return (
        r.find(qn("w:drawing")) is None
        and r.find(qn("w:fldChar")) is None
        and r.find(qn("w:instrText")) is None
        and bool(run.text)
    )


def _detect_heading_level(p) -> str | None:
    style_name = (p.style.name if p.style else "") or ""
    text = p.text.replace(ZWS, "").strip()

    if "Title" in style_name or re.match(r"^บทที่\s+\S", text):
        return "title"
    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) < 90 and p.runs and all(r.bold for r in p.runs if r.text.strip()):
        return "title"
    if "Heading 3" in style_name or (re.match(r"^\d+\.\d+\.\d+\s+\S", text) and len(text) < 110):
        return "h3"
    if "Heading 2" in style_name or (re.match(r"^\d+\.\d+\s+\S", text) and len(text) < 110):
        return "h2"
    if "Heading 1" in style_name or "Heading" in style_name or (re.match(r"^\d+\.\s+\S", text) and len(text) < 110):
        return "h1"
    if p.runs and all(r.bold for r in p.runs if r.text.strip()) and len(text) < 90:
        return "h1"
    return None


def fix_existing_docx(
    input_path: str | Path,
    output_path: str | Path,
    top: float | None = None,
    bottom: float | None = None,
    left: float | None = None,
    right: float | None = None,
    font_name: str = "TH Sarabun New",
    body_size_pt: float | None = None,
    first_line_indent_cm: float | None = None,
    footer_text: str | None = None,
    preset: str = "proposal",
    thai_numerals: bool = False,
) -> Path:
    """
    Upgrade an existing .docx file with Thai A4 margins, thaiDistribute alignment
    on body prose, Thai heading hierarchy (title/h1/h2/h3), w:bidi='th-TH',
    TH Sarabun New complex script fonts, clean ZWS, and fixed-width tables.
    """
    cfg = PRESETS.get(preset, PRESETS["proposal"])
    top_v = cfg["top"] if top is None else top
    bottom_v = cfg["bottom"] if bottom is None else bottom
    left_v = cfg["left"] if left is None else left
    right_v = cfg["right"] if right is None else right
    size_v = cfg["body_size_pt"] if body_size_pt is None else body_size_pt
    indent_v = cfg["indent"] if first_line_indent_cm is None else first_line_indent_cm

    doc = Document(str(input_path))
    content_width_cm = configure_thai_a4_margins(
        doc,
        top=top_v,
        bottom=bottom_v,
        left=left_v,
        right=right_v,
        font_name=font_name,
        body_size_pt=size_v,
    )

    for p in doc.paragraphs:
        raw_text = p.text.strip()
        if not raw_text:
            continue

        h_level = _detect_heading_level(p)
        is_centered = p.alignment == WD_ALIGN_PARAGRAPH.CENTER

        if h_level is not None:
            h_cfg = TYPOGRAPHY_HIERARCHY[h_level]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if h_cfg["align"] == "CENTER" else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            if h_cfg["indent_cm"] > 0:
                p.paragraph_format.left_indent = Cm(h_cfg["indent_cm"])
        elif not is_centered and len(raw_text) >= 40:
            set_thai_justify(p)
            p.paragraph_format.line_spacing = 1.15
            if indent_v > 0 and not p.paragraph_format.left_indent:
                p.paragraph_format.first_line_indent = Cm(indent_v)

        for run in p.runs:
            if _run_has_only_text(run):
                updated = insert_thai_zwsp_clean(run.text)
                if thai_numerals:
                    updated = to_thai_numerals(updated)
                run.text = updated
            if h_level is not None:
                h_cfg = TYPOGRAPHY_HIERARCHY[h_level]
                cur_size = h_cfg["size_pt"]
                eff_italic = bool(run.italic) if h_level == "h3" else False
                eff_bold = (not eff_italic) if (h_level == "h3" and eff_italic and not run.bold) else True
            else:
                cur_size = run.font.size.pt if run.font.size else size_v
                eff_bold = run.bold
                eff_italic = run.italic

            apply_thai_run_style(
                run,
                size_pt=cur_size,
                bold=eff_bold,
                italic=eff_italic,
                color_hex=None,
                font_name=font_name,
            )

    for table in doc.tables:
        table.autofit = False
        tblPr = table._tbl.tblPr
        for old_tag in ("w:tblW", "w:tblLayout"):
            old = tblPr.find(qn(old_tag))
            if old is not None:
                tblPr.remove(old)
        tblPr.append(
            parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(round(content_width_cm * 567))}" w:type="dxa"/>')
        )
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

        num_cols = len(table.columns)
        if num_cols > 0:
            raw_widths = [c.width.cm if c.width else (content_width_cm / num_cols) for c in table.columns]
            total_raw = sum(raw_widths) or content_width_cm
            scaled_widths = [w * (content_width_cm / total_raw) for w in raw_widths]
        else:
            scaled_widths = []

        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
            if r_idx == 0 and trPr.find(qn("w:tblHeader")) is None:
                trPr.append(OxmlElement("w:tblHeader"))

            for c_idx, cell in enumerate(row.cells):
                if c_idx < len(scaled_widths):
                    w_cm = scaled_widths[c_idx]
                    cell.width = Cm(w_cm)
                    tcPr = cell._tc.get_or_add_tcPr()
                    old_tcW = tcPr.find(qn("w:tcW"))
                    if old_tcW is not None:
                        tcPr.remove(old_tcW)
                    tcPr.append(
                        parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(round(w_cm * 567))}" w:type="dxa"/>')
                    )
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        if _run_has_only_text(run):
                            updated = insert_thai_zwsp_clean(run.text)
                            if thai_numerals:
                                updated = to_thai_numerals(updated)
                            run.text = updated
                        cur_size = run.font.size.pt if run.font.size else max(size_v - 2.5, 13.0)
                        apply_thai_run_style(
                            run,
                            size_pt=cur_size,
                            bold=run.bold,
                            italic=run.italic,
                            color_hex=None,
                            font_name=font_name,
                        )

    if cfg.get("page_num_pos") == "header_right" and footer_text is None:
        add_thai_page_number_header(doc.sections[0], font_name=font_name, thai_numerals=thai_numerals)
    elif footer_text is not None:
        ft = to_thai_numerals(footer_text) if thai_numerals else footer_text
        add_thai_page_number_footer(doc.sections[0], footer_text=ft or None, font_name=font_name)

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def check_docx(docx_path: str | Path) -> tuple[bool, list[str]]:
    """Audit a .docx file against Thai Word justification, margins, and table widths."""
    doc = Document(str(docx_path))
    sec = doc.sections[0]
    issues: list[str] = []

    w_cm = round(sec.page_width.cm, 2)
    h_cm = round(sec.page_height.cm, 2)
    l_cm = round(sec.left_margin.cm, 2)
    r_cm = round(sec.right_margin.cm, 2)
    t_cm = round(sec.top_margin.cm, 2)
    b_cm = round(sec.bottom_margin.cm, 2)

    if not (20.8 <= w_cm <= 21.2 and 29.5 <= h_cm <= 29.9):
        issues.append(f"Page size is {w_cm}x{h_cm} cm, expected A4 (21.0x29.7 cm)")
    if l_cm < 2.5:
        issues.append(f"Left margin (กั้นหน้า) is {l_cm} cm, expected >= 2.5 cm (recommended 3.0 cm)")
    if r_cm < 1.5:
        issues.append(f"Right margin (กั้นหลัง) is {r_cm} cm, expected >= 1.5 cm (recommended 2.0 cm)")

    thai_justify_count = sum(1 for p in doc.paragraphs if p.alignment == WD_ALIGN_PARAGRAPH.THAI_JUSTIFY)
    if thai_justify_count == 0:
        issues.append("No paragraphs use THAI_JUSTIFY (thaiDistribute) alignment")

    printable_cm = round(w_cm - l_cm - r_cm, 2)
    max_dxa = int(round(printable_cm * 567)) + 60
    bad_table_align = 0

    for t_idx, table in enumerate(doc.tables, start=1):
        tblPr = table._tbl.tblPr
        if tblPr is not None:
            tblW = tblPr.find(qn("w:tblW"))
            if tblW is not None and tblW.get(qn("w:type")) == "dxa":
                w_val = int(tblW.get(qn("w:w")) or 0)
                if w_val > max_dxa:
                    issues.append(
                        f"Table {t_idx} width ({round(w_val / 567, 2)} cm) overflows printable width ({printable_cm} cm)"
                    )
        if table.rows:
            row0_dxa = 0
            has_dxa_cells = False
            for cell in table.rows[0].cells:
                tcPr = cell._tc.tcPr
                if tcPr is not None:
                    tcW = tcPr.find(qn("w:tcW"))
                    if tcW is not None and tcW.get(qn("w:type")) == "dxa":
                        has_dxa_cells = True
                        row0_dxa += int(tcW.get(qn("w:w")) or 0)
            if has_dxa_cells and row0_dxa > max_dxa:
                issues.append(
                    f"Table {t_idx} column sum ({round(row0_dxa / 567, 2)} cm) overflows printable width ({printable_cm} cm)"
                )

        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.alignment in (WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.DISTRIBUTE):
                        bad_table_align += 1

    if bad_table_align > 0:
        issues.append(f"Found {bad_table_align} table cell paragraph(s) using JUSTIFY/DISTRIBUTE instead of LEFT")

    bidi_langs: set[str] = set()
    cs_fonts: set[str] = set()
    zws_paras = 0
    double_zws = 0
    thai_paras = 0

    for p in doc.paragraphs:
        if _THAI_RUN_RE.search(p.text):
            thai_paras += 1
            if ZWS in p.text:
                zws_paras += 1
            if (ZWS * 2) in p.text:
                double_zws += 1
        for run in p.runs:
            rPr = run._r.rPr
            if rPr is not None:
                rf = rPr.find(qn("w:rFonts"))
                if rf is not None and rf.get(qn("w:cs")):
                    cs_fonts.add(rf.get(qn("w:cs")))
                lg = rPr.find(qn("w:lang"))
                if lg is not None and lg.get(qn("w:bidi")):
                    bidi_langs.add(lg.get(qn("w:bidi")))

    if "th-TH" not in bidi_langs:
        issues.append(f"Missing w:lang w:bidi='th-TH' on runs (found: {sorted(bidi_langs)})")
    if not cs_fonts:
        issues.append("Missing Complex Script font w:rFonts w:cs")
    if thai_paras > 0 and zws_paras == 0:
        issues.append("Thai paragraphs contain no Zero-Width Space (U+200B) word-break cues")
    if double_zws > 0:
        issues.append(f"Found {double_zws} paragraph(s) with duplicate ZWS (\\u200b\\u200b)")

    report = [
        f"File: {docx_path}",
        f"Page: A4 ({w_cm} x {h_cm} cm) | Margins L={l_cm}, R={r_cm}, T={t_cm}, B={b_cm} cm (printable {printable_cm} cm)",
        f"THAI_JUSTIFY paragraphs: {thai_justify_count} | Thai paras with ZWS: {zws_paras}/{thai_paras}",
        f"Tables: {len(doc.tables)} | Complex-script fonts: {sorted(cs_fonts)} | w:bidi: {sorted(bidi_langs)}",
    ]
    if issues:
        report.append("FAILED:")
        report.extend(f"  - {item}" for item in issues)
        return False, report
    report.append("ALL CHECKS PASSED")
    return True, report


def main():
    parser = argparse.ArgumentParser(description="Thai Microsoft Word (.docx) justification tool")
    sub = parser.add_subparsers(dest="command", required=True)

    p_fix = sub.add_parser("fix", help="Fix margins, thaiDistribute, headings, w:bidi, and ZWS on an existing .docx")
    p_fix.add_argument("input", help="Input .docx path")
    p_fix.add_argument("-o", "--output", required=True, help="Output .docx path")
    p_fix.add_argument("--preset", choices=sorted(PRESETS.keys()), default="proposal", help="Thai layout preset")
    p_fix.add_argument("--font", default="TH Sarabun New", help="Thai font name")
    p_fix.add_argument("--size", type=float, default=None, help="Body font size in pt")
    p_fix.add_argument("--left", type=float, default=None, help="Left margin (กั้นหน้า) in cm")
    p_fix.add_argument("--right", type=float, default=None, help="Right margin (กั้นหลัง) in cm")
    p_fix.add_argument("--top", type=float, default=None, help="Top margin in cm")
    p_fix.add_argument("--bottom", type=float, default=None, help="Bottom margin in cm")
    p_fix.add_argument("--indent", type=float, default=None, help="First-line indent in cm (e.g. 1.25)")
    p_fix.add_argument("--footer", default=None, help="Optional footer text with dynamic page numbers")
    p_fix.add_argument("--thai-numerals", action="store_true", help="Convert digits to Thai numerals (๐-๙)")

    p_check = sub.add_parser("check", help="Audit an existing .docx for Thai justification, margins & table overflow")
    p_check.add_argument("input", help="Input .docx path")

    args = parser.parse_args()
    if args.command == "fix":
        out = fix_existing_docx(
            args.input,
            args.output,
            top=args.top,
            bottom=args.bottom,
            left=args.left,
            right=args.right,
            font_name=args.font,
            body_size_pt=args.size,
            first_line_indent_cm=args.indent,
            footer_text=args.footer,
            preset=args.preset,
            thai_numerals=args.thai_numerals,
        )
        print(f"Fixed and saved: {out}")
    elif args.command == "check":
        ok, lines = check_docx(args.input)
        print("\n".join(lines))
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
