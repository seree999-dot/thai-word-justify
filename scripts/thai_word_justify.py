#!/usr/bin/env python3
"""
thai_word_justify.py — Library and CLI for Thai Microsoft Word (.docx)
margins (กั้นหน้า-กั้นหลัง), thaiDistribute justification, w:bidi="th-TH",
TH Sarabun New complex-script styling, page numbering, fixing, and auditing.
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


def _append_word_field(paragraph, field_name: str, size_pt: float, color_hex: str, font_name: str):
    run_begin = paragraph.add_run()
    apply_thai_run_style(run_begin, size_pt=size_pt, color_hex=color_hex, font_name=font_name)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(fld_begin)

    run_instr = paragraph.add_run()
    apply_thai_run_style(run_instr, size_pt=size_pt, color_hex=color_hex, font_name=font_name)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    apply_thai_run_style(run_sep, size_pt=size_pt, color_hex=color_hex, font_name=font_name)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(fld_sep)

    run_val = paragraph.add_run("1")
    apply_thai_run_style(run_val, size_pt=size_pt, color_hex=color_hex, font_name=font_name)

    run_end = paragraph.add_run()
    apply_thai_run_style(run_end, size_pt=size_pt, color_hex=color_hex, font_name=font_name)
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

    _append_word_field(p, "PAGE", size_pt=size_pt, color_hex=color_hex, font_name=font_name)

    r_mid = p.add_run(" / ")
    apply_thai_run_style(r_mid, size_pt=size_pt, italic=True, color_hex=color_hex, font_name=font_name)

    _append_word_field(p, "NUMPAGES", size_pt=size_pt, color_hex=color_hex, font_name=font_name)


def _run_has_only_text(run) -> bool:
    r = run._r
    return (
        r.find(qn("w:drawing")) is None
        and r.find(qn("w:fldChar")) is None
        and r.find(qn("w:instrText")) is None
        and bool(run.text)
    )


def _is_heading_paragraph(p) -> bool:
    style_name = (p.style.name if p.style else "") or ""
    if "Heading" in style_name or "Title" in style_name or "Subtitle" in style_name:
        return True
    if p.runs and all(r.bold for r in p.runs if r.text.strip()) and len(p.text.strip()) < 90:
        return True
    return False


def fix_existing_docx(
    input_path: str | Path,
    output_path: str | Path,
    top: float = 2.5,
    bottom: float = 2.0,
    left: float = 3.0,
    right: float = 2.0,
    font_name: str = "TH Sarabun New",
    body_size_pt: float = 16,
    first_line_indent_cm: float = 0.0,
    footer_text: str | None = None,
) -> Path:
    """
    Upgrade an existing .docx file with Thai A4 margins, thaiDistribute alignment
    on body prose, w:bidi='th-TH', TH Sarabun New complex script fonts, clean ZWS,
    and fixed-width tables matching the printable width.
    """
    doc = Document(str(input_path))
    content_width_cm = configure_thai_a4_margins(
        doc,
        top=top,
        bottom=bottom,
        left=left,
        right=right,
        font_name=font_name,
        body_size_pt=body_size_pt,
    )

    for p in doc.paragraphs:
        raw_text = p.text.strip()
        if not raw_text:
            continue

        is_heading = _is_heading_paragraph(p)
        is_centered = p.alignment == WD_ALIGN_PARAGRAPH.CENTER

        if is_heading:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.keep_with_next = True
        elif not is_centered and len(raw_text) >= 40:
            set_thai_justify(p)
            p.paragraph_format.line_spacing = 1.15
            if first_line_indent_cm > 0 and not p.paragraph_format.left_indent:
                p.paragraph_format.first_line_indent = Cm(first_line_indent_cm)

        for run in p.runs:
            if _run_has_only_text(run):
                run.text = insert_thai_zwsp_clean(run.text)
            cur_size = run.font.size.pt if run.font.size else (18 if is_heading else body_size_pt)
            apply_thai_run_style(
                run,
                size_pt=cur_size,
                bold=run.bold,
                italic=run.italic,
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
                            run.text = insert_thai_zwsp_clean(run.text)
                        cur_size = run.font.size.pt if run.font.size else max(body_size_pt - 2.5, 13.0)
                        apply_thai_run_style(
                            run,
                            size_pt=cur_size,
                            bold=run.bold,
                            italic=run.italic,
                            color_hex=None,
                            font_name=font_name,
                        )

    if footer_text is not None:
        add_thai_page_number_footer(doc.sections[0], footer_text=footer_text or None, font_name=font_name)

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def check_docx(docx_path: str | Path) -> tuple[bool, list[str]]:
    """Audit a .docx file against Thai Word justification and margin standards."""
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

    bad_table_align = 0
    for table in doc.tables:
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
        f"Page: A4 ({w_cm} x {h_cm} cm) | Margins L={l_cm}, R={r_cm}, T={t_cm}, B={b_cm} cm",
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

    p_fix = sub.add_parser("fix", help="Fix margins, thaiDistribute, w:bidi, and ZWS on an existing .docx")
    p_fix.add_argument("input", help="Input .docx path")
    p_fix.add_argument("-o", "--output", required=True, help="Output .docx path")
    p_fix.add_argument("--font", default="TH Sarabun New", help="Thai font name")
    p_fix.add_argument("--size", type=float, default=16.0, help="Body font size in pt")
    p_fix.add_argument("--left", type=float, default=3.0, help="Left margin (กั้นหน้า) in cm")
    p_fix.add_argument("--right", type=float, default=2.0, help="Right margin (กั้นหลัง) in cm")
    p_fix.add_argument("--top", type=float, default=2.5, help="Top margin in cm")
    p_fix.add_argument("--bottom", type=float, default=2.0, help="Bottom margin in cm")
    p_fix.add_argument("--indent", type=float, default=0.0, help="First-line indent in cm (e.g. 1.25)")
    p_fix.add_argument("--footer", default=None, help="Optional footer text with dynamic page numbers")

    p_check = sub.add_parser("check", help="Audit an existing .docx for Thai justification & margins")
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
        )
        print(f"Fixed and saved: {out}")
    elif args.command == "check":
        ok, lines = check_docx(args.input)
        print("\n".join(lines))
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
