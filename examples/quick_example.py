#!/usr/bin/env python3
"""
Minimal standalone example creating a publication-grade Thai A4 .docx
with thai-word-justify.
"""

import sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls

SKILL_SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SKILL_SCRIPTS))

from thai_word_justify import (  # noqa: E402
    add_thai_page_number_footer,
    apply_thai_run_style,
    configure_thai_a4_margins,
    insert_thai_zwsp_clean,
    set_thai_justify,
)


def main():
    out_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample_thai_proposal.docx")
    doc = Document()

    content_width_cm = configure_thai_a4_margins(
        doc, top=2.5, bottom=2.0, left=3.0, right=2.0, font_name="TH Sarabun New", body_size_pt=16
    )
    add_thai_page_number_footer(doc.sections[0], footer_text="เอกสารตัวอย่างภาษาไทย (TH Sarabun New)")

    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h1.paragraph_format.keep_with_next = True
    h1.paragraph_format.space_after = Pt(6)
    apply_thai_run_style(h1.add_run(insert_thai_zwsp_clean("1. ตัวอย่างการจัดกั้นหน้า–กั้นหลังภาษาไทย")), size_pt=18, bold=True, color_hex="#1E2761")

    body_text = (
        "การจัดทำเอกสารภาษาไทยใน Microsoft Word จำเป็นต้องตั้งค่าระยะขอบมาตรฐานเอสี่ "
        "กั้นหน้าสามเซนติเมตร กั้นหลังสองเซนติเมตร และใช้การจัดแนวแบบกระจายแบบไทย "
        "(thaiDistribute) ร่วมกับการแทรกตัวคั่นคำที่มองไม่เห็นและการระบุภาษาไทยระดับโครงสร้าง "
        "เพื่อให้ขอบด้านซ้ายและด้านขวาเรียบเสมอกันทุกบรรทัดโดยไม่ดึงตัวอักษรให้ถ่างผิดปกติ"
    )
    p = doc.add_paragraph()
    set_thai_justify(p)
    p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.space_after = Pt(8)
    apply_thai_run_style(p.add_run(insert_thai_zwsp_clean(body_text)), size_pt=16)

    widths = [4.0, 6.0, content_width_cm - 10.0]
    table = doc.add_table(rows=2, cols=3)
    table.autofit = False
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(round(content_width_cm * 567))}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    rows_data = [
        ["รายการ", "ค่ามาตรฐาน", "หมายเหตุ"],
        ["ระยะขอบซ้าย / ขวา", "3.0 ซม. / 2.0 ซม.", "พื้นที่เนื้อหาสุทธิ 16.0 ซม."],
    ]
    for r_idx, row_vals in enumerate(rows_data):
        row = table.rows[r_idx]
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        if r_idx == 0:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = Cm(widths[c_idx])
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            apply_thai_run_style(cp.add_run(insert_thai_zwsp_clean(val)), size_pt=14, bold=(r_idx == 0))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"Created example: {out_path}")


if __name__ == "__main__":
    main()
