#!/usr/bin/env python3
"""
Minimal standalone example creating a Thai A4 .docx replicating the
Thai แบบที่ 1 heading numbering, cascading indents, and body prose indents.
"""

import sys
from pathlib import Path
from docx import Document

SKILL_SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SKILL_SCRIPTS))

from thai_word_justify import (  # noqa: E402
    add_thai_body,
    add_thai_heading,
    add_thai_page_number_footer,
    configure_thai_a4_margins,
)


def main():
    out_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample_thai_proposal.docx")
    doc = Document()

    configure_thai_a4_margins(
        doc, top=2.5, bottom=2.0, left=3.0, right=2.0, font_name="TH Sarabun New", body_size_pt=16
    )
    add_thai_page_number_footer(doc.sections[0], footer_text="ตัวอย่างการจัดหัวข้อภาษาไทย (แบบที่ 1)")

    add_thai_heading(doc, "แบบที่ 1", level="title")

    add_thai_heading(doc, "1.  หัวข้อใหญ่", level="h1")
    add_thai_body(
        doc,
        "ย่อหน้าแรกใต้หัวข้อใหญ่ระดับที่ 1 ใช้ระยะเยื้องบรรทัดแรก 1.0 เซนติเมตร (5 เคาะ) ให้ตรงแนวกับหัวข้อย่อย 1.1 "
        "และจัดแนวแบบกั้นหน้ากั้นหลังภาษาไทย (thaiDistribute) เพื่อให้ขอบซ้ายและขอบขวาเสมอกันทุกบรรทัด",
        under_level="h1",
    )

    add_thai_heading(doc, "1.1  หัวข้อย่อยระดับที่ 1", level="h2")
    add_thai_body(
        doc,
        "ย่อหน้าใต้หัวข้อย่อยระดับที่ 1 ใช้ระยะเยื้องบรรทัดแรก 2.0 เซนติเมตร (10 เคาะ) เยื้องลึกเข้าไปจากแนวหัวข้อ 1.1 "
        "โดยบรรทัดถัดไปของย่อหน้าชิดขอบซ้ายและจัดแนวแบบกั้นหน้ากั้นหลังอย่างเรียบร้อย",
        under_level="h2",
    )

    add_thai_heading(doc, "1.1.1  หัวข้อย่อยระดับที่ 2", level="h3")
    add_thai_heading(doc, "1.1.2  หัวข้อย่อยระดับที่ 2", level="h3")
    add_thai_heading(doc, "1.1.2.1  หัวข้อย่อยระดับที่ 3", level="h4")
    add_thai_heading(doc, "1)  หัวข้อย่อยระดับที่ 4", level="h5")
    add_thai_heading(doc, "2)  หัวข้อย่อยระดับที่ 4", level="h5")
    add_thai_heading(doc, "1.1.2.2  หัวข้อย่อยระดับที่ 3", level="h4")
    add_thai_heading(doc, "1.1.3  หัวข้อย่อยระดับที่ 2", level="h3")
    add_thai_heading(doc, "1.2  หัวข้อย่อยระดับที่ 1", level="h2")

    add_thai_heading(doc, "2.  หัวข้อใหญ่", level="h1")
    add_thai_heading(doc, "2.1  หัวข้อย่อยระดับที่ 1", level="h2")
    add_thai_heading(doc, "2.1.1  หัวข้อย่อยระดับที่ 2", level="h3")
    add_thai_heading(doc, "2.1.1.1  หัวข้อย่อยระดับที่ 3", level="h4")
    add_thai_heading(doc, "2.2  หัวข้อย่อยระดับที่ 1", level="h2")
    add_thai_heading(doc, "2.3  หัวข้อย่อยระดับที่ 1", level="h2")

    add_thai_heading(doc, "3.  หัวข้อใหญ่", level="h1")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"Created example: {out_path}")


if __name__ == "__main__":
    main()
