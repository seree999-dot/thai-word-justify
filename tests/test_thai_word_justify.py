import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import thai_word_justify as twj  # noqa: E402


class TestThaiWordJustify(unittest.TestCase):
    def test_insert_thai_zwsp_clean_idempotent_and_punctuation(self):
        sample = "ระบบติดตามการฝึกงานของนักศึกษา (Student) ดีมากๆ"
        first = twj.insert_thai_zwsp_clean(sample)
        second = twj.insert_thai_zwsp_clean(first)
        self.assertIn("\u200b", first)
        self.assertEqual(first, second)
        self.assertNotIn("\u200b\u200b", first)
        self.assertNotIn("\u200bๆ", first)
        self.assertIn("(Student)", first)

    def test_to_thai_numerals(self):
        self.assertEqual(twj.to_thai_numerals("หน้า 12 / 34"), "หน้า ๑๒ / ๓๔")

    def test_normalize_thai_heading_number(self):
        self.assertEqual(twj.normalize_thai_heading_number("1. หัวข้อใหญ่"), "1.  หัวข้อใหญ่")
        self.assertEqual(twj.normalize_thai_heading_number("1.1. หัวข้อย่อยระดับที่ 1"), "1.1  หัวข้อย่อยระดับที่ 1")
        self.assertEqual(twj.normalize_thai_heading_number("1.1.1. หัวข้อย่อยระดับที่ 2"), "1.1.1  หัวข้อย่อยระดับที่ 2")
        self.assertEqual(twj.normalize_thai_heading_number("1.1.2.1. หัวข้อย่อยระดับที่ 3"), "1.1.2.1  หัวข้อย่อยระดับที่ 3")
        self.assertEqual(twj.normalize_thai_heading_number("1) หัวข้อย่อยระดับที่ 4"), "1)  หัวข้อย่อยระดับที่ 4")

    def test_presets_geometry(self):
        doc = Document()
        self.assertEqual(
            twj.configure_thai_a4_margins(
                doc,
                top=twj.PRESETS["proposal"]["top"],
                bottom=twj.PRESETS["proposal"]["bottom"],
                left=twj.PRESETS["proposal"]["left"],
                right=twj.PRESETS["proposal"]["right"],
            ),
            16.0,
        )
        self.assertEqual(
            twj.configure_thai_a4_margins(
                doc,
                top=twj.PRESETS["academic"]["top"],
                bottom=twj.PRESETS["academic"]["bottom"],
                left=twj.PRESETS["academic"]["left"],
                right=twj.PRESETS["academic"]["right"],
            ),
            14.65,
        )

    def test_heading_hierarchy_levels(self):
        doc = Document()
        twj.configure_thai_a4_margins(doc)
        p_t = twj.add_thai_heading(doc, "แบบที่ 1", level="title")
        p_1 = twj.add_thai_heading(doc, "1. หัวข้อใหญ่", level="h1")
        p_b1 = twj.add_thai_body(doc, "ย่อหน้าแรกใต้หัวข้อใหญ่ที่มีความยาวเพียงพอ", under_level="h1")
        p_2 = twj.add_thai_heading(doc, "1.1. หัวข้อย่อยระดับที่ 1", level="h2")
        p_b2 = twj.add_thai_body(doc, "ย่อหน้าใต้หัวข้อย่อยระดับที่ 1 ที่เยื้องสิบเคาะ", under_level="h2")
        p_3 = twj.add_thai_heading(doc, "1.1.1. หัวข้อย่อยระดับที่ 2", level="h3")
        p_4 = twj.add_thai_heading(doc, "1.1.2.1. หัวข้อย่อยระดับที่ 3", level="h4")
        p_5 = twj.add_thai_heading(doc, "1) หัวข้อย่อยระดับที่ 4", level="h5")

        self.assertEqual(p_t.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(p_t.runs[0].font.size.pt, 20.0)
        self.assertTrue(p_t.runs[0].bold)

        # Visual start = left_indent + first_line_indent
        for para, exp_start, exp_hang in [
            (p_1, 0.0, 0.65),
            (p_2, 1.0, 0.80),
            (p_3, 1.8, 1.00),
            (p_4, 2.8, 1.20),
            (p_5, 4.0, 0.65),
        ]:
            li = para.paragraph_format.left_indent.cm
            fli = para.paragraph_format.first_line_indent.cm
            self.assertAlmostEqual(li + fli, exp_start, places=2)
            self.assertAlmostEqual(fli, -exp_hang, places=2)

        self.assertAlmostEqual(p_b1.paragraph_format.first_line_indent.cm, 1.0, places=2)
        self.assertAlmostEqual(p_b2.paragraph_format.first_line_indent.cm, 2.0, places=2)

        with tempfile.TemporaryDirectory() as tmp:
            raw_path = Path(tmp) / "raw_h.docx"
            fixed_path = Path(tmp) / "fixed_h.docx"
            d2 = Document()
            d2.add_paragraph("แบบที่ 1")
            d2.add_paragraph("1. หัวข้อใหญ่")
            d2.add_paragraph("เนื้อหาใต้หัวข้อใหญ่ที่มีความยาวเพียงพอสำหรับการจัดแนวแบบกั้นหน้าและกั้นหลังให้สวยงาม")
            d2.add_paragraph("1.1. หัวข้อย่อยระดับที่ 1")
            d2.add_paragraph("เนื้อหาใต้หัวข้อย่อยระดับที่ 1 ที่เยื้องสิบเคาะและมีรายละเอียดต่อเนื่องหลายบรรทัด")
            d2.add_paragraph("1.1.1. หัวข้อย่อยระดับที่ 2")
            d2.add_paragraph("1.1.2.1. หัวข้อย่อยระดับที่ 3")
            d2.add_paragraph("1) หัวข้อย่อยระดับที่ 4")
            d2.save(str(raw_path))

            twj.fix_existing_docx(raw_path, fixed_path)
            df = Document(str(fixed_path))
            self.assertEqual(df.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertEqual(df.paragraphs[0].runs[0].font.size.pt, 20.0)
            self.assertTrue(df.paragraphs[1].text.startswith("1.  "))
            self.assertAlmostEqual(df.paragraphs[2].paragraph_format.first_line_indent.cm, 1.0, places=2)
            self.assertTrue(df.paragraphs[3].text.startswith("1.1  "))
            self.assertAlmostEqual(
                df.paragraphs[3].paragraph_format.left_indent.cm + df.paragraphs[3].paragraph_format.first_line_indent.cm,
                1.0,
                places=2,
            )
            self.assertAlmostEqual(df.paragraphs[4].paragraph_format.first_line_indent.cm, 2.0, places=2)
            self.assertTrue(df.paragraphs[5].text.startswith("1.1.1  "))
            self.assertTrue(df.paragraphs[6].text.startswith("1.1.2.1  "))
            self.assertTrue(df.paragraphs[7].text.startswith("1)  "))
            ok, lines = twj.check_docx(fixed_path)
            self.assertTrue(ok, "\n".join(lines))

    def test_multiline_heading_hanging_indent(self):
        doc = Document()
        twj.configure_thai_a4_margins(doc)
        p_long_h2 = twj.add_thai_heading(
            doc,
            "1.1  วัตถุประสงค์เชิงยุทธศาสตร์ของการพัฒนาระบบสารสนเทศเพื่อติดตามการฝึกงานนักศึกษา",
            level="h2",
        )
        p_long_h4 = twj.add_thai_heading(
            doc,
            "1.1.2.1  แนวทางการควบคุมสิทธิ์และการตรวจสอบข้อมูลย้อนหลังสำหรับผู้ประสานงาน",
            level="h4",
        )
        self.assertAlmostEqual(p_long_h2.paragraph_format.left_indent.cm, 1.80, places=2)
        self.assertAlmostEqual(p_long_h2.paragraph_format.first_line_indent.cm, -0.80, places=2)
        self.assertAlmostEqual(p_long_h4.paragraph_format.left_indent.cm, 4.00, places=2)
        self.assertAlmostEqual(p_long_h4.paragraph_format.first_line_indent.cm, -1.20, places=2)

    def test_fix_and_check_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            raw_path = Path(tmp) / "raw.docx"
            fixed_path = Path(tmp) / "fixed.docx"
            doc = Document()
            doc.add_paragraph(
                "การบริหารจัดการการฝึกงานของนักศึกษาในปัจจุบันยังพึ่งพาเอกสารกระดาษ "
                "ทำให้ตรวจสอบสถานะได้ยากและล่าช้ากว่ากำหนดการที่วางไว้"
            )
            doc.save(str(raw_path))

            twj.fix_existing_docx(raw_path, fixed_path, preset="official", thai_numerals=True)
            ok, lines = twj.check_docx(fixed_path)
            self.assertTrue(ok, "\n".join(lines))

    def test_check_detects_trailing_dot_subheading(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_path = Path(tmp) / "trailing_dot.docx"
            doc = Document()
            twj.configure_thai_a4_margins(doc)
            p_h = doc.add_paragraph()
            twj.apply_thai_run_style(p_h.add_run(twj.insert_thai_zwsp_clean("1.1. วัตถุประสงค์ที่มีจุดเกินท้ายเลขข้อ")), size_pt=16, bold=True)
            twj.add_thai_body(doc, "เนื้อหาปกติสำหรับทดสอบการตรวจสอบจุดเกินท้ายเลขหัวข้อแบบที่ 1", under_level="h2")
            doc.save(str(bad_path))

            ok, lines = twj.check_docx(bad_path)
            self.assertFalse(ok)
            self.assertTrue(any("trailing dot" in line for line in lines), "\n".join(lines))

    def test_check_detects_table_overflow(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_path = Path(tmp) / "overflow.docx"
            doc = Document()
            twj.configure_thai_a4_margins(doc)
            p = doc.add_paragraph()
            twj.set_thai_justify(p)
            twj.apply_thai_run_style(p.add_run(twj.insert_thai_zwsp_clean("ข้อความภาษาไทยตัวอย่างสำหรับทดสอบตารางล้นขอบ")))

            table = doc.add_table(rows=1, cols=2)
            tblPr = table._tbl.tblPr
            tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(19.0 * 567)}" w:type="dxa"/>'))
            for cell in table.rows[0].cells:
                cell.width = Cm(9.5)
            doc.save(str(bad_path))

            ok, lines = twj.check_docx(bad_path)
            self.assertFalse(ok)
            self.assertTrue(any("overflows printable width" in line for line in lines), "\n".join(lines))

    def test_check_detects_missing_thai_justify(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_path = Path(tmp) / "no_justify.docx"
            doc = Document()
            twj.configure_thai_a4_margins(doc)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            twj.apply_thai_run_style(p.add_run(twj.insert_thai_zwsp_clean("ข้อความที่จัดชิดซ้ายธรรมดา")))
            doc.save(str(bad_path))

            ok, lines = twj.check_docx(bad_path)
            self.assertFalse(ok)
            self.assertTrue(any("THAI_JUSTIFY" in line for line in lines), "\n".join(lines))


if __name__ == "__main__":
    unittest.main()
