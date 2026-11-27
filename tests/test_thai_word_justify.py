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
