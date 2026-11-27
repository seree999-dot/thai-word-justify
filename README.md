# thai-word-justify

Agent Skill สำหรับจัดหน้าเอกสาร Microsoft Word (`.docx`) ภาษาไทยให้กั้นหน้า–กั้นหลังเรียบเสมอระดับสิ่งพิมพ์ราชการ/องค์กร ใช้งานได้ทั้งใน **Google Gemini (Chat / Web / Gems)** และ **Google Antigravity / Gemini CLI**

> Publication-grade Thai Microsoft Word (`.docx`) justification skill (`thaiDistribute` / `THAI_JUSTIFY`, standard Thai A4 `3.0 / 2.0 cm` margins, `TH Sarabun New 16pt`, `w:bidi="th-TH"`, punctuation-safe ZWSP, dynamic page numbers, and MS Word / LibreOffice visual QA).

---

## เลือกแพ็กเกจให้เหมาะกับสภาพแวดล้อม (2 Editions)

| แพ็กเกจ | ไฟล์ดาวน์โหลด | โครงสร้าง | เหมาะสำหรับ |
| :--- | :--- | :--- | :--- |
| **Gemini / Script-Free Edition** | [`thai-word-justify-gemini.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-gemini.zip) *(หรือ `thai-word-justify.zip`)* | **Pure Markdown 100%** (ไม่มีไฟล์สคริปต์ `.py` — อยู่ในโฟลเดอร์ `gemini-edition/thai-word-justify/`) | **Google Gemini (Chat / Web / Gems)** และสภาพแวดล้อมที่ไม่อนุญาตให้มี executable scripts รวมถึงใช้ใน **Google Antigravity** ได้ทันที |
| **Full CLI Edition** | [`thai-word-justify-cli.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-cli.zip) | มีสคริปต์ Python CLI (`fix`, `check`) และ `render_docx_to_png.py` ตรวจภาพผ่าน MS Word/LibreOffice | **Google Antigravity**, **Gemini CLI** และเครื่องที่รัน Python ได้ |

---

## ความสามารถหลัก (Features)

- **กั้นหน้า–กั้นหลังภาษาไทยแท้ (`thaiDistribute`)**: ใช้ `WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` ร่วมกับ `w:lang w:bidi="th-TH"` บนย่อหน้าบรรยายและ Callout เพื่อไม่ให้ช่องว่างหรือตัวอักษรถ่างผิดปกติ
- **ตัดคำไทยอย่างปลอดภัย (`U+200B`)**: แทรก Zero-Width Space แบบ Idempotent (ป้องกันซ้ำซ้อนและไม่ตัดหน้า `ๆ`, `ฯ`)
- **ระยะขอบ A4 มาตรฐานไทย**: บน `2.5 ซม.`, ล่าง `2.0 ซม.`, ซ้าย/กั้นหน้า `3.0 ซม.`, ขวา/กั้นหลัง `2.0 ซม.` (ความกว้างพื้นที่เนื้อหาสุทธิ `16.0 ซม.`)
- **ฟอนต์ Complex Script (`TH Sarabun New`)**: ตั้งค่า `w:ascii`, `w:hAnsi`, `w:cs`, `w:szCs` และ `w:bCs` ครบถ้วน
- **ตารางมาตรฐานและเลขหน้า**: ตารางกว้าง `16.0 ซม.` ชิดขอบซ้าย–ขวาพอดี (`tblHeader` + `cantSplit`) พร้อมท้ายกระดาษ `หน้า X / Y`

---

## การติดตั้ง (Installation)

### 1. สำหรับ Google Gemini (Chat / Web / Gems)
ดาวน์โหลด **`thai-word-justify-gemini.zip`** (ไม่มีไฟล์สคริปต์รันแยก) จาก [GitHub Releases](https://github.com/seree999-dot/thai-word-justify/releases/latest) แล้วอัปโหลดติดตั้งใน Google Gemini ได้ทันที

### 2. สำหรับ Google Antigravity / Gemini CLI
```bash
# Google Antigravity (Global)
git clone https://github.com/seree999-dot/thai-word-justify.git ~/.gemini/config/skills/thai-word-justify

# หรือ Gemini CLI (User-level)
git clone https://github.com/seree999-dot/thai-word-justify.git ~/.gemini/skills/thai-word-justify
```

ติดตั้ง Python dependencies เพิ่มเติมหากใช้ Full CLI Edition:
```bash
pip install -r requirements.txt
```

---

## การใช้งานแบบ CLI (Full CLI Edition)

```bash
# ปรับกั้นหน้า-กั้นหลังไฟล์ Word ที่มีอยู่
python scripts/thai_word_justify.py fix input.docx -o output.docx --indent 1.25 --footer "ชื่อเอกสาร"

# ตรวจสอบมาตรฐานเอกสาร (exit code 0 = ผ่าน)
python scripts/thai_word_justify.py check output.docx

# เรนเดอร์เป็นภาพ PNG ทุกหน้าเพื่อตรวจผลจริง
python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150
```

ดูตัวอย่างการสร้างเอกสารใหม่ได้ที่ [`examples/quick_example.py`](./examples/quick_example.py) และคู่มือ OOXML ได้ที่ [`references/ooxml-thai-guide.md`](./references/ooxml-thai-guide.md)

## License

MIT License — see [LICENSE](./LICENSE).
