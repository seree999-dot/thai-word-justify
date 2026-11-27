# thai-word-justify

Agent Skill สำหรับจัดหน้าเอกสาร Microsoft Word (`.docx`) ภาษาไทยให้กั้นหน้า–กั้นหลังเรียบเสมอ และจัดลำดับตัวเลขหัวข้อ–ระยะเยื้องแบบ **แบบที่ 1** ระดับสิ่งพิมพ์ราชการ/วิชาการ ใช้งานได้ทั้งใน **Google Gemini (Chat / Web / Gems)** และ **Google Antigravity / Gemini CLI**

> Publication-grade Thai Microsoft Word (`.docx`) justification skill (`thaiDistribute` / `THAI_JUSTIFY`, standard Thai A4 `3.0 / 2.0 cm` margins, `TH Sarabun New 16pt`, Thai **แบบที่ 1** heading numbering & cascading indentation, `w:bidi="th-TH"`, punctuation-safe ZWSP, layout presets, table-overflow auditing, and MS Word / LibreOffice visual QA).

---

## ผลลัพธ์จริงใน Microsoft Word (Before / After)

![Before and After Comparison in Microsoft Word](assets/before_after.png)

---

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

---

## เลือกแพ็กเกจให้เหมาะกับสภาพแวดล้อม (2 Editions)

| แพ็กเกจ | ไฟล์ดาวน์โหลด | โครงสร้าง | เหมาะสำหรับ |
| :--- | :--- | :--- | :--- |
| **Gemini / Script-Free Edition** | [`thai-word-justify-gemini.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-gemini.zip) *(หรือ `thai-word-justify.zip`)* | **Pure Markdown 100%** (ไม่มีไฟล์สคริปต์ `.py` — พร้อม Inline `python-docx` Recipe ใน `references/ooxml-thai-guide.md`) | **Google Gemini (Chat / Web / Gems)** และสภาพแวดล้อมที่ไม่อนุญาตให้มี executable scripts รวมถึงใช้ใน **Google Antigravity** ได้ทันที |
| **Full CLI Edition** | [`thai-word-justify-cli.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-cli.zip) | มีสคริปต์ Python CLI (`fix`, `check`), `add_thai_heading` / `add_thai_body`, Presets และ `render_docx_to_png.py` ตรวจภาพผ่าน MS Word/LibreOffice | **Google Antigravity**, **Gemini CLI** และเครื่องที่รัน Python ได้ |

---

## Presets มาตรฐานเอกสารไทย (3 รูปแบบ)

| Preset | ระยะขอบ (ซ้าย / ขวา / บน / ล่าง) | ความกว้างพื้นที่เนื้อหา | เลขหน้า |
| :--- | :--- | :--- | :--- |
| `proposal` *(ค่าเริ่มต้น)* | `3.0 / 2.0 / 2.5 / 2.0 ซม.` | `16.0 ซม.` | กึ่งกลางท้ายกระดาษ (`หน้า X / Y`) |
| `official` *(หนังสือราชการ)* | `3.0 / 2.0 / 2.5 / 2.0 ซม.` | `16.0 ซม.` | มุมขวาบน (รองรับเลขไทย `๐–๙`) |
| `academic` *(รายงาน/วิทยานิพนธ์)* | `3.81 / 2.54 / 2.54 / 2.54 ซม.` | `14.65 ซม.` | มุมขวาบน |

> **Mandatory Final Step**: เมื่อสร้างหรือปรับแก้เอกสารเสร็จสมบูรณ์ Skill จะตรวจสอบไฟล์ `.docx` และแสดงลิงก์/ไฟล์ Microsoft Word (`.docx`) ให้ดาวน์โหลดได้ทันทีในขั้นตอนสุดท้ายเสมอ

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

## การใช้งานแบบ CLI และทดสอบ (Full CLI Edition)

```bash
# ปรับกั้นหน้า-กั้นหลังและจัดหัวข้อแบบที่ 1 อัตโนมัติ (proposal / official / academic)
python scripts/thai_word_justify.py fix input.docx -o output.docx --preset proposal --footer "ชื่อเอกสาร"
python scripts/thai_word_justify.py fix input.docx -o official.docx --preset official --thai-numerals

# ตรวจสอบมาตรฐานเอกสารและตารางล้นขอบ (exit code 0 = ผ่าน)
python scripts/thai_word_justify.py check output.docx

# เรนเดอร์เป็นภาพ PNG ทุกหน้าเพื่อตรวจผลจริง
python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150

# รัน Unit Tests
python -m unittest discover -s tests -v
```

ดูตัวอย่างการสร้างเอกสาร **แบบที่ 1** ได้ที่ [`examples/quick_example.py`](./examples/quick_example.py) และคู่มือ OOXML + Inline Recipe ได้ที่ [`references/ooxml-thai-guide.md`](./references/ooxml-thai-guide.md)

## License

MIT License — see [LICENSE](./LICENSE).
