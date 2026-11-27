# thai-word-justify

Agent Skill สำหรับจัดหน้าเอกสาร Microsoft Word (`.docx`) ภาษาไทยให้กั้นหน้า–กั้นหลังเรียบเสมอระดับสิ่งพิมพ์ราชการ/องค์กร ใช้งานได้ทั้งใน **Google Gemini (Chat / Web / Gems)** และ **Google Antigravity / Gemini CLI**

> Publication-grade Thai Microsoft Word (`.docx`) justification skill (`thaiDistribute` / `THAI_JUSTIFY`, standard Thai A4 `3.0 / 2.0 cm` margins, `TH Sarabun New 16pt`, Thai heading hierarchy, `w:bidi="th-TH"`, punctuation-safe ZWSP, layout presets, table-overflow auditing, and MS Word / LibreOffice visual QA).

---

## ผลลัพธ์จริงใน Microsoft Word (Before / After)

![Before and After Comparison in Microsoft Word](assets/before_after.png)

---

## ลำดับขนาดฟอนต์และการจัดวางมาตรฐานภาษาไทย (`TH Sarabun New`)

| ระดับ (Level) | ขนาด (Size) | น้ำหนัก (Style) | การจัดวางและระยะเยื้อง (Alignment & Indent) |
| :--- | :--- | :--- | :--- |
| **ชื่อเรื่อง / ชื่อบท** (หัวข้อใหญ่กลางหน้ากระดาษ) | **18–20 pt** *(ค่าเริ่มต้น 20 pt)* | ตัวหนา (**Bold**) | จัดกึ่งกลางหน้ากระดาษ (`CENTER`) |
| **หัวข้อหลัก ระดับที่ 1** (เช่น `1. ความเป็นมา`) | **16–18 pt** *(ค่าเริ่มต้น 18 pt)* | ตัวหนา (**Bold**) | พิมพ์ชิดขอบซ้าย (`LEFT`, indent `0 cm`) |
| **หัวข้อรอง ระดับที่ 2** (เช่น `1.1 วัตถุประสงค์`) | **16 pt** | ตัวหนา (**Bold**) | พิมพ์ชิดซ้ายหรือเยื้องเข้า 1 ระดับ (`LEFT`, default indent `0.75 cm`) |
| **หัวข้อย่อย ระดับที่ 3** (เช่น `1.1.1 ด้านงบประมาณ`) | **16 pt** | ตัวหนา หรือ ตัวเอน (**Bold / Italic**) | เยื้องลึกเข้าไปจากหัวข้อรอง (`LEFT`, default indent `1.5 cm`) |
| **เนื้อหาปกติ** (Body prose) | **16 pt** | ตัวธรรมดา (**Regular**) | กั้นหน้า–กั้นหลัง (`THAI_JUSTIFY` / `thaiDistribute`), ย่อหน้าแรก `1.25 cm` |

---

## เลือกแพ็กเกจให้เหมาะกับสภาพแวดล้อม (2 Editions)

| แพ็กเกจ | ไฟล์ดาวน์โหลด | โครงสร้าง | เหมาะสำหรับ |
| :--- | :--- | :--- | :--- |
| **Gemini / Script-Free Edition** | [`thai-word-justify-gemini.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-gemini.zip) *(หรือ `thai-word-justify.zip`)* | **Pure Markdown 100%** (ไม่มีไฟล์สคริปต์ `.py` — พร้อม Inline `python-docx` Recipe ใน `references/ooxml-thai-guide.md`) | **Google Gemini (Chat / Web / Gems)** และสภาพแวดล้อมที่ไม่อนุญาตให้มี executable scripts รวมถึงใช้ใน **Google Antigravity** ได้ทันที |
| **Full CLI Edition** | [`thai-word-justify-cli.zip`](https://github.com/seree999-dot/thai-word-justify/releases/latest/download/thai-word-justify-cli.zip) | มีสคริปต์ Python CLI (`fix`, `check`), `add_thai_heading`, Presets (`proposal`, `official`, `academic`) และ `render_docx_to_png.py` ตรวจภาพผ่าน MS Word/LibreOffice | **Google Antigravity**, **Gemini CLI** และเครื่องที่รัน Python ได้ |

---

## Presets มาตรฐานเอกสารไทย (3 รูปแบบ)

| Preset | ระยะขอบ (ซ้าย / ขวา / บน / ล่าง) | ความกว้างพื้นที่เนื้อหา | เลขหน้า |
| :--- | :--- | :--- | :--- |
| `proposal` *(ค่าเริ่มต้น)* | `3.0 / 2.0 / 2.5 / 2.0 ซม.` | `16.0 ซม.` | กึ่งกลางท้ายกระดาษ (`หน้า X / Y`) |
| `official` *(หนังสือราชการ)* | `3.0 / 2.0 / 2.5 / 2.0 ซม.` | `16.0 ซม.` | มุมขวาบน (รองรับเลขไทย `๐–๙`) |
| `academic` *(รายงาน/วิทยานิพนธ์)* | `3.81 / 2.54 / 2.54 / 2.54 ซม.` | `14.65 ซม.` | มุมขวาบน |

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
# ปรับกั้นหน้า-กั้นหลังและระดับหัวข้อด้วย Preset (proposal / official / academic)
python scripts/thai_word_justify.py fix input.docx -o output.docx --preset proposal --footer "ชื่อเอกสาร"
python scripts/thai_word_justify.py fix input.docx -o official.docx --preset official --thai-numerals

# ตรวจสอบมาตรฐานเอกสารและตารางล้นขอบ (exit code 0 = ผ่าน)
python scripts/thai_word_justify.py check output.docx

# เรนเดอร์เป็นภาพ PNG ทุกหน้าเพื่อตรวจผลจริง
python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150

# รัน Unit Tests
python -m unittest discover -s tests -v
```

ดูตัวอย่างการสร้างเอกสารใหม่ได้ที่ [`examples/quick_example.py`](./examples/quick_example.py) และคู่มือ OOXML + Inline Recipe ได้ที่ [`references/ooxml-thai-guide.md`](./references/ooxml-thai-guide.md)

## License

MIT License — see [LICENSE](./LICENSE).
