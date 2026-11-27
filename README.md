# thai-word-justify

Agent Skill สำหรับจัดหน้าเอกสาร Microsoft Word (`.docx`) ภาษาไทยให้กั้นหน้า–กั้นหลังเรียบเสมอระดับสิ่งพิมพ์ราชการ/องค์กร ใช้งานได้กับ **Google Antigravity** และ **Google Gemini**

> Publication-grade Thai Microsoft Word (`.docx`) justification skill (`thaiDistribute` / `THAI_JUSTIFY`, standard Thai A4 `3.0 / 2.0 cm` margins, `TH Sarabun New 16pt`, `w:bidi="th-TH"`, punctuation-safe ZWSP, dynamic page numbers, and MS Word / LibreOffice visual QA).

---

## ความสามารถหลัก (Features)

- **กั้นหน้า–กั้นหลังภาษาไทยแท้ (`thaiDistribute`)**: ใช้ `WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` ร่วมกับ `w:lang w:bidi="th-TH"` บนย่อหน้าบรรยายและ Callout เพื่อไม่ให้ช่องว่างหรือตัวอักษรถ่างผิดปกติ
- **ตัดคำไทยอย่างปลอดภัย (`U+200B`)**: ใช้ `pythainlp` (`newmm`) แทรก Zero-Width Space แบบ Idempotent (ป้องกันซ้ำซ้อนและไม่ตัดหน้า `ๆ`, `ฯ`)
- **ระยะขอบ A4 มาตรฐานไทย**: บน `2.5 ซม.`, ล่าง `2.0 ซม.`, ซ้าย/กั้นหน้า `3.0 ซม.`, ขวา/กั้นหลัง `2.0 ซม.` (ความกว้างพื้นที่เนื้อหาสุทธิ `16.0 ซม.`)
- **ฟอนต์ Complex Script (`TH Sarabun New`)**: ตั้งค่า `w:ascii`, `w:hAnsi`, `w:cs`, `w:szCs` และ `w:bCs` ครบถ้วน
- **ตารางมาตรฐานและเลขหน้า**: ตารางกว้าง `16.0 ซม.` ชิดขอบซ้าย–ขวาพอดี (`tblHeader` + `cantSplit`) พร้อมท้ายกระดาษ `หน้า X / Y`
- **Visual QA อัตโนมัติ**: เรนเดอร์ `.docx` เป็น PDF/PNG ผ่าน Microsoft Word (Windows) หรือ LibreOffice (macOS/Linux) เพื่อตรวจหน้าจริงทุกหน้า

---

## การติดตั้ง (Installation)

### 1. ติดตั้งใน Google Antigravity

**วิธี A — โคลนจาก GitHub:**
```bash
# Global (ใช้ได้ทุกโปรเจกต์)
git clone https://github.com/seree999-dot/thai-word-justify.git ~/.gemini/config/skills/thai-word-justify

# หรือเฉพาะ Workspace ปัจจุบัน
git clone https://github.com/seree999-dot/thai-word-justify.git .agents/skills/thai-word-justify
```

**วิธี B — แตกไฟล์ `thai-word-justify.zip` (จาก GitHub Releases):**
- **Windows (PowerShell):**
  ```powershell
  Expand-Archive .\thai-word-justify.zip -DestinationPath "$env:USERPROFILE\.gemini\config\skills" -Force
  ```
- **macOS / Linux:**
  ```bash
  unzip -o thai-word-justify.zip -d ~/.gemini/config/skills/
  ```

### 2. ติดตั้งใน Google Gemini (Gemini CLI / Gemini Agent Skills)

**วิธี A — โคลนจาก GitHub (User-level Gemini CLI):**
```bash
git clone https://github.com/seree999-dot/thai-word-justify.git ~/.gemini/skills/thai-word-justify
```

**วิธี B — แตกไฟล์ `thai-word-justify.zip` (จาก GitHub Releases):**
- **Windows (PowerShell):**
  ```powershell
  Expand-Archive .\thai-word-justify.zip -DestinationPath "$env:USERPROFILE\.gemini\skills" -Force
  ```
- **macOS / Linux:**
  ```bash
  unzip -o thai-word-justify.zip -d ~/.gemini/skills/
  ```
- สำหรับ Gemini Web / แชท: สามารถอัปโหลดไฟล์ `thai-word-justify.zip` หรือ `SKILL.md` แนบเข้าในสนทนาได้โดยตรง

### 3. ติดตั้ง Python Dependencies
```bash
pip install -r requirements.txt
```

---

## การใช้งานแบบ CLI (Usage)

```bash
# ปรับกั้นหน้า-กั้นหลังไฟล์ Word ที่มีอยู่
python scripts/thai_word_justify.py fix input.docx -o output.docx --indent 1.25 --footer "ชื่อเอกสาร"

# ตรวจสอบมาตรฐานเอกสาร (exit code 0 = ผ่าน)
python scripts/thai_word_justify.py check output.docx

# เรนเดอร์เป็นภาพ PNG ทุกหน้าเพื่อตรวจผลจริง
python scripts/render_docx_to_png.py output.docx ./preview_png --dpi 150
```

ดูตัวอย่างการสร้างเอกสารใหม่ตั้งแต่ต้นได้ที่ [`examples/quick_example.py`](./examples/quick_example.py) และโครงสร้าง OOXML ได้ที่ [`references/ooxml-thai-guide.md`](./references/ooxml-thai-guide.md)

## License

MIT License — see [LICENSE](./LICENSE).
