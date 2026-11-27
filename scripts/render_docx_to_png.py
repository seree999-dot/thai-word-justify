#!/usr/bin/env python3
"""
render_docx_to_png.py — Render a .docx file to PDF -> PNGs for visual QA.
Uses Microsoft Word COM on Windows when available, with automatic LibreOffice
(soffice) headless fallback on macOS/Linux or non-Word systems.
"""

import argparse
import platform
import shutil
import subprocess
from pathlib import Path
import fitz  # PyMuPDF


def export_docx_to_pdf_via_word(docx_path: Path, pdf_path: Path):
    docx_str = str(docx_path).replace("'", "''")
    pdf_str = str(pdf_path).replace("'", "''")
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
$docFile = '{docx_str}'
$confirmConv = $false
$readOnly = $true
$saveChanges = $false
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open([ref]$docFile, [ref]$confirmConv, [ref]$readOnly)
    $doc.Fields.Update() | Out-Null
    $doc.ExportAsFixedFormat('{pdf_str}', 17)
}} finally {{
    if ($doc -ne $null) {{
        $doc.Close([ref]$saveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }}
    if ($word -ne $null) {{
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
    res = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if res.returncode != 0:
        raise RuntimeError(f"Word COM export failed:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}")


def export_docx_to_pdf_via_libreoffice(docx_path: Path, out_dir: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Neither Microsoft Word COM nor LibreOffice (soffice) is available to render DOCX.")
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if res.returncode != 0:
        raise RuntimeError(f"LibreOffice export failed:\nSTDOUT: {res.stdout}\nSTDERR: {res.stderr}")
    return out_dir / f"{docx_path.stem}.pdf"


def export_docx_to_pdf(docx_path: Path, pdf_path: Path) -> str:
    if platform.system() == "Windows":
        try:
            export_docx_to_pdf_via_word(docx_path, pdf_path)
            return "Microsoft Word"
        except Exception as word_err:
            if shutil.which("soffice"):
                generated = export_docx_to_pdf_via_libreoffice(docx_path, pdf_path.parent)
                if generated != pdf_path and generated.exists():
                    generated.replace(pdf_path)
                return "LibreOffice (fallback)"
            raise word_err
    generated = export_docx_to_pdf_via_libreoffice(docx_path, pdf_path.parent)
    if generated != pdf_path and generated.exists():
        generated.replace(pdf_path)
    return "LibreOffice"


def render_pdf_to_pngs(pdf_path: Path, output_dir: Path, dpi: int = 150) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    for old in output_dir.glob("page_*.png"):
        old.unlink()

    doc = fitz.open(str(pdf_path))
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    generated = []
    for idx, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_png = output_dir / f"page_{idx:02d}.png"
        pix.save(str(out_png))
        generated.append(out_png)
    doc.close()
    return generated


def main():
    parser = argparse.ArgumentParser(description="Render DOCX to PNG pages for visual QA")
    parser.add_argument("input_docx", help="Path to input .docx")
    parser.add_argument("output_dir", help="Directory for output PNGs and PDF")
    parser.add_argument("--dpi", type=int, default=150, help="DPI for PNG rendering")
    args = parser.parse_args()

    docx_path = Path(args.input_docx).resolve()
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / (docx_path.stem + ".pdf")

    engine = export_docx_to_pdf(docx_path, pdf_path)
    pngs = render_pdf_to_pngs(pdf_path, out_dir, dpi=args.dpi)
    print(f"Rendered {len(pngs)} page(s) via {engine}:")
    for p in pngs:
        print(str(p))


if __name__ == "__main__":
    main()
