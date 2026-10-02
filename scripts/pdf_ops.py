#!/usr/bin/env python3
"""PDF：合并/拆分/取文本/取图/转图/OCR/转 Word。

两条路各管一段，别混：
- **PyMuPDF / poppler / pypdf**：读、拆、合、渲染、取图 —— 快、不损失。
- **pdf2docx**：PDF → 可编辑 Word。LO 自己干不了这件事（`pdf → docx` 会报 `no export filter`）。
  保真度有限：实测简单表格常被还原成文本行，复杂版式必须人工核。
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import convert as lo


def _need(cmd: str) -> str:
    p = shutil.which(cmd)
    if not p:
        raise RuntimeError(f"缺少命令：{cmd}")
    return p


def info(path: str | Path) -> dict:
    import pymupdf
    p = Path(path).expanduser().resolve()
    with pymupdf.open(str(p)) as doc:
        first = doc[0]
        return {
            "path": str(p), "pages": doc.page_count, "encrypted": bool(doc.is_encrypted),
            "size_pt": [round(first.rect.width, 1), round(first.rect.height, 1)],
            "metadata": {k: v for k, v in (doc.metadata or {}).items() if v},
            "has_text_layer": bool(first.get_text().strip()),
        }


def merge(paths: list[str | Path], out: str | Path, *, bookmarks: bool = True) -> Path:
    from pypdf import PdfWriter
    w = PdfWriter()
    for f in paths:
        f = Path(f).expanduser().resolve()
        w.append(str(f))
        if bookmarks:
            w.add_outline_item(f.stem, len(w.pages) - 1)
    out = Path(out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("wb") as fh:
        w.write(fh)
    return out


def split(path: str | Path, out_dir: str | Path, *, pages_per_file: int = 1,
          ranges: list[tuple[int, int]] | None = None) -> list[Path]:
    """按每 N 页切，或按 1-based 闭区间 [(1,3),(4,4)] 切。"""
    from pypdf import PdfReader, PdfWriter
    src = Path(path).expanduser().resolve()
    out_dir = Path(out_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(src))
    n = len(reader.pages)
    if ranges:
        chunks = [(a, b) for a, b in ranges]
    else:
        chunks = [(i, min(i + pages_per_file - 1, n)) for i in range(1, n + 1, pages_per_file)]
    made: list[Path] = []
    for a, b in chunks:
        w = PdfWriter()
        for i in range(a - 1, min(b, n)):
            w.add_page(reader.pages[i])
        out = out_dir / f"{src.stem}_p{a}-{b}.pdf"
        with out.open("wb") as fh:
            w.write(fh)
        made.append(out)
    return made


def to_images(path: str | Path, out_dir: str | Path, *, dpi: int = 110,
              pages: tuple[int, int] | None = None, fmt: str = "png") -> list[Path]:
    """用 poppler 渲染页面为图片（缩略图/目视核对都靠它）。"""
    src = Path(path).expanduser().resolve()
    out_dir = Path(out_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [_need("pdftoppm"), "-r", str(dpi), f"-{fmt}"]
    if pages:
        cmd += ["-f", str(pages[0]), "-l", str(pages[1])]
    cmd += [str(src), str(out_dir / src.stem)]
    subprocess.run(cmd, check=True, capture_output=True, timeout=300)
    return sorted(out_dir.glob(f"{src.stem}*.{fmt}"))


def extract_images(path: str | Path, out_dir: str | Path, *, min_size: int = 64) -> list[Path]:
    import pymupdf
    src = Path(path).expanduser().resolve()
    out_dir = Path(out_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    made: list[Path] = []
    with pymupdf.open(str(src)) as doc:
        for pno in range(doc.page_count):
            for i, img in enumerate(doc[pno].get_images(full=True)):
                xref = img[0]
                pix = pymupdf.Pixmap(doc, xref)
                if pix.width < min_size or pix.height < min_size:
                    continue
                if pix.n - pix.alpha >= 4:
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                out = out_dir / f"{src.stem}_p{pno + 1}_{i + 1}.png"
                pix.save(str(out))
                made.append(out)
    return made


def ocr(path: str | Path, *, lang: str = "chi_sim+eng", dpi: int = 200,
        pages: tuple[int, int] | None = None) -> str:
    """扫描件/图片型 PDF 取字。先渲染再 tesseract，中文用 chi_sim。"""
    import tempfile

    tess = _need("tesseract")
    src = Path(path).expanduser().resolve()
    tmp = Path(tempfile.mkdtemp(prefix="office-docs-ocr-"))
    try:
        imgs = to_images(src, tmp, dpi=dpi, pages=pages)
        out: list[str] = []
        for img in imgs:
            proc = subprocess.run([tess, str(img), "stdout", "-l", lang],
                                  capture_output=True, text=True, timeout=300)
            out.append(proc.stdout)
        return "\n".join(out)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def to_docx(path: str | Path, out: str | Path | None = None, *, start: int = 0,
            end: int | None = None) -> Path:
    """PDF → 可编辑 docx（pdf2docx）。注意：表格/多栏保真度有限，务必人工核。"""
    Converter = lo.quiet_import("pdf2docx").Converter
    src = Path(path).expanduser().resolve()
    out = Path(out).expanduser() if out else src.with_suffix(".docx")
    out.parent.mkdir(parents=True, exist_ok=True)
    cv = Converter(str(src))
    try:
        cv.convert(str(out), start=start, end=end)
    finally:
        cv.close()
    if not out.exists():
        raise RuntimeError(f"pdf2docx 没产出文件：{out}")
    return out
