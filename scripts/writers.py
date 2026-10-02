#!/usr/bin/env python3
"""生成与修改：docx / xlsx / pptx。

设计取舍：**中文字体一律显式指定**。LO 在缺微软字体的机器上会做替换（宋体→Noto Sans CJK），
显式写 "Noto Sans CJK SC" 至少让渲染可预期，而不是碰运气等替换。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

CJK_FONT = "Noto Sans CJK SC"      # 本机实测存在；无微软字体时的中文字体基线
CJK_SERIF = "Noto Serif CJK SC"


# ---------- docx ----------

def _set_docx_run_font(run, name: str = CJK_FONT, size: float | None = None, bold: bool | None = None) -> None:
    from docx.oxml.ns import qn
    run.font.name = name
    try:
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
    except AttributeError:  # 老版本 python-docx 兜底
        rfonts = run._element.rPr.rFonts
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    if size is not None:
        from docx.shared import Pt
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def new_docx(spec: dict | str | Path, out: str | Path) -> Path:
    """按 JSON 规格建 Word。spec 可为 dict 或 json 路径。

    {"title": "标题", "font": "Noto Sans CJK SC",
     "blocks": [{"h": 1, "text": "章"}, {"p": "段落"}, {"bullets": ["a", "b"]},
                {"table": {"rows": [["甲","乙"],["1","2"]], "header": true}},
                {"pagebreak": true}]}
    """
    from docx import Document
    from docx.enum.text import WD_BREAK
    from docx.shared import Pt

    if isinstance(spec, (str, Path)):
        spec = json.loads(Path(spec).read_text(encoding="utf-8"))
    font = spec.get("font", CJK_FONT)
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(10.5)

    if spec.get("title"):
        doc.add_heading(spec["title"], level=0)

    for blk in spec.get("blocks", []):
        if "h" in blk:
            h = doc.add_heading("", level=int(blk["h"]))
            _set_docx_run_font(h.add_run(blk.get("text", "")), font, 16 - 2 * int(blk["h"]), True)
        elif "p" in blk:
            p = doc.add_paragraph()
            _set_docx_run_font(p.add_run(blk["p"]), font, blk.get("size", 10.5))
        elif "bullets" in blk:
            for item in blk["bullets"]:
                p = doc.add_paragraph(style="List Bullet")
                _set_docx_run_font(p.add_run(item), font, 10.5)
        elif "table" in blk:
            t = blk["table"]
            rows = t.get("rows", [])
            if not rows:
                continue
            table = doc.add_table(rows=len(rows), cols=len(rows[0]))
            table.style = "Table Grid"
            header = bool(t.get("header", True))
            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    cell = table.cell(i, j)
                    cell.text = ""
                    _set_docx_run_font(cell.paragraphs[0].add_run(str(val)), font, 10.5, header and i == 0)
        elif blk.get("pagebreak"):
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif "text" in blk:
            _set_docx_run_font(doc.add_paragraph().add_run(blk["text"]), font, 10.5)

    out = Path(out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out


def docx_replace(path: str | Path, mapping: dict[str, str], out: str | Path | None = None,
                 *, merge: bool = True) -> dict:
    """docx 查找替换，覆盖正文/表格/文本框/页眉页脚。

    踩过的坑：Word 把一句话拆成多个 `<w:r>`，逐 run 找词会**静默漏词**。所以这里先
    `docx_runs.merge_runs()` 合并碎片，再在**渲染文本**上替换 —— 跨 run 也能命中，
    且保留各 run 原有字体。细节与反例见 `scripts/docx_runs.py` 抬头。
    """
    _here = str(Path(__file__).resolve().parent)
    if _here not in sys.path:          # 直接从别处 import writers 时也要能找到同目录模块
        sys.path.insert(0, _here)
    import docx_runs
    return docx_runs.replace_in_file(path, mapping, out, merge=merge)


# ---------- xlsx ----------

def new_xlsx(spec: dict | str | Path, out: str | Path) -> Path:
    """{"sheets": [{"name": "S1", "rows": [[..],[..]], "header": true, "widths": [12, 20],
                    "freeze": "A2"}]}  单元格字符串以 = 开头会被当公式。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    if isinstance(spec, (str, Path)):
        spec = json.loads(Path(spec).read_text(encoding="utf-8"))
    wb = Workbook()
    wb.remove(wb.active)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    for s in spec.get("sheets", [{"name": "Sheet1", "rows": []}]):
        ws = wb.create_sheet(title=str(s.get("name", "Sheet1"))[:31])
        header = bool(s.get("header", True))
        for r_i, row in enumerate(s.get("rows", []), start=1):
            for c_i, val in enumerate(row, start=1):
                if isinstance(val, str) and val.startswith("="):
                    ws.cell(r_i, c_i, val)          # 公式原样写入，数值交给 LO 重算
                else:
                    ws.cell(r_i, c_i, val)
                if header and r_i == 1:
                    cell = ws.cell(r_i, c_i)
                    cell.font, cell.fill = head_font, head_fill
                    cell.alignment = Alignment(horizontal="center")
        for i, w in enumerate(s.get("widths", []), start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        if s.get("freeze"):
            ws.freeze_panes = s["freeze"]
    out = Path(out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return out


# ---------- pptx ----------

def _pptx_textbox(slide, left, top, width, height, *, title: bool, font: str):
    from pptx.util import Emu
    box = slide.shapes.add_textbox(Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    tf = box.text_frame
    tf.word_wrap = True
    return tf


def new_pptx(spec: dict | str | Path, out: str | Path) -> Path:
    """{"title": "封面", "size": "16:9", "font": "Noto Sans CJK SC",
        "slides": [{"title": "页1", "bullets": ["a","b"]},
                   {"title": "页2", "table": {"rows": [["甲","乙"],["1","2"]]}}]}
    全部用空白版式 + 文本框，字体完全由我们指定（不赌主题字体）。"""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    if isinstance(spec, (str, Path)):
        spec = json.loads(Path(spec).read_text(encoding="utf-8"))
    prs = Presentation()
    if spec.get("size", "16:9") == "16:9":
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    else:
        prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
    font = spec.get("font", CJK_FONT)
    # 字号可被规格覆盖（中文汇报 PPT 惯例：封面 32pt、页标题 28pt、正文 18pt）
    cover_sz = float(spec.get("title_size", 40))
    page_sz = float(spec.get("slide_title_size", 28))
    body_sz = float(spec.get("body_size", 18))
    table_sz = float(spec.get("table_size", 14))
    W, H = prs.slide_width, prs.slide_height

    def add_title_slide(title: str, subtitle: str = "") -> None:
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tf = _pptx_textbox(s, W * 0.1, H * 0.35, W * 0.8, H * 0.2, title=True, font=font)
        p = tf.paragraphs[0]
        r = p.add_run(); r.text = title
        r.font.size, r.font.bold, r.font.name = Pt(cover_sz), True, font
        if subtitle:
            tf2 = _pptx_textbox(s, W * 0.1, H * 0.55, W * 0.8, H * 0.15, title=False, font=font)
            r2 = tf2.paragraphs[0].add_run(); r2.text = subtitle
            r2.font.size, r2.font.name = Pt(body_sz), font

    if spec.get("title"):
        add_title_slide(spec["title"], spec.get("subtitle", ""))

    for sl in spec.get("slides", []):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        if sl.get("title"):
            tf = _pptx_textbox(s, W * 0.06, H * 0.06, W * 0.88, H * 0.16, title=True, font=font)
            r = tf.paragraphs[0].add_run(); r.text = sl["title"]
            r.font.size, r.font.bold, r.font.name = Pt(page_sz), True, font
        if sl.get("bullets"):
            tf = _pptx_textbox(s, W * 0.08, H * 0.28, W * 0.84, H * 0.62, title=False, font=font)
            for i, b in enumerate(sl["bullets"]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                r = p.add_run(); r.text = f"• {b}"
                r.font.size, r.font.name = Pt(body_sz), font
        if sl.get("table"):
            rows = sl["table"]["rows"]
            if rows:
                shape = s.shapes.add_table(len(rows), len(rows[0]),
                                           int(W * 0.08), int(H * 0.28), int(W * 0.84), int(H * 0.5))
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        cell = shape.table.cell(i, j)
                        cell.text = str(val)
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                r.font.size, r.font.name = Pt(table_sz), font
        if sl.get("notes"):
            s.notes_slide.notes_text_frame.text = sl["notes"]

    out = Path(out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return out
