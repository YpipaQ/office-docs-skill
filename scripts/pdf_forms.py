#!/usr/bin/env python3
"""非表单域 PDF 的填写：抽结构 → 自检框位 → 按框写字 → 出校验图。

和 `pdf-forms`（AcroForm 真表单域）的分工：
- 文件里**有**表单域（AcroForm） → 用 `pdf-forms --fill`，字段名直接写。
- 文件是**扫描件/静态排版**（中国绝大多数表格都是这种，框是画上去的线） → 用本模块：
  先在 PDF 里抽出文字标签、横线、竖线、勾选框（`pdf-structure`），
  据此写出"哪个框填什么"的 JSON，再由 `pdf-fill` 逐框写字。

为什么不能"看着图估坐标"：出图是栅格，估出来的像素坐标换算回 PDF 点会带 1-2% 的
系统偏差，一页里能差出半个字。PDF 里的线是**矢量**，`get_drawings()` 给的坐标是精确值，
所以优先用结构，视觉只作兜底。

自检（这一步别省）：
- `check_boxes()` 在填之前就查「框太小装不下」「框与框/标签重叠」「框跑到页外」；
- 填完看 `insert_textbox` 的返回值：**负数就是溢出**，直接报出来；
- `--preview` 把每个框画成红框、标签框画成蓝框渲染成 PNG —— 给人（或模型）看一眼，
  比读坐标快得多。
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

CM2PT = 72.0 / 2.54

ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
CJK_FONT = "china-s"      # PyMuPDF 内置简体中文字体（宋）；无需外部字体文件


def _pt(v: float, unit: str) -> float:
    return v * CM2PT if unit == "cm" else float(v)


def to_rect(box, unit: str = "cm"):
    import pymupdf
    x0, y0, x1, y1 = (_pt(v, unit) for v in box)
    return pymupdf.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))


def load_spec(spec) -> dict:
    if isinstance(spec, (str, Path)):
        return json.loads(Path(spec).read_text(encoding="utf-8"))
    return spec


# ---------- 抽结构 ----------

def _lines_and_boxes(page):
    """从矢量绘图里取横线/竖线/矩形。返回 (hlines, vlines, boxes)，都是 point 坐标。"""
    hl, vl, bx = [], [], []

    def add_line(p1, p2):
        if abs(p1.y - p2.y) <= 1.5 and abs(p1.x - p2.x) > 20:
            hl.append([min(p1.x, p2.x), (p1.y + p2.y) / 2, max(p1.x, p2.x)])
        elif abs(p1.x - p2.x) <= 1.5 and abs(p1.y - p2.y) > 12:
            vl.append([(p1.x + p2.x) / 2, min(p1.y, p2.y), max(p1.y, p2.y)])

    for d in page.get_drawings():
        for item in d.get("items", []):
            kind = item[0]
            if kind == "l":
                add_line(item[1], item[2])
            elif kind == "re":
                r = item[1]
                bx.append([r.x0, r.y0, r.x1, r.y1])
            elif kind == "qu":
                # 四边形（有些表格线是填充路径画的）
                pts = item[1]
                for i in range(len(pts) - 1):
                    add_line(pts[i], pts[i + 1])
        r = d.get("rect")
        if r is not None and d.get("fill") is not None and (r.width < 30 and r.height < 30):
            bx.append([r.x0, r.y0, r.x1, r.y1])
    return _merge(hl, axis=1), _merge(vl, axis=0), bx


def _merge(lines, axis: int, tol: float = 2.0):
    """把同一条线拆成的多段并起来（axis=1 横线按 y 归一，axis=0 竖线按 x 归一）。"""
    out = []
    for ln in sorted(lines, key=lambda l: (round(l[axis]), l[1 - axis] if axis == 1 else l[1])):
        for m in out:
            if abs(m[axis] - ln[axis]) <= tol and ln[1] <= m[2] + tol and m[1] <= ln[2] + tol:
                m[1], m[2] = min(m[1], ln[1]), max(m[2], ln[2])
                break
        else:
            out.append(list(ln))
    return out


def _row_bands(hlines, min_h=6.0, max_h=80.0):
    """相邻两条横线之间的"行"（表单填写区的基本单位）。"""
    ys = sorted({round(l[1], 1) for l in hlines})
    bands = []
    for a, b in zip(ys, ys[1:]):
        if min_h <= b - a <= max_h:
            bands.append({"top": a, "bottom": b, "height": round(b - a, 1)})
    return bands


def _checkboxes(boxes, labels):
    out = []
    for x0, y0, x1, y1 in boxes:
        w, h = x1 - x0, y1 - y0
        if 3 <= w <= 28 and 3 <= h <= 28 and 0.55 <= (w / h if h else 0) <= 1.8:
            out.append({"bbox": [round(v, 1) for v in (x0, y0, x1, y1)],
                        "center": [round((x0 + x1) / 2, 1), round((y0 + y1) / 2, 1)],
                        "source": "vector", "size": [round(w, 1), round(h, 1)]})
    for lb in labels:                       # 中文表格常直接打成 □ / ☐ 字符
        if any(ch in lb["text"] for ch in "□☐▢◻"):
            out.append({"bbox": lb["bbox"], "center": [(lb["bbox"][0] + lb["bbox"][2]) / 2,
                                                       (lb["bbox"][1] + lb["bbox"][3]) / 2],
                        "source": "char", "label": lb["text"]})
    return out


def extract_structure(pdf, pages: list[int] | None = None) -> dict:
    import pymupdf
    doc = pymupdf.open(pdf)
    try:
        if doc.needs_pass:
            raise ValueError("PDF 有密码，先 pdf-decrypt")
        out = {"file": str(Path(pdf).name), "pages": []}
        for i, page in enumerate(doc):
            if pages and (i + 1) not in pages:
                continue
            labels = []
            for blk in page.get_text("dict").get("blocks", []):
                if blk.get("type") != 0:
                    continue
                for line in blk.get("lines", []):
                    text = "".join(s.get("text", "") for s in line.get("spans", [])).strip()
                    if not text:
                        continue
                    sizes = [s.get("size", 0) for s in line.get("spans", [])]
                    labels.append({"text": text,
                                   "bbox": [round(v, 1) for v in line["bbox"]],
                                   "size": round(max(sizes) if sizes else 0, 1)})
            hlines, vlines, boxes = _lines_and_boxes(page)
            labels.sort(key=lambda l: (round(l["bbox"][1], 1), l["bbox"][0]))
            out["pages"].append({
                "page": i + 1,
                "size_pt": [round(page.rect.width, 1), round(page.rect.height, 1)],
                "size_cm": [round(page.rect.width / CM2PT, 2), round(page.rect.height / CM2PT, 2)],
                "labels": labels,
                "h_lines": [[round(v, 1) for v in l] for l in sorted(hlines, key=lambda l: l[1])],
                "v_lines": [[round(v, 1) for v in l] for l in sorted(vlines, key=lambda l: l[0])],
                "checkboxes": _checkboxes(boxes, labels),
                "row_bands": _row_bands(hlines),
            })
        out["note"] = ("坐标单位是 point（PDF 原点在左上角，向下为 y+）。"
                       "把要填的框写进 spec 的 fields[].box 再跑 pdf-fill；"
                       "本机表格用 cm 更顺手：spec 里写 \"unit\": \"cm\"。")
        return out
    finally:
        doc.close()


# ---------- 自检 ----------

_MEASURE_CACHE: dict = {}


def _measure(text: str, font_size: float, box_w: float, fontname: str = CJK_FONT) -> dict:
    """把文字在「和真填时同一个引擎」里排一遍，量出行数与实际高度（point）。

    这是本模块最贵的一课，试了三种量法：
    1. 按"中文 1em、西文 0.55em"估 —— 错；
    2. `pymupdf.Font("china-s").text_length()` —— **也错**：它给 6.06pt/字符，
       而 insert_textbox 真正用的那个 Heiti CID 字体是 **11pt/字符（全角）**，
       18 位身份证号实际 198pt、量出来 109pt，于是自检放行、填下去换行溢出；
    3. 干脆**空跑一次排版**（同字体同字号同框宽），数它排了几行、占多高 —— 精确，就这么办。
    """
    key = (text, round(font_size, 2), round(box_w, 2), fontname)
    if key in _MEASURE_CACHE:
        return _MEASURE_CACHE[key]
    try:
        import pymupdf
        doc = pymupdf.open()
        page = doc.new_page(width=max(box_w, 10) + 8, height=14000)
        rect = pymupdf.Rect(2, 2, 2 + max(box_w, 1), 13998)
        page.insert_textbox(rect, text, fontname=fontname, fontsize=font_size)
        boxes = [ln["bbox"] for blk in page.get_text("dict")["blocks"] for ln in blk.get("lines", [])]
        doc.close()
        if not boxes:
            raise ValueError("空跑没排出任何行")
        res = {"lines": len(boxes),
               "height": round(max(b[3] for b in boxes) - min(b[1] for b in boxes), 2),
               "width": round(max(b[2] - b[0] for b in boxes), 2)}
    except Exception:
        # 兜底：内置 CJK 字体一律按全角算
        width = font_size * len(text)
        lines = max(1, math.ceil(width / max(box_w - 2, 1)))
        res = {"lines": lines, "height": round(lines * font_size * 1.35, 2), "width": round(width, 2)}
    _MEASURE_CACHE[key] = res
    return res


def _line_height(font_size: float, fontname: str = CJK_FONT) -> float:
    """一行占的高度。量不出来时的兜底（正常走 `_measure`）。"""
    return font_size * 1.35


def fit_size(text: str, font_size: float, box_w: float, box_h: float,
             fontname: str = CJK_FONT, min_size: float = 5.0, step: float = 0.25,
             max_lines: int = 0) -> float:
    """`autofit` 用：在框里塞得下的前提下，取最大的那个字号。

    为什么需要它：内置中文字体是**全角**的，18 位身份证号在 11pt 下要 198pt 宽，
    普通表格的格子根本装不下 —— 这时要么换行（难看且可能出格），要么按框缩字号。

    `max_lines > 0` 时连行数一起卡（表格格子通常要**一行**放完）。
    """
    size = font_size
    while size > min_size:
        m = _measure(text, size, box_w, fontname)
        if m["height"] <= box_h and (max_lines <= 0 or m["lines"] <= max_lines):
            return round(size, 2)
        size -= step
    return round(min_size, 2)


def check_boxes(spec, *, tol: float = 0.5) -> dict:
    """填之前先查框：装不下、重叠、跑出页外。所有量都在 spec 自己的单位里（默认 cm）。"""
    spec = load_spec(spec)
    unit = spec.get("unit", "cm")
    pages = {p["page"]: p for p in spec.get("pages", [])}
    errors, warnings = [], []

    def ov(a, b) -> bool:
        return not (a[2] <= b[0] + tol or b[2] <= a[0] + tol or
                    a[3] <= b[1] + tol or b[3] <= a[1] + tol)

    fields = spec.get("fields", [])
    for f in fields:
        if "box" not in f:
            errors.append({"field": f.get("label") or f.get("text", "?"), "why": "缺 box"})
            continue
        x0, y0, x1, y1 = f["box"]
        w, h = abs(x1 - x0), abs(y1 - y0)
        fs = float(f.get("font_size", 12))
        u2pt = CM2PT if unit == "cm" else 1.0     # spec 单位 → point
        name = f.get("label") or f.get("text", "?")
        if w <= 0 or h <= 0:
            errors.append({"field": name, "why": "框宽/高为 0"})
            continue
        m = _measure(str(f.get("text", "")), fs, w * u2pt, f.get("font", CJK_FONT))
        need_h = m["height"] / u2pt               # 需要的高度，换回 spec 单位
        if f.get("autofit"):
            # 开了 autofit 就等于"字号可以缩"：只在缩到底仍装不下时才报错
            ml = int(f.get("max_lines", 0))
            fit = fit_size(str(f.get("text", "")), fs, w * u2pt, h * u2pt,
                           f.get("font", CJK_FONT), float(f.get("min_font_size", 5.0)),
                           max_lines=ml)
            fm = _measure(str(f.get("text", "")), fit, w * u2pt, f.get("font", CJK_FONT))
            if fm["height"] > h * u2pt + tol or (ml and fm["lines"] > ml):
                errors.append({"field": name,
                               "why": (f"开 autofit 缩到最小字号 {fit:g}pt 仍装不下"
                                       + (f"（要求 {ml} 行内放完，实际要 {fm['lines']} 行）" if ml else "")),
                               "text": f.get("text")})
            continue
        if h + tol < need_h:
            errors.append({"field": name, "why": f"框高 {h:.2f}{unit} 装不下 {m['lines']} 行 "
                                                 f"{fs:g}pt 字（实测需 {need_h:.2f}{unit}）",
                           "text": f.get("text"), "hint": "给它加 \"autofit\": true，或调大框/减小字号"})
        pg = pages.get(f.get("page", 1))
        if pg and "size_cm" in pg:
            pw, ph = (pg["size_cm"] if unit == "cm" else pg.get("size_pt"))
            if x0 < -tol or y0 < -tol or x1 > pw + tol or y1 > ph + tol:
                errors.append({"field": name, "why": f"框跑到页面外（页 {pw}×{ph}{unit}）"})

    for i, a in enumerate(fields):
        for b in fields[i + 1:]:
            if a.get("page", 1) != b.get("page", 1) or "box" not in a or "box" not in b:
                continue
            if ov(a["box"], b["box"]):
                errors.append({"field": f"{a.get('label')} ↔ {b.get('label')}",
                               "why": "两个填写框重叠，字会叠在一起"})
        lb = a.get("label_box")
        if lb and "box" in a and ov(lb, a["box"]):
            warnings.append({"field": a.get("label"), "why": "label_box 与填写框重叠："
                                                            "确认框位是否从标签右侧开始"})

    return {"ok": not errors, "errors": errors, "warnings": warnings,
            "fields": len(fields), "unit": unit}


# ---------- 填写 ----------

def fill(pdf, spec, out, *, preview: str | Path | None = None, dpi: int = 150,
         force: bool = False) -> dict:
    import pymupdf
    spec = load_spec(spec)
    unit = spec.get("unit", "cm")

    if not force:
        chk = check_boxes(spec)
        if not chk["ok"]:
            return {"ok": False, "error": "框位自检没过，先修 spec（或 --force 强行填）", **chk}

    doc = pymupdf.open(pdf)
    try:
        if doc.needs_pass:
            return {"ok": False, "error": "PDF 有密码，先 pdf-decrypt"}
        overflow, placed = [], 0
        for f in spec.get("fields", []):
            if "box" not in f or f.get("text") in (None, ""):
                continue
            idx = int(f.get("page", 1)) - 1
            if not (0 <= idx < doc.page_count):
                overflow.append({"field": f.get("label"), "why": f"没有第 {idx + 1} 页"})
                continue
            page = doc[idx]
            rect = to_rect(f["box"], unit)
            fs = float(f.get("font_size", 12))
            font = f.get("font", CJK_FONT)
            if f.get("autofit"):
                fs = fit_size(str(f["text"]), fs, rect.width, rect.height, font,
                              float(f.get("min_font_size", 5.0)),
                              max_lines=int(f.get("max_lines", 0)))
            valign = f.get("valign", "middle")
            if valign in ("middle", "bottom"):
                # 垂直居中/贴底：**整体下移**这个框，不要缩小它。
                # 踩过的坑：缩小后的框（32pt→23pt）会让 insert_textbox 误判成"要换行"，
                # 于是明明装得下也报溢出（实测 left=-8.3）。
                need = _measure(str(f["text"]), fs, rect.width, f.get("font", CJK_FONT))["height"]
                off = max(0.0, rect.height - need)
                off = off / 2 if valign == "middle" else off
                rect = pymupdf.Rect(rect.x0, rect.y0 + off, rect.x1, rect.y1 + off)
            color = f.get("color", [0, 0, 0.65])
            left = page.insert_textbox(
                rect, str(f["text"]), fontname=f.get("font", CJK_FONT), fontsize=fs,
                color=tuple(color), align=ALIGN.get(f.get("align", "left"), 0), overlay=True)
            placed += 1
            if left < 0:
                overflow.append({"field": f.get("label") or f.get("text"),
                                 "why": f"字放不下，超出约 {abs(left):.1f}pt（框高不足或换行太多）"})
        out = Path(out).expanduser()
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))

        res = {"ok": not overflow, "out": str(out), "fields_placed": placed,
               "overflow": overflow, "bytes": out.stat().st_size}
        if preview:
            res["preview"] = str(_preview(out, spec, unit, preview, dpi))
        if overflow:
            res["error"] = "有字段溢出：调大框、减小字号或缩短文字后重填"
        return res
    finally:
        doc.close()


def _preview(out_pdf, spec, unit, png_path, dpi: int = 150) -> Path:
    """把填写框画成红框、标签框画成蓝框，渲染成 PNG —— 一眼就能看出填歪没填歪。

    画在**副本**上，交付的 PDF 不会被污染。
    """
    import pymupdf
    doc = pymupdf.open(out_pdf)
    try:
        for f in spec.get("fields", []):
            idx = int(f.get("page", 1)) - 1
            if not (0 <= idx < doc.page_count):
                continue
            page = doc[idx]
            if f.get("label_box"):
                page.draw_rect(to_rect(f["label_box"], unit), color=(0, 0, 1), width=0.6)
            if "box" in f:
                page.draw_rect(to_rect(f["box"], unit), color=(1, 0, 0), width=0.8)
        png = Path(png_path).expanduser()
        png.parent.mkdir(parents=True, exist_ok=True)
        if doc.page_count == 1:
            doc[0].get_pixmap(dpi=dpi).save(str(png))
        else:                                   # 多页：一页一张，加 -p1 后缀
            for i in range(doc.page_count):
                doc[i].get_pixmap(dpi=dpi).save(str(png.with_name(f"{png.stem}-p{i + 1}{png.suffix}")))
        return png
    finally:
        doc.close()


def main(argv: list[str] | None = None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="pdf_forms.py", description="非表单域 PDF 的抽结构/自检/填写")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("structure")
    p.add_argument("pdf"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--pages", help="如 1-3")
    p = sub.add_parser("check")
    p.add_argument("spec")
    p = sub.add_parser("fill")
    p.add_argument("pdf"); p.add_argument("spec"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--preview", help="把框位画到图上再看一眼（PNG 路径）")
    p.add_argument("--dpi", type=int, default=150)
    p.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "structure":
            pages = None
            if a.pages:
                pages = (list(range(int(a.pages.split("-")[0]), int(a.pages.split("-")[1]) + 1))
                         if "-" in a.pages else [int(x) for x in a.pages.split(",")])
            r = extract_structure(a.pdf, pages)
            Path(a.out).write_text(json.dumps(r, ensure_ascii=False, indent=1), encoding="utf-8")
            print(json.dumps({"ok": True, "out": a.out, "pages": len(r["pages"]),
                              "labels": sum(len(p["labels"]) for p in r["pages"]),
                              "checkboxes": sum(len(p["checkboxes"]) for p in r["pages"])},
                             ensure_ascii=False))
            return 0
        if a.cmd == "check":
            r = check_boxes(a.spec)
            print(json.dumps(r, ensure_ascii=False, indent=1))
            return 0 if r["ok"] else 3
        r = fill(a.pdf, a.spec, a.out, preview=a.preview, dpi=a.dpi, force=a.force)
        print(json.dumps(r, ensure_ascii=False, indent=1))
        return 0 if r.get("ok") else 3
    except Exception as exc:
        print(json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
