#!/usr/bin/env python3
"""office-docs-skill 统一入口。

成功判据（沿用本工作区口径）：**退出码 0 且 stdout 是 `{"ok": true, ...}`**。
所有子命令都支持 `--json` 输出（其实默认就是 JSON），便于脚本/智能体直接解析。

用法速查：
  office_cli.py doctor
  office_cli.py info  a.docx
  office_cli.py text  a.pdf
  office_cli.py convert a.docx --to pdf
  office_cli.py docx-new --spec s.json -o out.docx
  office_cli.py docx-replace a.docx --set 旧词=新词
  office_cli.py xlsx-new --spec s.json -o out.xlsx
  office_cli.py pptx-new --spec s.json -o out.pptx
  office_cli.py table a.xlsx -o out.csv
  office_cli.py thumb a.pptx -o thumbs/
  office_cli.py pdf-merge a.pdf b.pdf -o all.pdf
  office_cli.py pdf-split a.pdf -o parts/ --every 2
  office_cli.py pdf-ocr scan.pdf -l chi_sim+eng
  office_cli.py pdf-to-docx a.pdf -o a.docx
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import cn_profiles           # noqa: E402
import convert as lo          # noqa: E402
import extras                 # noqa: E402
import pdf_ops                # noqa: E402
import readers                # noqa: E402
import writers                # noqa: E402


def emit(obj: dict, code: int = 0) -> int:
    obj.setdefault("ok", code == 0)
    print(json.dumps(obj, ensure_ascii=False, indent=1))
    return code


def fail(msg: str, code: int = 2) -> int:
    print(json.dumps({"ok": False, "error": msg}, ensure_ascii=False, indent=1))
    return code


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="office_cli.py", description="文档读写/转换/生成（LibreOffice + Python）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("doctor", help="环境自检（工具/库/中文字体）")

    p = sub.add_parser("info", help="看文件真实类型与结构")
    p.add_argument("file")

    p = sub.add_parser("text", help="抽文本")
    p.add_argument("file")
    p.add_argument("-o", "--out", help="写入文件而不是打印")
    p.add_argument("--max", type=int, default=None, help="最多取这么多字符")

    p = sub.add_parser("table", help="表格文件 -> UTF-8 CSV")
    p.add_argument("file")
    p.add_argument("-o", "--out", required=True)

    p = sub.add_parser("convert", help="LibreOffice 转换（并发安全）")
    p.add_argument("file")
    p.add_argument("--to", required=True, help="目标扩展名，如 pdf/docx/xlsx/csv/pptx")
    p.add_argument("-o", "--outdir")
    p.add_argument("--filter", dest="filter_name", help="强制指定 LO 过滤器名（少见）")
    p.add_argument("--timeout", type=int, default=lo.DEFAULT_TIMEOUT)

    for name in ("docx-new", "xlsx-new", "pptx-new"):
        q = sub.add_parser(name, help=f"按 JSON 规格生成 {name.split('-')[0]}")
        q.add_argument("--spec", required=True, help="规格 JSON 路径，或用 - 从 stdin 读")
        q.add_argument("-o", "--out", required=True)
        if name == "docx-new":
            q.add_argument("--profile", choices=sorted(cn_profiles.PROFILES),
                           help="中文文书版式：gongwen 公文(GB/T 9704) / contract 合同 / report 报告")
        if name == "pptx-new":
            q.add_argument("--profile", choices=["brief"], help="汇报 PPT 版式（16:9、标题32pt、正文18pt）")
        if name in ("docx-new", "pptx-new"):
            q.add_argument("--portable", action="store_true",
                           help="字体名换成本机可用的 Noto CJK（只为本机出图；不加则保留国标字体名）")

    p = sub.add_parser("docx-replace", help="docx 查找替换（正文/表格/文本框/页眉页脚；跨 run 也命中）")
    p.add_argument("file")
    p.add_argument("--set", action="append", required=True, metavar="旧=新")
    p.add_argument("-o", "--out")
    p.add_argument("--no-merge", action="store_true",
                   help="不先合并 run 碎片（默认先合并：Word 常把一句话拆成多个 run，逐 run 找词会漏）")

    p = sub.add_parser("docx-merge-runs", help="把 docx 里被拆碎的 run 合并后另存（改前预处理）")
    p.add_argument("file")
    p.add_argument("-o", "--out")

    p = sub.add_parser("thumb", help="任意文档 -> PNG 缩略图（docx/xlsx/pptx/pdf 均可）")
    p.add_argument("file")
    p.add_argument("-o", "--outdir", required=True)
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--pages", help="如 1-3")

    p = sub.add_parser("pdf-info")
    p.add_argument("file")

    p = sub.add_parser("pdf-merge")
    p.add_argument("files", nargs="+")
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--no-bookmarks", action="store_true")

    p = sub.add_parser("pdf-split")
    p.add_argument("file")
    p.add_argument("-o", "--outdir", required=True)
    p.add_argument("--every", type=int, default=1, help="每 N 页一个文件")
    p.add_argument("--ranges", help="如 1-3,4-4（与 --every 二选一）")

    p = sub.add_parser("pdf-images")
    p.add_argument("file")
    p.add_argument("-o", "--outdir", required=True)

    p = sub.add_parser("pdf-ocr")
    p.add_argument("file")
    p.add_argument("-l", "--lang", default="chi_sim+eng")
    p.add_argument("--dpi", type=int, default=200)
    p.add_argument("--pages", help="如 1-2")
    p.add_argument("-o", "--out")

    p = sub.add_parser("pdf-to-docx")
    p.add_argument("file")
    p.add_argument("-o", "--out")
    p.add_argument("--start", type=int, default=0)
    p.add_argument("--end", type=int, default=None)

    p = sub.add_parser("pdf-rotate", help="旋转整份或指定页（角度为 90 的倍数）")
    p.add_argument("file")
    p.add_argument("--angle", type=int, required=True)
    p.add_argument("--pages", help="如 1-3 或 1,3（默认全篇）")
    p.add_argument("-o", "--out")

    p = sub.add_parser("pdf-watermark", help="盖文字水印（支持中文，走 PyMuPDF 内置 CJK 字体）")
    p.add_argument("file")
    p.add_argument("--text", required=True)
    p.add_argument("--angle", type=int, default=45)
    p.add_argument("--size", type=int, default=60)
    p.add_argument("--opacity", type=float, default=0.12)
    p.add_argument("-o", "--out")

    p = sub.add_parser("pdf-encrypt", help="加密（默认 AES-256）")
    p.add_argument("file")
    p.add_argument("--password", required=True)
    p.add_argument("--owner-password")
    p.add_argument("-o", "--out")

    p = sub.add_parser("pdf-decrypt", help="解密")
    p.add_argument("file")
    p.add_argument("--password", required=True)
    p.add_argument("-o", "--out")

    p = sub.add_parser("pdf-forms", help="列表单字段；给了 --fill 就填")
    p.add_argument("file")
    p.add_argument("--fill", action="append", metavar="字段=值")
    p.add_argument("-o", "--out")

    p = sub.add_parser("ooxml-check", help="OOXML 包体检：交付前确认 Word/Excel/PPT 能正常打开")
    p.add_argument("file")
    p.add_argument("--quick", action="store_true", help="跳过 r:id 断链检查")

    p = sub.add_parser("pdf-structure", help="抽静态 PDF 的表单结构（标签/横竖线/勾选框/行）→ JSON")
    p.add_argument("file")
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--pages", help="如 1-3 或 1,3")

    p = sub.add_parser("pdf-boxcheck", help="填之前自检框位：装不下/重叠/出页（给 pdf-fill 的 spec）")
    p.add_argument("spec")

    p = sub.add_parser("pdf-fill", help="往静态 PDF 的指定框里写字（非 AcroForm 表格）")
    p.add_argument("file")
    p.add_argument("spec", help="JSON：{\"unit\":\"cm\",\"fields\":[{page,box,text,font_size,...}]}")
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--preview", help="把框位画到图上输出 PNG（红=填写框，蓝=标签框）")
    p.add_argument("--dpi", type=int, default=150)
    p.add_argument("--force", action="store_true", help="自检不过也强行填")

    p = sub.add_parser("xlsx-recalc", help="用 LibreOffice 把公式算一遍并写回（openpyxl 只写公式不算值）")
    p.add_argument("file")
    p.add_argument("--timeout", type=int, default=60)
    p.add_argument("--force", action="store_true", help="即使会毁掉外链/VBA 也照算")

    p = sub.add_parser("xlsx-chart", help="给工作表加图表")
    p.add_argument("file")
    p.add_argument("--type", required=True, choices=["bar", "line", "pie"])
    p.add_argument("--data", required=True, help="如 Sheet1!B2:B5")
    p.add_argument("--cats", help="分类轴，如 Sheet1!A2:A5")
    p.add_argument("--title")
    p.add_argument("--anchor", default="E2")
    p.add_argument("-o", "--out")

    p = sub.add_parser("docx-review", help="盘点 Word 修订/批注（只读）")
    p.add_argument("file")

    # ── 表单模板逆向（form_template.py）───────────────────────────────
    p = sub.add_parser("form-learn", help="从一张表单照片/扫描件学出模板 profile（网格线＋OCR 文字块）")
    p.add_argument("image")
    p.add_argument("-o", "--out", required=True, help="输出 profile JSON；可人工修订")
    p = sub.add_parser("form-render", help="按规格渲染**干净空白模板** docx（规格 JSON 见 templates/）")
    p.add_argument("spec")
    p.add_argument("-o", "--out", required=True)
    p = sub.add_parser("form-verify", help="渲染→出 PDF→量表格线位→和 spec.verify 比（出图自检）")
    p.add_argument("spec")
    p.add_argument("-o", "--out", help="docx 输出路径（默认 spec 同名 .verify.docx）")
    p.add_argument("--keep-pdf", help="PDF 落到指定路径")
    p.add_argument("--tol", type=float, help="覆盖 spec 里的 tol_cm")
    p = sub.add_parser("form-blank", help="用 profile 从扫描件出【去掉后填字迹】的空白（栅格）")
    p.add_argument("image"); p.add_argument("profile"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--template", help="重排模板的渲染图（PNG，同 profile 页面尺寸）：比 OCR 块可靠")
    p = sub.add_parser("form-extract", help="用 profile 从扫描件切出**只含后填字迹**的层")
    p.add_argument("image"); p.add_argument("profile"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--tight", action="store_true", help="裁掉空白")
    p.add_argument("--template", help="重排模板的渲染图（PNG，同 profile 页面尺寸）")

    a = ap.parse_args(argv)

    def page_range(s: str | None) -> tuple[int, int] | None:
        if not s:
            return None
        lo_, hi = (s.split("-") + [s])[:2] if "-" in s else (s, s)
        return int(lo_), int(hi)

    def page_list(s: str | None) -> list[int] | None:
        if not s:
            return None
        if "-" in s:
            a_, b_ = s.split("-")[:2]
            return list(range(int(a_), int(b_) + 1))
        return [int(x) for x in s.split(",") if x.strip()]

    def spec_path(s: str) -> dict:
        return json.loads(sys.stdin.read()) if s == "-" else json.loads(Path(s).read_text(encoding="utf-8"))

    try:
        if a.cmd == "doctor":
            d = lo.doctor()
            return emit(d, 0 if d["ok"] else 1)

        if a.cmd == "info":
            return emit(readers.info_of(a.file))

        if a.cmd == "text":
            txt = readers.text_of(a.file, max_chars=a.max)
            if a.out:
                Path(a.out).write_text(txt, encoding="utf-8")
                return emit({"out": a.out, "chars": len(txt)})
            print(txt)
            return 0

        if a.cmd == "table":
            rows = readers.read_table(a.file)
            out = readers.write_csv(rows, a.out)
            return emit({"out": str(out), "rows": len(rows), "cols": len(rows[0]) if rows else 0})

        if a.cmd == "convert":
            out = lo.convert(a.file, a.to, a.outdir, filter_name=a.filter_name, timeout=a.timeout)
            return emit({"out": str(out), "bytes": out.stat().st_size})

        if a.cmd in {"docx-new", "xlsx-new", "pptx-new"}:
            spec = spec_path(a.spec)
            prof = getattr(a, "profile", None)
            if a.cmd == "docx-new" and prof:
                out = cn_profiles.build_docx(spec, a.out, prof, portable=a.portable)
            elif a.cmd == "pptx-new" and prof:
                out = cn_profiles.build_pptx(spec, a.out, prof, portable=a.portable)
            else:
                fn = {"docx-new": writers.new_docx, "xlsx-new": writers.new_xlsx,
                      "pptx-new": writers.new_pptx}[a.cmd]
                out = fn(spec, a.out)
            return emit({"out": str(out), "bytes": out.stat().st_size, "profile": prof,
                         "portable": bool(getattr(a, "portable", False))})

        if a.cmd == "docx-replace":
            mapping: dict[str, str] = {}
            for kv in a.set:
                if "=" not in kv:
                    return fail(f"--set 需要 旧=新 形式：{kv!r}")
                k, v = kv.split("=", 1)
                mapping[k] = v
            r = writers.docx_replace(a.file, mapping, a.out, merge=not a.no_merge)
            if not r.get("hits"):
                r["note"] = ("一处都没命中 —— 先 `office_cli.py text <file>` 看实际文字，"
                             "注意全角/半角、空格与换行符；run 碎片已自动合并，不是这个原因")
            return emit(r, 0 if r.get("hits") else 4)   # 一处没命中 = 没改到东西，不算成功

        if a.cmd == "docx-merge-runs":
            import docx_runs
            return emit(docx_runs.merge_runs_in_file(a.file, a.out))

        if a.cmd == "thumb":
            pdf = Path(a.file).expanduser()
            if readers.kind_of(pdf) != "pdf":
                pdf = lo.convert(pdf, "pdf", a.outdir)
            imgs = pdf_ops.to_images(pdf, a.outdir, dpi=a.dpi, pages=page_range(a.pages))
            return emit({"pdf": str(pdf), "images": [str(i) for i in imgs], "count": len(imgs)})

        if a.cmd == "pdf-info":
            return emit(pdf_ops.info(a.file))

        if a.cmd == "pdf-merge":
            out = pdf_ops.merge(a.files, a.out, bookmarks=not a.no_bookmarks)
            return emit({"out": str(out), "inputs": len(a.files)})

        if a.cmd == "pdf-split":
            ranges = None
            if a.ranges:
                ranges = []
                for part in a.ranges.split(","):
                    lo_, hi = part.split("-")
                    ranges.append((int(lo_), int(hi)))
            made = pdf_ops.split(a.file, a.outdir, pages_per_file=a.every, ranges=ranges)
            return emit({"outdir": a.outdir, "files": [str(x) for x in made], "count": len(made)})

        if a.cmd == "pdf-images":
            imgs = pdf_ops.extract_images(a.file, a.outdir)
            return emit({"images": [str(i) for i in imgs], "count": len(imgs)})

        if a.cmd == "pdf-ocr":
            txt = pdf_ops.ocr(a.file, lang=a.lang, dpi=a.dpi, pages=page_range(a.pages))
            if a.out:
                Path(a.out).write_text(txt, encoding="utf-8")
                return emit({"out": a.out, "chars": len(txt)})
            print(txt)
            return 0

        if a.cmd == "pdf-to-docx":
            out = pdf_ops.to_docx(a.file, a.out, start=a.start, end=a.end)
            return emit({"out": str(out), "bytes": out.stat().st_size,
                         "warning": "pdf2docx 保真度有限：表格/多栏常被还原成文本，务必人工核对"})

        if a.cmd == "pdf-rotate":
            out = extras.rotate(a.file, a.angle, a.out, pages=page_list(a.pages))
            return emit({"out": str(out), "angle": a.angle})

        if a.cmd == "pdf-watermark":
            out = extras.watermark(a.file, a.text, a.out, angle=a.angle,
                                   fontsize=a.size, opacity=a.opacity)
            return emit({"out": str(out), "text": a.text})

        if a.cmd == "pdf-encrypt":
            out = extras.encrypt(a.file, a.password, a.out, owner_password=a.owner_password)
            return emit({"out": str(out), "algorithm": "AES-256",
                         "warning": "密码只在这里用一次；本工具不保存、不记录密码"})

        if a.cmd == "pdf-decrypt":
            out = extras.decrypt(a.file, a.password, a.out)
            return emit({"out": str(out)})

        if a.cmd == "pdf-forms":
            if a.fill:
                values: dict[str, str] = {}
                for kv in a.fill:
                    if "=" not in kv:
                        return fail(f"--fill 需要 字段=值 形式：{kv!r}")
                    k, v = kv.split("=", 1)
                    values[k] = v
                return emit(extras.form_fill(a.file, values, a.out))
            return emit({"fields": extras.form_fields(a.file)})

        if a.cmd == "ooxml-check":
            import ooxml_check
            r = ooxml_check.check(a.file, deep=not a.quick)
            return emit(r, 0 if r.get("ok") else (1 if "error" in r else 3))

        if a.cmd in {"pdf-structure", "pdf-boxcheck", "pdf-fill"}:
            import pdf_forms as pf
            if a.cmd == "pdf-structure":
                pages = None
                if a.pages:
                    pages = (list(range(int(a.pages.split("-")[0]), int(a.pages.split("-")[1]) + 1))
                             if "-" in a.pages else [int(x) for x in a.pages.split(",") if x.strip()])
                r = pf.extract_structure(a.file, pages)
                Path(a.out).write_text(json.dumps(r, ensure_ascii=False, indent=1), encoding="utf-8")
                return emit({"out": a.out, "pages": len(r["pages"]),
                             "labels": sum(len(p["labels"]) for p in r["pages"]),
                             "h_lines": sum(len(p["h_lines"]) for p in r["pages"]),
                             "checkboxes": sum(len(p["checkboxes"]) for p in r["pages"]),
                             "note": "把要填的框写进 spec.fields[].box 再跑 pdf-fill"})
            if a.cmd == "pdf-boxcheck":
                r = pf.check_boxes(a.spec)
                return emit(r, 0 if r["ok"] else 3)
            r = pf.fill(a.file, a.spec, a.out, preview=a.preview, dpi=a.dpi, force=a.force)
            return emit(r, 0 if r.get("ok") else 3)

        if a.cmd == "xlsx-recalc":
            import xlsx_recalc
            r = xlsx_recalc.recalc(a.file, a.timeout, force=a.force)
            return emit(r, 0 if r.get("ok") else (1 if "error" in r else 3))

        if a.cmd == "xlsx-chart":
            out = extras.chart(a.file, a.type, a.data, a.cats, a.out,
                               title=a.title, anchor=a.anchor)
            return emit({"out": str(out), "has_chart": extras.has_chart(out)})

        if a.cmd == "docx-review":
            return emit(extras.review(a.file))

        if a.cmd.startswith("form-"):
            # cv2 / form_template 只在用得上时导入：其余子命令不依赖 OpenCV
            import cv2
            import form_template as ft
            if a.cmd == "form-learn":
                img, info = ft.rectify(cv2.imread(a.image))
                prof = ft.learn(img, already_rectified=True)
                prof["page"]["rectify"] = info
                return emit({"out": str(ft.save_profile(prof, a.out)),
                             "blocks": len(prof["blocks"]),
                             "grid_h": len(prof["grid"]["h"]), "grid_v": len(prof["grid"]["v"]),
                             "warning": "OCR 只当起点：文字块请人工校对后再用"})
            if a.cmd == "form-render":
                spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
                out = ft.render_spec(spec, a.out)
                return emit({"out": str(out), "bytes": out.stat().st_size})
            if a.cmd == "form-verify":
                spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
                if a.tol:
                    spec.setdefault("verify", {})["tol_cm"] = a.tol
                out = a.out or str(Path(a.spec).with_suffix(".verify.docx"))
                r = ft.verify_spec(spec, out, a.keep_pdf)
                # 布局没落到位就让信封 ok=false，别让调用方只看 docx 存在就当成功
                return emit(r, 0 if r.get("ok") else 3)
            if a.cmd == "form-blank":
                prof = ft.load_profile(a.profile)
                out = Path(a.out).expanduser(); out.parent.mkdir(parents=True, exist_ok=True)
                tpl = cv2.imread(a.template) if getattr(a, "template", None) else None
                cv2.imwrite(str(out), ft.blank_raster(ft.prepare(cv2.imread(a.image), prof), prof, template_img=tpl))
                return emit({"out": str(out)})
            if a.cmd == "form-extract":
                prof = ft.load_profile(a.profile)
                tpl = cv2.imread(a.template) if getattr(a, "template", None) else None
                layer = ft.to_layer(ft.split_ink(ft.prepare(cv2.imread(a.image), prof), prof, tpl)["handwriting"],
                                    "tight" if a.tight else None)
                out = Path(a.out).expanduser(); out.parent.mkdir(parents=True, exist_ok=True)
                cv2.imwrite(str(out), layer)
                return emit({"out": str(out)})

    except FileNotFoundError as exc:
        return fail(f"文件不存在：{exc}", 2)
    except Exception as exc:  # 统一成 ok=false，别让智能体去猜 traceback
        return fail(f"{type(exc).__name__}: {exc}", 1)

    return fail(f"未知子命令：{a.cmd}")


if __name__ == "__main__":
    sys.exit(main())
