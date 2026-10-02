#!/usr/bin/env python3
"""OOXML 包体检：交付前先确认这个 docx/xlsx/pptx 能被 Word/Excel/PPT 正常打开。

为什么需要：docx/xlsx/pptx 本质是 **ZIP ＋ 一堆 XML**。python-docx / openpyxl 写出来的文件
通常没问题，但只要动过 XML（解压改字、拼装、第三方工具产出），就可能出现下面这些症状：
文件照样打得开、本机 LibreOffice 也照渲染，**到 Word 里却弹"发现无法读取的内容，是否修复"**，
修完图片/超链接/批注不见了 —— 光看渲染图完全看不出来。

本工具查的都是"Word 会判为损坏"的那几类硬伤：
1. ZIP 本身破损 / 条目重复 / 路径穿越（`..`、绝对路径）或 symlink 条目（外部来的文件不可信）；
2. `[Content_Types].xml` 缺失、内容类型没盖住某个部件，或 Override 指向不存在的部件；
3. `.rels` 关系表里 Target 指不到实际存在的部件；
4. 部件里用了 `r:id` / `r:embed` / `r:link`，但对应关系**不存在**（图片/超链接断链的头号原因）；
5. 某个 XML 部件本身不是良构 XML。

外加：用对应的库**再打开一次**（docx → python-docx、xlsx → openpyxl、pptx → python-pptx），
能打开才算过。查的是"能不能开"，不是"排版对不对" —— 排版另见 `form-verify` / `thumb`。
"""
from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

REQUIRED = {
    "docx": ("word/document.xml",),
    "xlsx": ("xl/workbook.xml",),
    "pptx": ("ppt/presentation.xml",),
}
RID_ATTR_RE = re.compile(rb'\sr:(?:id|embed|link|pict|dm|lo|qs|cs|href)="([^"]+)"')


def _kind(names: set[str], ct: dict) -> str:
    if "word/document.xml" in names:
        return "docx"
    if "xl/workbook.xml" in names:
        return "xlsx"
    if "ppt/presentation.xml" in names:
        return "pptx"
    main = ct.get("overrides", {}).get("/word/document.xml")
    if main:
        return "docx"
    return "unknown"


def _parse_content_types(data: bytes) -> dict:
    import xml.etree.ElementTree as ET
    root = ET.fromstring(data)
    defaults, overrides = {}, {}
    for el in root:
        tag = el.tag.split("}")[-1]
        if tag == "Default":
            defaults[el.get("Extension", "").lower()] = el.get("ContentType", "")
        elif tag == "Override":
            overrides[el.get("PartName", "")] = el.get("ContentType", "")
    return {"defaults": defaults, "overrides": overrides}


def check(path: str | Path, *, deep: bool = True) -> dict:
    p = Path(path).expanduser().resolve()
    if not p.exists():
        return {"ok": False, "error": f"文件不存在：{p}"}

    errors: list[str] = []
    warnings: list[str] = []

    try:
        zf = zipfile.ZipFile(p)
    except zipfile.BadZipFile as exc:
        return {"ok": False, "error": f"不是有效的 ZIP/OOXML 包：{exc}", "file": str(p)}

    with zf:
        infos = zf.infolist()
        names = [i.filename for i in infos]

        # 1. ZIP 层
        bad = zf.testzip()
        if bad:
            errors.append(f"ZIP 内条目损坏（CRC 校验失败）：{bad}")
        seen: dict[str, int] = {}
        for i in infos:
            key = i.filename
            seen[key] = seen.get(key, 0) + 1
            if key.startswith("/") or key.startswith("\\"):
                errors.append(f"非法条目路径（绝对路径）：{key}")
            if ".." in key.replace("\\", "/").split("/"):
                errors.append(f"非法条目路径（路径穿越）：{key}")
            if (i.external_attr >> 16) & 0o170000 == 0o120000:
                errors.append(f"symlink 条目（外部文件不可信，应删除）：{key}")
        dupes = [k for k, n in seen.items() if n > 1]
        if dupes:
            errors.append(f"重复条目 {len(dupes)} 个：{dupes[:5]}")

        name_set = {n for n in names if not n.endswith("/")}

        # 2. Content_Types
        if "[Content_Types].xml" not in name_set:
            errors.append("缺 [Content_Types].xml（Word 会直接判损坏）")
            ct = {"defaults": {}, "overrides": {}}
        else:
            try:
                ct = _parse_content_types(zf.read("[Content_Types].xml"))
            except Exception as exc:
                errors.append(f"[Content_Types].xml 解析失败：{type(exc).__name__}: {exc}")
                ct = {"defaults": {}, "overrides": {}}

        def typed(part: str) -> bool:
            ext = part.rsplit(".", 1)[-1].lower() if "." in part else ""
            return part in ct["overrides"] or ext in ct["defaults"]

        untitled = [n for n in name_set
                    if n != "[Content_Types].xml" and "." in n and not typed(n)]
        # 图片/媒体没有显式 Default 时最常见，单独提示
        if untitled:
            errors.append(f"{len(untitled)} 个部件没有内容类型（Word 打不开这些部件）：{untitled[:5]}")

        for part in ct["overrides"]:
            if part.lstrip("/") not in name_set:
                errors.append(f"Content_Types 里的 Override 指向不存在的部件：{part}")

        kind = _kind(name_set, ct)
        if kind == "unknown":
            errors.append("认不出是 docx/xlsx/pptx（缺 word/document.xml、xl/workbook.xml、"
                          "ppt/presentation.xml 三选一）")
        else:
            for req in REQUIRED[kind]:
                if req not in name_set:
                    errors.append(f"{kind} 缺必需部件：{req}")
            if "_rels/.rels" not in name_set:
                warnings.append("缺 _rels/.rels（包级关系表）—— 有些阅读器会拒绝打开")

        # 3/4. 关系表
        rels_parts = [n for n in name_set if n.endswith(".rels")]
        if not rels_parts:
            warnings.append("包里一个 .rels 都没有，很不正常")
        rel_map: dict[str, set[str]] = {}
        for rp in rels_parts:
            base = posixpath.dirname(posixpath.dirname(rp))          # _rels 的上一级
            try:
                import xml.etree.ElementTree as ET
                root = ET.fromstring(zf.read(rp))
            except Exception as exc:
                errors.append(f"{rp} 解析失败：{type(exc).__name__}: {exc}")
                continue
            ids = set()
            for rel in root:
                rid, target, mode = rel.get("Id"), rel.get("Target"), rel.get("TargetMode")
                if rid:
                    ids.add(rid)
                if not target or mode == "External":
                    continue
                resolved = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
                if resolved not in name_set and target not in name_set:
                    errors.append(f"{rp} 的关系 {rid} 指向不存在的部件：{target}")
            rel_map[rp] = ids

        # 5. 部件内的 r:id 是否都有对应关系
        if deep:
            for part in name_set:
                if not part.endswith(".xml") or part.endswith(".rels"):
                    continue
                rp = posixpath.join(posixpath.dirname(part), "_rels",
                                    posixpath.basename(part) + ".rels")
                try:
                    data = zf.read(part)
                except Exception:
                    continue
                used = {m.decode("utf-8", "ignore") for m in RID_ATTR_RE.findall(data)}
                if not used:
                    continue
                have = rel_map.get(rp, set())
                missing = sorted(used - have)
                if missing:
                    errors.append(f"{part} 用到关系 {missing[:5]}，但 {rp} 里没有"
                                  "（图片/超链接断链，Word 会提示修复）")

        # 6. XML 良构
        xml_bad = []
        for n in name_set:
            if not (n.endswith(".xml") or n.endswith(".rels") or n.endswith(".vml")):
                continue
            try:
                import xml.etree.ElementTree as ET
                ET.fromstring(zf.read(n))
            except Exception as exc:
                xml_bad.append(f"{n}（{type(exc).__name__}: {exc}）")
        if xml_bad:
            errors.append(f"{len(xml_bad)} 个 XML 部件不是良构 XML：{xml_bad[:3]}")

        media = [n for n in name_set if n.startswith(("word/media/", "xl/media/", "ppt/media/"))]
        has_vba = any(n.lower().endswith("vbaproject.bin") for n in name_set)
        stats = {"parts": len(name_set), "xml_parts": sum(1 for n in name_set if n.endswith(".xml")),
                 "rels": len(rels_parts), "media": len(media), "vba": has_vba,
                 "encrypted_entries": sum(1 for i in infos if i.flag_bits & 0x1)}

    # 7. 用对应的库再打开一次
    lib_ok, lib_note = None, ""
    try:
        if kind == "docx":
            from docx import Document
            d = Document(str(p))
            lib_ok, lib_note = True, f"python-docx 打开：{len(d.paragraphs)} 段 / {len(d.tables)} 表"
        elif kind == "xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(p, read_only=True)
            lib_ok, lib_note = True, f"openpyxl 打开：{len(wb.sheetnames)} 个工作表"
            wb.close()
        elif kind == "pptx":
            from pptx import Presentation
            pr = Presentation(str(p))
            lib_ok, lib_note = True, f"python-pptx 打开：{len(pr.slides)} 页"
    except Exception as exc:
        lib_ok, lib_note = False, f"{type(exc).__name__}: {exc}"
        errors.append(f"用库打开失败：{lib_note}")

    return {
        "ok": not errors,
        "file": str(p),
        "kind": kind,
        "errors": errors,
        "warnings": warnings,
        "stats": stats,
        "lib": {"ok": lib_ok, "note": lib_note},
        "boundary": "本检查判的是「包结构能不能打开」，不判排版与内容对不对",
    }


def main(argv: list[str] | None = None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="ooxml_check.py", description="OOXML 包结构体检（docx/xlsx/pptx）")
    ap.add_argument("file")
    ap.add_argument("--quick", action="store_true", help="不查 r:id 断链（快一点）")
    a = ap.parse_args(argv)
    try:
        r = check(a.file, deep=not a.quick)
    except Exception as exc:
        r = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    print(json.dumps(r, ensure_ascii=False, indent=1))
    return 0 if r.get("ok") else (1 if "error" in r else 3)


if __name__ == "__main__":
    sys.exit(main())
