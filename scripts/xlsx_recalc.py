#!/usr/bin/env python3
"""xlsx 公式重算 ＋ 错误闸门。

为什么必须有这一步（openpyxl 的硬伤）：
openpyxl 写公式只是把 `=SUM(B2:B9)` 当**字符串**塞进 XML，**不产生缓存值**。
于是任何按缓存值读的东西 —— pandas、`load_workbook(data_only=True)`、预览器、
以及 Word/PPT 里的链接 —— 读到的全是 `None`。文件本身没坏，但"算出来的数"是空的。

做法：让 LibreOffice 真的把每个公式算一遍，并把结果**存回文件**（宏 `calculateAll()` + `store()`），
然后逐个单元格扫 Excel 的 7 类错误串。公式能算出 `#NAME?` 通常意味着 LO 不支持这个函数
（见 `references/formats.md` 的"LibreOffice 不认哪些函数"）。

三个"看起来成功其实失败"的坑（都已固化）：
- LO 退出码 0 但**根本没重写文件**（例如已有另一个 LO 实例占着）⇒ 用 mtime+size 指纹判定。
- `errors_found` 也可能退出码 0 ⇒ **别看退出码，看 `status`**。
- 工作簿外链（`='[1]Sheet'!A1`）：openpyxl 存盘会剥掉缓存值，LO 再去解析就得到 `#NAME?`，
  且**外链被永久删除** ⇒ 默认拒绝重算，除非 `--force`。

本工具的成功判据沿用工作区口径：**退出码 0 且 `ok:true`**。发现公式错误 → `ok:false`，退出码 3。
"""
from __future__ import annotations

import contextlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import convert as lo

EXCEL_ERRORS = ("#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A", "#GETTING_DATA")
# LO 在中文界面下可能把错误串本地化后再存盘；两种都扫
LOCALIZED = ("#名称?", "#引用!", "#值!", "#除零!", "#数字!", "#空值!")
MAX_SHOWN = 100

# LibreOffice 的 Basic 宏：算全表 + 原地存盘 + 退出。除了这次调用，不会留在用户配置里。
# 踩过的坑：只 store() 不 terminate()，soffice 会**永远不退出**（实测挂满 45s 超时），
# 而文件其实已经写好了 —— 看起来像失败，其实只是进程没走。
MACRO_XBA = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub RecalcAllAndStore()
    ThisComponent.calculateAll()
    ThisComponent.store()
    StarDesktop.terminate()
End Sub
</script:module>
"""

SCRIPT_URL = ("vnd.sun.star.script:Standard.Module1.RecalcAllAndStore"
              "?language=Basic&location=application")

EXTERNAL_REF_RE = re.compile(r"\[\d+\]")


class RecalcError(RuntimeError):
    pass


def _stamp(path: Path) -> tuple[int, int]:
    st = path.stat()
    return st.st_mtime_ns, st.st_size


def _has_vba(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as z:
            return any(n.lower() == "xl/vbaproject.bin" for n in z.namelist())
    except (zipfile.BadZipFile, OSError):
        return False


def formulas_of(path: Path) -> tuple[int, dict[str, str]]:
    """返回 (公式格数, {坐标: 公式串})。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    n, cells = 0, {}
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("="):
                        n += 1
                        cells[f"{name}!{cell.coordinate}"] = v
    finally:
        wb.close()
    return n, cells


def external_links_at_risk(path: Path) -> list[str]:
    """找出「引用了外部工作簿、且缓存值已经丢了」的公式格 —— 重算会把它们变成 #NAME? 并删掉外链。"""
    try:
        with zipfile.ZipFile(path) as z:
            if not any(n.startswith("xl/externalLinks/") for n in z.namelist()):
                return []
    except (zipfile.BadZipFile, OSError):
        return []

    from openpyxl import load_workbook
    wb_f = load_workbook(path, data_only=False)
    wb_v = load_workbook(path, data_only=True)
    at_risk: list[str] = []
    try:
        names = [n for n, dn in wb_f.defined_names.items()
                 if isinstance(getattr(dn, "value", None), str)
                 and EXTERNAL_REF_RE.search(dn.value)]
        name_re = re.compile(r"\b(" + "|".join(re.escape(n) for n in names) + r")\b") if names else None
        for sheet in wb_f.sheetnames:
            ws, cached = wb_f[sheet], wb_v[sheet]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    if EXTERNAL_REF_RE.search(v) or (name_re and name_re.search(v)):
                        if cached[cell.coordinate].value is None:
                            at_risk.append(f"{sheet}!{cell.coordinate}")
    finally:
        wb_f.close()
        wb_v.close()
    return at_risk


def _kill_profile(profile: Path) -> None:
    """收掉挂着这个 profile 的 soffice 残留。

    踩过的坑：宏跑完不退出的那次超时，进程被 kill 后**孤儿 soffice.bin 还活着**，
    并且**占着那个 xlsx 的文件锁** —— 于是下一次重算在 headless 下卡在"文档已被占用"上，
    看起来像随机挂死。所以超时后必须按 profile 路径精确清场（不用 `pkill -x soffice.bin`，
    那会连用户自己开着的 LibreOffice 一起杀掉）。
    """
    with contextlib.suppress(Exception):
        subprocess.run(["pkill", "-f", str(profile)], capture_output=True, timeout=10)


def _run_macro(path: Path, timeout: int) -> dict:
    """起一个**独立 profile** 的 LO，跑重算宏。profile 用完即删，不污染用户配置。"""
    if lo.SOFFICE is None:
        raise RecalcError("找不到 soffice/libreoffice —— 重算必须靠 LibreOffice")
    profile = lo._profile_dir()
    hung, rc, tail = False, None, ""
    try:
        # 先让 LO 把 profile 造出来，再往里塞宏（顺序反了宏会被 profile 初始化覆盖）
        init = subprocess.run(
            [lo.SOFFICE, "--headless", "--norestore", "--nolockcheck", "--nologo",
             "--terminate_after_init", f"-env:UserInstallation=file://{profile}"],
            capture_output=True, text=True, timeout=timeout)
        macro_dir = profile / "user" / "basic" / "Standard"
        if not macro_dir.exists():
            raise RecalcError("LibreOffice 没能建出可用 profile，公式**没有**重算"
                              f"（exit={init.returncode}）")
        (macro_dir / "Module1.xba").write_text(MACRO_XBA, encoding="utf-8")

        cmd = [lo.SOFFICE, "--headless", "--norestore", "--nolockcheck", "--nologo",
               f"-env:UserInstallation=file://{profile}", SCRIPT_URL, str(path)]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            rc, tail = proc.returncode, (proc.stderr or proc.stdout or "").strip()
        except subprocess.TimeoutExpired:
            # 宏里的 store() 往往已经落盘了，只是进程没走 ⇒ 不立刻判死，让调用方看文件指纹
            hung = True
    finally:
        _kill_profile(profile)
        shutil.rmtree(profile, ignore_errors=True)

    if not hung and rc != 0:
        raise RecalcError(f"LibreOffice 重算失败（exit={rc}）：{' / '.join(tail.splitlines()[-3:])}")
    return {"hung": hung, "exit": rc}


def _scan_errors(path: Path) -> dict:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    found: dict[str, list[str]] = {}
    total = 0
    try:
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str):
                        for err in EXCEL_ERRORS + LOCALIZED:
                            if err in v:
                                found.setdefault(err, []).append(f"{sheet}!{cell.coordinate}")
                                total += 1
                                break
    finally:
        wb.close()
    return {"total": total, "by_type": found}


def recalc(path: str | Path, timeout: int = 60, *, force: bool = False) -> dict:
    """重算并原地存回。返回带 `ok` 的信封（不是抛异常）。"""
    p = Path(path).expanduser().resolve()
    if not p.exists():
        return {"ok": False, "error": f"文件不存在：{p}"}
    if not os.access(p, os.W_OK):
        return {"ok": False, "error": f"文件不可写：{p}（重算是原地重写，先复制一份）"}
    if p.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".ods"):
        return {"ok": False, "error": f"不是表格格式：{p.suffix}（本工具只重算 xlsx/xlsm/ods）"}

    try:
        total_formulas, cells = formulas_of(p)
    except Exception as exc:
        return {"ok": False, "error": f"读不出公式：{type(exc).__name__}: {exc}"}

    if total_formulas == 0:
        # 一个公式都没有：不值得起一次 LibreOffice
        r = _scan_errors(p)
        return {"ok": r["total"] == 0, "status": "success" if r["total"] == 0 else "errors_found",
                "total_formulas": 0, "total_errors": r["total"],
                "error_summary": _summary(r), "skipped": "文件里没有公式，未重算（纯值文件本就是好的）"}

    if _has_vba(p) and not force:
        return {"ok": False, "error": ("这个工作簿带 VBA 宏（xl/vbaProject.bin）；LibreOffice 存盘可能"
                                       "改写/丢掉宏工程。先复制一份单独重算，或加 --force 自行承担")}

    if not force:
        try:
            at_risk = external_links_at_risk(p)
        except Exception as exc:
            return {"ok": False, "error": f"外链预检失败：{type(exc).__name__}: {exc}"}
        if at_risk:
            return {"ok": False,
                    "error": (f"拒绝重算：工作簿引用了别的文件，且 {len(at_risk)} 个外链格的缓存值已丢。"
                              "重算会把它们解析成 #NAME? 并**永久删除外链**。"
                              "请先把这些格的值从原件抄出来，或加 --force 接受损失。"),
                    "external_link_cells": at_risk[:MAX_SHOWN],
                    "external_link_cells_truncated": max(0, len(at_risk) - MAX_SHOWN)}

    before = _stamp(p)
    started = time.monotonic()
    try:
        run = _run_macro(p, timeout)
    except RecalcError as exc:
        return {"ok": False, "error": str(exc)}

    if _stamp(p) == before:
        extra = ("（进程超时没退出，文件也没写成：多半是另有 LibreOffice 占着这个文件，"
                 "关掉再试）" if run["hung"] else
                 "（多半是另有 LibreOffice 实例占着文件，关掉其他 LibreOffice 再试）")
        return {"ok": False, "error": f"LibreOffice 退出了但**没有重写文件** —— 公式等于没算。{extra}"}

    r = _scan_errors(p)
    try:
        after_formulas, _ = formulas_of(p)
    except Exception:
        after_formulas = total_formulas
    empty = _cached_none(p)
    res = {
        "ok": r["total"] == 0 and not empty,
        "status": "success" if (r["total"] == 0 and not empty) else "errors_found",
        "file": str(p),
        "total_formulas": after_formulas,
        "total_errors": r["total"],
        "error_summary": _summary(r),
        "seconds": round(time.monotonic() - started, 1),
    }
    if empty:
        res["cached_but_empty"] = empty[:MAX_SHOWN]
        res["cached_but_empty_truncated"] = max(0, len(empty) - MAX_SHOWN)
        res["note"] = ("这些公式算完仍是空值：公式本身可能返回空串，或 LO 没认出来"
                       "（写回时函数名会被改成小写，是个线索）")
    if run["hung"]:
        res["note_slow"] = ("LibreOffice 跑完宏后没自己退出（已按 profile 清场），"
                            "但文件确实被重写了，结果有效")
    if r["total"] == 0 and not empty:
        res["reminder"] = "重算只证明公式**算得出来**，不证明**算得对**：范围错一位照样干净。抽查 2-3 个值。"
    return res


def _cached_none(path: Path) -> list[str]:
    """算完之后缓存值仍是 None 的公式格（`=""` 这类正常返回空，所以只提示不判死）。"""
    from openpyxl import load_workbook
    wb_f = load_workbook(path, data_only=False)
    wb_v = load_workbook(path, data_only=True)
    out: list[str] = []
    try:
        for sheet in wb_f.sheetnames:
            ws, cached = wb_f[sheet], wb_v[sheet]
            if not hasattr(ws, "iter_rows"):
                continue
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("=") and cached[cell.coordinate].value is None:
                        out.append(f"{sheet}!{cell.coordinate}")
    finally:
        wb_f.close()
        wb_v.close()
    return out


def _summary(r: dict) -> dict:
    s = {}
    for err, locs in r["by_type"].items():
        s[err] = {"count": len(locs), "locations": locs[:MAX_SHOWN]}
        if len(locs) > MAX_SHOWN:
            s[err]["locations_truncated"] = len(locs) - MAX_SHOWN
    return s


def main(argv: list[str] | None = None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="xlsx_recalc.py", description="用 LibreOffice 重算 xlsx 公式并扫错误")
    ap.add_argument("file")
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--force", action="store_true", help="即使会毁掉外链/VBA 也照算")
    a = ap.parse_args(argv)
    try:
        r = recalc(a.file, a.timeout, force=a.force)
    except Exception as exc:
        r = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    print(json.dumps(r, ensure_ascii=False, indent=1))
    if r.get("ok"):
        return 0
    return 3 if "error" not in r else 1


if __name__ == "__main__":
    sys.exit(main())
