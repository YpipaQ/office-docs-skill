#!/usr/bin/env python3
"""LibreOffice headless 转换封装：并发安全、有超时、退出码与产物双重校验。

踩过的坑（都固化在这里）：
- 同一个 UserInstallation profile 不能让两个 soffice 并发跑 ⇒ 每次调用开独立 profile。
- `soffice ... | tail` 的 `$?` 是 tail 的，pipe 会吞掉真实退出码 ⇒ 这里直接用 subprocess，
  不看 shell 管道；判定以"产物存在且非空"为准。
- 文本类导出不指定 filter 选项时中文常变乱码 ⇒ csv/txt 默认钉住 UTF-8（charset 76）。
- PDF 逆向（pdf→docx/odt）LO 没有导出过滤器，会报 `no export filter` ⇒ 走 pdf2docx，别绕 LO。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

# 扩展名 -> LO 过滤器名。只有 LO 的自动推断会选错或需要钉死参数时才需要显式指定。
TARGET_FILTERS: dict[str, str] = {
    "doc": "MS Word 97",
    "dot": "MS Word 97",
    "xls": "MS Excel 97",
    "ppt": "MS PowerPoint 97",
    "docx": "Office Open XML Text",
    "xlsx": "Calc MS Excel 2007 XML",
    "pptx": "Impress Office Open XML",
    "odt": "writer8",
    "ods": "calc8",
    "odp": "impress8",
    "pdf": None,  # writer_pdf_Export / calc_pdf_Export / impress_pdf_Export 由 LO 按文档类型选
    "csv": "Text - txt - csv (StarCalc)",
    "txt": "Text (encoded)",
    "html": "HTML (StarWriter)",
}

# LO 过滤器参数。csv: 字段分隔符 44(,)、文本定界符 34(")、字符集 76(UTF-8)、首行 1
FILTER_OPTIONS: dict[str, str] = {
    "csv": "44,34,76,1",
    "txt": "76,0,0,0",
}

DEFAULT_TIMEOUT = 180


class ConvertError(RuntimeError):
    pass


def _profile_dir() -> Path:
    """每次调用一个独立 profile（并发安全的唯一办法）。可用环境变量复用/覆盖。"""
    fixed = os.environ.get("OFFICE_DOCS_LO_PROFILE")
    if fixed:
        p = Path(fixed).expanduser()
        p.mkdir(parents=True, exist_ok=True)
        return p
    return Path(tempfile.gettempdir()) / f"office-docs-skill-lo-{os.getpid()}-{uuid.uuid4().hex[:6]}"


def convert(
    src: str | Path,
    to: str,
    out_dir: str | Path | None = None,
    *,
    filter_name: str | None = None,
    timeout: int = DEFAULT_TIMEOUT,
    keep_profile: bool = False,
) -> Path:
    """把 src 转成 to 扩展名，返回产物路径。产物不存在或为空即报错。"""
    if SOFFICE is None:
        raise ConvertError("找不到 soffice/libreoffice —— 先装 libreoffice")
    src = Path(src).expanduser().resolve()
    if not src.exists():
        raise ConvertError(f"源文件不存在：{src}")
    to = to.lower().lstrip(".")
    out_dir = Path(out_dir).expanduser() if out_dir else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    target = out_dir / (src.stem + "." + to)
    if target.exists():
        target.unlink()  # 避免把上一轮的旧产物误判为本次成功

    filt = filter_name if filter_name is not None else TARGET_FILTERS.get(to, "")
    if filt is None:
        spec = to
    elif filt and to in FILTER_OPTIONS:
        spec = f"{to}:{filt}:{FILTER_OPTIONS[to]}"
    elif filt:
        spec = f"{to}:{filt}"
    else:
        spec = to

    profile = _profile_dir()
    cmd = [
        SOFFICE, "--headless", "--norestore", "--nolockcheck", "--nologo",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", spec, "--outdir", str(out_dir), str(src),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        raise ConvertError(f"soffice 超时（{timeout}s）：{src.name} -> {to}") from exc
    finally:
        if not keep_profile:
            shutil.rmtree(profile, ignore_errors=True)

    if not target.exists() or target.stat().st_size == 0:
        tail = (proc.stderr or proc.stdout or "").strip().splitlines()[-3:]
        raise ConvertError(f"转换失败（exit={proc.returncode}）：{src.name} -> {to}；{' / '.join(tail)}")
    return target


def available_targets() -> list[str]:
    return sorted(TARGET_FILTERS)


def quiet_import(name: str):
    """导入会把噪声写到 stdout 的库时用（例：pdf2docx → fitz，会 print 弃用警告）。

    本工作区口径：CLI 的 stdout 必须是纯 JSON，所以这类噪声必须挡在 import 里。
    """
    import contextlib
    import importlib
    import io

    with contextlib.redirect_stdout(io.StringIO()):
        return importlib.import_module(name)


def doctor() -> dict:
    """环境自检：工具、库、中文字体。"""
    def _v(mod: str) -> str | None:
        try:
            m = quiet_import(mod)
            return getattr(m, "__version__", "ok")
        except Exception:
            return None

    tools = {name: shutil.which(name) or "" for name in
             ("soffice", "libreoffice", "pandoc", "pdftotext", "pdftoppm", "tesseract")}
    libs = {name: _v(name) for name in
            ("docx", "openpyxl", "pptx", "xlrd", "odf", "pymupdf", "pypdf", "pdf2docx")}
    fonts_zh = 0
    try:
        import subprocess as _sp
        out = _sp.run(["fc-list", ":lang=zh", "family"], capture_output=True, text=True, timeout=15)
        fonts_zh = len({ln.split(",")[0].strip() for ln in out.stdout.splitlines() if ln.strip()})
    except Exception:
        fonts_zh = -1
    missing = [k for k, v in libs.items() if v is None]
    aliases = _font_aliases()
    return {
        "soffice": tools.get("soffice") or "",
        "tools": tools,
        "libs": libs,
        "missing_libs": missing,
        "cjk_fonts": fonts_zh,
        "font_aliases": aliases,
        "font_aliases_bad": [k for k, v in aliases.items() if not v["ok"]],
        "ok": bool(tools.get("soffice")) and not missing,
    }


# Windows/常见中文字体名 → 本机应落到的字库（判据取 fc-match 输出里的子串）
_FONT_EXPECT = {
    "宋体": "文鼎", "SimSun": "文鼎",
    "仿宋": "cwTeX", "FangSong": "cwTeX",
    "楷体": "霞鹜文楷", "KaiTi": "霞鹜文楷",
    "黑体": "Noto Sans CJK", "SimHei": "Noto Sans CJK",
    "微软雅黑": "Noto Sans CJK",
}


def _font_aliases() -> dict:
    """查「文档里写宋体时，本机实际用哪套字」。

    踩过的坑：不装 fontconfig 别名时 fc-match 宋体 会落到 **Noto Sans CJK（无衬线）** ——
    宋体渲染成黑体。而 docx 里存的仍是「宋体」，所以在 Windows 上打开是对的、
    本机出 PDF 是错的，光看文件名根本发现不了。装别名见 scripts/install_fonts.sh。
    """
    import subprocess as _sp
    out = {}
    for name, expect in _FONT_EXPECT.items():
        try:
            r = _sp.run(["fc-match", name], capture_output=True, text=True, timeout=10)
            got = (r.stdout or "").strip()
        except Exception:
            got = ""
        out[name] = {"match": got, "expect": expect, "ok": expect in got}
    return out


if __name__ == "__main__":
    import json
    print(json.dumps(doctor(), ensure_ascii=False, indent=2))
    sys.exit(0)
