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


# Windows/常见字体名 → 本机应落到的**字族关键字**。
#
# 判据是"落到了对的一类字"，不是"落到了某一个特定字库"：宋体只要落在衬线宋、仿宋落在仿宋、
# 黑体落在无衬线黑体就算过；英文字体同理（Arial→无衬线、Times New Roman→衬线、
# Courier New→等宽），**字形不必与微软原字体等同**（用户口径 2026-10-02）。
# 关键字匹配 fc-match 的**完整 family 字段**（英文名与中文名都认），不匹配文件名。
# 键与 assets/99-cjk-windows-alias.conf 里的 <family> 一一对应；
# selftest 有一条回归检查，防止"加了别名却忘了加期望"。
_FONT_EXPECT: dict[str, tuple[str, ...]] = {
    # 中文字体名：还必须支持中文（见 _font_aliases 的 cjk_ok）
    "宋体": ("AR PL SungtiL", "AR PL UMing", "Noto Serif CJK"),
    "SimSun": ("AR PL SungtiL", "AR PL UMing", "Noto Serif CJK"),
    "仿宋": ("cwTeX", "FangSong", "仿宋"),
    "FangSong": ("cwTeX", "FangSong", "仿宋"),
    "楷体": ("Kai", "楷"),
    "KaiTi": ("Kai", "楷"),
    "黑体": ("Noto Sans CJK", "WenQuanYi", "文泉驿", "Hei", "黑"),
    "SimHei": ("Noto Sans CJK", "WenQuanYi", "文泉驿", "Hei", "黑"),
    "微软雅黑": ("Noto Sans CJK", "WenQuanYi", "文泉驿", "Hei", "黑"),
    "方正小标宋简体": ("AR PL SungtiL", "AR PL UMing", "Noto Serif CJK"),
    # 英文常用族：只要求落在对的类别（衬线/无衬线/等宽），不要求支持中文
    "Arial": ("Liberation Sans", "Nimbus Sans", "Helvetica", "DejaVu Sans", "FreeSans"),
    "Helvetica": ("Liberation Sans", "Nimbus Sans", "DejaVu Sans", "FreeSans"),
    "Times New Roman": ("Liberation Serif", "Nimbus Roman", "DejaVu Serif", "FreeSerif", "Noto Serif"),
    "Georgia": ("Noto Serif", "DejaVu Serif", "Liberation Serif", "FreeSerif"),
    "Calibri": ("Carlito",),
    "Cambria": ("Caladea",),
    "Verdana": ("DejaVu Sans", "Noto Sans", "Liberation Sans", "FreeSans"),
    "Tahoma": ("DejaVu Sans", "Noto Sans", "Liberation Sans", "FreeSans"),
    "Courier New": ("Liberation Mono", "Nimbus Mono", "DejaVu Sans Mono", "FreeMono"),
}

# 这些是中文字体名（要求支持中文）；表里其余按英文族处理
_FONT_ZH = frozenset({
    "宋体", "SimSun", "仿宋", "FangSong", "楷体", "KaiTi",
    "黑体", "SimHei", "微软雅黑", "方正小标宋简体",
})


def _font_aliases() -> dict:
    """查「文档里写某字体名时，本机实际用哪套字」，并判断是否落到了对的一类字。

    判据（同时成立才算 ok）：
      1. fc-match 能返回字体；
      2. 落到的 family 里含该字体名对应的**字族关键字**（见 `_FONT_EXPECT`）；
      3. **只有中文名字**还要额外验证该字体支持中文（fontconfig lang 含 `zh`，或 family 含 CJK 字样）。
         英文字体名不强求中文字形 —— 中英混排时 LibreOffice 会自己走 CJK 回退链。

    踩过的坑（判据本身）：早先拿中文期望串（`文鼎`/`霞鹜文楷`）去比 `fc-match` 的**输出行**，
    而输出里是**英文** family 名（`AR PL SungtiL GB`/`LXGW WenKai`）—— 于是别名配得完全正确的
    机器也判 `bad`，`selftest` 恒 FAIL 一项。改成比 family 字段后，判据与实际配置对得上。
    另一类真实坏法仍会被抓到：不装别名时 `fc-match 宋体` 落到 Noto Sans CJK（无衬线），
    family 里没有宋体关键字 → 判 bad。
    """
    import subprocess as _sp
    out = {}
    for name, want in _FONT_EXPECT.items():
        families: list[str] = []
        langs = ""
        try:
            r = _sp.run(["fc-match", "-f", "%{family}|%{lang}", name],
                        capture_output=True, text=True, timeout=10)
            fam_raw, _, langs = (r.stdout or "").partition("|")
            families = [f.strip() for f in fam_raw.split(",") if f.strip()]
        except Exception:
            pass
        # `%{family}` 会给整个字族链（首选在前）。任一条命中关键字即算这一类字。
        blob = " ".join(families)
        alias_ok = any(w in blob for w in want)
        zh_ok = "zh" in {t.split("-")[0] for t in langs.split("|") if t} or "CJK" in blob
        is_zh = name in _FONT_ZH
        first = families[0] if families else ""
        out[name] = {
            "match": first,
            "families": families,
            "expect": " / ".join(want),
            "alias_ok": alias_ok,
            "cjk_ok": zh_ok,
            "script": "zh" if is_zh else "latin",
            "ok": bool(families) and alias_ok and (zh_ok if is_zh else True),
        }
    return out


if __name__ == "__main__":
    import json
    print(json.dumps(doctor(), ensure_ascii=False, indent=2))
    sys.exit(0)
