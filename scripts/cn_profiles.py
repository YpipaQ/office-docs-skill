#!/usr/bin/env python3
"""中文公文 / 商业文书的排版实现（GB/T 9704-2012 及行业惯例）。

设计取舍（重要）：
- **docx 里写国标字体名**（仿宋_GB2312 / 楷体_GB2312 / 方正小标宋简体）：文件拿到装了这些字体的
  Windows/WPS 上打开就是合规的。本机没有这些商业字体，本机渲染靠 `--portable` 换成 Noto CJK，
  只用于看版式。
- **不写系统 fontconfig**：本模块不碰 `~/.config`，字体替换只在本文件内按开关生效。
- 行距用**固定值**（公文 29 磅 = 版心 225mm ÷ 22 行），字号按磅，不用"三号"这种字符串。
- 首行缩进写 `w:firstLineChars="200"`（Word 里显示为"2 字符"）＋ `w:firstLine` 磅值兜底。
"""
from __future__ import annotations

import json
from pathlib import Path

# 中国字号 → 磅
CN_SIZE = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}

# 国标/惯例字体名 vs 本机可用替代
FONTS = {
    "fangsong": "仿宋_GB2312", "kai": "楷体_GB2312", "hei": "黑体",
    "song": "宋体", "xbs": "方正小标宋简体", "yahei": "微软雅黑",
    "latin": "Times New Roman",
}
PORTABLE = {
    "fangsong": "Noto Serif CJK SC", "kai": "Noto Serif CJK SC", "hei": "Noto Sans CJK SC",
    "song": "Noto Serif CJK SC", "xbs": "Noto Serif CJK SC", "yahei": "Noto Sans CJK SC",
    "latin": "Liberation Serif",
}

# 版心 156×225 mm；每面 22 行 ⇒ 固定行距 225mm=637.8pt ÷ 22 ≈ 29pt
GONGWEN_LINE = 29.0

PROFILES: dict[str, dict] = {
    "gongwen": {                     # GB/T 9704-2012 标准公文
        "desc": "党政机关公文（标准版式）",
        "margins_cm": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
        "body": ("fangsong", CN_SIZE["三号"]),
        "line_exact": GONGWEN_LINE,
        "indent_chars": 2,
        "title": ("xbs", CN_SIZE["二号"], False),
        "title_line": 40.0,          # 标题行距建议固定值 40 磅
        "h_levels": [("hei", CN_SIZE["三号"]), ("kai", CN_SIZE["三号"]),
                     ("fangsong", CN_SIZE["三号"]), ("fangsong", CN_SIZE["三号"])],
        "page_number": ("song", CN_SIZE["四号"], "gongwen"),
        "latin": "latin",
    },
    "contract": {                    # 合同 / 协议（惯例）
        "desc": "合同、协议（行业惯例）",
        "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
        "body": ("song", CN_SIZE["小四"]),
        "line_multiple": 1.5,
        "indent_chars": 2,
        "title": ("hei", CN_SIZE["二号"], True),
        "h_levels": [("hei", CN_SIZE["小四"]), ("song", CN_SIZE["小四"]),
                     ("song", CN_SIZE["小四"]), ("song", CN_SIZE["小四"])],
        "page_number": ("song", CN_SIZE["五号"], "plain"),
        "latin": "latin",
        "numbering": "clause",       # 第一条 / 第二条
    },
    "report": {                      # 报告 / 方案（惯例）
        "desc": "报告、方案、总结（行业惯例）",
        "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 2.8},
        "body": ("song", CN_SIZE["小四"]),
        "line_multiple": 1.5,
        "indent_chars": 2,
        "title": ("hei", CN_SIZE["小二"], True),
        "h_levels": [("hei", CN_SIZE["小三"]), ("hei", CN_SIZE["四号"]),
                     ("song", CN_SIZE["小四"]), ("song", CN_SIZE["小四"])],
        "page_number": ("song", CN_SIZE["五号"], "plain"),
        "latin": "latin",
    },
}


# ---------------- 基础工具 ----------------

def cn_num(n: int) -> str:
    """1..99 的中文数字（公文层次/条款序号用）。"""
    d = "零一二三四五六七八九"
    if n <= 0:
        return str(n)
    if n < 10:
        return d[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + d[n % 10]
    if n < 100:
        return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")
    return str(n)


def _font(profile: dict, key: str, portable: bool) -> str:
    return (PORTABLE if portable else FONTS).get(key, key)


def _set_run(run, name: str, size: float | None = None, bold: bool | None = None) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def _indent(par, *, first_chars: float = 0, size: float = 12, right_chars: float = 0) -> None:
    """首行缩进按"字符"写（Word 显示 2 字符），同时给磅值兜底（LibreOffice 认这个）。"""
    from docx.oxml.ns import qn
    ppr = par._p.get_or_add_pPr()
    ind = ppr.get_or_add_ind()
    if first_chars:
        ind.set(qn("w:firstLineChars"), str(int(first_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(size * first_chars * 20)))   # twips
    if right_chars:
        ind.set(qn("w:rightChars"), str(int(right_chars * 100)))
        ind.set(qn("w:right"), str(int(size * right_chars * 20)))


def _par(doc, profile: dict, font_key: str, size: float, *, portable: bool,
         align=None, first_chars: float = 0, right_chars: float = 0,
         line_exact: float | None = None, line_multiple: float | None = None,
         bold: bool | None = None, space_before: float = 0, space_after: float = 0):
    from docx.shared import Pt
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    if line_exact:
        pf.line_spacing = Pt(line_exact)
    elif line_multiple:
        pf.line_spacing = line_multiple
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_chars or right_chars:
        _indent(p, first_chars=first_chars, size=size, right_chars=right_chars)
    return p, _font(profile, font_key, portable)


def _add_page_number(doc, section, profile: dict, *, portable: bool) -> None:
    """公文书眉：4 号宋体、形如 — 1 —、奇页居右 / 偶页居左；普通文书：五号、居中。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    key, size, style = profile["page_number"]
    name = _font(profile, key, portable)
    try:
        doc.settings.odd_and_even_pages_header_footer = True
    except Exception:
        pass

    def fill(footer, align):
        p = footer.paragraphs[0]
        p.alignment = align
        for text in ("— ", None, " —"):
            if text is None:
                fld = OxmlElement("w:fldSimple")
                fld.set(qn("w:instr"), " PAGE ")
                r = OxmlElement("w:r")
                rpr = OxmlElement("w:rPr")
                sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2)))
                rpr.append(sz); r.append(rpr)
                fld.append(r); p._p.append(fld)
            else:
                run = p.add_run(text)
                _set_run(run, name, size)

    if style == "gongwen":
        fill(section.footer, WD_ALIGN_PARAGRAPH.RIGHT)
        if hasattr(section, "even_page_footer"):
            fill(section.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT)
        from docx.shared import Cm
        section.footer_distance = Cm(2.8)     # 一字线上距版心下边缘 7mm ⇒ 距页底 3.5-0.7
    else:
        fill(section.footer, WD_ALIGN_PARAGRAPH.CENTER)


# ---------------- 生成器 ----------------

def build_docx(spec: dict | str | Path, out: str | Path, profile_name: str = "gongwen",
               *, portable: bool = False) -> Path:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Mm, Pt

    if isinstance(spec, (str, Path)):
        spec = json.loads(Path(spec).read_text(encoding="utf-8"))
    profile = PROFILES.get(profile_name)
    if profile is None:
        raise ValueError(f"未知 profile：{profile_name}（可选 {', '.join(PROFILES)}）")

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    m = profile["margins_cm"]
    sec.top_margin, sec.bottom_margin = Cm(m["top"]), Cm(m["bottom"])
    sec.left_margin, sec.right_margin = Cm(m["left"]), Cm(m["right"])

    body_key, body_size = profile["body"]
    body_name = _font(profile, body_key, portable)
    normal = doc.styles["Normal"]
    normal.font.name = body_name
    normal.font.size = Pt(body_size)
    from docx.oxml.ns import qn
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), body_name)

    def body_par(text: str, *, font_key: str = body_key, size: float = body_size,
                 bold: bool | None = None, first_chars: float | None = None, **kw):
        p, name = _par(doc, profile, font_key, size, portable=portable,
                       line_exact=profile.get("line_exact"),
                       line_multiple=None if profile.get("line_exact") else profile.get("line_multiple"),
                       first_chars=profile["indent_chars"] if first_chars is None else first_chars,
                       **kw)
        run = p.add_run(text)
        _set_run(run, name, size, bold)
        return p

    # 标题（2 号小标宋，居中）
    if spec.get("title"):
        tk, tsz, tbold = profile["title"]
        p, name = _par(doc, profile, tk, tsz, portable=portable, align=WD_ALIGN_PARAGRAPH.CENTER,
                       first_chars=0, line_exact=profile.get("title_line"),
                       line_multiple=None if profile.get("title_line") else profile.get("line_multiple"),
                       space_before=tsz, space_after=tsz)
        _set_run(p.add_run(spec["title"]), name, tsz, tbold)

    # 主送机关（顶格、全角冒号）
    if spec.get("recipient"):
        p, name = _par(doc, profile, body_key, body_size, portable=portable,
                       line_exact=profile.get("line_exact"),
                       line_multiple=None if profile.get("line_exact") else profile.get("line_multiple"),
                       first_chars=0)
        _set_run(p.add_run(spec["recipient"]), name, body_size)

    # 正文各块（层次自动编号：一、/（一）/1./（1）或 第一条）
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    clause_mode = profile.get("numbering") == "clause"
    for blk in spec.get("blocks", []):
        if "h" in blk and "text" in blk:          # 兼容旧写法
            blk = {f"h{int(blk['h'])}": blk["text"], **{k: v for k, v in blk.items() if k not in ("h", "text")}}
        for lvl in (1, 2, 3, 4):
            key = f"h{lvl}"
            if key not in blk:
                continue
            counters[lvl] += 1
            for deeper in range(lvl + 1, 5):
                counters[deeper] = 0
            text = blk[key]
            if clause_mode:
                label = f"第{cn_num(counters[1])}条" if lvl == 1 else ""
                text = f"{label}　{text}" if label else text
                hk, hsz = profile["h_levels"][min(lvl - 1, 3)]
            else:
                prefix = {1: f"{cn_num(counters[1])}、", 2: f"（{cn_num(counters[2])}）",
                          3: f"{counters[3]}. ", 4: f"（{counters[4]}）"}[lvl]
                text = prefix + text
                hk, hsz = profile["h_levels"][min(lvl - 1, 3)]
            body_par(text, font_key=hk, size=hsz, bold=(profile_name == "contract" and lvl == 1))
            break
        else:
            if "p" in blk:
                body_par(blk["p"])
            elif "bullets" in blk:
                for b in blk["bullets"]:
                    body_par("— " + b)
            elif "table" in blk:
                rows = blk["table"].get("rows", [])
                if rows:
                    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
                    t.style = "Table Grid"
                    for i, row in enumerate(rows):
                        for j, val in enumerate(row):
                            cell = t.cell(i, j)
                            cell.text = ""
                            run = cell.paragraphs[0].add_run(str(val))
                            _set_run(run, body_name, body_size - 1.5,
                                     bold=bool(blk["table"].get("header", True)) and i == 0)
            elif blk.get("pagebreak"):
                doc.add_page_break()

    # 附件说明（正文下空一行、"附件："左空二字）
    if spec.get("attachments"):
        lines = spec["attachments"]
        first = f"附件：{lines[0]}" if len(lines) > 1 else f"附件：{lines[0]}"
        body_par(first, first_chars=2)
        for extra in lines[1:]:
            body_par(extra, first_chars=0)

    # 发文机关署名 + 成文日期（盖章时日期右空四字）
    sig = spec.get("signature") or {}
    if sig:
        body_par("", first_chars=0)
        date = sig.get("date", "")
        org = sig.get("org", "")
        right = 4 if sig.get("seal", True) else 2
        if org:
            p, name = _par(doc, profile, body_key, body_size, portable=portable,
                           align=WD_ALIGN_PARAGRAPH.RIGHT, right_chars=right,
                           line_exact=profile.get("line_exact"),
                           line_multiple=None if profile.get("line_exact") else profile.get("line_multiple"))
            _set_run(p.add_run(org), name, body_size)
        if date:
            p, name = _par(doc, profile, body_key, body_size, portable=portable,
                           align=WD_ALIGN_PARAGRAPH.RIGHT, right_chars=right,
                           line_exact=profile.get("line_exact"),
                           line_multiple=None if profile.get("line_exact") else profile.get("line_multiple"))
            _set_run(p.add_run(date), name, body_size)

    # 附注（左空二字加圆括号）
    if spec.get("note"):
        body_par(spec["note"], first_chars=2)

    _add_page_number(doc, sec, profile, portable=portable)

    out = Path(out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out


def build_pptx(spec: dict | str | Path, out: str | Path, profile_name: str = "brief",
               *, portable: bool = False) -> Path:
    """汇报 PPT：16:9、标题 32pt 加粗、正文 18pt、每页要点 ≤6 条、页脚页码。"""
    import writers as w

    if isinstance(spec, (str, Path)):
        spec = json.loads(Path(spec).read_text(encoding="utf-8"))
    spec = dict(spec)
    spec.setdefault("size", "16:9")
    spec["font"] = PORTABLE["yahei"] if portable else FONTS["yahei"]
    # 中文汇报惯例：封面 32pt、页标题 28pt、正文 18pt、表内 14pt
    spec.setdefault("title_size", 32)
    spec.setdefault("slide_title_size", 28)
    spec.setdefault("body_size", 18)
    spec.setdefault("table_size", 14)
    for sl in spec.get("slides", []):
        if len(sl.get("bullets", [])) > 6:
            sl["bullets"] = sl["bullets"][:6] + [f"…（本页还有 {len(sl['bullets']) - 6} 条，已截断）"]
    return w.new_pptx(spec, out)
