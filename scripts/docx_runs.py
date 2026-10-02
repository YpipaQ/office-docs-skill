#!/usr/bin/env python3
"""docx 文本层手术：run 合并 ＋ 按「渲染出来的字」做查找替换。

为什么需要这个模块（实测踩坑）：
- Word/LibreOffice 会把一句话拆成很多 `<w:r>`（rsid 修订号、拼写检查标记、编辑历史），
  于是你**在屏幕上看到**的「北京市朝阳区」在 `word/document.xml` 里可能是
  `<w:t>北京</w:t>` ＋ `<w:t>市朝</w:t>` ＋ `<w:t>阳区</w:t>` 三段。
  python-docx 的 `paragraph.text` 拼出的是「北京市朝阳区」，但逐 run 找 `旧词`
  一个 run 都不命中 ⇒ **查找替换静默漏词**，文件看上去没坏，其实没改。
- `<w:t>` 不带 `xml:space="preserve"` 时，XML 规则会把它**两侧的空白吃掉**：
  `<w:t>Hello </w:t><w:t>world</w:t>` 渲染出来是 `Helloworld`。所以匹配必须按
  **渲染后的字符串**算，不能按元素里的原始字节算。

做法：
1. `merge_runs()` —— 把相邻、`<w:rPr>` 完全相同的 run 并成一个，并合并同一 run 内
   相邻的 `<w:t>`/`<w:delText>`；渲染结果不变（这是可证明的等价变换）。
2. `replace()` —— 在合并后的**渲染文本**上定位，再把改动摊回各 `<w:t>`；
   跨 run 的命中也能改，且**保留每个 run 原有的字体/加粗**（不是整段重建）。

刻意不做的：
- 跨 `<w:ins>`/`<w:del>` 边界的 run 不合并（那会改写修订结构，把两次修订揉成一次）。
- 不动 `<w:delText>` 里的文字（那是被删掉的内容，改了没有意义还破坏修订记录）。
- 只有明确要求才清 rsid 属性（默认清，因为它只是编辑痕迹、不参与渲染）。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_RPR = W + "rPr"
_T = W + "t"
_DELTEXT = W + "delText"
_R = W + "r"
_P = W + "p"
_DEL = W + "del"
_PROOFERR = W + "proofErr"


# ---------- 基础 ----------

def _tag(el) -> str:
    return el.tag if isinstance(el.tag, str) else ""


def _rpr_xml(run) -> str | None:
    """run 的格式指纹：`<w:rPr>` 的序列化。没有 rPr 返回 None。"""
    rpr = run.find(_RPR)
    if rpr is None:
        return None
    from lxml import etree
    return etree.tostring(rpr, encoding="unicode")


def _render(text: str, preserve: bool) -> str:
    """`<w:t>` 内容 → 实际画到纸上的字。"""
    return text if preserve else text.strip(" \t\r\n")


def _is_preserve(el) -> bool:
    return el.get(XML_SPACE) == "preserve"


def _set_text(el, text: str) -> None:
    """写回 `<w:t>`，并按需打/摘 xml:space（两侧有空白就必须 preserve，否则会被吃掉）。"""
    el.text = text
    if text != text.strip(" \t\r\n") or "\n" in text or "\t" in text:
        el.set(XML_SPACE, "preserve")
    elif el.get(XML_SPACE) is not None:
        del el.attrib[XML_SPACE]


def _iter_runs(p):
    """段落里直接挂在 p 下的 `<w:r>`（不进 w:hyperlink 内部 —— 那里的 run 换段要丢链接）。"""
    for child in p:
        if _tag(child) == _R:
            yield child


def _inside_del(run, stop) -> bool:
    node = run.getparent()
    while node is not None and node is not stop:
        if _tag(node) == _DEL:
            return True
        node = node.getparent()
    return False


# ---------- 合并 ----------

def strip_proof_errors(node) -> int:
    """删掉 `<w:proofErr>`（拼写/语法检查标记）：它夹在 run 之间，挡住合并，且不参与渲染。"""
    n = 0
    for el in list(node.iter(_PROOFERR)):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            n += 1
    return n


def strip_rsid_attrs(node) -> int:
    """删掉所有 `w:rsid*` 属性（编辑会话编号，纯痕迹）。"""
    n = 0
    for el in node.iter():
        for attr in list(el.attrib):
            if "rsid" in attr.lower():
                del el.attrib[attr]
                n += 1
    return n


def _consolidate_in_run(run) -> int:
    """把同一个 run 内相邻的 `<w:t>`（或 `<w:delText>`）并成一个，按渲染文本拼接。"""
    merged = 0
    for tag in (_T, _DELTEXT):
        kids = [c for c in run if _tag(c) == tag]
        for i in range(len(kids) - 1, 0, -1):
            cur, prev = kids[i], kids[i - 1]
            if prev.getnext() is not cur:      # 中间夹了别的东西，不动
                continue
            text = _render(prev.text or "", _is_preserve(prev)) + _render(cur.text or "", _is_preserve(cur))
            _set_text(prev, text)
            run.remove(cur)
            merged += 1
    return merged


def _merge_in_paragraph(p) -> int:
    """相邻同格式 run 合并。返回合并次数。"""
    merged = 0
    run = next(iter(_iter_runs(p)), None)
    while run is not None:
        while True:
            nxt = run.getnext()
            # 跳过纯空白文本节点；碰到元素就停（bookmark/hyperlink 等一律不跨）
            while nxt is not None and not isinstance(nxt.tag, str):
                if (nxt.text or "").strip():
                    break
                nxt = nxt.getnext()
            if nxt is None or _tag(nxt) != _R or _rpr_xml(run) != _rpr_xml(nxt):
                break
            for child in list(nxt):
                if _tag(child) != _RPR:
                    run.append(child)
            p.remove(nxt)
            merged += 1
        merged += _consolidate_in_run(run)
        run = _next_run(run)
    return merged


def _next_run(run):
    nxt = run.getnext()
    while nxt is not None:
        if _tag(nxt) == _R:
            return nxt
        nxt = nxt.getnext()
    return None


def merge_runs(node) -> int:
    """node：`<w:document>` / `<w:body>` / 页眉页脚根元素 / 也可直接给 `<w:p>`。返回合并次数。"""
    if _tag(node) == _P:
        return _merge_in_paragraph(node)
    total = strip_proof_errors(node)
    strip_rsid_attrs(node)
    for p in node.iter(_P):
        total += _merge_in_paragraph(p)
    return total


# ---------- 替换 ----------

def _text_elements(p) -> list[str]:
    """段落里参与「显示文字」的 `<w:t>` 列表（跳过修订删除里的 delText）。"""
    out = []
    for run in _iter_runs(p):
        if _inside_del(run, p):
            continue
        for child in run:
            if _tag(child) == _T:
                out.append(child)
    return out


def _spans(elems) -> tuple[str, list[tuple]]:
    """拼出渲染文本，并记录每个 `<w:t>` 覆盖的字符区间 [a, b)。"""
    parts, spans, pos = [], [], 0
    for el in elems:
        txt = _render(el.text or "", _is_preserve(el))
        spans.append((el, pos, pos + len(txt)))
        parts.append(txt)
        pos += len(txt)
    return "".join(parts), spans


def _find_matches(text: str, mapping: dict) -> list[tuple[int, int, str]]:
    """在 text 上找所有命中，返回 [(start, end, repl)]，互不重叠（先长后短、先左后右）。"""
    found: list[tuple[int, int, str]] = []
    for old in sorted(mapping, key=len, reverse=True):
        new = mapping[old]
        if not old:
            continue
        start = text.find(old)
        while start != -1:
            end = start + len(old)
            if not any(s < end and start < e for s, e, _ in found):
                found.append((start, end, new))
            start = text.find(old, start + 1)
    found.sort()
    return found


def _apply(el, a: int, b: int, edits: list[tuple[int, int, str]]) -> str:
    """把命中摊回单个 `<w:t>`：区间内的换成新串，区间外的原样保留。"""
    out, cur = [], a
    for s, e, new in edits:
        if e <= cur:
            continue
        if s >= b:
            break
        if s > cur:
            out.append((el.text or "")[cur - a:s - a])
        if cur <= s:                      # 只有命中起点所在的那个元素负责吐出替换串（且只吐一次）
            out.append(new)
        cur = min(b, e)
    if cur < b:
        out.append((el.text or "")[cur - a:b - a])
    return "".join(out)


def _replace_in_paragraph(p, mapping: dict) -> int:
    elems = _text_elements(p)
    if not elems:
        return 0
    text, spans = _spans(elems)
    if not text:
        return 0
    edits = _find_matches(text, mapping)
    if not edits:
        return 0
    for el, a, b in spans:
        if a == b:
            continue
        touching = [e for e in edits if e[0] < b and a < e[1]]
        if not touching:
            continue
        new = _apply(el, a, b, touching)
        if new != (el.text or ""):
            _set_text(el, new)
    return len(edits)


# ---------- 对外 ----------

def replace_in_document(doc, mapping: dict, *, merge: bool = True) -> dict:
    """在 python-docx 的 Document 上做替换。覆盖正文/表格/文本框/页眉页脚（含嵌套表格）。

    顶层表格、嵌套表格、文本框（w:txbxContent）都在 w:body 的 iter('w:p') 里，一次遍历全中。
    """
    body = doc.element.body
    stats = {"runs_merged": merge_runs(body) if merge else 0, "hits": 0}
    for p in body.iter(_P):
        stats["hits"] += _replace_in_paragraph(p, mapping)
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            try:
                root = part._element
            except Exception:
                continue
            if merge:
                stats["runs_merged"] += merge_runs(root)
            for p in root.iter(_P):
                stats["hits"] += _replace_in_paragraph(p, mapping)
    stats["revisions"] = _count_revisions(doc)
    return stats


def _count_revisions(doc) -> int:
    n = 0
    for tag in ("ins", "del", "moveFrom", "moveTo"):
        n += len(list(doc.element.body.iter(W + tag)))
    return n


def replace_in_file(path, mapping: dict, out=None, *, merge: bool = True) -> dict:
    from docx import Document
    src = Path(path).expanduser().resolve()
    doc = Document(str(src))
    stats = replace_in_document(doc, mapping, merge=merge)
    dst = Path(out).expanduser() if out else src.with_name(src.stem + "-replaced" + src.suffix)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    res = {"out": str(dst), **stats}
    if stats["revisions"]:
        res["warning"] = (f"文档里有 {stats['revisions']} 处修订标记；替换只改可见文字，"
                          "不改修订结构，改完请用 docx-review 盘一次")
    return res


def merge_runs_in_file(path, out=None) -> dict:
    """把 docx 里的 run 碎片合并后另存（改前先做这一步，替换/正则/XML 手改都更省事）。"""
    from docx import Document
    src = Path(path).expanduser().resolve()
    doc = Document(str(src))
    before = sum(1 for _ in doc.element.body.iter(_R))
    merged = merge_runs(doc.element.body)
    for section in doc.sections:
        for part in (section.header, section.footer):
            try:
                merged += merge_runs(part._element)
            except Exception:
                continue
    after = sum(1 for _ in doc.element.body.iter(_R))
    dst = Path(out).expanduser() if out else src.with_name(src.stem + "-merged" + src.suffix)
    doc.save(str(dst))
    return {"out": str(dst), "runs_before": before, "runs_after": after, "merged": merged}


def main(argv: list[str] | None = None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="docx_runs.py", description="docx run 合并 / 渲染文本替换")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("merge", help="合并 run 碎片后另存")
    p.add_argument("file")
    p.add_argument("-o", "--out")
    p = sub.add_parser("replace", help="查找替换（按渲染文本，跨 run 也命中）")
    p.add_argument("file")
    p.add_argument("--set", action="append", required=True, metavar="旧=新")
    p.add_argument("-o", "--out")
    p.add_argument("--no-merge", action="store_true", help="不先合并 run（默认先合并）")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "merge":
            print(json.dumps(merge_runs_in_file(a.file, a.out), ensure_ascii=False, indent=1))
            return 0
        mapping = {}
        for kv in a.set:
            k, v = kv.split("=", 1)
            mapping[k] = v
        r = replace_in_file(a.file, mapping, a.out, merge=not a.no_merge)
        print(json.dumps({"ok": True, **r}, ensure_ascii=False, indent=1))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
