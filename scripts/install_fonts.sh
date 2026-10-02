#!/usr/bin/env bash
# 装中文字库 + Windows 字体名别名 —— 让本机 LibreOffice 出的 PDF 与 Word 上打开的样子一致。
#
# 为什么需要：只有 Noto/文泉驿时，fontconfig 会把「宋体」兜到 Noto Sans CJK（**无衬线**）
# —— 宋体渲染成黑体，公文排版就废了。而 docx 里存的字名仍是「宋体」，
# 在 Windows 上打开是对的、本机出 PDF 是错的，光看文件名发现不了。
#
# 判据（2026-10-02 口径）：只要求**落到对的一类字** —— 宋体→衬线宋、仿宋→仿宋、
# 楷体→楷、黑体→无衬线黑体；字形与微软原字体**不必等同**。
# 期望值只有一处：`scripts/convert.py` 的 `_FONT_EXPECT`（本脚本向它查，不再各写一份）。
#
# 用法：
#   sudo bash install_fonts.sh          # 装字库 + 别名，然后 fc-cache
#   sudo bash install_fonts.sh --check  # 只体检，不改系统
#
# 幂等：重复跑只是重装/覆盖别名，不出错。
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 文档目录（SKILL.md + assets/ 所在）：脚本既在技能里跑，也在工作区的薄壳位置跑
DOCS=""
for cand in "$HERE/.." "$HERE"; do
  if [[ -f "$cand/SKILL.md" && -f "$cand/assets/99-cjk-windows-alias.conf" ]]; then DOCS="$(cd "$cand" && pwd)"; break; fi
done
CONF_SRC="${DOCS:+$DOCS/assets/99-cjk-windows-alias.conf}"
CONF_DST=/etc/fonts/conf.d/99-cjk-windows-alias.conf
PKGS=(fonts-arphic-gbsn00lp fonts-arphic-gkai00mp fonts-cwtex-fs fonts-lxgw-wenkai)

# 期望值从 convert.py 取，避免"脚本一套、体检脚本另一套"地漂移
expect_block() {
  python3 - "$1" <<'PY'
import importlib.util, os, sys

docs = sys.argv[1]
src = os.path.join(docs, "scripts", "convert.py") if docs else ""
if not os.path.exists(src):
    sys.exit(3)                      # 找不到真身：让 bash 侧退回自带的兜底表
spec = importlib.util.spec_from_file_location("_office_convert", src)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
for name, want in mod._FONT_EXPECT.items():
    print(f"{name}|{' / '.join(want)}")
PY
}

# 兜底：技能不在（只拷了薄壳）时用这张表；正常路径永远走上面的单一来源
FALLBACK_WANT=(
  "宋体|AR PL SungtiL / AR PL UMing / Noto Serif CJK"
  "SimSun|AR PL SungtiL / AR PL UMing / Noto Serif CJK"
  "仿宋|cwTeX / FangSong / 仿宋"
  "FangSong|cwTeX / FangSong / 仿宋"
  "楷体|Kai / 楷"
  "KaiTi|Kai / 楷"
  "黑体|Noto Sans CJK / WenQuanYi / 文泉驿 / Hei / 黑"
  "SimHei|Noto Sans CJK / WenQuanYi / 文泉驿 / Hei / 黑"
  "微软雅黑|Noto Sans CJK / WenQuanYi / 文泉驿 / Hei / 黑"
  "方正小标宋简体|AR PL SungtiL / AR PL UMing / Noto Serif CJK"
)

check() {
  echo "── 字库 ──"
  for p in "${PKGS[@]}"; do
    if dpkg -s "$p" >/dev/null 2>&1; then echo "  ✓ $p"; else echo "  ✗ $p 未装"; fi
  done

  local want=() line
  if [[ -n "$DOCS" ]]; then
    while IFS= read -r line; do [[ -n "$line" ]] && want+=("$line"); done < <(expect_block "$DOCS" 2>/dev/null || true)
  fi
  if [[ ${#want[@]} -eq 0 ]]; then
    want=("${FALLBACK_WANT[@]}")
    echo "  （技能目录不在，用自带期望表）"
  else
    echo "  （期望值取自 $DOCS/scripts/convert.py）"
  fi

  echo "── 别名（判据：落到对的一类字 + 支持中文）──"
  local bad=0
  for w in "${want[@]}"; do
    local name="${w%%|*}" expect="${w#*|}"
    # family 字段不含文件名，用 family[0] 给人的那一段
    local got; got="$(fc-match -f '%{family[0]}' "$name" 2>/dev/null || true)"
    if [[ -z "$got" ]]; then
      printf '  ✗ %-16s -> （fc-match 取不到）\n' "$name"; bad=$((bad+1)); continue
    fi
    local hit=0 k
    IFS='/' read -ra KEYS <<< "$expect"
    for k in "${KEYS[@]}"; do
      k="${k#"${k%%[![:space:]]*}"}"; k="${k%"${k##*[![:space:]]}"}"
      [[ -z "$k" ]] && continue
      [[ "$got" == *"$k"* ]] && { hit=1; break; }
    done
    if [[ $hit -eq 1 ]]; then
      printf '  ✓ %-16s -> %s\n' "$name" "$got"
    else
      printf '  ✗ %-16s -> %s   （期望含 %s）\n' "$name" "$got" "$expect"; bad=$((bad+1))
    fi
  done
  [[ -f "$CONF_DST" ]] && echo "  ✓ 别名文件已装：$CONF_DST" || echo "  ✗ 别名文件缺失：$CONF_DST"
  return "$bad"
}

if [[ "${1:-}" == "--check" ]]; then
  if check; then echo "字体就绪"; else echo "字体未就绪（上表 ✗ 项）"; fi
  exit 0
fi

if [[ $EUID -ne 0 ]]; then
  echo "需要 root：sudo bash $0" >&2
  exit 1
fi

echo "== 1/3 装字库 =="
export DEBIAN_FRONTEND=noninteractive
apt-get install -y "${PKGS[@]}"

echo "== 2/3 装 fontconfig 别名 =="
install -m 644 "$CONF_SRC" "$CONF_DST"

echo "== 3/3 刷新字体缓存 =="
fc-cache -f >/dev/null

echo
if check; then
  echo "✅ 字体就绪"
else
  echo "⚠️  仍有别名没落到预期的一类字，检查 $CONF_DST" >&2
  exit 1
fi
