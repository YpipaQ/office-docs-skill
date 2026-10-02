#!/usr/bin/env bash
# 装中文字库 + Windows 字体名别名 —— 让本机 LibreOffice 出的 PDF 真正像 Windows 上打开的样子。
#
# 为什么需要：本机只有 Noto Sans/Serif CJK 和文泉驿。文档里写「宋体」时，fontconfig 会把
# 它兜到 Noto Sans CJK（**无衬线**）—— 宋体变黑体，公文排版就废了。实测证据：
#   fc-match 宋体  -> NotoSansCJK-Regular.ttc（错）
#   装完后         -> gbsn00lp.ttf「文鼎ＰＬ简报宋」（对）
#
# 用法：
#   sudo bash install_fonts.sh          # 装字库 + 别名，然后 fc-cache
#   sudo bash install_fonts.sh --check  # 只体检，不改系统
#
# 幂等：重复跑只是重装/覆盖别名，不出错。
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONF_SRC="$HERE/../assets/99-cjk-windows-alias.conf"
CONF_DST=/etc/fonts/conf.d/99-cjk-windows-alias.conf
PKGS=(fonts-arphic-gbsn00lp fonts-arphic-gkai00mp fonts-cwtex-fs fonts-lxgw-wenkai)

check() {
  echo "── 字库 ──"
  for p in "${PKGS[@]}"; do
    if dpkg -s "$p" >/dev/null 2>&1; then echo "  ✓ $p"; else echo "  ✗ $p 未装"; fi
  done
  echo "── 别名（期望值见括号）──"
  local want=(
    "宋体|文鼎"
    "SimSun|文鼎"
    "仿宋|cwTeX"
    "FangSong|cwTeX"
    "楷体|霞鹜文楷"
    "KaiTi|霞鹜文楷"
    "黑体|Noto Sans CJK"
    "SimHei|Noto Sans CJK"
    "微软雅黑|Noto Sans CJK"
    "方正小标宋简体|文鼎"
  )
  local bad=0
  for w in "${want[@]}"; do
    local name="${w%%|*}" expect="${w##*|}"
    local got; got="$(fc-match "$name" 2>/dev/null)"
    if [[ "$got" == *"$expect"* ]]; then
      printf '  ✓ %-16s -> %s\n' "$name" "$got"
    else
      printf '  ✗ %-16s -> %s   （期望含 %s）\n' "$name" "$got" "$expect"; bad=$((bad+1))
    fi
  done
  [[ -f "$CONF_DST" ]] && echo "  ✓ 别名文件已装：$CONF_DST" || echo "  ✗ 别名文件缺失：$CONF_DST"
  return "$bad"
}

if [[ "${1:-}" == "--check" ]]; then
  check && echo "字体就绪" || echo "字体未就绪（上表 ✗ 项）"
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
  echo "⚠️  仍有别名没落到预期字库，检查 $CONF_DST" >&2
  exit 1
fi
