# 格式 / 过滤器 / 字体 / 规格 JSON —— 速查

> 本机实测日期：2026-09-30（Debian 13 trixie，LibreOffice 25.2.3.2，Python 3.13.5）。

## 1. 支持矩阵（本机实测）

| 格式 | 读 | 写 | 转换 | 说明 |
| --- | --- | --- | --- | --- |
| `.docx` / `.dotx` | ✅ python-docx | ✅ python-docx | ✅ | 主力格式 |
| `.doc`（97-2003） | ✅ LO | ✅ LO | ✅ | 编程处理前先转 docx |
| `.xlsx` / `.xlsm` | ✅ openpyxl | ✅ openpyxl | ✅ | 公式算值要过 LO |
| `.xls`（97-2003） | ✅ xlrd | ✅ LO | ✅ | xlrd 只能读 |
| `.csv` / `.tsv` | ✅ 标准库 | ✅ 标准库 | ✅ | 一律 UTF-8 |
| `.ods` / `.odt` / `.odp` | ✅ odfpy / LO | ✅ odfpy / LO | ✅ | |
| `.pptx` / `.potx` | ✅ python-pptx | ✅ python-pptx | ✅ | |
| `.ppt`（97-2003） | ✅ LO | ✅ LO | ✅ | |
| `.pdf` | ✅ PyMuPDF | ✅ LO 导出 | ✅ | 逆向见下 |
| `.wps` / `.et` / `.dps`（**金山**） | ❌ | ❌ | ❌ | 私有格式，本机不支持 |
| `.wps`（**微软 Works**） | ⚠️ LO 的 MS_Works 过滤器 | ❌ | ❌ | 与金山同名不同物，别混淆 |

## 2. LibreOffice 过滤器名（自动推断不对时才需要）

| 目标 | 过滤器名 | 备注 |
| --- | --- | --- |
| doc | `MS Word 97` | |
| xls | `MS Excel 97` | |
| ppt | `MS PowerPoint 97` | |
| docx | `Office Open XML Text` | |
| xlsx | `Calc MS Excel 2007 XML` | |
| pptx | `Impress Office Open XML` | |
| odt / ods / odp | `writer8` / `calc8` / `impress8` | |
| csv | `Text - txt - csv (StarCalc)` ＋ 参数 `44,34,76,1` | `,` 分隔、`"` 定界、**UTF-8**、首行 |
| txt | `Text (encoded)` ＋ 参数 `76,0,0,0` | UTF-8 |
| pdf | 由 LO 按文档类型选（writer/calc/impress_pdf_Export） | 不要手写过滤器名 |

`--convert-to "csv:Text - txt - csv (StarCalc):44,34,76,1"` —— 不钉 charset 的话中文经常会乱。

**LO 做不到的**：`pdf → docx/odt`，会报 `Error: no export filter`（PDF 只能导进 Draw）。
要 PDF → Word 必须用 pdf2docx（`pdf-to-docx` 子命令）。

**LO 算不出来的公式**（写进交付文件就是烘死的 `#NAME?`）：`XLOOKUP / XMATCH / SORT / FILTER /
UNIQUE / SEQUENCE` 任何前缀都不认；`TEXTJOIN / CONCAT / IFS / SWITCH / MAXIFS / MINIFS` 必须写成
`_xlfn.TEXTJOIN` 这种带前缀形式，裸写就报 `#NAME?`；写回后函数名变小写 = LO 没解析成功。
公式重算与错误扫描见 [`quality-gates.md`](quality-gates.md) 第 ② 节。

## 3. 中文字体替换（本机无微软字体）

| 请求的字体 | 实际替换为 |
| --- | --- |
| 宋体 / 黑体 / 微软雅黑 | Noto Sans CJK SC |
| SimSun | Noto Serif CJK SC |
| Times New Roman | Liberation Serif |
| Arial | Liberation Sans |
| Calibri | Carlito |
| Cambria | Caladea |
| 本机自带中文字族 | 24 个（Noto Sans/Serif CJK 全系 + 文泉驿 + Droid Fallback） |

生成文件时**显式写 `Noto Sans CJK SC`**（`writers.CJK_FONT`），别赌替换结果。
对外交付要声明"字体替换、分页可能与 Word 有差异"。

## 4. 规格 JSON

### docx-new
```json
{"title": "文档标题", "font": "Noto Sans CJK SC",
 "blocks": [
   {"h": 1, "text": "一级标题"},
   {"p": "正文段落", "size": 10.5},
   {"bullets": ["要点一", "要点二"]},
   {"table": {"rows": [["列A","列B"],["甲","1"]], "header": true}},
   {"pagebreak": true},
   {"text": "无样式段落"}
 ]}
```

### xlsx-new
```json
{"sheets": [
  {"name": "数据", "header": true, "freeze": "A2", "widths": [14, 10],
   "rows": [["项目","值"],["单价",3.5],["数量",4],["合计","=B2*B3"]]}
]}
```
单元格字符串以 `=` 开头 → 当公式写入；**openpyxl 不写缓存值**（读回来是 `None`），所以交付前跑
`office_cli.py xlsx-recalc b.xlsx` 让 LO 算一遍并写回，顺便扫 7 类错误（实测 `=B2*B3` → 14）。
只想看一眼值也可以 `convert b.xlsx --to csv` 后读 CSV，但那**不会写回 xlsx**。

### pptx-new
```json
{"title": "封面标题", "subtitle": "副标题", "size": "16:9",
 "slides": [
   {"title": "第一页", "bullets": ["要点 A", "要点 B"], "notes": "演讲者备注"},
   {"title": "数据页", "table": {"rows": [["季度","营收"],["Q3","1234"]]}}
 ]}
```
全部用空白版式 + 文本框，字体由调用方钉死（不依赖主题字体）。
缩略图：`office_cli.py thumb a.pptx -o thumbs/`。

## 5. LibreOffice 调用姿势

```bash
soffice --headless --norestore --nolockcheck --nologo \
        -env:UserInstallation=file:///tmp/lo-$$ \
        --convert-to pdf --outdir /out /in/a.docx
```
- `-env:UserInstallation` **必须每次唯一**，否则并发调用互相踢；
- 判成败看产物是否生成且非空，不要只看返回码（`dconf-CRITICAL` 之类的噪声不影响结果）；
- 首次调用要建 profile，约 1–2 秒，之后每次调用仍有该开销（本工具用完即删）。
