# office-docs-skill

**Word / Excel / PowerPoint / PDF 的读、写、改、转、生成 —— 既能当 Agent Skill，也能直接当 CLI 用。**
它把 LibreOffice headless 与 Python 库包在一起，因此在**没装 Microsoft Office（也没装 WPS）的机器上**照样能干活。

[English](README.en.md) ｜ MIT

---

## 依赖

| 类别 | 需要什么 |
| --- | --- |
| 系统命令 | `soffice`（LibreOffice）、`pdftoppm` / `pdftotext`（poppler-utils）、`tesseract`（中文 OCR 需 `chi_sim`） |
| Python 包 | `python-docx`、`openpyxl`、`python-pptx`、`xlrd`、`odfpy`、`pymupdf`、`pypdf`、`pdf2docx`、`cryptography`（仅 PDF AES 加解密） |
| 可选 | `pandoc`（Markdown ↔ docx 等） |
| 中文排版（可选） | CJK 字库 ＋ fontconfig 别名；`scripts/install_fonts.sh` 负责安装与体检 |
| 常用英文字体（可选） | 同一份别名文件把 Arial / Helvetica / Times New Roman / Georgia / Calibri / Cambria / Verdana / Tahoma / Courier New 接到本机的度量兼容字族 |

先跑 `python3 scripts/office_cli.py doctor` —— 它会把上面每一项列成 `ok` 或缺什么。
除此之外没有别的前提：不需要安装步骤、没有常驻服务、没有配置文件。

## 快速开始

```bash
python3 scripts/office_cli.py doctor          # 环境自检；{"ok": true} 才算齐
python3 scripts/selftest.py                   # 端到端自检；打印 N/N，全过才退出 0
bash    scripts/install_fonts.sh --check      # 中文字体/别名体检（只读）

python3 scripts/office_cli.py info  a.xlsx    # 按内容判真实类型与结构，不信扩展名
python3 scripts/office_cli.py text  a.pdf     # 抽文本
python3 scripts/office_cli.py convert a.docx --to pdf
python3 scripts/office_cli.py pptx-new --spec s.json -o out.pptx
python3 scripts/office_cli.py thumb out.pptx -o thumbs/
python3 scripts/office_cli.py pdf-ocr scan.pdf -l chi_sim+eng
```

**判据**：命令成功 = **退出码 0 且 stdout 是 `{"ok": true, ...}`**。出错时同一个信封里带错误信息，
退出码非 0。

## 交付前的四道闸门

"转换成功"不等于"文件是对的"。这四道检查专门挡那些**打开文件时看不出来**的坏法：

```bash
python3 scripts/office_cli.py ooxml-check out.docx    # ① 包结构：内容类型 / 关系表 / r:id / XML 良构
python3 scripts/office_cli.py form-verify spec.json   # ② 几何：渲染后量表格线位，和 spec.tol_cm 比
python3 scripts/office_cli.py xlsx-recalc out.xlsx    # ③ 公式：用 LibreOffice 重算并扫 #REF!/#NAME?
python3 scripts/office_cli.py pdf-boxcheck spec.json  # ④ 框位：填 PDF 之前查装不下/重叠/出页
python3 scripts/office_cli.py pdf-fill a.pdf spec.json -o out.pdf --preview p.png
```

原理、踩坑与 spec 格式见 [`references/quality-gates.md`](references/quality-gates.md)。

## 能做什么

| 方向 | 要点 |
| --- | --- |
| 读 | docx/xlsx/pptx/pdf/odt/ods 与旧版 doc/xls/ppt 的文本与表格；结构信息（页数、表数、幻灯片数）；按内容判真实格式 |
| 转 | 走 LibreOffice headless 任意格式互转，每次调用开独立 profile（可放心并发） |
| 生成 | `docx-new`、`xlsx-new`、`pptx-new` 按 JSON 规格生成；中文公文（`gongwen`）、合同、报告、汇报四种版式 |
| 改 | 查找替换覆盖正文/表格/文本框/页眉页脚；**先合并 run 再按渲染文本找**，跨 run 也命中 |
| Excel | 把公式算成真实值；加图表；导出 UTF-8 CSV |
| PDF | 合并、拆分、取图、旋转、水印、AES-256 加解密、OCR、转可编辑 Word、渲染成 PNG |
| 表单 | 列表单字段并填 AcroForm；静态 PDF 则先抽结构、自检框位，再按框填字并出红框校验图 |
| 表单逆向 | 从扫描件/照片学模板 profile（`form-learn`）、按规格渲染干净空白模板（`form-render`）、几何自检（`form-verify`）、切出字迹层（`form-extract` / `form-blank`） |
| 盘点 | Word 修订与批注清单（`docx-review`，只读） |

全部子命令见 [`SKILL.md`](SKILL.md)；格式、过滤器与规格 JSON 见
[`references/formats.md`](references/formats.md)。

## 作为 Agent Skill 使用

`SKILL.md` 是技能正文：front matter 里的 `name` / `description` 决定智能体什么时候加载它。
把整个目录放进你的技能目录并在那里启用 —— 例如 Claude Code 的 `~/.claude/skills/office-docs-skill/`、
DSH 的 `~/.dsh/skills/office-docs-skill/`。

`examples/gongwen-demo.json` 是一份可直接跑的公文规格（标题/主送机关/四层序号/附件/署名/附注）；
`examples/gongwen-demo.docx`、`gongwen-demo-portable.docx`、`gongwen-demo-portable.pdf` 与
`gongwen-demo-portable-1.png`／`-2.png` 是它的产物，用来目视核对版式。

`templates/` 放表单规格（JSON）：一张空白模板 ＋ 配套的扫描 profile，供 `form-render` / `form-verify` 使用。

## 边界

- **微软 Works / 金山 WPS 的私有格式**（`.wps`/`.et`/`.dps`）不支持。LibreOffice 里的 `MS_Works`
  过滤器是**微软 Works**，与金山 WPS 无关。
- **PDF → Word 是逆向重建，不是转换**：文字通常还在，但表格常被还原成文本行、多栏会串。
  重要文档必须人工核对，或让对方给原始 docx。
- **中文字形是"替身"**（没装微软字体时）：宋体 → AR PL SungtiL GB、仿宋 → cwTeX 仿宋、
  楷体 → 霞鹜文楷、黑体 → Noto Sans CJK。**只要求落到对的一类字**（宋体落衬线宋、黑体落无衬线黑体），
  字形与字宽与原字体有差，**分页不会与 Windows 上的 Word 逐像素一致**。
  docx 里照写正常字体名（`宋体`、`仿宋_GB2312`、`Arial`、`Times New Roman`）即可，两端都能解析。
  体检口径见 `doctor` 的 `font_aliases`：每个字体名给出落到的字族、是否对了类别、中文名还验是否支持中文。
- **旧版二进制格式**（doc/xls/ppt）只能靠 LibreOffice 读；要编程处理先转成 OOXML。
- **PDF 加密需要 `cryptography`**。表单填写只覆盖 AcroForm 字段；XFA（`/XFA`）不支持。
- **OCR 自动学的模板必须人工过一遍**：淡线或手机拍的表单，`form-learn` 会读错字。
  手写与印刷在同一行重叠时，改用"规格重排"的模板，别硬分扫描件。
- **排版类改动一律用 `form-verify` 量，不靠肉眼。**

## 目录结构

```
SKILL.md                  # 技能入口：规则、决策表、命令参考
README.md / README.en.md  # 中文 / 英文说明
scripts/
  office_cli.py           # 统一入口（JSON 信封，含 ok 字段）
  convert.py              # LibreOffice headless 封装 + doctor（并发、超时、产物校验）
  readers.py              # 抽文本 / 看结构 / 读表，按内容判类型
  writers.py              # 生成与修改 docx/xlsx/pptx
  cn_profiles.py          # 中文公文 / 合同 / 报告 / 汇报 PPT 版式
  pdf_ops.py              # PDF 合并、拆分、取图、渲染、OCR、转 Word
  extras.py               # PDF 旋转/水印/AES/表单，Excel 图表，Word 修订批注盘点
  docx_runs.py            # run 合并 + 按渲染文本查找替换
  xlsx_recalc.py          # LibreOffice 宏重算公式 + 错误扫描
  pdf_forms.py            # 静态 PDF：抽结构 / 框位自检 / 按框填字 / 出校验图
  ooxml_check.py          # OOXML 包结构闸门
  form_template.py        # 表单模板逆向：网格/OCR profile、规格渲染、几何自检
  install_fonts.sh        # 装中文字库 + Windows 字体名别名（--check 只体检）
  selftest.py             # 端到端自检，真跑真断言
assets/                   # fontconfig 别名文件（install_fonts.sh 装它）
templates/                # 表单规格（JSON）与配套扫描 profile
references/
  formats.md              # 格式、过滤器名、字体替换、规格 JSON 全表
  quality-gates.md        # 四道交付闸门：原理、坑、spec 格式
  cn-doc-formats.md       # 中文公文/商业文书排版速查
  form-reverse.md         # 表单逆向：profile 字段、填写区、字号反推
examples/                 # 公文 demo 规格与渲染产物
```

## 许可

MIT

---

> **说明**：本项目由智能体在日常使用中逐步积累而成 —— 命令、坑与判据都来自实际跑过的任务，
> 而非先写文档再实现。
