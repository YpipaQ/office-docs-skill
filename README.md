# office-docs

**本机的文档读写/转换技能**（DSH Agent Skill + 可直接当 CLI 用）。这台机器没有 WPS、也没有微软 Office，
本项目用 LibreOffice headless ＋ Python 库把 Word / Excel / PPT / PDF 的读、写、改、转、生成、渲染包起来。

**原创实现**：不依赖、不包含、不派生任何第三方技能包的材料（详见下方"来源与许可"）。

## 怎么跑

```bash
python3 scripts/office_cli.py doctor          # 环境自检，{"ok": true} 才算齐
python3 scripts/selftest.py                   # 端到端自检（31 项，真跑真断言；退出码 0 = 全过）
bash    scripts/install_fonts.sh --check      # 中文字体别名体检（只读，不改系统）
python3 scripts/office_cli.py info  a.xlsx
python3 scripts/office_cli.py text  a.pdf
python3 scripts/office_cli.py convert a.docx --to pdf
python3 scripts/office_cli.py pptx-new --spec s.json -o out.pptx
python3 scripts/office_cli.py thumb out.pptx -o thumbs/
python3 scripts/office_cli.py pdf-ocr scan.pdf -l chi_sim+eng
```

**交付前四道闸门**（原理、踩坑与 spec 格式见 [references/quality-gates.md](references/quality-gates.md)）：

```bash
python3 scripts/office_cli.py ooxml-check out.docx    # ① 包结构：内容类型/关系表/r:id 断链/XML 良构
python3 scripts/office_cli.py form-verify spec.json   # ② 几何：渲染后量表格线位，超差即 ok:false
python3 scripts/office_cli.py xlsx-recalc out.xlsx    # ③ 公式：LO 算出值写回 + 扫 #REF!/#NAME?
python3 scripts/office_cli.py pdf-boxcheck spec.json  # ④ 框位：装不下/重叠/出页
python3 scripts/office_cli.py pdf-fill a.pdf spec.json -o out.pdf --preview p.png
```

约定：**成功 = 退出码 0 且 stdout 是 `{"ok": true, ...}`**（本工作区口径）。所有子命令见
[`SKILL.md`](SKILL.md) 与 [`references/formats.md`](references/formats.md)。

## 作为技能使用

`SKILL.md` 是技能正文（YAML front matter 的 `name` / `description` 决定何时被触发）。
本机已把本项目软链到 `~/.dsh/skills/office-docs`，在「会话技能」里勾选即可加载。

`examples/gongwen-demo.json` 是一份可直接跑的公文规格（含标题/主送机关/四层序号/附件/署名/附注），
`gongwen-demo.docx` 与 `gongwen-demo-portable-1.png`（渲染成图）是它的产物，用来目视核对版式。

## 依赖

| 类别 | 需要什么 |
| --- | --- |
| 系统命令 | `soffice`（libreoffice）、`pdftoppm`/`pdftotext`（poppler-utils）、`tesseract`（含 `chi_sim`） |
| 中文字库 | `fonts-arphic-gbsn00lp`（宋）、`fonts-cwtex-fs`（仿宋）、`fonts-lxgw-wenkai`（楷）、`fonts-arphic-gkai00mp`（楷，备用）＋ fontconfig 别名；一键 `sudo bash scripts/install_fonts.sh` |
| Python（`~/.local`，`pip --user --break-system-packages`） | python-docx、openpyxl、python-pptx、xlrd、odfpy、pymupdf、pypdf、pdf2docx、cryptography（PDF AES 加解密用） |
| 可选 | pandoc（markdown ↔ docx 等） |

`python3 scripts/office_cli.py doctor` 会把上面逐项列出来。

## 当前状态

- 端到端自检 **31/31 通过**（生成 docx/xlsx/pptx → LO 转 PDF/CSV → 抽文本 → 替换 → 缩略图 →
  PDF 合并拆分 → 图片型 PDF 的 OCR → PDF 转 Word → 旋转/水印 → AES 加解密 → 表单填值 → Excel 图表 →
  Word 修订盘点 → 中文公文版式 → 合同/报告版式 → 汇报 PPT 版式 → **表单模板逆向（规格渲染 + 几何自检 + 填写区切字迹）**
  → **四道交付闸门（跨 run 替换、run 合并、公式重算、OOXML 包体检、静态 PDF 填字）**）。
- **交付闸门**：`ooxml-check`（包结构）、`form-verify`（几何）、`xlsx-recalc`（公式值）、
  `pdf-boxcheck` + `pdf-fill --preview`（填字框位）。四道各有真跑用例，原理见
  [references/quality-gates.md](references/quality-gates.md)。
- **表单模板逆向**走两条路，互不替代：
  1. **规格 → 干净空白模板**（`templates/*.json` → `form-render` → docx/PDF）：文字人工校读、排版通用渲染，
     原件丢了也能重新出一份能打印的空白件；
  2. **扫描件 → 字迹层 / 扫描件空白**（`form-learn` 出 profile，人工校正 `fill_regions` 后
     `form-extract` / `form-blank`）：同一张表单登记一次，之后每张同款都能切。
  规格路带**几何自检** `form-verify`：渲染 → 出 PDF → 读矢量线位 → 与 `spec.verify` 比，超差即 `ok:false`。
  实测边界：手机拍的淡线表单，OCR 自动学的模板**不够准**（会读错字），profile 必须人工过一遍；
  手写与印刷在同一行重叠时（如填空数字），单靠扫描件无法完全分开 —— 此时以"规格重排"的模板为准。
  逆向方法与量测踩坑见 [references/form-reverse.md](references/form-reverse.md)。
- 中文版式见 [references/cn-doc-formats.md](references/cn-doc-formats.md)：字号磅值对照、公文逐要素字体、
  页边距/版心、对稿清单、商业文书惯例。`--profile gongwen|contract|report|brief`。中文字体名已由 fontconfig 别名接好，**不需要** `--portable`。
- 已知边界（都写在 `SKILL.md` 第四节）：金山 `.wps/.et/.dps` 不支持；PDF→Word 保真度有限；
  本机用文鼎宋/cwTeX 仿宋/霞鹜文楷/Noto Sans 替代微软字体，**字形字宽有差，分页与 Word 不会逐像素一致**；
  PDF 加解密依赖 `cryptography`；表单只覆盖 AcroForm（XFA 不支持）；Word 修订只盘点、不自动改。

## 最后验证

2026-09-30（自检 **31/31**；LibreOffice 25.2.3.2 / Python 3.13.5 / Debian 13 / 中文字体别名 9/9）。

## 来源与许可

全部代码为本项目原创，只调用公开 API（python-docx / openpyxl / python-pptx / PyMuPDF / pypdf /
odfpy / pdf2docx / cryptography 与 LibreOffice、poppler、tesseract 的命令行）。
**未复制、未派生任何第三方技能材料**，因此本项目自身不携带他人许可限制。

> 说明：后来见过一份**声明为 Proprietary 的第三方文档技能包**（宿主共享目录里），只作**思路参考**
> （例如"每次写入后都要有一道校验"的组织方式、LibreOffice 函数支持范围这类**事实**）。
> 其代码与文字**一律未复制**：本项目的实现、模块划分、CLI 与文档都是自己的。
> 判断依据记在记忆区该日笔记的"第三轮"一节。

## 过程与结论

过程、踩坑与验证边界记录在记忆区（工作区外，绝对路径）：
`/home/ypipaq/Work/memory/2026-09-30/本机文档处理能力与官方技能许可证.md`。
