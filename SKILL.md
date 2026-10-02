---
name: office-docs
description: 当用户要读、写、改、转、生成 Word（docx/dotx/doc）、Excel（xlsx/xlsm/xls/csv/ods）、PowerPoint（pptx/ppt/odp）、PDF 文档时使用。触发包括：从文档里抽文字或表格、按模板生成报告/演示文稿/表格、Word 查找替换、Excel 公式算值、PDF 合并拆分/转图片/OCR/转可编辑 Word、给文档出缩略图核对排版、判断文档真实格式。本机无 WPS、无 MS Office，全部靠 LibreOffice headless ＋ python-docx/openpyxl/python-pptx/PyMuPDF/pdf2docx/poppler/tesseract 实现。金山私有格式 .wps/.et/.dps 不在支持范围。
---

# office-docs —— 本机文档读写与转换

**一句话**：这台机器没装 WPS、也没装微软 Office，但 Word / Excel / PPT / PDF 照样能读出来、改掉、
转格式、从零生成、渲染成图核对。

## 一、先跑 doctor（任何文档活之前）

```bash
python3 scripts/office_cli.py doctor      # {"ok": true, ...} 才算环境齐
```

`ok:false` 时看 `missing_libs` 与 `tools`：缺库用
`python3 -m pip install --user --break-system-packages <包>`（本机 PEP 668，**不动** dist-packages），
缺 `soffice` 就得装 libreoffice。

### 中文字体（先看 `font_aliases_bad`，空数组才算好）

本机**有**宋体/仿宋/楷体字库，但 Windows 字体名要靠一层 fontconfig 别名接过去；没这层别名时
`宋体` 会被兜到 **Noto Sans CJK（无衬线）**——宋体渲染成黑体，而 docx 里存的字名仍是「宋体」，
在 Windows 上打开是对的、本机出 PDF 是错的，**光看文件发现不了**。

```bash
bash scripts/install_fonts.sh --check   # 体检（只读）
sudo bash scripts/install_fonts.sh      # 装 4 套字库 + 别名 + fc-cache（幂等）
```

映射：宋体/SimSun/方正小标宋 → 文鼎PL简报宋（衬线）· 仿宋/FangSong → cwTeX 仿宋 ·
楷体/KaiTi → 霞鹜文楷 · 黑体/SimHei/微软雅黑/等线 → Noto Sans CJK。
**docx 里照写「宋体」「仿宋_GB2312」即可**，本机与 Windows 两端都对，不需要换字体名。

## 二、决策表：要做什么 → 走哪条路

| 任务 | 用什么 | 备注 |
| --- | --- | --- |
| 抽文字/表格看内容 | `office_cli.py text <file>` | docx/xlsx/pptx/pdf/odt/ods/ods/老版 doc/xls/ppt 都行 |
| 看结构（页数、表数、幻灯片数、sheet 名） | `office_cli.py info <file>` | 按**内容**判类型，不信扩展名 |
| 改格式 / 导出 PDF | `office_cli.py convert a.docx --to pdf` | LibreOffice headless，任意 Office 格式互转 |
| 表格文件 → CSV | `office_cli.py table a.xlsx -o out.csv` | 一律 UTF-8；中文不乱码 |
| Excel 公式要**算出来的值** | `office_cli.py xlsx-recalc a.xlsx` | openpyxl 只写公式**不算值**（读回来是 None）⇒ 交付前必须重算；顺手扫 `#REF!/#NAME?` 等 7 类错误。只想看值可以 `convert --to csv`（不算回 xlsx） |
| 生成 Word | `office_cli.py docx-new --spec s.json -o a.docx` | 规格见下 |
| 生成**中文公文**（GB/T 9704-2012） | `docx-new --spec s.json -o a.docx --profile gongwen` | 版心 3.7/3.5/2.8/2.6 cm、三号仿宋、29 磅固定行距、层次自动编号（黑/楷/仿）、页码 `— 1 —` |
| 生成**合同 / 报告** | `--profile contract` / `--profile report` | 小四宋体 1.5 倍行距；合同条款自动"第一条、第二条" |
| 生成**汇报 PPT** | `pptx-new --spec s.json -o a.pptx --profile brief` | 16:9、封面 32pt、正文 18pt、单页要点自动截到 6 条 |
| 本机出图（旧环境兜底） | 上面任一命令再加 `--portable` | 把字体名换成 Noto CJK。**本机已装字库+别名，正常不需要**；只在没有别名的机器上兜底 |
| Word 查找替换 | `office_cli.py docx-replace a.docx --set 旧=新` | 覆盖正文/表格/文本框/页眉页脚；**先合并 run 再按渲染文本找**，跨 run 也命中（一处没命中 → 退出码 4）。要手改 XML 先跑 `docx-merge-runs` 把碎片并了 |
| 生成 PPT | `office_cli.py pptx-new --spec s.json -o a.pptx` | 封面＋要点＋表格＋备注 |
| 核对排版 | `office_cli.py thumb a.pptx -o thumbs/` | 任意文档 → PDF → PNG |
| PDF 合并/拆分 | `pdf-merge` / `pdf-split` | 拆分支持 `--every N` 或 `--ranges 1-3,4-4` |
| PDF 取图 | `pdf-images` | 逐页抽内嵌图 |
| 扫描件取字 | `pdf-ocr scan.pdf -l chi_sim+eng` | 先渲染再 tesseract |
| PDF → 可编辑 Word | `pdf-to-docx a.pdf -o a.docx` | **保真度有限，必须人工核** |
| PDF 旋转 | `pdf-rotate a.pdf --angle 90 [--pages 1-3]` | 角度须为 90 的倍数 |
| PDF 水印 | `pdf-watermark a.pdf --text 机密` | 中文走 PyMuPDF 内置 `china-s` 字体 |
| PDF 加密/解密 | `pdf-encrypt a.pdf --password X` / `pdf-decrypt` | AES-256；密码只用一次、不落盘 |
| PDF 表单 | `pdf-forms a.pdf` / `pdf-forms a.pdf --fill 姓名=张三` | AcroForm 字段；缺 `/DR` 的表单已自动补 |
| **静态 PDF（扫描件/画线表格）填字** | `pdf-structure a.pdf -o s.json` → 写 `spec.json` → `pdf-fill a.pdf spec.json -o out.pdf --preview p.png` | 非 AcroForm；先 `pdf-boxcheck spec.json` 查"装不下/重叠/出页"，填完看 `--preview` 图上红框对不对 |
| 交付前确认 Word/Excel 能打开 | `office_cli.py ooxml-check a.docx` | 查 ZIP/内容类型/关系表/`r:id` 断链/XML 良构；这几类坏法**本机渲染看不出来**，到 Word 才弹"修复" |
| Excel 加图表 | `xlsx-chart a.xlsx --type bar --data '数据!B2:B3' --cats '数据!A2:A3'` | bar/line/pie |
| Word 修订批注盘点 | `docx-review a.docx` | 只读：插入/删除/格式变更条数＋批注作者正文 |
| **从用过的表单逆向出空白模板** | `form-render templates/xxx.json -o blank.docx` | **规格驱动**：模板文字人工校读一次，排版通用渲染 |
| **确认模板排对没排对** | `form-verify templates/xxx.json` | 渲染→出 PDF→**量每条表格线的 cm**→和 `spec.verify` 比；不达标退出码 3 |
| 学一张表单的结构 | `form-learn photo.jpg -o form.json` | OCR 只当起点（实测会读错），profile 要人工校对 |
| 用 profile 切字迹 / 出扫描件空白 | `form-extract photo.jpg form.json -o hw.png` / `form-blank …` | 靠 `fill_regions`（填写区）判定；参数见 `references/form-reverse.md` |

规格 JSON 例子（`docx-new` / `pptx-new` / `xlsx-new` 通用格式见 `references/formats.md`）：

```json
{"title": "季报", "blocks": [
  {"h": 1, "text": "一、结论"},
  {"p": "本季度营收…"},
  {"bullets": ["要点一", "要点二"]},
  {"table": {"rows": [["季度", "营收"], ["Q3", "1234"]]}}
]}
```

## 三、四条铁律（都是踩出来的）

1. **soffice 不能共用 profile 并发**：同一个 `UserInstallation` 起两个进程会互相踢。本工具每次调用开
   独立临时 profile（`convert.py` 内建），所以可以放心并发调用；自己手写 `soffice` 命令时必须带
   `-env:UserInstallation=file:///tmp/xxx`。
2. **别用管道判退出码**：`soffice … | tail` 的 `$?` 是 `tail` 的，失败也会报 0。本 CLI 的判据是
   **退出码 0 且 stdout 的 `{"ok": true}`**；自己写 shell 时用 `${PIPESTATUS[0]}`。
3. **bash 工具的 `/tmp` 不跨调用共享**：脚本中间产物要么写在当前工作目录，要么在同一次调用里用完。
4. **docx 的表格列宽不钉死就不生效**：只写 `cell.width = Cm(x)`，LibreOffice/Word 照样按内容自动适配
   ——实测分隔线跑到 11.05cm（要的是 5.50cm）。必须同时给 `tblLayout=fixed` + `tblW` + **重建 `tblGrid`
   的 `gridCol`**（`form_template._fix_table_layout` 已封装）。同理，要把表格顶边钉到某个 cm，标题段
   得用**固定行距 `line_pt`**，靠默认行距会被字体度量顶开 0.5cm 以上。
   **结论：排版类改动一律 `form-verify` 量一遍，别用肉眼当判据。**

## 四、交付前的四道闸门（"改完就算完"是错的）

| 闸门 | 命令 | 挡住什么 |
| --- | --- | --- |
| ① 包结构 | `ooxml-check a.docx` | 内容类型/关系表/`r:id` 断链 —— 本机照渲染，**到 Word 才弹"修复"** |
| ② 排版几何 | `form-verify spec.json` | 表格线位漂移、列宽没钉死（`tol_cm` 内才算过） |
| ③ 公式与值 | `xlsx-recalc a.xlsx` | openpyxl 写的公式**没有缓存值**（读回来全 None）、`#REF!/#NAME?` |
| ④ 填字框位 | `pdf-boxcheck spec.json` ＋ `pdf-fill --preview` | 装不下/重叠/出页；红框校验图一眼看出填歪 |

共同点：**别用"文件存在"当成功**。原理、坑与 spec 格式全表 → [`references/quality-gates.md`](references/quality-gates.md)。

## 五、能力与边界（诚实版，别过度承诺）

- **金山私有的 `.wps/.et/.dps` 打不开**。LibreOffice 注册表里的 `MS_Works`（`UserData=WPS`）是
  **微软 Works** 的 .wps，与金山 WPS 无关，别当支持。
- **中文字体已就位但仍是"替身"**：宋体→文鼎PL简报宋、仿宋→cwTeX 仿宋、楷体→霞鹜文楷、黑体→Noto Sans
  CJK。**字形、字宽与微软原字体不完全一致**，因此分页/行距不会与 Word 像素级一致；对外交付要说明。
  装了 WPS 或微软字体的机器上，同一个 docx 会渲染得更接近原件。
- **PDF → Word 是逆向重建**：pdf2docx 实测能把文字取回，但表格常被还原成文本行、多栏会串。
  重要文档必须人工核对，或让对方给原始 docx。
- **CSV 导出必须钉 UTF-8**（本工具已固化 `charset=76`），否则中文变乱码。
- **旧版二进制格式**（doc/xls/ppt）只能靠 LibreOffice 读；要编程处理先 `convert` 成 OOXML。
- **PDF 加密要 `cryptography`**：缺了就报 `DependencyError`（本机已装）。表单填写已自动补 `/DR`，
  但只覆盖 AcroForm 文本框/勾选框，XFA 表单（`/XFA`）不支持。
- **Word 修订/批注只做盘点**（`docx-review`）；替换正文时若文档里有修订标记，返回值带 `warning` 与
  `revisions` 条数 —— **替换只改可见文字，不动修订结构**。`ooxml-check` 判"能不能打开"，不判排版对不对；
  `pdf-fill` 写的是文字层，不是真实表单域（要真表单域只能对本来就是 AcroForm 的文件用 `pdf-forms`）。

## 六、目录

```
scripts/
  office_cli.py   统一入口（JSON 信封，ok 字段）
  convert.py      LibreOffice headless 封装 + doctor（并发安全、超时、产物校验）
  readers.py      抽文本 / 看结构 / 读表（按内容判类型）
  writers.py      生成与修改 docx/xlsx/pptx（中文字体显式指定）
  cn_profiles.py  中文公文/合同/报告/汇报 PPT 版式（GB/T 9704-2012 与行业惯例）
  pdf_ops.py      PDF 合并/拆分/取图/渲染/OCR/转 Word
  extras.py       PDF 旋转/水印/AES 加解密/表单，Excel 图表，Word 修订批注盘点
  docx_runs.py    run 合并 + 按渲染文本查找替换（跨 run 命中；`xml:space` 语义）
  xlsx_recalc.py  LO 宏重算公式并写回 + 7 类错误扫描 + 外链/VBA 拒跑
  pdf_forms.py    静态 PDF：抽结构 / 框位自检 / 按框填字 / 出红框校验图
  ooxml_check.py  OOXML 包结构闸门（内容类型、关系表、r:id 断链、XML 良构）
  form_template.py 表单模板逆向：rectify/网格/OCR profile、规格渲染空白模板、verify 几何自检、填写区切字迹
  install_fonts.sh 装中文字库 + Windows 字体名别名（`--check` 只体检）
  selftest.py     端到端自检（31 项，真跑真断言）
assets/           99-cjk-windows-alias.conf —— fontconfig 别名（install_fonts.sh 装它）
templates/        表单规格（JSON，数据）：`住所经营场所证明.json` 等；配套 `*.scan-profile.json`
references/formats.md        格式、过滤器名、字体替换、规格 JSON 全表（含 LO 公式支持）
references/quality-gates.md  四道交付闸门：run 碎片、公式重算、OOXML 硬伤、静态 PDF 填字 spec
references/cn-doc-formats.md 中文公文/商业文书排版速查（字号磅值、逐要素字体、对稿清单）
references/form-reverse.md   表单逆向：profile 字段、fill_regions 怎么标、字号怎么反推
```

改完代码先跑 `python3 scripts/selftest.py`（退出码 0 = 31/31 通过），再交付。
