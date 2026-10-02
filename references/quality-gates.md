# 交付前的四道闸门（怎么做、为什么、坑在哪）

> 结论前置：**"命令退出 0" 不等于"文件是对的"**。四道闸门分别管包结构、几何、公式值、填字位置；
> 全过才交付。`scripts/selftest.py` 里四道都有真跑用例，改代码后先跑它。

| 闸门 | 命令 | 挡住什么 | 典型症状（不查就等着出事） |
| --- | --- | --- | --- |
| ① 包结构 | `office_cli.py ooxml-check a.docx` | ZIP 破损 / 内容类型没盖住 / 关系表断链 / `r:id` 断链 / XML 不良构 | 本机 LibreOffice 照渲染，**到 Word 弹"发现无法读取的内容"**，修完图片或超链接没了 |
| ② 排版几何 | `office_cli.py form-verify spec.json` | 表格列宽没钉死、标题行距把表格顶边顶开 | 分隔线跑到 11.05cm（要 5.50cm），肉眼"差不多"其实差半个字 |
| ③ 公式与值 | `office_cli.py xlsx-recalc a.xlsx` | openpyxl 写的公式**没有缓存值**；`#REF!/#NAME?` 等 7 类错误 | pandas/`data_only=True`/预览器读回来全是 `None` |
| ④ 填字框位 | `pdf-boxcheck spec.json` ＋ `pdf-fill --preview` | 框太小装不下、框与框重叠、框跑出页 | 字叠在一起、第二行跑出格子；**看图 3 秒就能发现** |

---

## ① docx 的 run 碎片：为什么"查找替换"会静默漏词

Word / LibreOffice 会把一句话拆成很多 `<w:r>`（`rsid` 修订号、拼写检查标记、编辑历史）。
屏幕上是一整句「北京市朝阳区」，`word/document.xml` 里可能是：

```xml
<w:r><w:t>北京</w:t></w:r><w:r><w:t>市朝</w:t></w:r><w:r><w:t>阳区</w:t></w:r>
```

于是**逐 run 找词一个都不命中**：文件看上去没坏，其实一个字没改。

还有一层：`<w:t>` 不带 `xml:space="preserve"` 时两侧空白会被吃掉
（`<w:t>Hello </w:t><w:t>world</w:t>` 渲染成 `Helloworld`），所以匹配必须按**渲染后的字符串**算。

`docx_runs.py` 的做法：

| 步骤 | 做什么 | 为什么 |
| --- | --- | --- |
| `strip_proof_errors()` | 删 `<w:proofErr>` | 它夹在 run 之间挡住合并，且不参与渲染 |
| `merge_runs()` | 相邻且 `<w:rPr>` 完全相同的 run 并成一个；同一 run 内相邻 `<w:t>` 合并 | 渲染结果不变的等价变换 |
| `_render()` | 非 `preserve` 的 `<w:t>` 两端空白按渲染语义去掉 | 匹配的是纸上的字，不是字节 |
| `replace_in_document()` | 在渲染文本上定位，再把改动摊回各 `<w:t>` | 跨 run 也能命中，且**保留每个 run 原有字体**，不是整段重建 |

刻意不做：跨 `<w:ins>/<w:del>` 边界不合并（会改写修订结构）；不动 `<w:delText>`（那是被删内容）；
文档里已有修订时返回值带 `warning` + `revisions`，改完请用 `docx-review` 盘一次。

**一处没命中就退出码 4**（`ok:false`）—— 静默"0 处替换"最容易被当成成功。

---

## ② xlsx 公式：openpyxl 不写缓存值

`ws["B1"] = "=SUM(A1:A3)"` 只是把公式当**字符串**塞进 XML，**不产生缓存值**。
于是 `load_workbook(data_only=True)`、pandas、大多预览器读到的都是 `None`。

```bash
python3 scripts/office_cli.py xlsx-recalc a.xlsx            # 原地重写，返回 JSON
python3 scripts/office_cli.py xlsx-recalc a.xlsx --force    # 明知会毁外链/VBA 也要算
```

实现：独立 profile 起 LibreOffice → Basic 宏 `calculateAll()` + `store()` + `StarDesktop.terminate()`
→ 扫每个单元格的 7 类错误串。三个"看起来成功其实失败"的坑：

- **宏里只 `store()` 不 `terminate()`**：soffice 永远不退出（实测挂满 45s 超时），而文件其实已经写好了；
  kill 掉的**孤儿进程还占着文件锁**，下一次重算会卡在"文档已被占用"上，像随机挂死。
  ⇒ 宏里必须 `terminate()`，超时后按 profile 路径精确清场（别用 `pkill -x soffice.bin`，会杀掉用户自己的 LO）。
- **退出码 0 但文件没被重写**：用 mtime+size 指纹判定，不看退出码。
- **外链**（`='[1]Sheet'!A1`）：openpyxl 存盘会剥掉缓存值，LO 再解析就得到 `#NAME?`，**外链被永久删除**
  ⇒ 默认拒绝重算并列出错在哪几格；带 VBA 的 `.xlsm` 同理，除非 `--force`。

`errors_found` 也退出码 3（不是 0）：**别拿退出码当结论，看 `status`**。
另外：重算只证明公式**算得出来**，不证明**算得对** —— 范围错一位照样干净，抽查 2–3 个值。

### LibreOffice 不认哪些函数（写在公式里就是 `#NAME?`，会烘进交付文件）

| 情况 | 例子 | 结论 |
| --- | --- | --- |
| Excel 2007 及以前 | `SUMIFS` `INDEX` `MATCH` `IFERROR` `SUMPRODUCT` | 直接用，无需前缀 |
| 2007 之后但 LO 认 | `TEXTJOIN` `CONCAT` `IFS` `SWITCH` `MAXIFS` `MINIFS` | 必须写成 `_xlfn.TEXTJOIN` 这种带前缀形式（Excel 内部就这么存，界面把前缀藏了）；写裸名 → `#NAME?` |
| LO 完全不认 | `XLOOKUP` `XMATCH` `SORT` `FILTER` `UNIQUE` `SEQUENCE` | 任何前缀都不行；查找用 `INDEX/MATCH`，排序/筛选/去重在 Python 里做完再写入 |
| 线索 | 写回后函数名**变成小写** | 说明 LO 没解析成功，旁边多半就是 `#NAME?` |

实测：`=TEXTJOIN(",",TRUE,A1:A3)` 裸写 → 重算后 `D1 = #NAME?`。

---

## ③ OOXML 包结构：`ooxml-check` 查什么

docx/xlsx/pptx 就是 **ZIP ＋ 一堆 XML**。本机渲染正常、到 Word 才报损坏的，基本是这几类：

| 查什么 | 判据 | 怎么破 |
| --- | --- | --- |
| ZIP 层 | CRC 校验、重复条目、`..`/绝对路径、symlink 条目 | 外部来的文件先删 symlink 条目再处理 |
| 内容类型 | 每个部件都被 `[Content_Types].xml` 的 Default 或 Override 覆盖；Override 不得指向不存在的部件 | 删部件忘了删 Override 是最常见的一种 |
| 关系表 | 每个 `.rels` 的 Target 指到的部件要真实存在 | 图片/样式表被删但关系还留着 |
| `r:id` 引用 | 部件里出现的 `r:id/r:embed/r:link/...` 必须在同目录 `.rels` 里有定义 | **图片、超链接断链的头号原因** |
| XML 良构 | 每个 `.xml/.rels` 能被解析 | 手改 XML 时打错一个标签 |
| 库兜底 | python-docx / openpyxl / python-pptx 能打开 | 前五项都过、库打不开说明还有别的问题 |

自己造坏样本自测（比等 Word 报错快）：

```bash
python3 - <<'PY'
import zipfile
# 删一个部件但保留 Content_Types 里的 Override → 必须被判 ok:false
with zipfile.ZipFile("good.docx") as zin, zipfile.ZipFile("broken.docx","w",zipfile.ZIP_DEFLATED) as zout:
    for n in zin.namelist():
        if n != "word/styles.xml":
            zout.writestr(n, zin.read(n))
PY
python3 scripts/office_cli.py ooxml-check broken.docx
```

**边界**：这只判"能不能打开"，不判排版与内容对不对（那是 ② 和 `thumb` 的活）。

---

## ④ 静态 PDF 填字：spec 格式与自检

中国绝大多数表格是**画线的静态版式**，没有 AcroForm 字段。流程是三步：

```bash
python3 scripts/office_cli.py pdf-structure form.pdf -o s.json      # 抽标签/横竖线/勾选框/行
# 人工（或按 s.json）写出 spec.json
python3 scripts/office_cli.py pdf-boxcheck spec.json               # 先自检，不通过就别填
python3 scripts/office_cli.py pdf-fill form.pdf spec.json -o out.pdf --preview p.png
```

spec：

```json
{"unit": "cm",
 "pages": [{"page": 1, "size_cm": [21.0, 29.7]}],
 "fields": [
   {"page": 1, "label": "姓名", "label_box": [2.1, 3.4, 3.6, 4.0],
    "box": [3.8, 3.4, 8.0, 4.0], "text": "张三", "font_size": 12,
    "align": "left", "valign": "middle"},
   {"page": 1, "label": "证件号码", "box": [3.8, 4.6, 9.0, 5.4],
    "text": "110105199001011234", "font_size": 11, "autofit": true, "max_lines": 1}
 ]}
```

| 字段 | 含义 |
| --- | --- |
| `unit` | `cm`（默认）或 `pt`；坐标原点在**左上角**，y 向下 |
| `box` | 填写区 `[x0,y0,x1,y1]` |
| `label_box` | 标签文字区（可选）：用来查"字会不会压到标签上" |
| `font_size` | 磅；`align` = left/center/right/justify；`valign` = top/middle/bottom |
| `autofit` ＋ `min_font_size` | 装不下就按框缩字号（5pt 起） |
| `max_lines` | 最多几行；表格格子通常写 1 |

踩过的坑：

- **内置中文字体（`china-s`）是全角的**：数字、字母也按全角排。18 位身份证号在 11pt 下要 198pt 宽，
  普通格子根本放不下 ⇒ 要么 `autofit`、要么 `max_lines: 1`。
- **量字号能不能装下，不能用 `Font("china-s").text_length()`**：它给 6.06pt/字符，
  而真正排版用的那个 Heiti CID 字体是 11pt/字符 —— 量出来 109pt、实际 198pt，自检放行、填下去换行溢出。
  正解是**空跑一次排版**（同字体同字号同框宽）数行数与高度（`pdf_forms._measure`）。
- **垂直居中不要缩小框**：把 32pt 高的框缩成 23pt 再 `insert_textbox`，会被误判成"要换行"而报溢出。
  正确做法是**整体下移**这个框，高度不变。
- `--preview` 出的是**红框＝填写框、蓝框＝标签框**的校验图；一眼能看出框错位、压字、串行。
  图是画在**副本**上的，交付的 PDF 干净。
- 填写写的是**文字层**（可选中、可搜索），不是真实表单域；要真表单域只能用于本来就是 AcroForm 的文件
  （`pdf-forms`）。

---

## 串成一条工作流

```
读/复原 → 生成或改 → ① 包结构 → ② 几何（排版类）→ ③ 公式（表格类）→ ④ 框位（填字类）→ 出图看一遍 → 交付
```

出图那步别省：`thumb a.pptx -o thumbs/` 或 `pdf-fill --preview`，**用眼睛（或让模型看图）复核一次**，
比读一堆坐标快，也是唯一能发现"内容放错地方"的手段。
