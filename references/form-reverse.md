# 表单逆向：从「用过的表单」还原空白模板

> 结论前置：**规格（spec）是主路，扫描件（profile）是辅路。** 规格渲染是声明式的、可自检的；
> 扫描件切分只能处理"印刷与手写不重叠"的部分。两条路可互相补，但别指望扫描件能恢复像素级原模板。

## 一、为什么不能靠像素特征分「印刷 / 手写」

2026-09-30 在一张真实表单照片上实测（矩形钢笔手写 + 激光印刷）：

| 特征 | 印刷 | 手写 | 能否分开 |
| --- | --- | --- | --- |
| 颜色 B−R | ≈ −12 | ≈ −12 | ✗ 完全重叠 |
| 笔画宽度 | 6.0 px | 4.4–8.4 px | ✗ 区间重叠 |
| 墨色深度 | — | — | ✗ 重叠 |
| 边缘锐度 | — | — | ✗ 重叠 |
| 3×3 开运算保留率 | 88% | 97% | ✗ 差 9 个点，不够 |
| **是否落在模板结构内** | 26% 的墨落在 OCR 词框外 | 86–100% 落在外面 | ✓ **唯一可用信号** |

所以流程是：**先有模板结构，再判定"落在结构外的墨"就是后填的**。没有模板就没有判据。

## 二、为什么扫描件不能还原「原模板几何」

`form_template.rectify()` 用纸的四角做透视校正。前提是**四个角都在画面里**。
实测那张照片：纸在**左、下、右三边都被裁掉**，只看得见上边与右上角 ⇒
校正后的"cm"只能保证**互相自洽的比例**，绝对尺度有约 ±10% 不确定。

判断方法（便宜）：逐行找纸张亮区的最长连续段。

```python
sm = cv2.GaussianBlur(gray, (0,0), 8)
_, th = cv2.threshold(sm.astype("uint8"), 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
# 每行取最长亮段 → 若某行 left==0 或 right==W-1，说明该边被裁
```

**四角不全 ⇒ 别声称精确 cm。** 此时正确的做法：
1. 行列尺寸取**实测比例**（横线间距 / 竖线间距的相互比值）；
2. 绝对数值按"能排进 A4 且视觉一致"来定；
3. 用 `form-verify` 把**渲染结果**钉在选定的目标 cm 上（自洽即可）。

绝对尺度若确实需要，只有两条路：重新拍一张四角齐全的、或者拿原件量。

## 三、量线位：什么方法在皱纸照片上能用

按可靠性从高到低（都在同一张照片上试过）：

| 方法 | 结果 |
| --- | --- |
| `cv2.adaptiveThreshold` + 覆盖率/最长暗串 | ✗ 光照不均 + 折痕，一条线都提不到 |
| HoughLinesP（容倾斜） | ✗ 只提到 3 条，且都是误检 |
| 形态学 blackhat + 阈值 | ✗ 被纸外深色背景淹没 |
| **窄条带一维亮度剖面**（`bg − img` 后找峰） | ✓ 可靠；见下 |
| **PDF 矢量线**（`form-verify` 走这条） | ✓✓ 最可靠，但没有原件时用不了 |

一维剖面用法：取一条**不含文字**的窄带（如左侧标签列的 x∈[1.7, 4.2]cm），
按行平均得到 1-D 曲线，`bg = cv2.medianBlur(img, 101)`，`d = bg − img`，
`d` 的局部极大就是线。**同一条线要在多个 x 位置各量一次**，因为皱纸照片有透视残留：
实测那张的竖分隔线从 y=5cm 的 5.26cm 漂到 y=18.5cm 的 4.28cm（约 4°），
单点测量会得出互相矛盾的结论。

## 四、字号怎么反推（照片没有参考物时）

用「**每行容纳多少个全角字**」反推，比直接量字高可靠得多——字高会被相邻行粘连撑大。

```python
# 一行印刷文字的字形跨度 ÷ 全角字数 = 字距（em）
#   须知正文 46 个全角字，可用宽 174mm  → 3.78mm/字 → 10.5pt（五号）
#   正文段  28 个全角字，可用宽 143mm  → 5.11mm/字 → 14pt（四号）
```

对应关系：`pt = mm / 0.3528`。中文全角字的**字距 = 字号**，所以这个方法误差很小。

`cv2.connectedComponentsWithStats` 数连通块再取中心距的做法**不可靠**：粘连与断笔会
让间距出现 1.5×/2× 的假峰，实测同一行给出 32.5px 与 26px 两种答案。

## 五、规格渲染的三个硬坑

1. **列宽不钉死就不生效**：只写 `cell.width = Cm(x)`，LibreOffice/Word 仍按内容自动适配。
   实测分隔线跑到 **11.05cm**（目标 5.30cm）。必须同时：
   `tblLayout=fixed` + `tblW` + **重建 `tblGrid` 里的 `gridCol`**。
   已封装成 `form_template._fix_table_layout(table, widths_cm)`。
2. **表格顶边会被标题段的行距顶开**：默认行距按字体度量伸缩，实测高 0.55cm。
   要钉位置就用**固定行距**：`{"line_pt": 30, "space_after_pt": 15.35}`。
3. **内容会悄悄溢出到第 2 页**：`pages` 必须显式断言。实测第一次渲染须知整行跑到第 2 页。

`form-verify` 把这三条一次性检查掉：渲染 → LO 出 PDF → 读**矢量线**（不是截图找线）→
和 `spec.verify.h_cm` / `v_cm` 比，超差或页数不对就 `ok:false`、退出码 3。

```bash
python3 scripts/office_cli.py form-verify templates/住所经营场所证明.json
# {"ok": true, "pages": 1, "h_max_delta": 0.06, "v_max_delta": 0.0, ...}
```

## 六、spec 字段速查

| 字段 | 位置 | 说明 |
| --- | --- | --- |
| `page.size` / `page.margins_cm` | 顶层 | A4 固定；margins 定表格左右边界 |
| `verify.pages` / `verify.h_cm` / `verify.v_cm` / `verify.tol_cm` | 顶层 | 自检目标（cm，从纸左上角）；公差默认 0.15 |
| `title.line_pt` / `space_before_pt` / `space_after_pt` | title | 钉标题块高度用 |
| `table.widths_cm` | table | **必填**，否则列宽不生效 |
| `rows[].height_cm` + `height_rule` | table | `exact` 钉死 / `atLeast` 只保底 |
| `cells[].valign` | cell | `top` / `center` / `bottom` |
| `cells[].padding_cm` | cell | 窄列要调小，默认左右各 0.19cm 会挤断行 |
| `cells[].merge_right` | cell | 横向合并的格数 |
| `cells[].paras[]` | cell | 段列表；每段可带 `align` / `font` / `size` / `bold` / `first_line_cm` / `space_after_pt` / `line_pt` |

## 七、profile（扫描件路）字段速查

`form-learn` 产出的 JSON，坐标都在 **rectify 后的页面系**（默认 A4@200dpi = 1654×2339 px，1px = 0.127mm）：

| 字段 | 说明 |
| --- | --- |
| `page.size` | [W, H] px |
| `grid.h` / `grid.v` | 检测到的横/竖线 `{pos, from, to}`；**淡线表单会漏检，人工补** |
| `words` | OCR 词框 `[x0,y0,x1,y1]` |
| `fill_regions` | **人工标**的填写区 `[x0,y0,x1,y1]`，`split_ink` 靠它判字迹 |
| `protect_margin` | 词框外扩多少 px 算"印刷字保护带" |
| `page.rectify` | rectify 的返回信息；`rectified:false` 表示没找到纸四角，坐标不可信 |

`fill_regions` 不要用 `grow_regions()` 自动长——实测会把印刷文字整片吞进来（8 → 140 个区域），
必须照着扫描件手工标。
