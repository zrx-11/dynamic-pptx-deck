# 版面与换图工具集（textfit / grid_flow / refill_images）与后处理检查

> 本文件规定 `scripts/textfit.py`、`scripts/grid_flow.py`、`scripts/refill_images.py`
> 三个脚本的用途、CLI、返回结构与在流水线四道闸中的位置；并记录 `audit_all.py`
> 三条后处理检查的判定口径、Morph 结构对齐核对结论。所有脚本走 python-pptx + EMU
> 路线，坐标口径 1 px @1280 画布 = 9525 EMU；不依赖 SVG 画布坐标，不保留上游
> 脚本中的路径占位。

## 1. 定位

| 脚本 | 落在哪道闸 / 工序 | 解决的问题 |
|---|---|---|
| `textfit.py` | 容量闸（装不装得下）+ 内容框闸（填没填满） | 框空 / 留白 / 塞不进：真实字体量测、缩字号→加高→减字顺序、空框给字数区间 |
| `grid_flow.py` | 兜底页排布 | 多项内容均衡网格：列数回推均衡、同行等高、余量按比例分配 |
| `refill_images.py` | 导入换图 | 成品页复用换图：只换 blob，不动几何 / 动画 / 转场 / 命名 |

上游参考为 SVG→PPTX 管线产物（1280×720 px、Windows 路径占位）；本路线三个脚本
均按 python-pptx / EMU 独立重写，未拷用上游坐标与路径。

## 2. textfit.py

### 2.1 断行规则

- CJK 按字断行；Latin 按词断行。
- 数字与单位整体不拆：`239 MB`、`1.4 GB`、`17,374,215`、`24.4%`、`0.9%` 等
  （千分位逗号、小数点、百分号与常见单位作为一个 token）。
- 行首禁则：`，。、；：？！）］｝」』·` 不落在行首；行尾禁则：`（［｛「『`
  不落在行尾。

### 2.2 适配顺序（严格按序，不跳步）

1. 缩字号至下限（默认 10 pt）；
2. 加高容器（只改高度，不改宽度）；
3. 报告 `reduce_chars` 并给出目标字数。

禁止改宽度；除字号与容器高度外，不改任何几何。

### 2.3 量测口径

- PIL + 真实字体量测，候选字体：Source Han Sans CN Regular → Noto Sans CJK SC
  → Noto Sans Mono CJK SC；pt→px 按 96/72。
- 不做「字数 × 固定宽度」估算。
- 跳过非文本槽位：含 `a:blip` 的形状、箭头（`*Arrow`）、自定义自由多边形、
  名为「图片占位」的图片槽位。

### 2.4 空框判定

文本框为空，或字数撑不满一行而容器可容 ≥2 行时，判 `empty_frame`（不是 `ok`），
并给出可补方案：`suggested_chars_min/max`、`capacity_lines`、`suggested_pt`。

### 2.5 返回结构

逐框 dict：

| 字段 | 含义 |
|---|---|
| `status` | `ok` / `shrunk` / `grown` / `reduce_chars` / `empty_frame` |
| `pt` | 判定字号（pt） |
| `lines` | 断行后的行列表 |
| `needed_px` | 该字号下所需高度（px） |
| `box_h_emu` | 容器高度 EMU（`grown` 时为建议加高后的高度） |
| `reasons[]` | 判定理由 |
| `suggestions` | 空框 / 减字方案：字数区间、容量行数、每行字数、建议字号 |

### 2.6 CLI

```bash
python3 scripts/textfit.py <pptx> [--page N] [--json out.json] [--apply out.pptx]
python3 scripts/textfit.py --selftest
```

`--apply` 只在 `shrunk`（改字号）与 `grown`（改高度）两种动作上真改文件；
`reduce_chars` 与 `empty_frame` 只报告，不改。

### 2.7 自检覆盖

`--selftest` 四条用例 + 一条反证，断言失败以非零退出码退出：

1. CJK 长句在窄槽内需缩字号（`shrunk`，字号小于起始值）；
2. 数字 + 单位不被拆开（五个样例 token 整词出现在断行结果中）；
3. 空文本框判 `empty_frame`，给出字数区间与容量行数；
4. 行首禁则不落在行首；
5. 反证：文字充足的框返回 `ok`，不得一律判空。

## 3. grid_flow.py

### 3.1 排布规则

1. **均衡列数**：列数先按 `max_cols` 上限取，再按行数回推。6 项 → 3×2；
   7 项 → 4+3；4 项 → 4×1；不允许「能塞 4 列就塞 4 列、最后一行只剩 1 个」。
2. **同行等高**：同一行槽位同一高度；多行时行高均分，不出现高低不齐。
3. **余量按比例分配**：剩余高度按行数均摊，取整余数逐行轮补；不得把余量
   全给单独一行拉伸。

### 3.2 返回结构

`{cols, rows, rects:[{x,y,w,h}], warnings[]}`，坐标均为 EMU。

### 3.3 CLI

```bash
python3 scripts/grid_flow.py --n 7 --area x,y,w,h --max-cols 4 [--json]
python3 scripts/grid_flow.py --selftest
```

`--area` 出入参为 EMU；默认区域 1280×720 px 折算
（12,192,000 × 6,858,000 EMU）。

### 3.4 自检覆盖

6 / 7 / 4 三个用例的列数、行数与行分配；5 项由朴素 4+1 纠正为 3+2 的反例；
同行等高与余量轮补（行高差不超过 1 EMU 取整误差）。

## 4. refill_images.py

### 4.1 支持的四类承载

- `bg`：幻灯片背景 `p:bg` 内的 `a:blipFill`；
- `root_pic`：根级 `p:pic`；
- `group_pic`：组合 `p:grpSp` 内的 `p:pic`；
- `root_fill` / `group_fill`：`a:blipFill` 形状（根级 / 组合内）。

### 4.2 不改清单

只替换图片关系指向的 blob（重写 `r:embed`）；以下内容一律不动：

- `xfrm` 几何与 `a:srcRect` 裁剪；
- 动画节点、切换节点；
- 原对象 `name`（Morph 靠 name 配对，改名会毁掉转场）。

图片不足时按顺序循环取用；报告写明循环轮次与重复槽位。

### 4.3 报告字段

逐槽位 `{page, slot, shape_name, kind, from_image, cycled, cycle}`：

| 字段 | 含义 |
|---|---|
| `page` / `slot` | 页码（1 起）/ 页内槽位序号（文档顺序） |
| `shape_name` | 原对象 name，保持不变 |
| `kind` | `bg` / `root_pic` / `group_pic` / `root_fill` / `group_fill` |
| `from_image` | 本次取用的图片文件名 |
| `cycled` / `cycle` | 是否循环取用 / 第几轮（1 起） |

### 4.4 CLI

```bash
python3 scripts/refill_images.py <pptx> --images a.jpg,b.jpg --out out.pptx
        [--page-range 5-14] [--report-json r.json]
python3 scripts/refill_images.py --selftest
```

### 4.5 自检覆盖

程序化构造含组合内 pic、组合内 blipFill 形状、根级 pic、背景填充的最小 pptx，
替换后断言：① blob 已变；② `xfrm` 逐字节未变；③ shape name 未变；
④ timing / transition 节点逐字节未变；⑤ 4 槽 1 图循环计数正确（后 3 槽标记循环）。

## 5. audit_all.py 三条后处理检查

三条检查同时进入 `--fast` 与全量两条路径；`--json` 新增
`checks.morph_missing` / `checks.image_placeholder` / `checks.letters_count`
与 `summary.extra_fail_pages`。

### 5.1 morph-missing

- 判定：manifest 页字段 `"transition": "morph"` 声明的页，或
  `--require-morph-pages 3,5,7` 显式指定的页，必须挂
  `http://schemas.microsoft.com/office/powerpoint/2015/09/main` 命名空间下的
  `p159:morph`；缺失即 FAIL。
- 以下均算缺失：只挂 `p:fade`、2010 / 2015/12 等其他命名空间的 morph、
  无 morph 节点。
- 未声明且未显式指定的页不参与本检查。

### 5.2 image-placeholder

- 判定：文本框（含组合内）文本命中「添加图片 / 插入图片 / 图片占位 /
  点击添加图片」，且该形状（或其所在组合）内没有 `a:blip`，即 FAIL。
- 仅名称命中、文本为空的图片槽位不在本检查口径内。

### 5.3 letters-count

- 判定：单字母文本对象（`text.strip()` 长度为 1 的 ASCII 字母）数量，与页声明
  manifest 字段 `letters_count` 不符即 FAIL。
- 跳过口径：无 `letters_count` 声明记 `skipped`，不算 FAIL；该页所有单字母对象
  都已移出画布（绝对 `top < 0`，组合内按 chOff/chExt 换算绝对坐标）视为布局本身，
  跳过不判 FAIL。

### 5.4 反例与自检

`--selftest` 在既有两条（达标样例退出 0；只挂 Morph、无页内动效反例退出 1
命中第 1 页）之外，新增三条反例，均断言退出码 1 且命中页号正确：

- 声明 `transition=morph` 但只挂 fade → `morph_missing=[1]`；
- 「点击添加图片」文本且无 blip → `image_placeholder=[1]`；
- 声明 `letters_count=3` 实际 2 个单字母 → `letters_count=[{page:1,
  expected:3, actual:2}]`；
- 另含两条跳过反证：无声明不判不符；字母全部出画不判不符。

## 6. Morph 结构对齐核对

对照上游 `ppt_morph.py` 与技能现有实现 `dynamic_deck.transition`：

| 核对项 | 上游 ppt_morph.py | 技能现有实现 | 结论 |
|---|---|---|---|
| 外层包裹 | `mc:AlternateContent` | `mc:AlternateContent` | 一致 |
| Choice 条件 | `Requires="p159"` | `Requires="p159"` | 一致 |
| morph 命名空间 | 2015/09/main | 2015/09/main（`P159`） | 一致 |
| morph 选项 | `option="byObject"` | `option="byObject"` | 一致 |
| Fallback | `p:transition` 内挂 `p:fade` | `mc:Fallback` 内 `p:transition` 挂 `p:fade` | 一致 |
| 速度属性 | `spd="slow"` | `spd="med"` + `p14:dur="1500"` | 有出入，以技能现有实现为准 |
| 挂载位置 | `cSld` 后插入 | `clrMapOvr` 后插入 | 有出入，以技能现有实现为准（结构体检校验子元素顺序兼容） |

核对结论：规范骨架（AlternateContent / Choice Requires / byObject /
Fallback fade / 2015-09 命名空间）两边一致；速度与挂载位置两处差异保留技能
现有写法，不改代码。

## 7. stage_offscreen 不采纳原因

上游 `stage_offscreen.py` 处理 SVG 画布：把整块停在 viewBox 外的根级 `<g>`
包成 morph staging 组，换取 SVG 质量门的出画豁免。本路线不产 SVG，出画对象在
pptx 中由绝对坐标与页内动画承担，不存在 viewBox 越界判定与对应豁免场景，
该脚本不移植。

## 8. 回归证据索引

- textfit 对 v2 旧稿的空框回归：`docs/textfit-regression.md`；
  逐框原始结果 `docs/textfit-v2-report.json`。
- 旧稿基线：v4 成品页版在 `--fast` 下通过；旧自绘稿因页内动效配额不足仍 FAIL。
