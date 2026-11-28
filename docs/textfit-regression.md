# textfit 空框回归（v2 稿）

## 1. 口径

- 输入稿：`deck-riemann-12p-tpl/out/一纸宣告与两千页重构-12页-成品页版-v2.pptx`（12 页）。
- 工具：`scripts/textfit.py`，PIL + 真实 CJK 字体量测；pt→px 按 96/72，1 px = 9525 EMU。
- 判定状态：`ok` / `shrunk` / `grown` / `reduce_chars` / `empty_frame`。
- 不计入文本容量闸的槽位：箭头（`rightArrow` 等）、自定义自由多边形、`图片占位` 图片槽位、含 `a:blip` 的形状。
- 逐框原始结果：`docs/textfit-v2-report.json`。

复现命令：

```bash
python3 scripts/textfit.py <v2.pptx> --json docs/textfit-v2-report.json
```

## 2. 已知空框判定结果

回归目标：P2 右两卡、P9 中栏三框、P11 中栏大框。六框全部判 `empty_frame`，无一条误判为 `ok`。

| 页 | 框名 | 判定 | 建议字数 min–max | 容量行数 | 建议字号 | 按 min 填充 | 按 max 填充 |
|---|---|---|---|---|---|---|---|
| P2 | 矩形: 圆角 11 | empty_frame | 65–130 字 | 5 行 | 18.0 pt | ok，3 行，18.0 pt | shrunk，6 行，17.0 pt |
| P2 | 矩形: 圆角 29 | empty_frame | 117–234 字 | 9 行 | 18.0 pt | ok，6 行，18.0 pt | shrunk，10 行，17.0 pt |
| P9 | 矩形: 圆角 10 | empty_frame | 98–196 字 | 14 行 | 18.0 pt | ok，9 行，18.0 pt | shrunk，16 行，16.5 pt |
| P9 | 矩形: 圆角 13 | empty_frame | 161–322 字 | 14 行 | 18.0 pt | ok，9 行，18.0 pt | shrunk，15 行，16.5 pt |
| P9 | 矩形: 圆角 40 | empty_frame | 75–150 字 | 15 行 | 18.0 pt | ok，9 行，18.0 pt | shrunk，17 行，16.5 pt |
| P11 | 矩形: 圆角 28 | empty_frame | 133–266 字 | 14 行 | 18.0 pt | ok，8 行，18.0 pt | shrunk，15 行，16.5 pt |

容量结论：

- 按建议字数下限填充：六框全部 `ok`，保持 18.0 pt 不缩字号。
- 按建议字数上限填充：六框均进入 `shrunk`（字号 16.5–17.0 pt），仍在容器内容纳，不触发 `grown` 或 `reduce_chars`。
- 即建议区间偏保守：下限是「18 pt 铺满」的容量，上限是「缩半档字号仍装得下」的容量。

## 3. 假阳性反证（有足够文字不得判空）

同一稿内已有文字的框返回 `ok`，不判 `empty_frame`：

| 页 | 框名 | 文字 | 字数 | 字号 | 判定 |
|---|---|---|---|---|---|
| P11 | 矩形: 圆角 55 | 仅输出标准三公理为通过 | 11 | 16.0 pt | ok |
| P11 | 矩形: 圆角 56 | 出现非常规公理名即不通过 | 12 | 16.0 pt | ok |
| P11 | 文本框 33 | lake exe cache get；lake build | 29 | 18.0 pt | ok |
| P2 | 文本框 25 | 第一部 9 章：4 页宣告说了什么（翻译导读） | 23 | 18.0 pt | ok |

全稿扫描中，判 `empty_frame` 的框字数均为 0；非空文本框无一进入 `empty_frame`（欠满分支未触发假阳性）。

## 4. 其余空框说明

除上述六张卡片外，v2 稿还扫出一批单行标签槽位（页内徽章 / 分组标题位置）文本为空，均判 `empty_frame`，建议字数 2–14 字、1 行；这些是 v2 未回填的标签位，不属于本次回归点名的卡片。箭头、自由多边形、图片占位槽位按口径跳过，不出现在文本容量报告里。
