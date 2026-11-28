# 自检与回归证据

## 1. 四个脚本 selftest 输出与退出码

命令均在技能仓库根目录执行，断言失败以非零退出码退出。

### 1.1 textfit.py

```bash
python3 scripts/textfit.py --selftest
```

输出：

```
selftest OK：①窄槽长句缩字号 status=shrunk pt=11.5；②数字单位整体不拆；③空框 empty_frame 字数区间 69–138、容量 6 行；④行首无禁则标点；充足文字样例 status=ok 非空框。
```

退出码：`0`。

覆盖用例：① CJK 长句窄槽缩字号；② `239 MB` / `1.4 GB` / `17,374,215` /
`24.4%` / `0.9%` 整体不拆；③ 空框 `empty_frame` + 字数区间；④ 行首禁则不顶行首；
反证：文字充足样例 `ok`。

### 1.2 grid_flow.py

```bash
python3 scripts/grid_flow.py --selftest
```

输出：

```
selftest OK：6 项 3×2、7 项 4+3、4 项 4×1；5 项由朴素 4+1 纠正为 3+2；同行等高，余量轮补不拉伸单行。
```

退出码：`0`。

### 1.3 refill_images.py

```bash
python3 scripts/refill_images.py --selftest
```

输出：

```
selftest OK：4 类承载全部替换；blob 已变；xfrm / name / timing / transition 逐字节未变；4 槽 1 图循环 3 次，重复槽位 [2,3,4]。
```

退出码：`0`。

断言：① blob 变了；② `xfrm` 未变；③ shape name 未变；④ 动画 / 切换 XML
逐字节未变；⑤ 4 槽 1 图循环计数正确（槽 2/3/4 标记循环）。

### 1.4 audit_all.py

```bash
python3 scripts/audit_all.py --selftest
```

输出：

```
selftest OK：达标样例退出 0；配额反例退出 1 命中 [1]；morph 缺失 / 图片占位 / 字母数不符三条反例均退出 1 且命中第 1 页；无声明与全部出画字母记 skipped 不判 FAIL，部分出画仍核对。
```

退出码：`0`。

## 2. audit_all 三条新检查反例命中

| 检查 | 反例构造 | 命中页 | 退出码 |
|---|---|---|---|
| morph-missing | manifest 声明 `transition=morph`，页内只挂 `p:fade` | `morph_missing=[1]` | 1 |
| image-placeholder | 文本「点击添加图片」，形状 / 组合内无 `a:blip` | `image_placeholder=[1]` | 1 |
| letters-count | 声明 `letters_count=3`，实际 2 个单字母 | `letters_count=[{page:1, expected:3, actual:2}]` | 1 |

跳过反证：无 `letters_count` 声明 → 不判不符；单字母全部出画（`top<0`）→
不判不符；部分出画（一个出画、一个在画布内）不豁免，仍按实际数量核对。既有两条自检（达标样例退出 0；只挂 Morph、无页内动效反例退出 1
命中第 1 页）继续通过。

显式指定路径冒烟：v4 稿 `--require-morph-pages 1,3,5` →
`morph_missing=[1,3,5]`，`summary.extra_fail_pages=[1,3,5]`，退出码 1。

## 3. 旧稿基线

证据 JSON：`docs/audit-v4-fast.json`、`docs/audit-v2-fast.json`、
`docs/audit-old-fast.json`。

| 稿件 | 模式 | 退出码 | 结论 |
|---|---|---|---|
| v4 成品页版 | `--fast` | 0 | PASS；三条后处理检查均为空，`extra_fail_pages=[]` |
| v2 成品页版 | `--fast` | 0 | PASS；三条后处理检查均为空 |
| 旧自绘稿（deck-riemann-12p） | `--fast` | 1 | FAIL；动效配额不达标页
`[1,2,3,4,5,6,7,9,10,11,12]`，与改动前一致 |

全量模式说明：v4 全量退出码 1，唯一问题是 `presentation.xml.rels` 中一条
未被 `sldIdLst` 引用的死关系 `rId7 → slides/slide1.xml`（该部件不存在，
ContentTypes 也无对应 Override）；该问题在改动前的审计版本上同样存在，
与本次新增三条检查无关。三条后处理检查在 v4 全量报告中均为空。旧自绘稿
全量仍 FAIL（页内动效配额不足）。

## 4. textfit 空框回归

见 `docs/textfit-regression.md`；逐框原始结果 `docs/textfit-v2-report.json`。

- 点名六框（P2 右两卡、P9 中栏三框、P11 中栏大框）全部判 `empty_frame`，
  给出字数区间与建议字号；按字数下限填充六框均 `ok`，按上限填充进入
  `shrunk` 但仍装得下。
- 同稿已有文字的框返回 `ok`，无非空框误判 `empty_frame`。
