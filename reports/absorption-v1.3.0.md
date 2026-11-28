# 上游 v1.3.0 吸收说明

> 上游版本：tag `v1.3.0`，commit `91746f6`。本轮只吸收与「python-pptx + EMU」路线兼容
> 的部分；上游的 SVG→PPTX 创作管线整体不引入。

## 一、新增效果（5 包，canonical 72 页 / physical 85 页）

| 包 | 页数 | 角色 | 材料门槛 | 硬约束 |
|---|---|---|---|---|
| 018 摄像机转场 | 7 | `cover` / `section_transition` | 1 张 16:9 横图（宽高比容差 0.08），副标题 ≤20 字 | 三层结构（全屏底图 / 中部横幅 / 底部装饰），中部横幅可缺省 |
| 019 情景演绎双版本 | 13×2 | 叙事演绎（开场帧 `cover`、对话帧 `quote`、招式卡 `case_card`） | `cutout`：透明 PNG 主体 5–6 个（≥1000px 宽）+ 背景 2–4 张；`plain`：普通照片 5–6 张（≥1600px） | 两版素材不可互换；抠图版每页多 1–2 张图；选页按用户素材定版本，不得默认其一 |
| 020 照片接力转场 | 18 | **仅** `cover` / `section_transition` | 16:9 照片；同一张照片两种承载（`p:pic` + 幻灯片背景 `blipFill`） | Morph 靠对象名 `图片 1/2/3` 与 `!!1/!!2/!!3` 配对；照片不足会重复 |
| 021 词云开场 | 1 | `cover` | 无媒体；标题 2–8 字、关键词 2–3 字 | 白字黑底，**强调色不可改**（改色像素差 0.00）；部分槽位初始 x 为负，属出画入场，静态稿不得判越界 |
| 022 园林摄影集 | 33 | **仅** `cover` / `section_transition`（p20 目录可作 `agenda`，p22–p32 正文可作 `card_grid`） | 16:9 照片流，最多 17 张 | 照片不足会循环重复，须先确认是否接受；正文槽 18pt 小字，长文案不适用 |

索引总量：physical **491 页**（每个 pptx 文件都计）/ canonical **478 页**（同一效果只计一版，
019 的 `plain` 版为重复 variant）；可换文案 386 → **442 页**。

## 二、移植的四条（按上游工具逐条对应）

| 上游 | 本仓库落点 | 做法 |
|---|---|---|
| `typeset.py` | `scripts/textfit.py`（496 行） | 重写为 python-pptx 口径：CJK 按字、Latin 按词、数字与单位不拆、行首行尾禁则；三段收敛「缩字号至下限 → 加高容器 → 报需减字」；宽度不可改；空框 / 欠满框判 `empty_frame` 并给字数区间 |
| `layout_flow.py` | `scripts/grid_flow.py`（193 行） | 列数先按上限、再按行数回推均衡（6→3×2、7→4+3、4→4×1，5 项由 4+1 纠正为 3+2）；同行等高；余量按行轮补不拉伸单行；输出 EMU |
| `flex_slots.py` | 并入索引与流水线文档（不单独建脚本） | 七条槽位边界写进 `layout-assets.md` 末尾固定段（扫描器常量，重扫不丢）与 `template-pipeline.md` 第 4.1 节；1 px = 9525 EMU |
| `postfix_deck.py` + `refill_images.py` | `scripts/audit_all.py` 三条检查 + `scripts/refill_images.py`（357 行） | Morph 命名空间缺失、占位文案无图、单字母数与声明不符三条判 FAIL；换图覆盖背景 / 根级 pic / 组合内 pic / blipFill 四类承载，只换 blob，不动 xfrm、动画、切换与 shape name |

## 三、只核对不采纳

- `ppt_morph.py`：与技能现有实现五项一致（`mc:AlternateContent`、`Choice Requires="p159"`、
  2015/09 命名空间、`option="byObject"`、Fallback 挂 fade）；两处出入以现有实现为准——速度
  属性（上游 `spd=slow`，技能 `spd=med` + `p14:dur=1500`）与挂载位置（上游 cSld 后，技能
  clrMapOvr 后）。不改代码。
- `stage_offscreen.py`：为 SVG viewBox 出画豁免而写，本路线不产 SVG，没有对应场景；引入
  只会多一条无效分支。
- 上游 `build_effect*.py` / `patch_*.py` / `prefix_texts9.py` / `remake_uni.py` 等：均为
  一次性造件脚本，不是可复用工具，不入库。

## 四、覆盖率变化

- 包：29 → **35**（master 28 → 34，filled_sample 1）。
- 页：406 → **491**（canonical 478）。
- 可用角色：7 → **8**（新增 `agenda` 1 页，来自 022 p20）。
- 增长最多：`section_transition` 4 → 46（可用 4 → 40，018/020/022 贡献），`cover`
  121 → 135（可用 115 → 126）。
- 新出成品页但不可换文案：`case_card` 10 页、`quote` 10 页（019，被判烘焙文字阻断）。
- 仍为 0 的 12 类：`table_list`、`timeline`、`two_column_compare`、`metric_band`、
  `trend_chart`、`compare_table`、`org_chart`、`risk_alert`、`case_card`、
  `conclusion_list`、`quote`、`closing`。

## 五、验证证据

| 项 | 结果 |
|---|---|
| `textfit.py --selftest` | 退出 0：窄槽缩字号、数字单位不拆、空框给区间、行首禁则生效 |
| `grid_flow.py --selftest` | 退出 0：6→3×2、7→4+3、4→4×1、5 项纠正、余量轮补 |
| `refill_images.py --selftest` | 退出 0：四类承载替换后 blob 变、xfrm/name/timing/transition 未变、循环计数正确 |
| `audit_all.py --selftest` | 退出 0：达标 0；配额反例 1；morph/占位/字母数三条反例各退出 1 且命中第 1 页 |
| textfit 回归（v2 稿已知空框） | 六框全部 `empty_frame`（P2 圆角 11/29、P9 圆角 10/13/40、P11 圆角 28），按建议字数下限填充后 `ok`；同页有字框返回 `ok`，无假阳性 |
| v4 稿 `--fast` | 退出 0，三条新检查为空 |
| 旧自绘稿 `--fast` | 退出 1，配额不达标页与改动前一致 |

## 六、已知问题（如实登记）

1. **v4 稿全量模式退出 1**：`presentation.xml.rels` 有一条未被 `sldIdLst` 引用的死关系
   （`rId7 → slides/slide1.xml`，部件不存在）。用改动前的审计脚本复跑同样报出，属既有现象，
   非本轮引入；本轮未放宽 schema 判定。`--fast` 通过。
2. **项目内审计副本口径落后**：`deck-riemann-12p-tpl/build/audit_all.py` 是早前拷到项目里的
   副本，对同一份 v4 稿返回 PASS，掩盖了上述 schema 问题。交付判定必须以技能内
   `scripts/audit_all.py` 为准；项目副本仅作图省事的中途回环。
