# v1.3.0 重扫验收记录（过程产物）

> 本文件为重扫过程的验证留痕，非交付索引本体；索引本体见
> `references/layout-assets.json` / `references/layout-assets.md`，覆盖率结论见
> `reports/layout-asset-coverage.md`，新包清单见 `reports/new-packs-v1.3.0.md`。

## 1. 扫描命令

```bash
ST=<skill-root>/dynamic-pptx-deck
ROOT=<repo>/ext-research/ppt-effects/effects
python3 "$ST/scripts/scan_layout_assets.py" --root "$ROOT" --skill-root "$ST"
```

- 扫描根：`ext-research/ppt-effects/effects`（含 `academic/` 子目录；`preview/` 图不作数据源）。
- 只改了 `scripts/scan_layout_assets.py`；`audit_all.py`、`slide_importer.py`、
  `dynamic_deck.py`、`textfit.py`、`grid_flow.py`、`SKILL.md`、`README.md` 未改。

## 2. 页数口径

| 口径 | 页数 | 说明 |
|---|---|---|
| physical | **491** | 每个 pptx 文件都计页；406（旧）+ 7 + 13×2 + 18 + 1 + 33 |
| canonical | **478** | 同一效果只计一版；019 的 plain 免抠图版为重复 variant，不计入 |
| variant 差值 | **13** | 即 019 plain 版 13 页重复 |

## 3. 两次重扫逐字节一致证据

连续两次执行同一命令后 `md5sum -c` 校验，四份产物全部 OK：

| 产物 | md5（两次一致） |
|---|---|
| `references/layout-assets.json` | `cfa8939acf1f45a11d18ae9b3d87dccb` |
| `references/layout-assets.md` | `80fc909a40bacfd06ba1becbadd964ac` |
| `reports/layout-asset-coverage.md` | `a12eca038e4de242c60058512b43770f` |
| `reports/new-packs-v1.3.0.md` | `63bcf0e4b3b741026f8d58459d0b8bbe` |

校验方式：首轮落盘后 `md5sum` 存 `/tmp/A.md5`，次轮重扫后 `md5sum -c /tmp/A.md5`，
四项均 `OK`；与上一轮基线 md5 也一致，证明重扫可复现、无随机序 / 时间戳漂移。

## 4. 验收清单逐条核对

| # | 验收项 | 结果 |
|---|---|---|
| 1 | JSON 记录数 = 491（physical），报告 canonical = 478，两数同列并解释差异 | 通过：491 / 478，差值 13 来自 019 plain 重复 variant |
| 2 | 019 共 26 条记录带 variant，cutout 13 / plain 13，无 null | 通过：`Counter({'plain': 13, 'cutout': 13})` |
| 3 | 021 记录 `lock_palette=true`，notes 含「强调色不可改」「初始出画槽位」 | 通过 |
| 4 | 020 / 022 不出现 card_grid / timeline 等内容页角色（022 的 agenda / 正文三栏除外并带边界说明） | 通过：020 = {cover, section_transition}；022 = {cover, section_transition, agenda, card_grid} |
| 5 | 槽位规则段出现在 `layout-assets.md` 末尾，重扫后仍在 | 通过：附录段含 7 条规则与 EMU 折算值 `3,619,500–11,047,800` |
| 6 | 两次重扫 md/json 逐字节一致 | 通过（见第 3 节） |
| 7 | 旧 406 页记录结构向后兼容，旧字段未删改 | 通过：旧包角色计数与上一轮基线完全一致（card_grid 170 / cover 121 / gallery 83 / process 20 / kpi 7 / section_transition 4 / thanks 1） |

## 5. 角色前后对照（成品页 / 可换文案）

| 角色 | 上一轮成品 | 上一轮可用 | 本轮成品 | 本轮可用 | 结论 |
|---|---|---|---|---|---|
| `cover` | 121 | 115 | 135 | 126 | 增加 |
| `section_transition` | 4 | 4 | 46 | 40 | 增加 |
| `kpi` | 7 | 7 | 7 | 7 | 不变 |
| `process` | 20 | 20 | 20 | 20 | 不变 |
| `card_grid` | 170 | 157 | 178 | 165 | 增加（022 正文三栏 8 页） |
| `agenda` | 0 | 0 | 1 | 1 | **本轮新覆盖**（022 p20 目录页） |
| `thanks` | 1 | 1 | 1 | 1 | 不变 |
| `gallery` | 83 | 82 | 83 | 82 | 不变 |
| `case_card` | 0 | 0 | 10 | 0 | 新出成品页（019 招式卡 ×2 版），全部被烘焙判定阻断，暂不可换文案 |
| `quote` | 0 | 0 | 10 | 0 | 新出成品页（019 对话帧 ×2 版），全部被烘焙判定阻断，暂不可换文案 |

### 仍为 0 的角色

- **完全无成品页（10）**：`table_list`、`timeline`、`two_column_compare`、
  `metric_band`、`trend_chart`、`compare_table`、`org_chart`、`risk_alert`、
  `conclusion_list`、`closing`。
- **有成品页但按可换文案口径仍为 0（2）**：`case_card`、`quote`（来源 019，
  physical 10 / canonical 5，阻断原因：图片覆盖率高触发烘焙文字判定）。

## 6. 三个硬约束页索引记录片段

### 6.1 021 词云开场（ref-026-p1）

- `role`: `cover`；`role_source`: `人工-覆盖`；`variant`: `null`
- `constraint`: `strong`；`lock_palette`: `true`
- `constraint_notes`：
  - 白字黑底，强调色不可改（spec 实测改色像素差 0.00）
  - 槽位：标题 115pt / 短句引子 80pt / 三句 72pt / 关键词 44pt×7
  - 部分槽位初始 x 为负（画布外左侧，靠动画入场）：初始出画槽位，禁止在静态稿里判为越界
  - 全部 12 个文本框都要换，漏一个会留下占位；文本框单行，文字变长加长框，不缩字号、不换行
- `role_boundary`：可承担单页词云开场；不能承担多段正文与数据展示；配色与出画位置锁定

### 6.2 019 情景演绎（ref-024-p1 cutout / ref-023-p1 plain）

- `constraint`: `strong`；`variant`: `cutout` / `plain`
- 选页口径：有透明 PNG 主体 5–6 个 + 背景 2–4 张走 cutout，只有普通照片 5–6 张走 plain；
  两版页数相同，抠图版每页比免抠版多 1–2 张图，选页时按素材情况判成本。
- 材料门槛：cutout 7–10（透明主体 5–6 + 背景 2–4）；plain 5–6（普通照片，建议 ≥1600px 宽）。
- `role_boundary`（按页型）：
  - 开场帧 cover：实为情景演绎页，可承担对话开场纯文字 + 背景帧，不能承担标准封面长文案排版
  - 对话帧 quote：实为情景演绎页，可承担中英两行台词叙事，不能承担标准引语页或流程步骤页
  - 招式卡 case_card：实为情景演绎页，可承担单主体招式 / 案例式卡片，最接近既有 case_card 键，不能承担标准并列卡片网格

### 6.3 022 园林摄影集

- p20 `agenda`：目录页（01–04 四章）；随摄影集风格，文字槽为 18pt 小字，长文案不适用。
- p22 `card_grid`：正文三栏页；随摄影集风格，文字槽为 18pt 小字，每段 ≤14 字，长文案不适用。
- p33 `section_transition`：结尾待续帧（未完待续 / TO BE CONTINUED）；不能承担标准致谢页，也不承担正文长文案。
- 包级硬约束（已写入照片流 / 章节帧 boundary）：需要 16:9 照片流，最多 17 张；不足会循环重复，须先确认用户是否接受重复。
