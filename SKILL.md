---
name: dynamic-pptx-deck
display_name: "高级动态PPT"
description: "把任意已有内容（文档、报告、笔记、大纲、旧 pptx 原稿）做成带动画与转场的原生可编辑 .pptx。默认路线是检索成品页资产、整页导入后只替换文案/数字/配色/图片并叠加原生动画；自绘版式仅在资产索引无匹配页时兜底。技能价值在原生动画编排、效果配方、跨包整页导入与校验工具链，不承担版式设计本身。覆盖三档交付取向、Morph 配对帧、容器—文字配对时序、页内动效配额与合并校验；结构化源随稿交付。"
---

# 高级动态PPT

## 定位

- **默认路线**：检索成品页资产 → 整页导入 → 只替换文案 / 数字 / 配色 / 图片 → 叠加原生
  动画。成品页的几何、层级与时序由源页锁定，导入后不强行复刻形状顶点。
- **兜底分支**：仅当资产索引无匹配页（缺口角色，或同角色无可换文案页）时，才走原生
  自绘构建。
- **不可替代能力**：原生动画编排、效果配方、跨包整页导入与校验工具链；本技能不做版式
  设计本身，也不产出 HTML / JS / PDF 替代 pptx。
- 交付物始终是可在 PowerPoint 2019 / 365 与 WPS 演示打开的 OOXML 文件；动画写在每页
  `ppt/slides/slideN.xml` 的 `<p:timing>`，转场写在 `<p:transition>`，节点写错会被
  PowerPoint 静默丢弃整页动画。

## 工作流

1. **整理内容大纲**：从素材提取章节、要点、数据与结论；不编造经历、统计数字与事实，
   素材缺失用【】占位；先定页数与每页信息块。动工前声明交付档位：`fidelity_first`
   成品优先（默认）、`balanced` 兼顾、`compat` 兼容优先；档位判定与阻断规则见
   `references/delivery-tiers.md`。
2. **查成品页资产索引**（版式来源先查资产，不先画版式）：读 `references/layout-assets.md`
   与人读分组、`references/layout-assets.json` 机读记录，按页角色选候选页；选定后记录
   `source_id`（索引 `id`）、`deck` 相对路径、页码、基调、槽位容量、烘焙文字与背景填充
   标记。索引 491 页（同一效果只计一版为 478 页），覆盖 8 个可换文案角色，其余为缺口
   角色；选页前先核 `materials` 与 `role_boundary`——`019` 按 `cutout` / `plain` 两版
   标注不可默认其一，`021` 配色与出画槽位锁定不可改，`020` / `022` 仅限封面与章节过渡
   且需足量 16:9 照片。无匹配页时转兜底分支（第 12 步）。
3. **过选页闸**：候选页排除摄影背景、锁定大图、烘焙文字与短槽容量不足项；细则见
   `references/template-pipeline.md` 第 2.1 节。
4. **导入前适配检查**：`inspect` 读结构快照，`check_slide` 对照目标画像出四维度适配
   报告；报告先行，未通过的页面不导入、不适配。口径见 `references/import-guide.md`
   第 2、3 节。
5. **整页导入**：用 `scripts/slide_importer.py` 的 `import`（或带 Profile 的 `adapt`）
   把源页整页复制进目标包，复用其 timing / transition；导入参数与接口见
   `references/import-guide.md` 第 4、11 节。
6. **只替换 `a:t` 与色值**：允许替换文案 / 数字、按重着色映射改角色色、按媒体映射换图
   （整册换图用 `scripts/refill_images.py`，覆盖组合内外四类图片承载，只换图不改几何、
   动画、切换与 shape name）；**禁止重排坐标、禁止改几何层级、禁止重建文本框**（重建会
   丢失动画绑定）。过容量闸：用 `scripts/textfit.py` 逐槽以真实字体实测宽高，按「缩字号
   至下限 → 加高容器 → 报需减字」顺序适配，不扩宽、不越下限；细则见
   `references/layout-tools.md` 第 2 节、`references/import-guide.md` 第 6、7 节与
   `references/template-pipeline.md` 第 4 节。
7. **内容框闸**：逐框判定填满（墨块覆盖率 ≥25%、无空框），未用槽位补文案或删形状；
   `textfit` 返回 `empty_frame` 的框必须按它给的字数区间补足；细则见
   `references/template-pipeline.md` 第 5 节。
8. **挂页内动画补配额**：按下方「页内动效配额」核对并补足；编排三段与容器—文字配对
   矩阵见 `references/animation-orchestration.md`。
9. **成稿闸**：重叠检测（文本框互压 + 文字超出框高）与内容框检测并列，任一命中即不得交付；
   细则见 `references/template-pipeline.md` 第 7 节。
10. **一次合并校验**：跑一次 `scripts/audit_all.py` 出合并报告与退出码；不要分别跑
   `dynamic_deck.py check`、`pairing_check.py`、`editability_check.py`、
   `effect_router.py` 四遍。单独调用这些脚本仅用于排障定位。
11. **一次 LibreOffice 导出**：导出 PDF / 静帧核对静态版式与文本落位；LO 不解释动画。
   只对 audit 报错的页面（配额不达标、可编辑性阻断、Morph 冲突）看图，修完重跑一次
   `audit_all.py` 复核。移植类与高版本效果须在目标放映软件 F5 复核。
12. **兜底分支（仅索引无匹配页时）**：用 `scripts/dynamic_deck.py` 原生构建，版式骨架与
   效果族候选查 `references/asset-index.md` / `asset-index.json`，配色基线与页型骨架见
   `references/layout-recipes.md`；多块并列内容的行列排布用 `scripts/grid_flow.py` 求
   均衡矩形（输出 EMU），不手工试列数。
13. **结构化源交付**：成稿同级落 `deck-manifest.json`，逐页登记 `index`、`page_type`、
    `content`、`assets`、`notes`；facts 清单与素材引用（生成器、参数、sha256）随稿交付；
    页级重生成复核见 `references/regeneration-guide.md`。

## 页内动效配额

配额写死，`audit_all.py` 逐页判定：

- **可感知运动对象 ≥3**：位移 / 缩放 / 旋转 / 擦除 / 路径 / fly / zoom 计入；纯 fade、
  opacity、appear 不计；退出类效果不计入可感知运动对象计数。
- **持续循环 ≥1**：环自转（spin + 无限重复）或光点路径巡游（path + 无限重复）。
- **效果总数 ≤12**；**cTn 节点 ≤60**。
- **Morph 配对对象**（同名同类型跨页出现）不得叠加页内位置类动画（位移 / 路径 / fly /
  wipe），页内只允许 fade / opacity。
- **不合格形态**：只靠 Morph 转场衔接、页内近乎静态（仅纯 fade 入场）的页面判不合格；
  页内动效应挂在非配对对象上（装饰环、光点、连线、数字、标题下划线、卡片内文字等）。

## 许可纪律

- 成品页资产属第三方素材：个人学习与校内汇报场景可用；公开放映或商业交付前，必须在
  交付说明 / README 声明资产来源与许可范围，未获授权则退回自绘分支。
- 资产索引只记录路径与结构特征，**不打包资产本体**（pptx / 静帧 / svg 均留在参考件
  目录）；许可分级与素材使用纪律见 `references/source-inventory.md` 第 1、2 节。

## 按需读取表

| 你要做什么 | 读哪份 reference 的哪一节 |
|---|---|
| 选成品页（检索候选、看容量与可换文案性） | `layout-assets.md`「怎么用这份索引」及角色分组；机读 `layout-assets.json` |
| 排版适配（断行禁则、缩字号 / 加高 / 减字三段、空框判定） | `layout-tools.md` 第 2 节 |
| 多块并列排布（均衡行列、同行等高、余量分配） | `layout-tools.md` 第 3 节 |
| 整册换图（组合内外四类承载、循环取用） | `layout-tools.md` 第 4 节 |
| 整页导入（适配检查、导入参数、文字替换） | `import-guide.md` 第 1、3、7、11 节 |
| 挂动画（编排三段、配对矩阵、反例形态） | `animation-orchestration.md` 第 1、2、3 节 |
| 查页型路由与效果预算 | `effect-routing.md` 第 2 节 |
| 跑校验（CLI、报告字段、配额判定、退出码） | `audit-all.md` |
| 判定交付档位与允许 / 禁止清单 | `delivery-tiers.md` 第 1–3 节 |
| 装饰层预渲染 | `decor-bake.md` 第 7 节 |
| 页级局部重生成 | `regeneration-guide.md` 第 3、6 节 |
| 全稿复用工序总览 | `template-pipeline.md` |
| 素材许可分级与使用纪律 | `source-inventory.md` 第 1、2 节 |
| 兜底自绘的版式骨架 / 效果族候选 | `asset-index.md` / `asset-index.json` |
| 配色基线、页型骨架与字号阶梯 | `layout-recipes.md` |
| 效果目录、转场清单、弃用清单 | `effects-catalog.md` |
| 完整效果配方 | `effect-recipes.md` |
| 可编辑性扣分码与评分表 | `editability-audit.md` |

## 脚本接口摘要

- **原生构建 `scripts/dynamic_deck.py`**：形状与文本 `text` / `rect` / `oval` / `donut` /
  `set_name`；自转组合体 `spin_group` / `spin_arc` / `spin_pie`（轴心锁定）。动画
  `animate(slide, groups, auto=0)`：`groups` 为点击批次列表，同组同时开始、靠 `delay`
  错开；`auto=N` 前 N 批进页自动播。元素元组：入场 / 退出 `(name, kind, delay_ms,
  dur_ms)`；自转 `(name, "spin", delay, dur[, by])`；钟摆 `(name, "swing", ...)`；
  放大回弹 / 单次放大 `grow` / `grow_once`；闪烁 `blink`；路径 `path`（直线 dx,dy 或
  折线串）。转场 `transition(slide, kind, adv_ms=毫秒, on_click=False)`。CLI：
  `python3 scripts/dynamic_deck.py check <deck.pptx>` / `demo <out.pptx>` / `selftest`。
- **跨包导入 `scripts/slide_importer.py`**：`SlideImporter(dst_pptx)` 提供
  `import_slide` / `check_slide` / `adapt_slide` / `replace_text` / `reorder` / `save`；
  模块级 `inspect` 读结构快照。CLI 子命令 `inspect` / `check` / `adapt` / `import` /
  `replace` / `reorder` / `selftest`；完整签名与快照键见 `import-guide.md` 第 11 节。
- **合并校验 `scripts/audit_all.py`**：一次调用合并结构、配对、可编辑性、schema 深扫、
  动效配额与三条后处理检查（`morph-missing` / `image-placeholder` / `letters-count`）；
  CLI `python3 scripts/audit_all.py <deck.pptx> [--manifest P]
  [--tier 档位] [--fast] [--require-morph-pages 3,5,7] [--json OUT] [--md OUT]`，
  `--selftest` 跑内联自检；口径见 `audit-all.md` 与 `layout-tools.md` 第 5 节。
- **排版适配 `scripts/textfit.py`**：以 PIL + 真实字体量测文本框，按「缩字号至下限 →
  加高容器 → 报需减字」收敛，空框 / 欠满框判 `empty_frame` 并给字数区间与建议字号；
  CLI `python3 scripts/textfit.py <deck.pptx> [--page N] [--min-pt 10] [--json OUT]
  [--apply OUT]`，`--selftest` 跑内联自检。
- **网格排布 `scripts/grid_flow.py`**：按 `max_cols` 上限与项数回推均衡行列（6→3×2、
  7→4+3、4→4×1），同行等高、余量轮补，输出 EMU 矩形；CLI
  `python3 scripts/grid_flow.py --n 7 --area x,y,w,h --max-cols 4 --json`，`--selftest`。
- **整册换图 `scripts/refill_images.py`**：覆盖背景 / 根级 `p:pic` / 组合内 `p:pic` /
  `a:blipFill` 四类承载，只重写图片 blob，不动 xfrm、动画、切换与 shape name；图片不足
  循环取用并在报告标注；CLI `python3 scripts/refill_images.py <deck.pptx> --images a.jpg,b.jpg
  --out OUT [--page-range 5-14] [--report-json OUT]`，`--selftest`。
- **配对检查 `scripts/pairing_check.py`**：仅标准库，核对同批绑定、延迟差 150–450ms、
  文字只挂 fade、退出顺序与 Morph 互斥；CLI
  `python3 scripts/pairing_check.py <deck.pptx> [--min-gap 150] [--max-gap 450]
  [--max-ctn 60] [--json OUT] [--md OUT]`，`--selftest` 跑内联自检。
- **效果路由 `scripts/effect_router.py`**：规划工具，识别页型与信息特征，输出批次拆法、
  预设族、强调对象、转场与路线建议；CLI
  `python3 scripts/effect_router.py <deck.pptx|page.json> [--page N] [--tier 档位]
  [--engine effect_first|compat] [--json OUT] [--md OUT]`。
- **索引扫描 `scripts/scan_layout_assets.py`**：一键重扫参考件目录，重建
  `references/layout-assets.json`、`layout-assets.md` 与
  `reports/layout-asset-coverage.md`；CLI `--root / --out / --md / --report /
  --skill-root`。

## 效果分级原则

- 实测可用（标准 DrawingML 安全族，2019 / 365 / WPS 均可解析）：`fade`、`fly_up`、
  `wipe_l`、`spin`。
- 结构移植（带真实行为节点，未在目标机复测）：其余 fly / wipe 方向、`zoom`、`appear`、
  `blink`、`path`。
- 高版本增强（effect_first 引擎族下发，弱引擎降级但不崩）：`zoom_rot`、`zoom_out`、
  `bounce_up`、`drop_in`、`grow_once`、p14 现代转场、Morph 配对补间；compat 引擎族
  不下发或仅作点缀。
- 弃用清单（已确认不生效，不要加回）：`box`、旧版 zoom、stretch、空壳预设、翻书页 /
  页面卷曲、内嵌视频；Morph 仅在配对成立时使用。引擎族口径见 `effects-catalog.md`
  第 10 节。

## 参考文档

- `references/template-pipeline.md`：成品页复用流水线完整工序（选页判据、导入参数、
  容量纪律、配额挂载、校验、导出与许可声明模板）。
- `references/layout-tools.md`：排版适配（断行禁则、三段收敛、空框判定）、均衡网格排布、
  整册换图三类工具的 CLI 与返回结构，外加上游工具的结构对齐核对与不采纳说明。
- `references/audit-all.md`：合并校验入口 CLI、报告字段、配额判定、`--fast` 取舍、
  退出码与既有脚本关系。
- `references/import-guide.md`：跨包整页导入口径——三路线判定、适配检查、处理范围、
  主题色解析、文字替换、组装纪律与接口 / CLI 参考。
- `references/animation-orchestration.md`：编排三段、容器—文字配对矩阵、判定条件与
  反例形态、配对检查器口径。
- `references/effect-routing.md`：20 类页型识别特征、骨架、效果族、禁用项与反例形态。
- `references/delivery-tiers.md`：三档交付取向的场景判定、允许 / 禁止清单、结构化源
  契约与分档验收阻断规则。
- `references/asset-index.md` / `asset-index.json`：20 个页角色的版式骨架与效果族候选
  库（兜底自绘分支选型用）；与成品页资产索引 `layout-assets.md` / `layout-assets.json`
  是两件事。
- `references/layout-assets.md` / `layout-assets.json`：成品页资产检索索引（默认复用
  路线选页用）；覆盖率报告见 `reports/layout-asset-coverage.md`。
- `references/effects-catalog.md`：效果目录、行为节点、转场清单、timing 树排错、方向位
  标记、轴心锁定、节奏规则与分层取舍。
- `references/effect-recipes.md`：18 套完整效果配方（页数规则、素材门槛、图层结构、
  移植约束）。
- `references/layout-recipes.md`：配色字体基线、页型构成与字号阶梯、通用骨架与母题、
  配方级版式与环境动效母题。
- `references/decor-bake.md`：装饰层预渲染配方、可复现参数与素材登记口径。
- `references/editability-audit.md`：可编辑性判定项、扣分码、评分表与修复建议。
- `references/regeneration-guide.md`：结构化源维护与页级局部重生成流程。
- `references/source-inventory.md`：素材许可分级、样稿预设统计归档与路线选型边界。
