---
name: dynamic-pptx-deck
display_name: "高级动态PPT"
description: "把任意已有内容（文档、报告、笔记、大纲、旧 pptx 原稿）做成带动画与转场的原生可编辑 .pptx。用于在 PowerPoint 2019/365 与 WPS 演示中编排入场/强调/退出/路径动画、页面转场与商务版式；覆盖五套主基线 + 11 套学术配色、通用页型骨架与母题、17 套效果配方、Morph 配对帧、三层渐进增强适配（L0 安全层 / L1 高版本增强 / L2 排除层）、timing 树结构排错、容器—文字配对时序核对与效果弃用清单；支持跨包幻灯片整页导入，复用参考件现成效果页的动画与转场，导入后仅做文字替换与页序编排。"
---

# 高级动态PPT

## 适用场景

- 输入是已有素材：Word/Markdown/PDF 文档、汇报笔记、要点大纲、旧版静态 pptx 等；
  交付物必须是原生可编辑 `.pptx`，目标放映环境覆盖 PowerPoint 2019/365 与 WPS 演示；
  不用 HTML/JS/PDF 替代 pptx。
- 需要把静态内容讲成有节奏的演示：对象分批入场、强调重点、路径漂移、页面转场、
  跨页母题连贯；按素材性质选版式骨架与配色基线，不绑定特定题材（汇报、总结、
  宣讲、路演、课程、答辩等通用）。
- 动画不在 python-pptx / PptxGenJS API 内，全部写在每页 `ppt/slides/slideN.xml` 的
  `<p:timing>`；转场写在 `<p:transition>`。节点写错 PowerPoint 会静默丢弃整页动画。

## 工作流

1. **整理内容大纲**：从素材中提取章节、要点、数据与结论；不编造经历、统计数字与
   事实，素材缺失的信息用【】占位；先定页数与每页承载的信息块，再谈视觉。
2. 读 `references/layout-recipes.md`：选定配色基线（A 浅色商务 / B 深色科技 / C 红金 /
   D 紫夜 / E 深绿青碧，或 11 套学术配色），按内容选版式骨架（封面、目录、章节过渡、卡片、双栏、
   时间线、表格、Q&A、数字带、轮播、模糊大字页、结尾）。
3. 路线选型：素材进入目标稿前先过适配检查，按目标基调、配色族、槽位与背景类型判定
   路线（整页导入 / 抽结构原生重建 / 阻断，判据见 `references/import-guide.md` 第 1 节）。
   参考件中已有满足需求的高冲击效果页、几何层级时序从零重建会失真、且冲突有确定性
   消解路径时，优先用 `scripts/slide_importer.py` 整页导入，复用其 timing/transition，
   导入后只做文字替换与页序编排；`useBgFill` 依赖重、烘焙像素主导或形状必须大改的页面
   走 `scripts/dynamic_deck.py` 原生构建。两条路线可同稿混用，适配检查、分工边界与
   阻断口径见 `references/import-guide.md` 第 1、3、9 节。
4. 先按页型排出转场分工表：主轴方向全稿统一，转场类型按页型轮换、相邻页不堆叠同一预设
   （分工表见 `references/layout-recipes.md` 第 7b 节）；再生成脚本 `build.py` 接入
   `scripts/dynamic_deck.py`：用封装函数画形状与文本，用 `animate` 编排动画批次，
   用 `transition` 写转场。
5. 动画纪律、方向位标记、轴心锁定、节奏规则见 `references/effects-catalog.md`；
   新页面动效先查效果目录，白名单外的预设不要加。
5b. 动画编排按三段走：① 常驻环境动效基线（环自转 / 光点往返，indefinite）；
   ② 分批入场（3—6 批，批内 delay 错开，换批用旧退出 + 新入场）；
   ③ 容器—文字配对（同批、文字 fade 晚 150—450ms、退出文字先容器后）。
   配对矩阵、判定条件与反例形态见 `references/animation-orchestration.md`。
6. 生成后跑 `dynamic_deck.py check` 做结构体检，再跑 `pairing_check.py` 核对
   容器—文字配对；用 LibreOffice 导出 PDF/截图核对静态版式。LO 不解释动画，
   移植类与导入页效果必须在目标机 F5 复核，结果记入验证记录。

## 脚本接口（scripts/dynamic_deck.py）

形状与文本：`text` / `rect` / `oval` / `donut` / `set_name`；
自转组合体：`spin_group` / `spin_arc` / `spin_pie`（轴心锁定，避免短弧绕自身中心偏离圆心）。

动画：`animate(slide, groups, auto=0)`
- `groups` 是点击批次列表；每批是元素元组列表。同组同时开始，先后靠各自 delay 错开。
- `auto=N`：前 N 批进页即自动播（delay=0），其余等点击。
- 元素元组：
  - 入场/退出：`(name, kind, delay_ms, dur_ms)`；退出含 `disappear`、`zoom_out`、`wipe_out_l/r/u/d`
  - 强调自转：`(name, "spin", delay, dur[, by])`；小角度钟摆回摆：`(name, "swing", delay, dur[, by])`
  - 放大回弹：`(name, "grow", delay, dur)`；单次放大：`(name, "grow_once", delay, dur)`
  - 闪烁：`(name, "blink", delay, dur)`；路径漂移：直线 `(name, "path", delay, dur, dx, dy)`，
    折线巡游 `(name, "path", delay, dur, 0, 0, "M 0 0 L x y …")`
  - 自动换片：`transition(slide, kind, adv_ms=2000, on_click=False)`

转场：`transition(slide, kind)`；可用清单、观感分组与选型纪律见 `references/effects-catalog.md` 第 5 节，
页型分工表见 `references/layout-recipes.md` 第 7b 节；17 套完整效果配方见 `references/effect-recipes.md`。
体检：`validate(path)`；CLI：`python3 dynamic_deck.py check <deck.pptx>` / `demo <out.pptx>`。

## 配对检查接口（scripts/pairing_check.py）

仅用标准库（zipfile + xml.etree），逐页核对容器—文字配对：同批绑定、延迟差
150—450ms、文字只挂 fade、退出文字先容器后 100—150ms，统计孤儿文字、cTn 预算、
转场与悬空 spTgt。

- CLI：`python3 pairing_check.py <deck.pptx> [--min-gap 150] [--max-gap 450] [--max-ctn 60] [--json OUT] [--md OUT]`；
  `--json` 出机器可读报告，`--md` 出人读表格。
- 退出码：0=全部 PASS；1=存在 FAIL；2=输入错误或包不可读。
- `python3 pairing_check.py --selftest`：内联最小 fixture 断言六类判定，含必然 FAIL
  的反例；口径见 `references/animation-orchestration.md` 第 3、5 节。

## 效果路由接口（scripts/effect_router.py）

仅用标准库（zipfile + xml.etree + json），解析每页形状 / 文本 / 数字 / 层级 / 时间词，
或接收 JSON 页描述；输出 EffectPlan：页型、信息特征、批次拆法、预设族、常驻环境动效、
强调对象、转场建议与路线建议（原生重建 / 整页导入 / 阻断），每个推荐项带命中的规则编号
（basis）；信号重叠或稀疏时给 `confidence: low` 与保守档，不猜。

- CLI：`python3 effect_router.py <deck.pptx|page.json> [--page N] [--json OUT] [--md OUT]`；
  `--json` 出机器可读计划，`--md` 出人读表格。
- 退出码：0=正常出计划；2=输入不可读。
- 批次按页型角色矩阵生成（KPI / 时间线 / 表格 / 卡片 / 双栏 / 引语 / 过渡 / 结尾各有独立
  角色、数量、预设族与延迟区间）；路由为原生重建时，批次结构（角色 + 数量 + 预设族 + 延迟
  区间）直接喂给 build 脚本；整页导入时给出 `check_slide` 参数与背景 / 重着色要点；阻断时
  写明缺什么（烘焙像素声明、字号档、授权素材、基调 recolor 路径等）。规则矩阵见
  `references/effect-routing.md`。命令行默认输出带批次角色 + 预设族摘要。
- 预算收敛：每页估算 `estimated_effects`（≤12）与 `estimated_ctn`（≤60），写入
  `budget.fits`；超预算时按角色给 `split_plan`（每批配额 `per_batch`、批数 `batch_count`），
  卡片容器与文字同批成对，编号圆牌随卡片 zoom 不另占节点；表格按行聚类、表头不进滚动批次。

## 跨包导入接口（scripts/slide_importer.py）

仅用标准库做文件级 zip 手术，不依赖 python-pptx；整页复制参考件的 slide part、rels 与媒体，
导入后只做文字替换与页序编排。完整口径见 `references/import-guide.md`。

- `TargetProfile(tone: str, bg_color: str, palette: dict, fonts: dict | None = None, baked_declarations: list | None = None, strict: bool = False)`：目标对象画像；`tone` 取 dark/light，`palette` 角色为 bg/card/ink/mut/accent1/accent2/accent3/line；`baked_declarations` 条目为 `{"media": partname 或 sha1, "baked_text": str, "action": replace|keep|block}`；`strict` 时未消解警告升级阻断；含 `to_dict` / `from_dict`。
- `SlideImporter(dst_pptx: str)`：目标包构造器。
  - `check_slide(src_pptx, src_index=0, target=None, target_texts=None) -> dict`：对照目标 Profile 出适配报告（槽位、基调配色、媒体背景、烘焙像素四维度），报告先行。
  - `adapt_slide(src_pptx, src_index=0, target=None, text_map=None, recolor_map=None, media_map=None, background=None, strict=None) -> int`：按消解结果适配写入，返回新页索引；target 必传（传 None 抛错），未消解冲突时写入前阻断。
  - `import_slide(src_pptx: str, src_index: int = 0, text_map: dict | None = None, by_name: bool = False) -> int`：导入参考页，返回新页索引；可随导入一并传入文字映射。
  - `replace_text(slide_index: int, mapping: dict, by_name: bool = False) -> None`：替换指定页文本；默认有序整串匹配，`by_name=True` 按形状名替换。
  - `reorder(order: list[int]) -> None`：按索引列表重排页序。
  - `save(path: str | None = None) -> None`：落盘；不传路径回写目标包。
- `inspect(src_pptx: str, src_index: int = 0) -> dict`：读取参考页结构快照（文本槽、媒体、主题色、关系类型，及 `background` / `useBgFill` / `baked_candidates` / `colors`）。
- 模块级工具：`recolor_srgb`（重着色并保留颜色变换）、`luminance`（亮度判定基调）、`hue_family`（色相族归并）、`estimate_slot`（槽位容量与溢出估算）、`center_crop`（居中裁切 srcRect）、`detect_background`（背景类型识别）、`find_bg_fill_shapes`（背景填充形状清单）。
- CLI：`inspect --src X.pptx [--index N]`；`check --src X.pptx [--index N] [--profile profile.json] [--target-texts texts.json]`；`adapt --src X.pptx [--index N] --dst D.pptx --profile profile.json [--out O.pptx] [--map map.json] [--recolor recolor.json] [--media media.json] [--strict]`；`import --src X.pptx --index N --dst D.pptx [--out O.pptx] [--map map.json] [--by-name]`；`replace --dst D.pptx --index N --map map.json [--by-name] [--out O.pptx]`；`reorder --dst D.pptx --order 0,3,1,2 [--out O.pptx]`。

## 效果分级原则

- 实测可用（安全打底）：`fade`、`fly_up`、`wipe_l`、`spin`。
- 结构移植（带真实行为节点，未在目标机复测）：其余 fly/wipe 方向、`zoom`、
  `zoom_rot`、`zoom_out`、`bounce_up`、`drop_in`、`appear`、`grow_once`、`blink`、`path`。
- 双端（PPT + WPS）适配分级见 effects-catalog 第 10 节：不确定目标软件时转场只用 ECMA
  基础组，主入场用安全层，p14 转场与高风险预设只作点缀或替换。
- 弃用清单（已确认不生效，不要加回）：`box`、旧版 zoom、stretch、空壳预设、Morph
  （仅 L1 配对成立时使用）、翻书页/页面卷曲、内嵌视频；详见 effects-catalog 第 4 节。

## 参考文档

- `references/effects-catalog.md`：效果目录、行为节点、转场清单、timing 树结构与排错顺序、方向位标记、轴心锁定、节奏规则、容器—文字配对时序纪律与页面活着基线。
- `references/animation-orchestration.md`：编排三段、容器—文字配对矩阵、判定条件与反例形态、配对检查器口径。
- `references/effect-routing.md`：内容识别与效果路由矩阵——页型命中条件、推荐效果、禁用项与反例形态、基调与规模约束、保守档口径。
- `references/layout-recipes.md`：配色字体基线（含 11 套学术配色）、页型构成与字号阶梯、通用页型骨架与母题、配方级版式（轮播/扇形/六边形/模糊大字等）、深色技术风装饰手法、环境动效母题、跨页流畅感编排、工程注意事项与许可边界。
- `references/source-inventory.md`：素材许可分级、样稿预设统计归档、路线选型与验证边界。
- `references/effect-recipes.md`：17 套效果配方（页数规则、素材门槛、图层结构、移植约束、跨配方通用规则）。
- `references/import-guide.md`：跨包幻灯片导入口径——适配前置总原则与三路线判定（整页导入 / 原生重建 / 阻断）、inspect 结构快照、check_slide 四维度适配检查、adapt_slide 适配动作、导入处理范围、主题色解析与重着色、文字替换两种模式、组装与路线混用边界、验证纪律与接口/CLI 参考。
