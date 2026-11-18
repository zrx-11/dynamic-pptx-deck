---
name: dynamic-pptx-deck
description: "把任意已有内容（文档、报告、笔记、大纲、旧 pptx 原稿）做成带动画与转场的原生可编辑 .pptx。用于在 PowerPoint 2019/365 与 WPS 演示中编排入场/强调/退出/路径动画、页面转场与商务版式；覆盖五套主基线 + 11 套学术配色、通用页型骨架与母题、Morph 配对帧配方、三层渐进增强适配（L0 安全层 / L1 高版本增强 / L2 排除层）、timing 树结构排错与效果弃用清单。"
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
   D 紫夜 / E 深绿，或 11 套学术配色），按内容选版式骨架（封面、目录、章节过渡、卡片、双栏、
   时间线、表格、Q&A、数字带、轮播、模糊大字页、结尾）。
3. 生成脚本 `build.py` 接入 `scripts/dynamic_deck.py`：用封装函数画形状与文本，
   用 `animate` 编排动画批次，用 `transition` 写转场。
4. 动画纪律、方向位标记、轴心锁定、节奏规则见 `references/effects-catalog.md`；
   新页面动效先查效果目录，白名单外的预设不要加。
5. 生成后跑 `dynamic_deck.py check` 做结构体检；用 LibreOffice 导出 PDF/截图核对
   静态版式。LO 不解释动画，移植类效果必须在目标机 F5 复核，结果记入验证记录。

## 脚本接口（scripts/dynamic_deck.py）

形状与文本：`text` / `rect` / `oval` / `donut` / `set_name`；
自转组合体：`spin_group` / `spin_arc` / `spin_pie`（轴心锁定，避免短弧绕自身中心偏离圆心）。

动画：`animate(slide, groups, auto=0)`
- `groups` 是点击批次列表；每批是元素元组列表。同组同时开始，先后靠各自 delay 错开。
- `auto=N`：前 N 批进页即自动播（delay=0），其余等点击。
- 元素元组：
  - 入场/退出：`(name, kind, delay_ms, dur_ms)`；退出含 `disappear`、`zoom_out`、`wipe_out_l/r/u/d`
  - 强调自转：`(name, "spin", delay, dur[, by])`
  - 放大回弹：`(name, "grow", delay, dur)`；单次放大：`(name, "grow_once", delay, dur)`
  - 闪烁：`(name, "blink", delay, dur)`；路径漂移：`(name, "path", delay, dur, dx, dy)`

转场：`transition(slide, kind)`；可用清单见 `references/effects-catalog.md` 第 5 节。
体检：`validate(path)`；CLI：`python3 dynamic_deck.py check <deck.pptx>` / `demo <out.pptx>`。

## 效果分级原则

- 实测可用（安全打底）：`fade`、`fly_up`、`wipe_l`、`spin`。
- 结构移植（模板/原生片反查，带真实行为节点，未复测）：其余 fly/wipe 方向、`zoom`、
  `zoom_rot`、`zoom_out`、`bounce_up`、`drop_in`、`appear`、`grow_once`、`blink`、`path`。
- 双端（PPT + WPS）适配分级见 effects-catalog 第 10 节：不确定目标软件时转场只用 ECMA
  基础组，主入场用安全层，p14 转场与高风险预设只作点缀或替换。
- 弃用清单（已确认不生效，不要加回）：`box`、旧版 zoom、stretch、空壳预设、Morph
  （仅 L1 配对成立时使用）、翻书页/页面卷曲、内嵌视频；详见 effects-catalog 第 4 节。

## 参考文档

- `references/effects-catalog.md`：效果目录、行为节点、转场清单、timing 树结构与排错顺序、方向位标记、轴心锁定、节奏规则。
- `references/layout-recipes.md`：配色字体基线（含 11 套学术配色）、通用页型骨架与母题、配方级版式（轮播/扇形/六边形/模糊大字等）、终端风装饰手法、环境动效母题、跨页流畅感编排、工程坑位与许可边界。
- `references/source-inventory.md`：开源项目台账、历史样稿拆解统计、路线选型与素材许可纪律。
