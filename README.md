# dynamic-pptx-deck · 高级动态PPT 技能

把任意已有内容（文档、报告、笔记、大纲、旧版静态 pptx）做成**带动画与转场的原生可编辑 `.pptx`**。
这是一个供 AI agent 调用的技能（skill）：它不生成 HTML 幻灯片，也不生成 PDF，交付物始终是
可以在 PowerPoint / WPS 里继续编辑的 OOXML 文件。

## 它解决什么问题

python-pptx、PptxGenJS 这类库都不封装对象动画与页面转场；动画必须写进每页
`ppt/slides/slideN.xml` 的 `<p:timing>`，转场写进 `<p:transition>`。节点写错时 PowerPoint
会静默丢弃整页动画，而且不同版本（2019 / 365 / WPS）对高级预设的支持程度不一样。

这个仓库把经过结构核对与分层验证的结论固定下来：

- 一套形状与文本封装函数，配合动画批次、转场、轴心锁定组合体；
- 一份效果白名单与弃用清单，只写带真实行为节点的预设；
- 一套双档适配模型：效果优先档默认面向 365 / 新版 WPS，兼容优先档仅在明确要求兼容 2019 / 旧 WPS 时启用；排除层不进稿；
- 五套主配色基线 + 11 套学术配色，通用页型骨架与配方级版式（封面、目录、卡片、双栏、时间线、表格、Q&A、数字带、轮播、模糊大字页、结尾）；
- 18 套效果配方：页数规则、素材门槛、图层结构、移植约束与跨配方通用规则；
- 跨包幻灯片整页导入：复用参考件现成效果页的 timing/transition，导入后仅做文字替换与页序编排，主题色解析为显式色值避免漂移。

## 不做什么

- 不交付 HTML/JS/PDF 替代 pptx；
- 不打包内嵌视频、翻书页/页面卷曲、3D、空壳预设与非商业素材；
- 不编造事实、经历与统计数字；缺失信息由调用方用占位符标出；
- 离线环境只能做 XML 结构体检与静态版式核对，不能代替目标机 F5 放映复核。

## 仓库结构

```
dynamic-pptx-deck/
├── SKILL.md                  技能入口，agent 首先读这个
├── README.md                 本文件
├── LICENSE                   MIT
├── .gitignore
├── agents/
│   ├── doubao.yaml           豆包面板显示名与触发配置
│   └── openai.yaml            同构配置
├── references/
│   ├── effects-catalog.md    效果目录、行为节点、转场清单、timing 排错、轴心锁定、节奏规则、双端分级
│   ├── animation-orchestration.md  编排三段、容器—文字配对矩阵、判定条件与反例形态、配对检查器口径
│   ├── effect-routing.md     内容识别与效果路由矩阵：20 类页型识别特征、骨架、效果族、禁用项与反例形态、基调与规模约束
│   ├── asset-index.md         双库索引（人读）：页角色 → 版式骨架 / 效果族候选清单（条件、代价、引擎、禁用场景）
│   ├── asset-index.json       双库索引（机读）：键为页角色，供路由器生成多方案候选
│   ├── layout-recipes.md      配色基线（含 11 套学术配色）、页型构成与字号阶梯、页型骨架与母题、配方级版式、工程注意事项、许可边界
│   ├── source-inventory.md    素材许可分级、样稿预设统计归档、路线选型与验证边界
│   ├── effect-recipes.md      18 套效果配方（页数规则、素材门槛、图层结构、移植约束、跨配方通用规则）
│   └── import-guide.md        跨包整页导入口径：选型、前置检查、处理范围、主题色解析、文字替换、验证清单
└── scripts/
    ├── dynamic_deck.py       形状/文本/组合体封装、animate、transition、validate、CLI check/demo
    ├── pairing_check.py     容器—文字配对时序核对（零标准库依赖）：同批、延迟差、fade-only、退出顺序、孤儿计数、目标与预算
    ├── effect_router.py      内容识别与效果路由（零标准库依赖）：页型/特征/批次/预设/路线计划，逐项带规则编号
    └── slide_importer.py     跨包导入：文件级 zip 手术、媒体哈希去重、主题色解析、文字替换、页序重排
```

## 安装

仓库本身就是 skill 目录，clone 到 agent 的用户级技能目录即可：

```bash
git clone <仓库地址> \
  ~/.doubao/agent_mode/workspace/.user_skills/dynamic-pptx-deck
```

目录名要和 `SKILL.md` 里的 `name: dynamic-pptx-deck` 一致。其他 agent（Claude Code、
Codex 等）放进它自己的 skills 路径，结构相同。

## 依赖

- Python 3.9+
- `python-pptx`、`lxml`（脚本运行必需）
- 静态预览（可选）：LibreOffice headless、`pdftoppm`；核对版式用，不参与生成

```bash
pip install python-pptx lxml
```

## 最小使用流程

1. 整理输入内容的章节、要点与结论；不补充素材里没有的事实。
2. 读 `references/layout-recipes.md` 选配色基线与页型骨架。
3. 写一个 `build.py`，`sys.path` 接入 `scripts/dynamic_deck.py`，用封装函数画形状、
   用 `animate` 编排批次、用 `transition` 写转场，保存为 pptx。
4. 结构体检：

```bash
python3 scripts/dynamic_deck.py check <deck.pptx>
```

5. 用 LibreOffice 导出 PDF/截图核对静态版式；移植类与高版本效果在目标软件 F5 复核。

自测样例：

```bash
python3 scripts/dynamic_deck.py demo /tmp/demo.pptx
python3 scripts/dynamic_deck.py check /tmp/demo.pptx
```

## 脚本接口摘要

形状与文本：`text` / `rect` / `oval` / `donut` / `set_name`；
自转组合体：`spin_group` / `spin_arc` / `spin_pie`（轴心锁定，避免短弧绕自身中心偏离圆心）。

动画：`animate(slide, groups, auto=0)`，`groups` 是点击批次列表，同组元素同时开始，
先后靠 `delay` 错开；`auto=N` 前 N 批进页即自动播。
体检 CLI `check <deck.pptx>` 退出码 0=全过、1=有问题、2=包不可读；`selftest` 跑内联构造器自检。

- 入场/退出：`(name, kind, delay_ms, dur_ms)`；退出含 `disappear`、`zoom_out`、`wipe_out_l/r/u/d`、`fly_out_l/r/u/d`
- 自转：`(name, "spin", delay, dur[, by])`
- 放大回弹 / 单次放大：`(name, "grow" | "grow_once", delay, dur)`
- 闪烁：`(name, "blink", delay, dur)`
- 路径漂移：直线 `(name, "path", delay, dur, dx, dy)`；折线巡游 `(name, "path", delay, dur, 0, 0, "M 0 0 L x y …")`

转场：`transition(slide, kind, adv_ms=毫秒数, on_click=False)`；`adv_ms` 为自动换片间隔，
清单见 `references/effects-catalog.md` 第 5 节。

容器—文字配对核对：`scripts/pairing_check.py` 仅用标准库，CLI
`python3 pairing_check.py <deck.pptx> [--min-gap 150] [--max-gap 450] [--max-ctn 60] [--json OUT] [--md OUT]`，
`--selftest` 跑内联 fixture 自检；退出码 0=全 PASS、1=有 FAIL、2=输入错误。

内容识别与效果路由：`scripts/effect_router.py` 仅用标准库，CLI
`python3 effect_router.py <deck.pptx|page.json> [--page N] [--engine effect_first|compat] [--json OUT] [--md OUT]`；
识别页型与信息特征，输出批次拆法、预设族、强调对象、转场与路线建议，逐项带规则编号；
报告顶层 `engine_profile` 标注档位（默认 `effect_first`，`compat` 走兼容优先安全族），批次带 `degrade` 降级提示；
退出码 0=正常出计划、2=输入不可读。规则矩阵见 `references/effect-routing.md`。

跨包导入：`scripts/slide_importer.py` 仅用标准库做文件级 zip 手术；`SlideImporter(dst_pptx)`
提供 `import_slide` / `replace_text` / `reorder` / `save`，模块级 `inspect` 读取参考页结构快照，
支持有序整串与按形状名两种文字替换模式。完整口径见 `references/import-guide.md`。

## 双档适配

- **效果优先档（默认）**：目标 365 / 新版 WPS；放开 Morph 配对补间、p14 现代转场与高版本强调预设，安全族仅作可选兜底。
- **兼容优先档（可选）**：明确要求兼容 2019 / 旧 WPS 时启用；只下发标准 DrawingML 安全族（入场、强调、退出与 ECMA 基础转场），高版本项不下发或仅作点缀。
- **共同底线**：文件结构合法、能打开、不弹修复、内容不丢、不崩；弱引擎上允许降级但不允许损坏。
- **排除层**：内嵌视频/mediacall、翻书页卷曲、3D、空壳预设、非商业素材，任何档位都不进模块。

动工前先声明目标引擎档位；未声明按效果优先档交付，并在说明里写清 2019 上的降级预期。跨软件交付前在 PowerPoint 与 WPS 各 F5 一次。

## 已知边界

- Morph 只在前后两页同名、同类型、数量一致的对象配对成立时才有补间效果；2019 另存时
  补间会被剥离，脚本用 `mc:Choice` + `mc:Fallback` 包裹，兜底 fade。
- 描边连续生长、翻书页、页面卷曲在 2019 对象动画里没有对应预设；用分段擦除/放大近似，
  不写伪节点。
- LibreOffice 只原样保留 XML 字符串，不能用来验证动画是否生效；filter 字符串被
  PowerPoint 忽略，渲染只由 `presetID + presetSubtype` 决定。

## License

MIT，见 [LICENSE](LICENSE)。`references/source-inventory.md` 记录素材许可分级与使用纪律；
本仓库不打包第三方样稿素材。
