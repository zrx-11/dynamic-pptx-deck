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
- 一套三层渐进适配模型：安全层双端通用，高版本增强不认则静默回退，排除层不进稿；
- 五套主配色基线 + 11 套学术配色，通用页型骨架与配方级版式（封面、目录、卡片、双栏、时间线、表格、Q&A、数字带、轮播、模糊大字页、结尾）；
- 17 套效果配方：页数规则、素材门槛、图层结构、移植约束与跨配方通用规则；
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
│   ├── layout-recipes.md      配色基线（含 11 套学术配色）、页型构成与字号阶梯、页型骨架与母题、配方级版式、工程注意事项、许可边界
│   ├── source-inventory.md    素材许可分级、样稿预设统计归档、路线选型与验证边界
│   ├── effect-recipes.md      17 套效果配方（页数规则、素材门槛、图层结构、移植约束、跨配方通用规则）
│   └── import-guide.md        跨包整页导入口径：选型、前置检查、处理范围、主题色解析、文字替换、验证清单
└── scripts/
    ├── dynamic_deck.py       形状/文本/组合体封装、animate、transition、validate、CLI check/demo
    └── slide_importer.py     跨包导入：文件级 zip 手术、媒体哈希去重、主题色解析、文字替换、页序重排
```

## 安装

仓库本身就是 skill 目录，clone 到 agent 的用户级技能目录即可：

```bash
git clone https://github.com/zrx-11/dynamic-pptx-deck.git \
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

5. 用 LibreOffice 导出 PDF/截图核对静态版式；移植类与 L1 效果在目标软件 F5 复核。

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

- 入场/退出：`(name, kind, delay_ms, dur_ms)`；退出含 `disappear`、`zoom_out`、`wipe_out_l/r/u/d`
- 自转：`(name, "spin", delay, dur[, by])`
- 放大回弹 / 单次放大：`(name, "grow" | "grow_once", delay, dur)`
- 闪烁：`(name, "blink", delay, dur)`
- 路径漂移：直线 `(name, "path", delay, dur, dx, dy)`；折线巡游 `(name, "path", delay, dur, 0, 0, "M 0 0 L x y …")`

转场：`transition(slide, kind, adv_ms=毫秒数, on_click=False)`；`adv_ms` 为自动换片间隔，
清单见 `references/effects-catalog.md` 第 5 节。

跨包导入：`scripts/slide_importer.py` 仅用标准库做文件级 zip 手术；`SlideImporter(dst_pptx)`
提供 `import_slide` / `replace_text` / `reorder` / `save`，模块级 `inspect` 读取参考页结构快照，
支持有序整串与按形状名两种文字替换模式。完整口径见 `references/import-guide.md`。

## 三层适配

- **L0 安全层**：2019 / 365 / WPS 通用的入场、强调、退出与 ECMA 基础转场。
- **L1 高版本增强**：p14 现代转场、Morph、移植类强调/入场；引擎不认时静默回退，不丢内容。
- **L2 排除层**：内嵌视频/mediacall、翻书页卷曲、3D、空壳预设、非商业素材，不进模块。

主入场永远用 L0；L1 只做同组点缀。跨软件交付前在 PowerPoint 与 WPS 各 F5 一次。

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
