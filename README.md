# dynamic-pptx-deck · 高级动态PPT 技能

把任意已有内容（文档、报告、笔记、大纲、旧版静态 pptx）做成**带动画与转场的 `.pptx`**。
这是一个供 AI agent 调用的技能（skill）：它不生成 HTML 幻灯片，也不生成 PDF，交付物始终是
可以在 PowerPoint / WPS 里打开的 OOXML 文件。成品优先为默认，可编辑性按档位可选，
可重生成性是硬要求：结构化源随稿交付，页面可凭源描述重建。

## 默认路线

**检索成品页资产 → 整页导入 → 只替换文案 / 数字 / 配色 / 图片 → 叠加原生动画。**
自绘版式仅在资产索引无匹配页（缺口角色，或同角色无可换文案页）时作为兜底分支。
技能价值在原生动画编排、效果配方、跨包整页导入与校验工具链，不承担版式设计本身。

## 它解决什么问题

python-pptx、PptxGenJS 这类库都不封装对象动画与页面转场；动画必须写进每页
`ppt/slides/slideN.xml` 的 `<p:timing>`，转场写进 `<p:transition>`。节点写错时 PowerPoint
会静默丢弃整页动画，而且不同版本（2019 / 365 / WPS）对高级预设的支持程度不一样。

这个仓库把经过结构核对与分层验证的结论固定下来：

- 一套形状与文本封装函数，配合动画批次、转场、轴心锁定组合体；
- 一份效果白名单与弃用清单，只写带真实行为节点的预设；
- 一套三档交付取向：成品优先档面向直接交付成品，兼顾档保留后续改稿条件，兼容优先档仅在
  明确要求兼容旧版本时启用；引擎族（高版本预设 / 标准安全族）与交付取向分别声明；
- 成品页资产检索索引：按页角色、槽位容量、烘焙文字与背景依赖选现成效果页；
- 跨包幻灯片整页导入：复用参考件现成效果页的 timing / transition，导入后仅做文字替换与
  页序编排，主题色解析为显式色值避免漂移；
- 合并校验入口：结构、配对、可编辑性、schema 深扫与页内动效配额一次出报告。

## 不做什么

- 不交付 HTML/JS/PDF 替代 pptx；
- 不打包第三方模板素材、照片、图标与主题文件本体；资产索引只记录路径与结构特征；
- 不编造事实、经历与统计数字；缺失信息由调用方用占位符标出；
- 离线环境只能做 XML 结构体检与静态版式核对，不能代替目标机 F5 放映复核。

## 仓库结构

```
dynamic-pptx-deck/
├── SKILL.md                  技能入口，agent 首先读这个
├── README.md                 本文件
├── LICENSE
├── .gitignore
├── agents/
│   ├── doubao.yaml           面板显示名与触发配置
│   └── openai.yaml           同构配置
├── references/
│   ├── template-pipeline.md  成品页复用流水线完整工序（选页、导入、容量、配额、校验、声明模板）
│   ├── audit-all.md          合并校验入口：CLI、报告字段、配额判定、--fast、退出码、脚本关系
│   ├── import-guide.md        跨包整页导入口径：选型、前置检查、处理范围、文字替换、组装纪律
│   ├── layout-assets.md      成品页资产索引（人读）：按角色分组的候选页、槽位与可换文案性
│   ├── layout-assets.json      成品页资产索引（机读）：一条一页，供检索与路由
│   ├── asset-index.md         版式骨架 / 效果族候选库（兜底自绘分支选型用，与上面两件事区分）
│   ├── asset-index.json         候选库机读版
│   ├── animation-orchestration.md  编排三段、容器—文字配对矩阵、判定条件与反例形态
│   ├── effect-routing.md      内容识别与效果路由矩阵：页型、批次、预设族、禁用项
│   ├── effects-catalog.md      效果目录、行为节点、转场清单、timing 排错、节奏规则
│   ├── effect-recipes.md       完整效果配方：页数规则、素材门槛、图层结构、移植约束
│   ├── layout-recipes.md       配色字体基线、页型骨架与字号阶梯、母题与环境动效
│   ├── delivery-tiers.md       三档交付取向：场景判定、允许 / 禁止清单、结构化源契约与阻断规则
│   ├── editability-audit.md    可编辑性审计：判定项、扣分码、评分表与修复建议
│   ├── decor-bake.md           装饰层预渲染配方、可复现参数与素材登记口径
│   ├── regeneration-guide.md   结构化源维护与页级局部重生成流程
│   └── source-inventory.md     素材许可分级、样稿预设统计归档与路线选型边界
├── reports/
│   └── layout-asset-coverage.md  成品页覆盖率 / 阻断清单 / 家族覆盖 / 优先复用建议
├── scripts/
│   ├── dynamic_deck.py        形状/文本/组合体封装、animate、transition、validate、CLI check/demo
│   ├── slide_importer.py      跨包导入：文件级 zip 手术、媒体哈希去重、主题色解析、文字替换、页序重排
│   ├── audit_all.py           合并校验入口：结构 / 配对 / 可编辑性 / schema / 动效配额
│   ├── pairing_check.py       容器—文字配对时序核对（标准库）：同批、延迟差、fade-only、退出顺序、互斥
│   ├── effect_router.py       内容识别与效果路由（标准库）：页型/特征/批次/预设/路线计划
│   ├── editability_check.py   可编辑性审计：分层位图 / 文字位图 / 烘焙层判定，按档位出阻断结论
│   ├── decor_bake.py          装饰层预渲染管线：渐变/光晕/噪点/暗角/纹理/模糊铺底，同 spec 同 seed 可复现
│   ├── page_regen.py          页级局部重生成回填：单页替换，其余 slide 部件逐字节不变
│   └── scan_layout_assets.py  成品页资产扫描器：一键重建索引 md/json 与覆盖率报告
└── tests/
    ├── page-type-samples.json
    └── run_sample_check.py
```

## 安装

仓库本身就是 skill 目录，clone 到 agent 的用户级技能目录即可：

```bash
git clone <仓库地址> <你的技能根目录>/dynamic-pptx-deck
```

目录名要和 `SKILL.md` 里的 `name: dynamic-pptx-deck` 一致。其他 agent 放进它自己的
skills 路径，结构相同。

## 依赖

- Python 3.9+
- `python-pptx`、`lxml`（脚本运行必需）
- 静态预览（可选）：LibreOffice headless、`pdftoppm`；核对版式用，不参与生成

```bash
pip install python-pptx lxml
```

## 一次完整流程

### 1. 重扫索引（可选）

索引已随仓库提供；参考件目录变动后重扫重建三份产物：

```bash
python3 scripts/scan_layout_assets.py --root <参考件根目录> \
  --out references/layout-assets.json \
  --md references/layout-assets.md \
  --report reports/layout-asset-coverage.md
```

不带参数时按脚本位置推断参考件根目录，并写入技能内默认位置。索引全部由脚本生成，
不要手工改表。

### 2. 建稿

先按内容大纲定页数与每页信息块，声明交付档位；再查 `references/layout-assets.md` 选页，
记录 `source_id`。适配检查通过后整页导入：

```bash
python3 scripts/slide_importer.py import --src <参考件.pptx> --index <N> \
  --dst <目标.pptx> [--out <out.pptx>] [--map <map.json>] [--by-name]
```

导入后只替换 `a:t` 文本与色值（可换文案 / 数字 / 配色 / 图片），不重排坐标、不改几何
层级；随后按页内动效配额挂动画。兜底自绘分支：写 `build.py` 接入
`scripts/dynamic_deck.py`，用封装函数画形状、用 `animate` 编排批次、用 `transition`
写转场。工序与纪律见 `references/template-pipeline.md`。

### 3. 合并校验

```bash
python3 scripts/audit_all.py <deck.pptx> --manifest deck-manifest.json --tier fidelity_first
```

一次调用合并结构、配对、可编辑性、schema 深扫与动效配额；不要分别跑各检查脚本。
之后一次 LibreOffice 导出 PDF / 静帧核对静态版式，只对 audit 报错的页面看图；修完重跑
一次校验复核。

自测样例：

```bash
python3 scripts/dynamic_deck.py demo /tmp/demo.pptx
python3 scripts/audit_all.py /tmp/demo.pptx
```

## 脚本接口摘要

- **原生构建 `dynamic_deck.py`**：形状与文本 `text` / `rect` / `oval` / `donut` /
  `set_name`；自转组合体 `spin_group` / `spin_arc` / `spin_pie`。动画
  `animate(slide, groups, auto=0)`，`groups` 是点击批次列表，同组元素同时开始，先后靠
  `delay` 错开；`auto=N` 前 N 批进页即自动播。转场 `transition(slide, kind,
  adv_ms=毫秒, on_click=False)`。CLI `check <deck.pptx>` / `demo <out.pptx>` / `selftest`。
- **跨包导入 `slide_importer.py`**：`SlideImporter(dst_pptx)` 提供 `import_slide` /
  `check_slide` / `adapt_slide` / `replace_text` / `reorder` / `save`；模块级 `inspect`
  读结构快照。CLI 子命令 `inspect` / `check` / `adapt` / `import` / `replace` /
  `reorder` / `selftest`。
- **合并校验 `audit_all.py`**：CLI `python3 scripts/audit_all.py <deck.pptx>
  [--manifest P] [--tier fidelity_first|balanced|compat] [--fast] [--json OUT]
  [--md OUT]，`--selftest` 跑内联自检；口径见 `references/audit-all.md`。
- **配对检查 `pairing_check.py`**：CLI `python3 scripts/pairing_check.py <deck.pptx>
  [--min-gap 150] [--max-gap 450] [--max-ctn 60] [--json OUT] [--md OUT]，`--selftest`
  跑内联自检。
- **效果路由 `effect_router.py`**：CLI `python3 scripts/effect_router.py
  <deck.pptx|page.json> [--page N] [--tier 档位] [--engine effect_first|compat]
  [--json OUT] [--md OUT]`。

## 交付档位

- **成品优先档 `fidelity_first`（默认）**：直接交付成品场景；装饰层可预渲染图像化，
  内容层保留原生文本，结构化源随稿交付；可编辑性低不阻断，仅结构化源缺失 / 不全阻断。
- **兼顾档 `balanced`**：成稿后续可能在编辑器内改文案、配色或局部排版；单页评分低于
  阈值（默认 70）或出现整页图片式幻灯片、文字位图化即阻断。
- **兼容优先档 `compat`**：明确要求兼容旧版本时启用；阻断项同兼顾档，且烘焙位图层也
  计入阻断；引擎族只下发标准 DrawingML 安全族，高版本项不下发或仅作点缀。
- **引擎族（独立维度）**：高版本 Morph 配对补间、p14 现代转场与强调预设在 `effect_first`
  引擎族下发，弱引擎降级但不崩；`compat` 引擎族只给标准安全族。档位判定、允许 / 禁止
  清单与分档阻断规则见 `references/delivery-tiers.md`。
- **共同底线**：文件结构合法、能打开、不弹修复、内容不丢、不崩；弱引擎上允许降级但不
  允许损坏。
- **排除层**：内嵌视频 / mediacall、翻书页卷曲、3D、空壳预设、非商业素材，任何档位都
  不进模块。

动工前先声明交付档位；未声明按成品优先档交付。跨软件交付前在 PowerPoint 与 WPS 各
F5 一次。

## 素材许可与交付声明

- 成品页资产属第三方素材：个人学习与校内汇报场景可用；公开放映或商业交付前，必须在
  交付说明 / README 声明资产来源与许可范围，未获授权则退回自绘分支。
- 许可分级规则与素材使用纪律见 `references/source-inventory.md` 第 1、2 节；公开 /
  商业交付随稿附的声明模板见 `references/template-pipeline.md` 第 10 节。
- 资产索引只记录路径与结构特征，不打包资产本体；样稿示例文字、占位图一律不沿用。
- 本仓库代码许可证见 [LICENSE](LICENSE)；第三方模板、照片、图标与主题文件不随仓库分发。

## 交付件模板

成品优先档交付时，成稿同级目录必须包含：

- 成稿 `*.pptx`；
- `deck-manifest.json`：`deck` 文件名、`tier` 档位，逐页 `index` / `page_type` /
  `content` / `assets` / `notes`，并指向 `facts`；
- facts 清单：文案、要点与数据口径，与成稿正文一一对应；
- 素材引用：`assets/` 下图片与烘焙素材，manifest 的 `assets[]` 逐条登记路径、生成器、
  生成参数与 `sha256`；
- 验收报告：合并校验结论（结构、配对、可编辑性、schema、动效配额）。

## 已知边界

- Morph 只在前后两页同名、同类型、数量一致的对象配对成立时才有补间效果；2019 另存时
  补间会被剥离，脚本用 `mc:Choice` + `mc:Fallback` 包裹，兜底 fade。
- 描边连续生长、翻书页、页面卷曲在 2019 对象动画里没有对应预设；用分段擦除 / 放大近似，
  不写伪节点。
- LibreOffice 只原样保留 XML 字符串，不能用来验证动画是否生效；filter 字符串被
  PowerPoint 忽略，渲染只由 `presetID + presetSubtype` 决定。
