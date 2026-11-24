# 跨包幻灯片导入指南

> 本文件规定把参考件中现成的高冲击效果页整页导入目标包的操作口径。导入工具为
> `scripts/slide_importer.py`，仅用标准库做文件级 zip 手术，不依赖 python-pptx。
> 选型与许可边界见 `source-inventory.md`；动画白名单、行为节点与引擎族适配见
> `effects-catalog.md`；配方形态与移植约束见 `effect-recipes.md`。

## 目录

1. 总原则与三路线判定
2. 前置检查（inspect）
3. 适配检查（check_slide）
4. 适配动作（adapt_slide）
5. 导入处理范围
6. 主题色解析与重着色口径
7. 文字替换（两种模式）
8. 组装与页序编排
9. 与原生构建路线的分工边界
10. 验证纪律
11. 接口与 CLI 参考

---

## 1. 总原则与三路线判定

- **适配前置**：素材进入目标 deck 前必须先过适配检查，按目标对象（基调、配色族、
  槽位、背景类型）决定动作；禁止未检查的原样导入。目标对象由 `TargetProfile` 描述，
  见第 11.1 节。
- **整页导入（import / adapt）**：参考件中已有满足需求的高冲击效果页，其几何、层级或
  时序由参考件锁定，从零重建会失真；timing / transition 锁定是整页导入的核心价值，
  且适配检查识别出的冲突都有确定性消解路径（recolor、裁切、背景替换、声明处置）时，
  走整页导入，复用其 timing 与 transition，导入后只做文字替换与页序编排，不强行复刻
  形状顶点。
- **抽结构原生重建（`dynamic_deck.py`）**：`useBgFill` 依赖重、烘焙像素主导、形状必须
  大改而 timing 无法保全时，只抽取结构与节奏，用原生封装按配方重建，不整页搬运。
- **阻断**：出现下列任一情形即阻断该素材，退回更换素材或原生重建，不产出与目标风格
  割裂的页面：
  - 烘焙像素未在 `baked_declarations` 中声明；
  - 存在未支持的关系类型；
  - 文本槽位在目标字号阶梯内无可行字号；
  - 基调冲突且无 recolor 路径；
  - `strict` 模式下存在未消解的警告。
- **许可前提**：导入页附带的模板素材、照片、图标与主题文件必须满足 `source-inventory.md`
  许可分级；非商业或个人学习许可的资产不进入公开交付稿。

## 2. 前置检查（inspect）

导入前先用 `inspect` 读取参考页结构快照，逐项核对：

- **文本槽**：`a:t` 文本整串清单与形状名；核对每个槽位对应标题档还是正文档
  （字号分档见 `effect-recipes.md` 4.7），目标文案长度与定宽文本框是否匹配。
- **媒体**：媒体清单与槽位；先比画幅比例再决定槽位，比例不符居中裁切，不拉伸。
- **主题色**：`schemeClr` / `srgbClr` 数量与变换节点；依赖主题色落位的页面在导入后
  按第 6 节口径核对颜色。
- **关系类型**：识别页面挂接的全部关系；未支持的关系类型在导入时显式报错，
  前置检查阶段先识别并剔除该页或替换素材。
- **快照键**：除文本槽、媒体、主题色、关系类型外，快照另含 `background`（背景类型）、
  `useBgFill`（幻灯片背景填充形状清单）、`baked_candidates`（疑似烘焙像素媒体）、
  `colors`（页面取色清单），供第 3 节适配检查使用。

## 3. 适配检查（check_slide）

`inspect` 是结构快照；`check_slide` 对照目标 `TargetProfile` 出适配报告，按四个维度
逐项核对。报告先行：先读报告结论，再决定导入、重建或阻断。

### 3.1 文本槽位

- 槽位容量对照目标文本长度与字号：按槽位宽高、目标字号与行距估算可容纳字数，
  估算口径区分 CJK 字宽与拉丁字宽；`wrap=none` 的定宽文本框按单行容量判定。
- 溢出判定：估算文本超出槽位容量即记 overflow，给出建议字号（在目标字号阶梯内下调）；
  下调后仍无可行字号，按第 1 节阻断。
- 估算性质：槽位估算是离线预判断，最终落位以渲染复核为准；估算通过不替代静帧与
  F5 复核。
- 槽位估算由模块工具 `estimate_slot` 承担。

### 3.2 明暗基调与配色族

- 源页色调判定：按页面取色判定 dark / light 基调；照片背景占主导、无法离线判色时记
  `unknown`，不强行归入任一基调，转人工核对。
- 冲突识别：源页基调、主色与目标基线冲突时列出冲突角色（bg / card / ink / mut /
  accent1 / accent2 / accent3 / line）与对应取值。
- recolor 映射：冲突角色按目标 `palette` 生成重着色建议映射；自动映射可人工覆盖，
  人工覆盖后仍需在报告中留痕并复核。
- 变换保留：重着色保留 `lumMod` / `lumOff` / `alpha` 等颜色变换子节点；遇到 `a:comp`
  补色场景给出警告，不静默改写其变换链。
- 色调与取色工具：`luminance`、`hue_family`、`recolor_srgb`；页面取色见快照 `colors`。

### 3.3 媒体与背景

- 画幅核对：媒体先比画幅比例再决定槽位；比例不符按槽位居中裁切（`srcRect`），
  禁止拉伸变形。
- 背景类型区分：整页图背景在 `p:bg/a:blipFill`；纯色背景在 `p:bg` 纯色填充；
  形状级图片填充挂在各自 `sp` / `pic` 下的 blipFill。背景识别由 `detect_background`
  承担，形状级 blip 由 `find_bg_fill_shapes` 识别。
- `useBgFill` 风险：幻灯片背景填充形状透出的是目标包 `p:bg` 本身；源页背景被替换后，
  这类形状会继承目标背景，导致透出关系错位。处置口径：随背景替换同步核对透出形状，
  无法保全透出效果的页面转原生重建；裁切工具为 `center_crop`。

### 3.4 烘焙像素

- 标记范围：整页照片类媒体，包括画面中不可编辑的像素文字、过期预览图类情形，
  一律列为烘焙像素候选，见快照 `baked_candidates`。
- 声明处置：每个候选媒体必须经 `baked_declarations` 显式声明动作——
  `replace`（替换为目标媒体或原生封装）、`keep`（保留原样并在报告中登记）、
  `block`（剔除该媒体）；未声明的烘焙像素一律阻断，不静默带入目标稿。
- 声明条目：`{"media": partname 或 sha1, "baked_text": str, "action": replace|keep|block}`；
  `baked_text` 记录媒体内固化的不可编辑文本，供替换与核对。

### 3.5 报告与 strict

- `check_slide` 返回 dict：按四维度列出槽位、基调、媒体背景、烘焙像素的警告与建议动作，
  附建议的重着色、裁切与背景处置口径。
- `strict=True`（或 Profile 内 `strict` 置位）时，未消解的警告升级为阻断；
  非 strict 时警告仅提示，由操作者复核后决定。

### 3.6 整页背景替换后的残留核对

`p:bg` 整页图片是最易替换的一类背景；替换同时会连带清掉烘焙在原照片上的像素文字。
换完必须逐项核对：

- 源媒体不残留：旧背景图不再被任何 slide / layout / master 的 rels 引用；
- 无悬空 rels：rels 里不存在指向已删除媒体的 rId，也没有未登记的新媒体；
- Content_Types 同步：新扩展名与默认条目 / Override 一致；
- `ppt/media` 无孤儿图片：入库媒体均有引用，未引用的旧件移出包内。

判定条件：替换后遍历所有 slide rels 与 media 目录，引用集合与文件集合一一对应；
任一媒体有文件无引用、或有引用无文件即判未清干净。反例形态：换了背景但旧照片
仍挂在 layout rels 里；新图扩展名未登记 Content_Types，目标机打不开。

### 3.7 成品优先档下的扁平化判定

交付档位为 `fidelity_first` 时，导入与重建过程中按下列边界判定对象能否拍平成图；
`balanced` 仅允许局部装饰审批，`compat` 不接受烘焙位图层。

- **允许拍平成图**：
  - 装饰层对象：渐变、光晕、噪点、纹理、柔化阴影、模糊、光斑、暗角；
  - 纯视觉性大字：封面主标、章节大数字等不承载事实信息的标题字，版式无法用文本
    实现时可转轮廓或位图；
  - 图表 / 表格的局部视觉修饰可扁平化，但原始数据必须保留在结构化源中。
- **禁止拍平成图**：
  - 正文、要点、数据、表格文字保持 `a:t` 文本节点；仅当版式无法用文本实现时才允许
    图片化，且必须登记；
  - 结构层（页序、页型、内容归属）永不图像化；
  - 不允许无结构化源的整页截图式导入。
- **拍平后结构化源登记**：
  - `deck-manifest.json` 页面条目的 `content` 保留该页全部文案归属，`assets` 列出
    拍平素材路径；
  - `assets[]` 逐条登记素材来源 / 生成器、生成参数与 `sha256`；固化在图内的不可编辑
    文本记入 `baked_text`；
  - facts 清单保留图表 / 表格的原始数据口径，不随扁平化丢失；
  - 分层原则与各档禁止项见 `delivery-tiers.md` 第 2、3 节。

## 4. 适配动作（adapt_slide）

- **写入前阻断**：适配报告存在未消解冲突时，`adapt_slide` 不进入写入流程；
  先消解冲突（补声明、出 recolor 映射、换媒体或转重建）再执行适配。
- **动作参数**：
  - `target`：目标 `TargetProfile`，提供基调、背景色、配色角色、字体与烘焙声明；必传，为 None 时抛错；
  - `text_map`：文本槽位映射，口径见第 7 节；
  - `recolor_map`：角色到目标色的重着色映射；
  - `media_map`：媒体替换映射，替换后按槽位居中裁切；
  - `background`：目标背景处置（整页背景替换或统一为目标纯色）；
  - `strict`：覆盖 Profile 的 strict 口径；不传时沿用 Profile 设置。
- **动作清单**：文本映射落槽；`recolor_map` 按角色重着色并保留变换节点；
  `media_map` 替换媒体并核对画幅；`background` 统一背景；烘焙媒体按声明的
  `replace` / `keep` / `block` 处置。
- **返回**：适配后新页在目标包中的索引（int）。
- 适配动作只改槽位内容、颜色、媒体与背景；形状位置、层级与动画时序不动
  （高节点密度页口径见 `effect-recipes.md` R16）。

## 5. 导入处理范围

- 复制源页 slide part 及其全部 rels；
- 媒体按内容哈希去重，重复媒体不重复入库，并补登记 `[Content_Types].xml`；
- 页面 layout 关系重指目标包的 blank layout；
- 备注页（notes）丢弃；
- 未支持的关系类型显式报错，不静默跳过。

## 6. 主题色解析与重着色口径

- 源页 `schemeClr` 按源 theme 解析为 `srgbClr` 后再进入目标包，避免颜色随目标主题漂移；
- 颜色变换子节点（`lumMod` / `lumOff` / `alpha` 等）保留；
- 文字替换与媒体替换时不改填充色，也不回写被解析掉的主题引用；
- 重着色按目标基线角色落位，冲突时按 `recolor_map` 调整形状填充，不动变换节点结构；
  `a:comp` 补色场景按第 3.2 节给出警告；
- 导入后核对源页主色、强调色与目标稿基线是否冲突；冲突无 recolor 路径时按第 1 节阻断。

### 6b. 主题 fillRef 解析陷阱

- 走 `p:style` 的 `fillRef` 取色的形状（尤其 `custGeom` 自由几何），在部分渲染器上
  会把填充色涂满包围盒，而不是沿路径填充；要跨端一致，必须在 `spPr` 里显式
  `solidFill`，不依赖主题引用。
- 判定方法：用同几何、只改填充色的对照实验，不能凭「原稿渲染正常」推断适配稿正常。
- 对照实验注意：原稿 fillRef 若是近黑色，洗色会被深色底掩盖，问题在浅色适配稿上
  才暴露；浅色 / 深色两种基调各做一次对照。

判定条件：目标稿每个自由几何形状的 `spPr` 内存在显式 `solidFill`；只挂 `fillRef`
且未做对照实验的形状，列为待核验项。反例形态：自由箭头 / 叶形在源稿深色底上看着
正常，换到浅色稿后整块包围盒被涂色，沿路径的镂空丢失。

## 7. 文字替换（两种模式）

- **有序整串匹配（默认）**：按文档顺序对 `a:t` 整串精确匹配，映射键为源文本整串；
  源串在页内重复出现时，按出现顺序对应映射值。
- **按形状名替换**：`by_name=True`，映射键为形状名；适用于同名槽位在多页间复用、
  或整串文本不便作为键的页面。
- 只改文本节点：保留 `rPr` 字号字体与段落结构，不重建文本框（重建会丢失动画绑定，
  见 `effect-recipes.md` 4.4）。
- 替换后核对词宽与换行：定宽文本框字号变大或字数变多会换行、挤乱位移落点；
  长文案先缩短，再按字号阶梯调整，必要时单独放大文本框；槽位溢出以 `estimate_slot`
  预判断，以渲染复核为准。
- **映射组织**：导入时可在 `import_slide` 直接传 `text_map`；也可先导入全部页，
  再用 `replace_text` 逐页改；映射表可落为 `map.json` 供 CLI 读取。

## 8. 组装与页序编排

- 导入时向目标包 `sldIdLst` 与 presentation rels 注册新页；
- `reorder(order)` 按目标页索引列表重排页序；导入页与原生页混排时，先定全稿页型
  分工表，再排索引；
- 转场分工与全稿主轴方向仍按 `layout-recipes.md` 7b 统一编排，导入页不脱离全稿方向；
- 删页需同步清理 `sldIdLst`、presentation rels、slide rels、备注页与 Content_Types 五处
  （约束沿用 `effect-recipes.md` R7 条目），导入路线不单独提供删页接口。

## 9. 与原生构建路线的分工边界

同一交付稿可混用两条路线，路线归属先按第 1 节判据判定：

- **原生页（`dynamic_deck.py`）**：承载事实内容——数据、结论、要点与素材纪律由原生构建
  保证正确；常规版式骨架与标准安全族动画走封装函数；`useBgFill` 依赖重、烘焙像素主导
  或形状必须大改的页面在此路线重建。
- **导入页（`slide_importer.py`）**：负责高冲击节奏——开场、章节过渡、图集、散点巡游、
  高节点密度阵列等；只复用已锁定的几何、层级与时序，且冲突均有确定性消解路径。
- 边界纪律：原生页不直接抄参考件顶点；导入页不承担事实性数字与结论，导入页上的
  事实内容经文字替换后仍需回素材核对。
- 混用稿的引擎族适配、转场白名单与弃用清单对两条路线同等生效。

## 10. 验证纪律

1. `check_slide` 报告先行：先出适配报告并消解冲突，再进入写入；报告未通过的页面
   不导入、不适配。
2. `adapt_slide` 写入前阻断：未消解冲突停在写入前，不产出半成品页面。
3. 跑 `dynamic_deck.py check` 做结构体检：spid 命中、`cTn` id 唯一、子元素顺序正确；
   结构错误导致的整页动画丢弃不属于特性回退，必须修掉。
4. LibreOffice 导出静帧只核对版式与文本落位；LO 不解释动画，不能据此判断动效是否生效。
5. 动画与转场必须在目标放映软件 F5 复核；移植类与高版本效果未复测前标注「未复测」。
6. 引擎族适配与排除层对导入页与原生页同等生效：弱引擎降级但不崩、不丢内容；导入页内的
   Morph、p14 转场与移植强调按 `effects-catalog.md` 第 10 节标注回退行为。
7. 媒体分组替换、画布外出画元素、同段文字大小副本、重影与串图等约束沿用
   `effect-recipes.md` 对应配方条目，导入后逐项核对。

## 11. 接口与 CLI 参考

### 11.1 TargetProfile

```
TargetProfile(tone: str, bg_color: str, palette: dict,
              fonts: dict | None = None,
              baked_declarations: list | None = None,
              strict: bool = False)
```

- `tone`：目标基调，取值 `dark` / `light`。
- `bg_color`：目标背景色。
- `palette`：配色角色映射，角色为 `bg` / `card` / `ink` / `mut` / `accent1` /
  `accent2` / `accent3` / `line`。
- `fonts`：字体族映射；不传时沿用目标稿现有字体。
- `baked_declarations`：烘焙像素声明列表，条目为
  `{"media": partname 或 sha1, "baked_text": str, "action": replace|keep|block}`。
- `strict`：为真时适配检查中未消解的警告升级为阻断。
- 方法：`to_dict()` / `from_dict()`，用于 Profile 序列化与落盘。

### 11.2 SlideImporter 方法

```
SlideImporter(dst_pptx: str)
```

| 方法 | 签名 | 口径 |
|---|---|---|
| `import_slide` | `(src_pptx: str, src_index: int = 0, text_map: dict \| None = None, by_name: bool = False) -> int` | 导入参考页，返回新页在目标包中的索引；可随导入一并传入文字映射 |
| `check_slide` | `(src_pptx: str, src_index: int = 0, target: TargetProfile \| None = None, target_texts: dict \| None = None) -> dict` | 对照目标 Profile 出适配报告；返回四维度警告、冲突与建议动作 |
| `adapt_slide` | `(src_pptx: str, src_index: int = 0, target: TargetProfile \| None = None, text_map: dict \| None = None, recolor_map: dict \| None = None, media_map: dict \| None = None, background: str \| None = None, strict: bool \| None = None) -> int` | 按报告消解结果适配并写入目标包，返回新页索引；target 必传（None 抛错），未消解冲突时写入前阻断 |
| `replace_text` | `(slide_index: int, mapping: dict, by_name: bool = False) -> None` | 替换指定页文本；默认有序整串匹配，`by_name=True` 按形状名替换 |
| `reorder` | `(order: list[int]) -> None` | 按目标页索引列表重排页序 |
| `save` | `(path: str \| None = None) -> None` | 落盘；不传路径时回写目标包 |

### 11.3 模块级函数

```
inspect(src_pptx: str, src_index: int = 0) -> dict
```

读取参考页结构快照（文本槽、媒体、主题色、关系类型，及 `background` / `useBgFill` /
`baked_candidates` / `colors`），供前置检查与适配检查使用。

适配与判定工具：

| 函数 | 口径 |
|---|---|
| `recolor_srgb` | 按目标角色色重着色 srgb 取值，保留 `lumMod` / `lumOff` / `alpha` 变换 |
| `luminance` | 计算颜色相对亮度，供 dark / light 基调判定 |
| `hue_family` | 归并色相族，供配色族冲突识别 |
| `estimate_slot` | 按槽位尺寸、字号、行距估算容量与溢出，区分 CJK / 拉丁字宽与 `wrap=none` 单行 |
| `center_crop` | 按槽位比例生成居中裁切 `srcRect`，不拉伸 |
| `detect_background` | 识别页面背景类型（整页图 `p:bg`、纯色、形状 blip） |
| `find_bg_fill_shapes` | 列出幻灯片背景填充（`useBgFill`）形状及其透出关系 |

### 11.4 inspect 快照键

| 键 | 含义 |
|---|---|
| `shape_text` | 按形状名归集的文本整串 |
| `runs` | 文档顺序的 `a:t` 整串清单 |
| `media` | 媒体清单（part、sha1、尺寸） |
| `scheme_count` | `schemeClr` 节点数量 |
| `background` | 背景类型与背景填充指向 |
| `useBgFill` | 幻灯片背景填充形状清单 |
| `baked_candidates` | 疑似烘焙像素媒体及固化文本线索 |
| `colors` | 页面取色清单（含角色与亮度线索） |

### 11.5 CLI

- 结构快照：`inspect --src X.pptx [--index N]`
- 适配检查：`check --src X.pptx [--index N] [--profile profile.json] [--target-texts texts.json]`
- 适配写入：`adapt --src X.pptx [--index N] --dst D.pptx --profile profile.json [--out O.pptx] [--map map.json] [--recolor recolor.json] [--media media.json] [--strict]`
- 整页导入：`import --src X.pptx --index N --dst D.pptx [--out O.pptx] [--map map.json] [--by-name]`
- 文字替换：`replace --dst D.pptx --index N --map map.json [--by-name] [--out O.pptx]`
- 页序重排：`reorder --dst D.pptx --order 0,3,1,2 [--out O.pptx]`
