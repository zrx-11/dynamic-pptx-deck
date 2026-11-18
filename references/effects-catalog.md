# 效果目录（PowerPoint 2019 原生 .pptx）

> 所有结论来自 PowerPoint ProPlus2019 实机探针与原生成片反查。目标机离线时只能做结构体检，
> 不能替代 F5 放映复核。LibreOffice 往返会原样保留 XML 字符串，**不能**用来验证动画是否生效。

## 目录

1. 环境事实与路线边界
2. 实测可用效果（安全打底）
3. 移植效果（结构可信，未复测）
4. 已弃用 / 不可用清单
5. 转场目录
6. timing 树结构与排错顺序
7. 方向位标记与位移起点
8. 自转轴心锁定（转盘/圆环）
9. 节奏编排规则

---

## 1. 环境事实与路线边界

- 交付物必须是原生可编辑 `.pptx`，用 PowerPoint 2019 放映；不用 HTML/JS/PDF 替代。
- 对象动画不在 python-pptx / PptxGenJS 任何 API 里，全部写在每页 `ppt/slides/slideN.xml`
  的 `<p:timing>` 节点；节点写错 PowerPoint 会**静默丢弃整页动画**。
- **Morph 分层结论**：2019 ProPlus 用 COM 另存后 `p14:morph`/`p159:morph` 补间整段被剥离，
  只剩空 transition；365 / 新版 WPS 才可能执行。按三层模型 Morph 归入 L1：用
  `mc:AlternateContent` 包裹，`mc:Choice` 写 `p159:morph`，`mc:Fallback` 写 `fade`，
  弱引擎自动走兜底。Morph 配对要求前后两页对应对象同名、同类型、数量一致；不要把 Morph
  作为对象首现的唯一手段。
- 2019 没有翻书页 / 页面卷曲转场（Keynote/WPS 专属）；最接近透视/三维的只有 doors 类
  p14 转场，不认时回退无转场、不伤内容。
- `p:animEffect@filter` 字符串被 PowerPoint 完全忽略，渲染只由 `presetID + presetSubtype`
  决定。模块里的 filter 仅作 LibreOffice 参考，不要据此判断方向。

## 2. 实测可用效果（安全打底）

| kind | presetID | subtype | 实机表现 | 行为节点 |
|---|---|---|---|---|
| `fade` | 10 | 0 | 淡入 | set 可见 + animEffect |
| `fly_up` | 2 | 4 | 自下而上滑入 | set 可见 + 两条 anim(ppt_x/ppt_y) |
| `wipe_l` | 22 | 8 | 自左向右擦除 | set 可见 + animEffect |
| `spin` | 8 | emph | 陀螺旋 | animRot by=21600000（=360°） |

实测踩坑记录：

- `box`(4) 渲染成瞬间出现：缺真正的缩放行为节点，只给 presetID 时 PowerPoint 只执行
  「设为可见」。zoom/stretch 同理。
- 旧版 wipe 若 subtype 写错（如 0），细线整条瞬间出现，不擦除。
- 自转组合体必须锁定轴心，否则短弧绕自身几何中心转、偏离圆心（见第 8 节）。

## 3. 移植效果（结构取自原生片反查，未在目标机复测）

### 入场

| kind | presetID | subtype | 表现 | 备注 |
|---|---|---|---|---|
| `fly_r` | 2 | 2 | 从右滑入 | anim x 起点 `1+#ppt_w/2` |
| `fly_l` | 2 | 8 | 从左滑入 | anim x 起点 `-#ppt_w/2` |
| `fly_d` | 2 | 1 | 从上滑入 | anim y 起点 `-#ppt_h/2` |
| `wipe_r` | 22 | 2 | 自右向左擦除 | |
| `wipe_u` | 22 | 1 | 自下向上擦除 | |
| `wipe_d` | 22 | 4 | 自上向下擦除 | |
| `appear` | 1 | 0 | 直接出现 | 仅 set 可见，无过渡 |
| `zoom` | 53 | 16 | 缩放淡入 | anim ppt_w/ppt_h 0→原值 + fade |
| `zoom_rot` | 31 | 0 | 缩放 + 90° 回正入场 | w/h 0→原值 + rotation 90→0 + fade；取自批量样稿反查 |
| `bounce_up` | 37 | 0 | 下方弹入，轻微过冲后落位 | y 两段：`#ppt_y+1`→`#ppt_y-.03`（decel），再 100ms 回 `#ppt_y` |
| `drop_in` | 47 | 0 | 上方小幅下落入场 | y `#ppt_y-.1`→`#ppt_y` + fade |

### 强调

| kind | presetID | 表现 | 行为节点 |
|---|---|---|---|
| `grow` | 26 | 放大回弹 105% | animScale by=105000，autoRev=1 |
| `grow_once` | 6 | 单次放大到 110%，不回弹 | animScale by=110000，dur=250，无 autoRev；取自模板反查 |
| `blink` | 35 | 隐藏-显现闪烁 | discrete str anim，visibility hidden→visible，grpId=1 |
| `path` | 42 | 相对路径往返漂移 | animMotion，origin=layout，relative，indefinite + autoRev |

> 野片路径族还出现 preset 35/63/64，均带 animMotion 真实节点，与 42 同族；
> 模块暂只封装 42，需要时按 animMotion 结构扩展。

### 退出

| kind | presetID | subtype | 表现 |
|---|---|---|---|
| `disappear` | 1 | 0 | 直接隐藏（set hidden） |
| `zoom_out` | 23 | 32 | 缩小到 0 退出（w/h 原值→0 + set hidden，真实行为节点） |
| `wipe_out_l/r/u/d` | 22 | 8/2/1/4 | 对应方向擦除退出，animEffect transition=out |

> 移植效果在目标机离线时无法 F5 复核；正式交付前若条件允许，优先用第 2 节实测效果打底，
> 移植效果作为点缀并在验证记录里标注「未复测」。
> 双端（PPT + WPS）交付时，`zoom_rot`/`bounce_up`/`drop_in`/`grow_once` 按第 10.2 节处理：
> 主入场退回安全层， exotic 项只做点缀或直接替换。

## 4. 已弃用 / 不可用清单

- `box`(4)、旧版 zoom、stretch：缺行为节点，放映等同出现。
- entr 14/10、16/21、12/1、12/4、21/1：仅 set + animEffect 空壳，filter 被忽略，放映等同出现
  （批量样稿二次反查确认）；entr 23/272 周边无 anim/animScale/animRot，同判空壳。
- entr 41/0（按字符 iterate 波浪）、45/0（sin 公式横向摆动）：结构复杂、依赖 iterate/fmla，
  移植风险高，暂不收录。
- Morph 不在本弃用清单：归入 L1 高版本增强，写法与配对要求见第 1、10 节。
  2026-11-19 四套外部动态模板（红盘、紫色、深绿山林）转场写 `p159:morph option="byObject"`，
  外层均为 `mc:Choice` + `mc:Fallback` 结构；2019 剥离 Choice 后走 Fallback。
- 旧版 preset 23/16：野片里带真实 w/h 缩放节点，等价模块 `zoom`（53/16）；不要按旧版空壳 zoom 判死。
- 翻书页 / 页面卷曲转场：2019 不支持（模板里的 `p15:prstTrans prst="pageCurlDouble"` 同属此类，不搬）。
- RHINE 模板内嵌的 AVI/MP4 与 mediacall：不搬入新稿。
- 本地 deck 里缺 animMotion 的 entr 42 是假阳性；只有带 animMotion 的 42 路径可信。
- 二次反查（ppt-effects 配方库，48 份参考片 865 页）补充判死：
  `emph 3/2`（无任何行为节点）、`emph 19/0`（仅 set 可见性）、`entr 18/12`（仅 set + animEffect 空壳）、
  `exit 16/37`（animEffect + set 空壳）；`exit 38/0` 为字符级 iterate 旋转退出，结构复杂、依赖
  `p:iterate` 与逐字 rotation，暂不收录，需要时手工从参考片反查。

## 5. 转场目录

ECMA 基础转场（2019 通用）：`fade`、`push_l`、`push_r`、`cover_l`、`cover_r`、
`wipe_l`、`wipe_r`、`wipe_u`、`wipe_d`、`split_h`、`pull`、`blinds_v`、`blinds_h`、
`comb`、`random`、`checker`、`wheel`、`circle`、`zoom_t`、`randombar_v`。

p14 预设转场（L1，2010+；不认则回退无转场，不伤内容）：`doors`、`gallery`、`flip`、
`switch`、`prism`、`conveyor`、`window`、`honeycomb`、`reveal`、`ripple`、`vortex`、
`glitter`、`ferris`。均取自 452 套模板原生片反查，未在目标机 F5 复测。

Morph（L1）：`morph`，`p159:morph option="byObject"`，mc:Choice/Fallback 包裹、兜底 fade；
仅在前后两页同名同类型对象配对成立时才有补间效果。

转场写入位置：`p:sld` 子元素顺序必须是 `cSld → clrMapOvr → transition → timing`；
Morph 的 `mc:AlternateContent` 包裹层占 transition 槽位，体检脚本按槽位归一识别。

## 6. timing 树结构与排错顺序

四层嵌套，缺一层 PowerPoint 就不播：

```
p:timing
└─ p:tnLst
   └─ p:par
      └─ p:cTn id=1 nodeType="tmRoot"
         └─ p:childTnLst
            └─ p:seq concurrent="1" nextAc="seek"
               └─ p:cTn id=2 nodeType="mainSeq"
                  └─ p:childTnLst
                     └─ p:par                    ← 一个 click group
                        └─ p:cTn fill="hold" delay="indefinite"
                           └─ p:childTnLst
                              └─ p:par          ← 同组容器 delay=0
                                 └─ p:childTnLst
                                    ├─ p:par  presetID  ← 第 1 个元素
                                    └─ p:par  presetID  ← 第 2 个元素
```

属性要点：

| 位置 | 属性 | 含义 |
|---|---|---|
| `p:cTn` | `id` | 整页唯一，按文档顺序递增 |
| `p:cTn` | `dur` | 毫秒；根节点 `indefinite` |
| `p:cTn` | `nodeType` | `tmRoot` / `mainSeq` 固定；组首 `clickEffect`，同组其余 `withEffect` |
| `p:cTn` | `presetClass` | `entr` 进入 / `emph` 强调 / `exit` 退出 / `path` 路径 |
| `p:cTn` | `presetID` | 决定渲染成什么效果 |
| `p:cTn` | `presetSubtype` | 决定方向，不能省 |
| `p:cond` | `delay` | 毫秒；click group `indefinite`=等点击，`0`=进页即播 |
| `p:spTgt` | `spid` | 必须是形状 id 整数，不是形状名 |

排错顺序（命中率从高到低）：

1. `p:spTgt/@spid` 是否命中本页 `p:cNvPr/@id`？错一个，那个动画消失。
2. `p:cTn/@id` 有没有重复？重复会让整棵树被丢弃。
3. `p:sld` 子元素顺序是否 `cSld → clrMapOvr → transition → timing`？
4. presetID / presetSubtype 是否在白名单内？
5. 跑 `python3 scripts/dynamic_deck.py check <deck.pptx>` 过结构。
6. 以上都对但仍不动 → 确认放映环境（WPS 与 PowerPoint 行为有差异，以目标机实测为准）。

## 7. 方向位标记与位移起点

subtype 位标记：`1=上 2=右 4=下 8=左`。

飞入位移起点（anim 两条，ppt_x / ppt_y）：

| 方向 | subtype | x 起点 | y 起点 |
|---|---|---|---|
| 上（自下而上） | 4 | `#ppt_x` | `1+#ppt_h/2` |
| 右（自右向左） | 2 | `1+#ppt_w/2` | `#ppt_y` |
| 左（自左向右） | 8 | `-#ppt_w/2` | `#ppt_y` |
| 下（自上而下） | 1 | `#ppt_x` | `-#ppt_h/2` |

## 8. 自转轴心锁定（转盘 / 圆环 / 扇面）

陀螺旋对不规则弧段默认绕其自身几何中心转，短弧质量分布不对称时枢轴落在弧段自身中心，
转起来偏离圆心（实拍故障：最内圈半环扫向中心而不是绕圆心）。

标准纠偏（模块 `spin_group` 已实现）：

1. 组内放一个透明整圆辅助形，把组合包围盒撑成对称正方形；
2. 自转件与辅助形都改本地坐标（off=0,0；ext=组尺寸）；
3. 陀螺旋挂到组合体上，枢轴 = 组合对称中心 = 转盘圆心；
4. 组合体在 `p:bldLst` 里写 `animBg="1"`。

组合体结构：`p:grpSp` 内含 `grpSpPr/a:xfrm`（off/ext 与 chOff/chExt 同尺寸），
子元素先辅助形后自转件。

## 9. 节奏编排规则

生命周期分类（对齐 ppt-master 方法论，落到本模块 kind）：

- **静态**：背景、母版装饰、页眉页脚，不进 timing。
- **入场**：`fade` / `fly_*` / `wipe_*` / `appear` / `zoom`。
- **强调**：`spin` / `grow` / `blink` / `path`。
- **退出**：`disappear` / `wipe_out_*`。

启动模式：

- 一个 click group 等一次点击；同组元素同时开始，先后靠各自 `delay` 错开。
- `auto=0` 起前 N 批进页即自动播（delay=0），其余 `indefinite` 等点击。
- 环境动效（转盘自转、漂浮点）放第一批自动播；内容条目分批点击出现。

编排纪律：

- 一页内方向要铺开：不要全页只从左飞入；标题淡入、卡片上滑、分隔线擦除、装饰自转。
- 强调动画只给需要被注意的对象（当前章节、关键数字、核心圆环），不要满页闪烁。
- 单页动画批次控制在 3–6 批；每批元素错开 120–300ms，时长 400–800ms。
- 退出动画仅在需要「换内容」的局部页使用；整页换内容交给转场。

## 10. 三层渐进增强适配模型（全量写入，接受回退）

跨软件事实（WPS 官方说明 + 本库实测反查）：基础形状、文本、常规动画与 ECMA 转场在
PowerPoint 2019 / 365 与 WPS 演示之间互通稳定；高级预设、p14 现代转场、Morph、3D、嵌入视频在
弱引擎上回退为「直接出现」或无转场，**内容不丢、不崩溃**，只是观感降级。交付策略不再锁定
单一版本：同一份 pptx 全量写入 L0 + L1，弱引擎静默回退；L2 排除层任何版本都不写。

### 10.1 L0 安全层（2019 / 365 / WPS 均按标准 DrawingML 解析）

- 入场：`fade`、`fly_up/l/r/d`、`wipe_l/r/u/d`、`appear`、`zoom`。
- 强调：`spin`、`grow`、`blink`、`path`。
- 退出：`disappear`、`wipe_out_l/r/u/d`。
- 转场：ECMA 基础组——`fade`、`push_l/r`、`cover_l/r`、`wipe_l/r/u/d`、`split_h`、
  `pull`、`blinds_v/h`、`comb`、`random`、`checker`、`wheel`、`circle`、`zoom_t`、`randombar_v`。

### 10.2 L1 高版本增强（写入同稿；不认则回退，不伤内容）

| 效果 | 分层 | 回退行为 | 编排约束 |
|---|---|---|---|
| `zoom_rot`（31，含 rotation 90→0） | 动画 | 可能映射不稳，退回出现/zoom | 不承担对象首现唯一手段 |
| `bounce_up`（37 弹跳）/ `drop_in`（47 下降） | 动画 | 版本差异，可能回退直接出现 | 主入场用 L0，弹跳/下落只做同组点缀 |
| `grow_once`（emph 6，animScale 无 autoRev） | 动画 | 旧版 WPS 识别不稳 | 强调场景用；需要回弹优先 `grow` |
| `zoom_out`（exit 23/32，w/h→0 缩小退出） | 动画 | 不认则退回 disappear | 局部换内容页点缀；整页换内容交给转场 |
| p14 转场（doors/gallery/flip/switch/prism/conveyor/window/honeycomb/reveal/ripple/vortex/glitter/ferris） | 转场 | 不认则整段回退无转场 | 可全稿使用；相邻页转场方向保持动线一致 |
| `morph`（p159:morph，mc:Choice/Fallback 包裹） | 转场 | 2019 剥离 Choice，走 Fallback fade | 前后两页对应对象必须同名、同类型、数量一致；不依赖它做首现 |

### 10.3 L2 排除层（任何交付都不写）

- 内嵌视频 / mediacall；翻书页 / 页面卷曲 / 3D 模型位相机动画；
- 空壳预设（第 4 节清单）；非商业/个人学习许可的模板素材与照片资产。

### 10.4 适配规则

1. 同一份 pptx 全量写入 L0 + L1；不按猜测的目标软件裁剪 L1 项。
2. 对象主入场优先 L0；L1 动画可作同组强调或点缀，避免某个对象只靠 L1 预设首现。
3. Morph 只用于前后页可配对的对象；配对不成立时不要写 morph 转场。
4. 生成后必须过 `validate` 结构体检：spid 命中、cTn id 唯一、子元素顺序正确；结构错误导致的
   「整页动画被丢弃」不属于特性回退，必须修掉。
5. 跨软件交付前在 PPT 2019、PPT 365、WPS 新版各 F5 一次；回退项记入验证记录（L0 / L1 标注）。
6. 字体按 `a:latin`/`a:ea` 族名记录；两端缺字体会回退系统字体，正式放映机预装字体或换同档黑体。
7. WPS 与 PPT 的动画窗格、计时入口位置不同，但 click/with/after 语义一致；自动批次与点击批次的
   分层写法在两端通用。

## 11. ppt-effects 配方库反查结论（2026-11-19）

来源：GitHub `HY-hue113/ppt-effects`（MIT，v1.2.0，17 个配方 + 工具区；48 份参考 pptx、865 页）。
批量统计 preset / 转场 / 行为节点后，可吸收与需排除项如下。

### 11.1 节点统计

- 转场：morph 388 次、fade 391 次，其余 wipe/gallery/prstTrans 各 1；该库主线是 Morph 序列 + 淡入兜底。
- 行为节点：set 251、anim 186、animEffect 149、animMotion 45、animRot 12、animScale 8；无 animColor。
- 路径族：`path 42/35/63/64` 同为 `animMotion` 相对路径（直线或折线，`pathEditMode="relative"`）；
  模块只封装 42，需要折线/弧线段落时按 `_path` 结构改 `path` 字符串即可，不必新增 kind。
- 旧版 zoom `entr 23/16` 带真实 w/h 缩放节点，等价模块 `zoom`（53/16）；`exit 23/32` 即本节
  新增的 `zoom_out`。

### 11.2 Morph 配对帧配方（L1，配对成立才写）

该库 17 个配方中绝大多数丝滑感来自 Morph，而非页内对象动画。可复用的帧序列模式：

1. **三帧起手**：隐藏态页（文字/面板在画布外负坐标）→ 落位帧（对象到目标位置）→ 停留帧
   （与落位帧完全相同）。停留帧必不可少，否则 Morph 一闪而过。
2. **越界准备帧**：下一段要推入的装饰/面板先放在画布外（负坐标或 y=-50cm），目标帧再推入；
   画布外元素是 Morph 推拉用的，**不能删**，只能改填充指向。
3. **配对纪律**：前后两页对应对象同名、同类型、数量一致；Morph 位移的对象不要再叠加页内动画，
   否则两套时间轴打架。
4. **时长与节奏**：配方常用 1.5–2 秒（spd=slow，p14:dur 1500–2000）；需要自动翻页时在页面
   换片时间上设置，不要靠动画时长硬凑。
5. **重影坑**：同一段文字若有大小两份文本框，副本必须移出画布（y=-40cm）隐藏，否则成稿出现
   「标题2 / 标题2」重影。
6. **媒体分组**：模糊遮罩类页面清晰层/模糊层/切片层是三组不同媒体，换图要按页分组替换，
   只换一张会串旧图。

### 11.3 参考片里出现但不收录的预设

见第 4 节补充清单：emph 3/2、emph 19/0、entr 18/12、exit 16/37 为空壳；exit 38/0 字符级
iterate 旋转退出结构复杂，暂不模块封装。
