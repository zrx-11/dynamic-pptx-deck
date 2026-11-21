# 效果目录（PowerPoint 2019 原生 .pptx）

> 结论基于 PowerPoint 2019 实机验证与样稿结构核对。目标机离线时只能做结构体检，
> 不能替代 F5 放映复核。LibreOffice 往返会原样保留 XML 字符串，**不能**用来验证动画是否生效。

## 目录

1. 环境事实与路线边界
2. 实测可用效果（标准安全族）
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
- **Morph 分层结论**：PowerPoint 2019 另存后 `p14:morph`/`p159:morph` 补间整段被剥离，
  只剩空 transition；365 / 新版 WPS 才可能执行。按双档模型 Morph 属效果优先档增强项：用
  `mc:AlternateContent` 包裹，`mc:Choice` 写 `p159:morph`，`mc:Fallback` 写 `fade`，
  弱引擎自动走兜底。Morph 配对要求前后两页对应对象同名、同类型、数量一致；对象名以 `!!`
  开头可强制配对（允许跨类型，如圆变方、图变图、字变字）；不要把 Morph 作为对象首现的唯一手段。
- 2019 没有翻书页 / 页面卷曲转场（Keynote/WPS 专属）；最接近透视/三维的只有 doors 类
  p14 转场，不认时回退无转场、不伤内容。
- `p:animEffect@filter` 字符串被 PowerPoint 完全忽略，渲染只由 `presetID + presetSubtype`
  决定。模块里的 filter 仅作 LibreOffice 参考，不要据此判断方向。

## 2. 实测可用效果（标准安全族）

| kind | presetID | subtype | 实机表现 | 行为节点 |
|---|---|---|---|---|
| `fade` | 10 | 0 | 淡入 | set 可见 + animEffect |
| `fly_up` | 2 | 4 | 自下而上滑入 | set 可见 + 两条 anim(ppt_x/ppt_y) |
| `wipe_l` | 22 | 8 | 自左向右擦除 | set 可见 + animEffect |
| `spin` | 8 | emph | 陀螺旋 | animRot by=21600000（=360°） |

注意事项：

- `box`(4) 渲染成瞬间出现：缺真正的缩放行为节点，只给 presetID 时 PowerPoint 只执行
  「设为可见」。zoom/stretch 同理。
- 旧版 wipe 若 subtype 写错（如 0），细线整条瞬间出现，不擦除。
- 自转组合体必须锁定轴心，否则短弧绕自身几何中心转、偏离圆心（见第 8 节）。

## 3. 移植效果（带真实行为节点，未在目标机复测）

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
| `zoom_rot` | 31 | 0 | 缩放 + 90° 回正入场 | w/h 0→原值 + rotation 90→0 + fade |
| `bounce_up` | 37 | 0 | 下方弹入，轻微过冲后落位 | y 两段：`#ppt_y+1`→`#ppt_y-.03`（decel），再 100ms 回 `#ppt_y` |
| `drop_in` | 47 | 0 | 上方小幅下落入场 | y `#ppt_y-.1`→`#ppt_y` + fade |

### 强调

| kind | presetID | 表现 | 行为节点 |
|---|---|---|---|
| `grow` | 26 | 放大回弹 105% | animScale by=105000，autoRev=1 |
| `grow_once` | 6 | 单次放大到 110%，不回弹 | animScale by=110000，dur=250，无 autoRev |
| `blink` | 35 | 隐藏-显现闪烁 | discrete str anim，visibility hidden→visible，grpId=1 |
| `path` | 42 | 相对路径往返漂移 | animMotion，origin=layout，relative，indefinite + autoRev |
| `swing` | 8 | 小角度钟摆回摆 | animRot by 默认 36°，repeatCount=indefinite + autoRev=1；仅左右往返，不整圈自转 |

> `swing` 即社区所称「钟摆」：陀螺旋强调 + 自动翻转 + 重复到页末，是**页内动画，不是转场**；
> PowerPoint 无原生钟摆转场。整圈自转用 `spin`，小幅度呼吸摆动用 `swing`。

> 路径族还存在 preset 35/63/64，均带 animMotion 真实节点，与 42 同族；
> 模块封装 42：直线往返传 dx/dy，多段巡游传 `M 0 0 L x y …` 路径字符串（见 effect-recipes 4.10）。

### 退出

| kind | presetID | subtype | 表现 |
|---|---|---|---|
| `disappear` | 1 | 0 | 直接隐藏（set hidden） |
| `zoom_out` | 23 | 32 | 缩小到 0 退出（w/h 原值→0 + set hidden，真实行为节点） |
| `wipe_out_l/r/u/d` | 22 | 8/2/1/4 | 对应方向擦除退出，animEffect transition=out |
| `fly_out_r/l/u/d` | 2 | 2/8/1/4 | 对应方向滑出画外：起手作者位 → 画外终值，dur 处 `style.visibility→hidden` 闭合 |

> 移植效果在目标机离线时无法 F5 复核；正式交付前若条件允许，优先用第 2 节实测效果，
> 移植效果作为点缀并在验证记录里标注「未复测」。
> 高版本项（`zoom_rot`/`bounce_up`/`drop_in`/`grow_once`）按第 10.2 节处理：
> 效果优先档下发并记降级提示；兼容档主入场退回安全族，高版本项只做点缀或直接替换。

## 4. 已弃用 / 不可用清单

- `box`(4)、旧版 zoom、stretch：缺行为节点，放映等同出现。
- entr 14/10、16/21、12/1、12/4、21/1、16/42：仅 set + animEffect 空壳，filter 被忽略，
  放映等同出现；entr 23/272 周边无 anim/animScale/animRot，同判空壳。
- entr 41/0（按字符 iterate 波浪）、45/0（sin 公式横向摆动）：结构复杂、依赖 iterate/fmla，
  移植风险高，暂不收录。
- Morph 不在本弃用清单：属效果优先档高版本增强项，写法与配对要求见第 1、10 节。
  `p159:morph option="byObject"` 外层为 `mc:Choice` + `mc:Fallback` 结构；低版本剥离
  Choice 后走 Fallback。
- 旧版 preset 23/16：带真实 w/h 缩放节点时等价模块 `zoom`（53/16）；不要按旧版空壳 zoom 判死。
- 翻书页 / 页面卷曲转场：2019 不支持（`p15:prstTrans prst="pageCurlDouble"` 同属此类，不搬）。
- 内嵌 AVI/MP4 与 mediacall：不搬入新稿。
- 缺 animMotion 的 entr 42 是假阳性；只有带 animMotion 的 42 路径可信。
- 退出 fly（exit presetID 2）必须在 dur 处挂 `style.visibility→hidden` 结束节点；
  缺该节点对象停在画外不消失，白名单校验不会单独报错，需人工核对结束行为。
- 补充判死：`emph 3/2`（无行为节点）、`emph 19/0`（仅 set 可见性）、
  `entr 18/12`（仅 set + animEffect 空壳）、`exit 16/37`（animEffect + set 空壳）；
  `exit 38/0` 为字符级 iterate 旋转退出，结构复杂、依赖 `p:iterate` 与逐字 rotation，
  暂不收录。

## 5. 转场目录

ECMA 基础转场（2019 通用）：`fade`、`push_l`、`push_r`、`cover_l`、`cover_r`、
`wipe_l`、`wipe_r`、`wipe_u`、`wipe_d`、`split_h`、`pull`、`blinds_v`、`blinds_h`、
`comb`、`random`、`checker`、`wheel`、`wheel4`、`wheel8`、`wedge`、`circle`、
`zoom_t`、`randombar_v`。

- `wheel` 辐条数由 `spokes` 属性决定：`wheel`=1 辐、`wheel4`=4 辐、`wheel8`=8 辐；
  原生轮辐只出规整放射状边界，做不出不规则扇面。
- `wedge` 楔入：单个扇形绕一圈扫过页面；`circle` 圆形：新页从画面中心涨开。
- 遮罩式转场与版面遮罩是两层：转场只发生在切页瞬间；内容嵌在圆内、圆外压暗的常驻圆环
  属于版面元素（见 layout-recipes 第 7b/9 节），二者不要混为一谈。
- 自动换片：`transition(slide, kind, adv_ms=毫秒数, on_click=False)`；`advTm` 写到 transition 根节点，
  Morph 的 Choice 与 Fallback 两处都会写入。图集/轮播类配方靠自动换片定节奏，不要靠动画时长硬凑。

按观感分组（先定观感，再定方向）：

- 方向推移：`push_l/r`、`cover_l/r`、`pull`、`conveyor`(高版本)——整页被推过去或盖过去。
- 擦除：`wipe_l/r/u/d`——边界扫过，适合流程、时间线与清单。
- 分割与条状：`split_h`、`blinds_v/h`、`comb`、`randombar_v`、`checker`。
- 几何遮罩：`wheel`/`wheel4`/`wheel8`、`wedge`、`circle`——辐条、扇形或圆形张开。
- 缩放与淡化：`zoom_t`、`fade`、`random`。

选型纪律：方向跟全稿主轴保持一致，类型按页型轮换；既不要把清单里的转场全堆进一份稿，
也不要全篇单一转场。页型分工表见 layout-recipes 第 7b 节。

p14 预设转场（高版本，2010+；不认则回退无转场，不伤内容）：`doors`、`gallery`、`flip`、
`switch`、`prism`、`conveyor`、`window`、`honeycomb`、`reveal`、`ripple`、`vortex`、
`glitter`、`ferris`。结构取自样稿核对，未在目标机 F5 复测。

Morph（高版本）：`morph`，`p159:morph option="byObject"`，mc:Choice/Fallback 包裹、兜底 fade；
仅在前后两页同名同类型对象配对成立时才有补间效果。命名空间硬事实：`p159` 必须绑定
`http://schemas.microsoft.com/office/powerpoint/2015/09/main`；写成 `2015/12/main` 时
PowerPoint 不认该前缀声明，整段跳过 Choice，只播 Fallback fade，校验器对 12 版记 ERROR。

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
转起来偏离圆心。

标准纠偏（模块 `spin_group` 已实现）：

1. 组内放一个透明整圆辅助形，把组合包围盒撑成对称正方形；
2. 自转件与辅助形都改本地坐标（off=0,0；ext=组尺寸）；
3. 陀螺旋挂到组合体上，枢轴 = 组合对称中心 = 转盘圆心；
4. 组合体在 `p:bldLst` 里写 `animBg="1"`。

组合体结构：`p:grpSp` 内含 `grpSpPr/a:xfrm`（off/ext 与 chOff/chExt 同尺寸），
子元素先辅助形后自转件。

## 9. 节奏编排规则

生命周期分类（落到本模块 kind）：

- **静态**：背景、母版装饰、页眉页脚，不进 timing。
- **入场**：`fade` / `fly_*` / `wipe_*` / `appear` / `zoom`。
- **强调**：`spin` / `grow` / `blink` / `path`。
- **退出**：`disappear` / `wipe_out_*` / `fly_out_*`。

启动模式：

- 一个 click group 等一次点击；同组元素同时开始，先后靠各自 `delay` 错开。
- `auto=0` 起前 N 批进页即自动播（delay=0），其余 `indefinite` 等点击。
- 环境动效（转盘自转、漂浮点）放第一批自动播；内容条目分批点击出现。

编排纪律：

- 一页内方向要铺开：不要全页只从左飞入；标题淡入、卡片上滑、分隔线擦除、装饰自转。
- 强调动画只给需要被注意的对象（当前章节、关键数字、核心圆环），不要满页闪烁。
- 单页动画批次控制在 3–6 批；每批元素错开 120–300ms，时长 400–800ms。
- 退出动画仅在需要「换内容」的局部页使用；整页换内容交给转场。

### 9a. 容器—文字配对时序纪律

- **同批绑定**：容器（rect / roundRect / 胶囊 / 行底 / 卡片）入场与其内部文字入场
  落在同一 click 批次（mainSeq 同一外层 `p:par`）；跨批即错配。
- **延迟区间**：文字入场延迟 − 容器入场延迟 ∈ [150, 450]ms；负值必判失败。
- **文字只挂 fade**：文字入场不带 `ppt_x` / `ppt_y` 位移、不带 `animMotion` 路径；
  位移类行为只给容器。
- **退出顺序**：两侧都有退出时，文字先消失，容器晚 100—150ms 再走。
- **按批拆组**：同批内错开的多张卡片可共享文字组；跨批的卡片不得共用一个组节点，
  单点延迟满足不了多张，第一击会看到空白卡。
- **强调归属**：`grow` / `blink` 绑数字或文字本身，不绑容器。
- **节点预算**：单页 cTn ≤ 60；预算不够时先砍页头标题滑入，不留孤儿文字。

### 9b. 每页「活着」基线

- 每页至少一个常驻环境动效：环自转或光点路径往返，进页即播，
  `repeatCount="indefinite"`。
- 关键数字配一次 `grow` 回弹；强调不长闪。
- 页内换批用「旧退出 + 新入场」，不做同帧全铺。
- 反例：整页只靠 fade 同帧铺满，放映时只有描边在动，判不达标。

配对矩阵、一眼识别反例形态与检查器口径，见 `animation-orchestration.md`。

## 10. 双档适配模型（效果优先默认 / 兼容优先可选）

跨软件事实：基础形状、文本、常规动画与 ECMA 转场在
PowerPoint 2019 / 365 与 WPS 演示之间互通稳定；高级预设、p14 现代转场、Morph、3D、嵌入视频在
弱引擎上回退为「直接出现」或无转场，**内容不丢、不崩溃**，只是观感降级。
动工前先声明目标引擎档位；未声明时按效果优先档交付，并在交付说明里写清 2019 上的
降级预期（降级但不崩）。

### 10.1 两档共同底线（任何档位都必须满足）

- 文件结构合法、能打开、不弹修复、内容不丢、不崩；
- 2019 上允许效果降级（退回出现 / fade / 无转场），但不允许因高版本写法导致文件损坏；
- 排除层非法内容任何档位都不写（见 10.4）。

### 10.2 效果优先档（默认；目标 PowerPoint 365 / 新版 WPS）

放开 Morph 跨页配对补间、p14 现代转场、`zoom_rot` / `bounce_up` / `drop_in` /
`grow_once`、p14 转场组合与艺术效果预烘焙后的高强度观感；安全族不再强制打底，
仅作可选兜底或局部稳定项。高版本项在弱引擎上的降级行为逐条记入批次降级提示，
不表述为「跨端一致」。

| 效果 | 档位归属 | 弱引擎回退行为 | 编排约束 |
|---|---|---|---|
| `zoom_rot`（31，含 rotation 90→0） | 高版本增强 | 可能映射不稳，退回出现/zoom | 不承担对象首现唯一手段 |
| `bounce_up`（37 弹跳）/ `drop_in`（47 下降） | 高版本增强 | 版本差异，可能回退直接出现 | 可作主入场；兼容档降级为安全族 fly/wipe |
| `grow_once`（emph 6，animScale 无 autoRev） | 高版本增强 | 旧版 WPS 识别不稳 | 强调场景用；需要回弹优先 `grow` |
| `zoom_out`（exit 23/32，w/h→0 缩小退出） | 高版本增强 | 不认则退回 disappear | 局部换内容页点缀；整页换内容交给转场 |
| p14 转场（doors/gallery/flip/switch/prism/conveyor/window/honeycomb/reveal/ripple/vortex/glitter/ferris） | 高版本增强 | 不认则整段回退无转场 | 可全稿使用；方向跟全稿主轴，类型按页型轮换，避免相邻页同类型堆叠 |
| `morph`（p159:morph，mc:Choice/Fallback 包裹） | 高版本增强 | 2019 剥离 Choice，走 Fallback fade | 前后两页对应对象必须同名、同类型、数量一致；配对对象不叠加页内位置动画；不依赖它做首现 |

### 10.3 兼容优先档（非默认；明确要求兼容 2019 / 旧 WPS 时启用）

仅当用户明确要求在 PowerPoint 2019 / 旧版 WPS 放映，或目标机版本未知且不能接受
降级观感时启用。口径为标准 DrawingML 安全族打底，高版本项仅作同组点缀或不下发。

- 入场：`fade`、`fly_up/l/r/d`、`wipe_l/r/u/d`、`appear`、`zoom`。
- 强调：`spin`、`swing`、`grow`、`blink`、`path`。
- 退出：`disappear`、`wipe_out_l/r/u/d`、`fly_out_l/r/u/d`。
- 转场：ECMA 基础组——`fade`、`push_l/r`、`cover_l/r`、`wipe_l/r/u/d`、`split_h`、
  `pull`、`blinds_v/h`、`comb`、`random`、`checker`、`wheel`、`wheel4`、`wheel8`、
  `wedge`、`circle`、`zoom_t`、`randombar_v`。

### 10.4 排除层（任何交付都不写）

- 内嵌视频 / mediacall；翻书页 / 页面卷曲 / 3D 模型位相机动画；
- 空壳预设（第 4 节清单）；非商业/个人学习许可的模板素材与照片资产。

### 10.5 适配规则

1. 动工前声明目标引擎档位；未声明按效果优先档，交付说明写明 2019 降级预期。
2. 效果优先档可全量写入高版本项；兼容档主入场用安全族，高版本预设只作点缀或不下发。
3. Morph 只用于前后页可配对的对象；配对不成立时不要写 morph 转场；配对对象不叠加
   页内位置动画，两套时间轴会打架。
4. 生成后必须过 `validate` 结构体检：spid 命中、cTn id 唯一、子元素顺序正确；结构错误导致的
   「整页动画被丢弃」不属于特性回退，必须修掉。
5. 跨软件交付前在 PPT 2019、PPT 365、WPS 新版各 F5 一次；回退项记入验证记录（档位标注）。
6. 字体按 `a:latin`/`a:ea` 族名记录；两端缺字体会回退系统字体，正式放映机预装字体或换同档黑体。
7. WPS 与 PPT 的动画窗格、计时入口位置不同，但 click/with/after 语义一致；自动批次与点击批次的
   分层写法在两端通用。

## 11. Morph 配对帧配方（高版本，配对成立才写）

丝滑的跨页位移主要来自 Morph 配对帧序列，而非页内对象动画。18 套完整效果配方（页数规则、
素材门槛、图层结构、移植约束）见 `effect-recipes.md`；本节只列跨配方通用的帧序列模式：

1. **三帧起手**：隐藏态页（文字/面板在画布外负坐标）→ 落位帧（对象到目标位置）→ 停留帧
   （与落位帧完全相同）。停留帧必不可少，否则 Morph 一闪而过。
2. **越界准备帧**：下一段要推入的装饰/面板先放在画布外（负坐标或 y=-50cm），目标帧再推入；
   画布外元素是 Morph 推拉用的，**不能删**，只能改填充指向。
3. **配对纪律**：前后两页对应对象同名、同类型、数量一致；Morph 位移的对象不要再叠加页内动画，
   否则两套时间轴打架。需要跨类型配对（圆变方、图变图、字变字）时，配对对象名以 `!!` 开头，
   可强制 Morph 按同名匹配，不再要求形状类型一致。
4. **时长与节奏**：常用 1.5–2 秒（p14:dur 1500–2000）；需要自动翻页时在页面
   换片时间上设置，不要靠动画时长硬凑。
5. **重影坑**：同一段文字若有大小两份文本框，副本必须移出画布（y=-40cm）隐藏，否则成稿出现
   「标题2 / 标题2」重影。
6. **媒体分组**：模糊遮罩类页面清晰层/模糊层/切片层是三组不同媒体，换图要按页分组替换，
   只换一张会串旧图。
