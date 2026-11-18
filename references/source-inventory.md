# 开源项目台账与样稿拆解记录

> 本文件是历史调研与反查结论的存档；技能交付不依赖本机会话路径。公开使用时只引用
> 结论、协议状态与许可纪律，不打包样稿素材。

## 1. 开源项目台账（调研快照，2026-11）

| 项目 | 协议 | 状态 | 可吸收内容 |
|---|---|---|---|
| hugohe3/ppt-master | MIT | 可用 | 生命周期分类（静态/入场/强调/退出）、节奏编排规范 |
| op7418/guizang-ppt-skill | AGPL-3.0 | 注意传染 | 结构思路可参考，产物需隔离协议 |
| zarazhangrui/frontend-slides | MIT | 可用 | Web 版式思路（本路线不交付 Web） |
| Anionex/banana-slides | AGPL-3.0 | 注意传染 | 同上 |
| presenton | Apache-2.0 | 可用 | 内容→版式映射思路 |
| arcsin1/Oh My PPT | Apache-2.0 | 可用 | 章节骨架 |
| NestMold/notrat-ppt-studio-skill | MIT（star 低） | 自担风险 | 本地脚本组织方式 |
| GordenSun/GordenPPTSkill | 素材仅限个人研究 | **不可用于公开演示** | 只看技法，不搬素材 |
| wzpan/BeamerStyleSlides | MIT | 可用 | 学术排版与版式复刻思路 |
| addsumtech/slides_maker | MIT | 可用 | 生成流程组织 |
| atomiechen/THU-PPT-Theme | CC BY-NC-SA 4.0 | **非商业** | 配色与版式思路参照，不搬资产 |

另检索过未深入采用：reveal.js、Slidev、ts-pptx、Achuan-2/SlideSCI、
Binaryify/open-kimi-ppt-skill。

## 2. 本地样稿拆解结论（历史反查）

- 批量静态样稿中真带动画的占比很低；空壳 timing 假阳性判定：只含 tmRoot 无效果节点
  不算带动画。
- 三套带动画 deck 预设统计：
  - deck1（21 页）：entr 42/0×19、10/0×19、2/4×5、2/8×4、53/16×1；
    beh set48 / animEffect39 / anim58；转场 fade×42。
  - deck2（30 页）：22/1×36、22/8×26、42/0×21、14/10×20、10/0×18、16/37×16、
    22/4×15、16/21×12、2/4×8、21/1×4、22/2×4、16/42×3；
    beh set192 / animEffect184 / anim60 / animScale1。
  - deck3（30 页）：42/0×36、22/8×34、22/4×24、22/1×19、16/21×19、22/2×14、
    14/10×14、16/37×8、21/1×5、16/42×5、2/4×5、6/32×3；
    beh set210 / animEffect202 / anim98。
- 注意：14/10、16/37、16/21、16/42、21/1、6/32 等预设在本技能白名单外，
  缺行为节点或已判空壳，不写回模块。
- 深色技术风模板：143 个效果节点；emph 35 是隐藏-显现闪烁（discrete str anim，
  visibility hidden→visible，grpId=1）；exit 1 是 set hidden；path 42 用 animMotion
  相对路径；fly 2/2 结构为 set 可见 + 两条 anim(ppt_x/ppt_y)，cTn 带 decel=71429。
- 四套追加动态样稿结论：
  - 蓝图：全屏照片底 + 同心环/辐条叠加，中心圆盘裁切照片；无转场；内嵌 mediacall 不搬。
  - 红盘：红金藏青配色；path 42 用量最大，fade/wipe/zoom/fly 次之；p159:morph 转场
    归入 L1；旧版 preset 23/16 带真实 w/h 缩放节点，等价模块 `zoom`。
  - 紫色：紫夜全屏底 + 圆环十字准星；仅 fly_up；转场同为 p159:morph。
  - 深绿山林：深青碧 `#437477/#549195` 配色；wipe 主导，path 35/63/64 带 animMotion；
    emph 6/0 即 grow_once；entr 23/272 空壳不移植；转场同为 p159:morph。

## 3. 路线选型记录

- 交付原生 `.pptx`；不再锁定单一放映版本，按 L0/L1/L2 三层全量写入（见 effects-catalog 第 10 节），
  弱引擎静默回退。
- Morph 在 2019 ProPlus 另存时补间被剥离；归入 L1，mc:Choice/Fallback 包裹、兜底 fade，
  仅在 365/新版 WPS 上可能生效；翻书页/卷曲为 Keynote/WPS 专属，2019 无对应 XML，列入 L2。
- 验证边界：离线 Linux 环境只有 LibreOffice 静态渲染能力，只能核对版式与 XML 结构，
  不代表 PowerPoint 放映动画；移植效果须在实机 F5 复核后再标「实测」。

## 4. 素材许可使用纪律

- 公开演示稿：只使用 MIT/Apache 规范思路与原生重画形状；
  不直接打包 Gorden/WPS 素材、不复制 CC BY-NC-SA 主题的版式资产。
- 样稿示例文字、占位图一律不沿用；只抽结构、配色、字体层级、装饰手法。
