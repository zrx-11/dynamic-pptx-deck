# 成品页资产目录（layout-assets）

> 本目录是「成品页资产索引」：资产本体（pptx / 静帧 png / svg）存放在
> 参考件目录内，**不随技能分发**；本文与配套 JSON 只记录相对路径与结构特征，
> 供「选页 → 整页导入 → 换文案」流水线按页角色、槽位与动画特征检索。

- 索引由 `scripts/scan_layout_assets.py` 一键重扫重建，不要手工改表。
- 页角色共 20 类；`router` 表示高置信自动路由，`人工` 表示 router 低置信后按规则兜底归类。
- 标记「仅可整页复用、不可换文案」的页面：大字已烘焙进位图，或强依赖幻灯片背景填充透出。
- `samples/` 下的实填样例以 `filled_sample` 标记纳入索引，用于核对槽位容量与替换边界；其余为 `master` 母版包。

## 怎么用这份索引

1. **按页角色查**：先在下文角色分组（机读版取 JSON 记录的 `role` 字段）定位目标页型；
   索引未覆盖的角色见文末「缺口角色」，该类页面走自绘兜底分支。
2. **看槽位与容量决定换文案可行性**：每行的「槽位数」「可替换字数」对应该页 `slots` 与
   `capacity_total`；目标文案先对照容量，容量不足时压文案或换页，不通过缩字号硬塞。
3. **看烘焙文字与背景填充决定能不能改字**：`baked_text` 为 `yes`、或 `bgFill` 标「是」
   （`usebgfill` 强依赖）的页面仅可整页复用、不可换文案；`suspect` 表示存在位图文字风险，
   导入前需人工核对。`无原生文本槽` 的页面只作视觉参考。
4. **资产本体不随技能分发**：索引只记录相对路径与结构特征，pptx / 静帧 png / svg 本体留在
   参考件目录；整页导入需在持有参考件的环境执行，操作口径见 `import-guide.md`。
5. **与 `asset-index.md` / `asset-index.json` 区分**：后者是 20 个页角色的版式骨架与效果族
   候选库，供自绘兜底分支选型；本文件是现成成品页的检索索引，供默认复用路线选页。
6. **估算与实测的界限**：索引里的 `capacity` / `capacity_total` 是离线估算；真正替换
   文案前须以实测容量表（`slot-capacity.json`）为准。`baked_text=yes`、`usebgfill=true`，
   或带摄影背景 / 锁定大图的页，不进入改动文案的候选。

## 统计概览

- 总页数：**406**
- 覆盖页角色数：**7 / 20**
- 可换文案页数（非烘焙文字、非背景填充依赖、有文本槽）：**386**

## 封面（cover）· 121 页

- 同构簇：**5 包 × 23 页**，代表包 `academic/templates/学术模板1/reference/学术蓝-导航栏1-可编辑模板.pptx`；成员：`学术蓝-导航栏1-可编辑模板.pptx`、`学术蓝-导航栏2-可编辑模板.pptx`、`学术蓝-导航栏3-可编辑模板.pptx`、`学术蓝-导航栏4-可编辑模板.pptx`、`学术蓝-导航栏5-可编辑模板.pptx`。JSON 保留全部逐页记录，文档仅取代表包行。

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-007-p3 | 005_三段上滚/reference | 6 | 269 | none（high） | 否 | 0 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p23 | academic/templates/学术模板1/reference（同构×5） | 6 | 78 | none（high） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p23 | academic/templates/学术模板1/reference | 6 | 78 | none（high） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p1 | academic/templates/学术模板1/reference（同构×5） | 8 | 51 | suspect（low） | 否 | 0 | light | 封面：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p2 | academic/templates/学术模板1/reference（同构×5） | 8 | 51 | none（high） | 否 | 0 | light | 封面：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p1 | academic/templates/学术模板1/reference | 8 | 51 | suspect（low） | 否 | 0 | light | 封面：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p2 | academic/templates/学术模板1/reference | 8 | 51 | none（high） | 否 | 0 | light | 封面：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p8 | 009_金边弧线/reference | 5 | 40 | suspect（low） | 否 | 0 | light | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p1 | 014_图片轮播/reference | 14 | 32 | none（high） | 否 | 31 | dark | 封面：14 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p3 | academic/templates/学术模板1/reference（同构×5） | 6 | 32 | suspect（low） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p4 | academic/templates/学术模板1/reference（同构×5） | 6 | 32 | suspect（low） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p3 | academic/templates/学术模板1/reference | 6 | 32 | suspect（low） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p4 | academic/templates/学术模板1/reference | 6 | 32 | suspect（low） | 否 | 0 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p5 | academic/templates/学术模板1/reference（同构×5） | 5 | 30 | suspect（low） | 否 | 0 | light | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p5 | academic/templates/学术模板1/reference | 5 | 30 | suspect（low） | 否 | 0 | light | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p1 | 009_金边弧线/reference | 3 | 24 | suspect（low） | 否 | 1 | light | 封面：3 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p4 | 009_金边弧线/reference | 3 | 24 | none（high） | 否 | 0 | dark | 封面：3 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p5 | 009_金边弧线/reference | 3 | 24 | none（high） | 否 | 0 | dark | 封面：3 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p6 | 009_金边弧线/reference | 3 | 24 | none（high） | 否 | 0 | dark | 封面：3 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p16 | 007_环形开场/reference | 4 | 22 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p20 | 007_环形开场/reference | 4 | 20 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p1 | 008_纯矢量动画/reference | 5 | 19 | none（high） | 否 | 0 | light | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p12 | 007_环形开场/reference | 4 | 18 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p8 | 007_环形开场/reference | 4 | 18 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p11 | 010_暗红塔影/reference | 4 | 18 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p17 | 009_金边弧线/reference | 2 | 16 | suspect（mid） | 否 | 1 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p2 | 010_暗红塔影/reference | 13 | 15 | suspect（low） | 否 | 3 | dark | 封面：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p2 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p3 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p4 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p5 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p6 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p7 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p8 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p9 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-014-p2 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-014-p3 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | light | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-014-p4 | 011_中式滑转/reference | 6 | 14 | none（high） | 否 | 3 | dark | 封面：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p33 | 008_纯矢量动画/reference | 7 | 12 | none（high） | 否 | 9 | light | 封面：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p1 | 010_暗红塔影/reference | 12 | 12 | none（high） | 否 | 1 | dark | 封面：12 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p3 | 008_纯矢量动画/reference | 2 | 11 | none（high） | 否 | 0 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p3 | 002_玻璃碎片/reference | 5 | 10 | none（high） | 否 | 0 | dark | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p3 | 002_玻璃碎片/reference | 5 | 10 | none（high） | 否 | 0 | dark | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-004-p3 | 002_玻璃碎片/reference | 5 | 10 | none（high） | 否 | 0 | dark | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p10 | 011_中式滑转/reference | 4 | 10 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-014-p5 | 011_中式滑转/reference | 4 | 10 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-016-p1 | 013_模糊遮罩/reference | 4 | 10 | none（high） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-016-p2 | 013_模糊遮罩/reference | 4 | 10 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-016-p3 | 013_模糊遮罩/reference | 4 | 10 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-016-p4 | 013_模糊遮罩/reference | 4 | 10 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-013-p1 | 011_中式滑转/reference | 4 | 8 | none（high） | 否 | 1 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-014-p1 | 011_中式滑转/reference | 4 | 8 | none（high） | 否 | 1 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p31 | 008_纯矢量动画/reference | 3 | 7 | none（high） | 否 | 1 | light | 封面：3 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p12 | 002_玻璃碎片/reference | 5 | 6 | suspect（low） | 否 | 0 | dark | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p19 | 002_玻璃碎片/reference | 5 | 6 | suspect（low） | 否 | 0 | dark | 封面：5 个可替换槽位，按槽位字数控制文案长度 |
| ref-008-p1 | 006_加载指针/reference | 1 | 5 | none（high） | 否 | 13 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-015-p2 | 012_六边形旋转/reference | 2 | 5 | none（high） | 否 | 2 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-018-p1 | 015_枫叶飘动/reference | 2 | 5 | none（high） | 否 | 6 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-001-p1 | 001_湖景斜切/reference | 1 | 4 | suspect（mid） | 否 | 0 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p2 | 002_玻璃碎片/reference | 1 | 4 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p4 | 002_玻璃碎片/reference | 2 | 4 | none（high） | 否 | 0 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p5 | 002_玻璃碎片/reference | 2 | 4 | none（high） | 否 | 0 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p13 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p14 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p15 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p16 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p17 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p18 | 002_玻璃碎片/reference | 4 | 4 | suspect（low） | 否 | 0 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p2 | 002_玻璃碎片/reference | 1 | 4 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p4 | 002_玻璃碎片/reference | 2 | 4 | none（high） | 否 | 0 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p5 | 002_玻璃碎片/reference | 2 | 4 | none（high） | 否 | 0 | dark | 封面：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-004-p2 | 002_玻璃碎片/reference | 1 | 4 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-006-p1 | 004_扇形展开/reference | 1 | 4 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-006-p2 | 004_扇形展开/reference | 1 | 4 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p4 | 007_环形开场/reference | 4 | 4 | none（high） | 否 | 4 | dark | 封面：4 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p10 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p11 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p6 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p7 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p8 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p9 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p10 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p11 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p6 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p7 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p8 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-003-p9 | 002_玻璃碎片/reference | 1 | 2 | none（high） | 否 | 0 | dark | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p10 | 008_纯矢量动画/reference | 1 | 2 | none（high） | 否 | 1 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p11 | 008_纯矢量动画/reference | 1 | 2 | none（high） | 否 | 1 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p12 | 008_纯矢量动画/reference | 1 | 2 | none（high） | 否 | 1 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p19 | 008_纯矢量动画/reference | 1 | 2 | none（high） | 否 | 3 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p13 | 009_金边弧线/reference | 2 | 16 | yes（mid） | 否 | 0 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |
| ref-011-p16 | 009_金边弧线/reference | 2 | 16 | yes（mid） | 否 | 0 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |
| ref-011-p3 | 009_金边弧线/reference | 2 | 16 | yes（mid） | 否 | 0 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |
| ref-011-p7 | 009_金边弧线/reference | 2 | 16 | yes（mid） | 否 | 0 | light | 封面：2 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |
| ref-011-p19 | 009_金边弧线/reference | 1 | 8 | yes（high） | 否 | 0 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |
| ref-001-p2 | 001_湖景斜切/reference | 1 | 4 | yes（high） | 否 | 0 | light | 封面：1 个可替换槽位，按槽位字数控制文案长度；仅可整页复用、不可换文案 |

## 章节过渡（section_transition）· 4 页

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-015-p3 | 012_六边形旋转/reference | 2 | 17 | none（high） | 否 | 0 | dark | 章节过渡：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-015-p4 | 012_六边形旋转/reference | 2 | 17 | none（high） | 否 | 0 | dark | 章节过渡：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-015-p5 | 012_六边形旋转/reference | 2 | 17 | none（high） | 否 | 0 | dark | 章节过渡：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-015-p6 | 012_六边形旋转/reference | 2 | 17 | none（high） | 否 | 0 | dark | 章节过渡：2 个可替换槽位，按槽位字数控制文案长度 |

## KPI 数字（kpi）· 7 页

- 同构簇：**5 包 × 23 页**，代表包 `academic/templates/学术模板1/reference/学术蓝-导航栏1-可编辑模板.pptx`；成员：`学术蓝-导航栏1-可编辑模板.pptx`、`学术蓝-导航栏2-可编辑模板.pptx`、`学术蓝-导航栏3-可编辑模板.pptx`、`学术蓝-导航栏4-可编辑模板.pptx`、`学术蓝-导航栏5-可编辑模板.pptx`。JSON 保留全部逐页记录，文档仅取代表包行。

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-010-p32 | 008_纯矢量动画/reference | 11 | 107 | none（high） | 否 | 6 | light | KPI 数字：11 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p6 | academic/templates/学术模板1/reference（同构×5） | 7 | 32 | none（high） | 否 | 0 | light | KPI 数字：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p6 | academic/templates/学术模板1/reference | 7 | 32 | none（high） | 否 | 0 | light | KPI 数字：7 个可替换槽位，按槽位字数控制文案长度 |

## 流程步骤（process）· 20 页

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-021-p1 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p10 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p11 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p12 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p13 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p14 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p15 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p16 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p17 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p18 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p19 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p2 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p20 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p3 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p4 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p5 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p6 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p7 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p8 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-021-p9 | 017_纵横导航栏/reference | 16 | 73 | none（high） | 否 | 0 | light | 流程步骤：16 个可替换槽位，按槽位字数控制文案长度 |

## 并列卡片（card_grid）· 170 页

- 同构簇：**5 包 × 23 页**，代表包 `academic/templates/学术模板1/reference/学术蓝-导航栏1-可编辑模板.pptx`；成员：`学术蓝-导航栏1-可编辑模板.pptx`、`学术蓝-导航栏2-可编辑模板.pptx`、`学术蓝-导航栏3-可编辑模板.pptx`、`学术蓝-导航栏4-可编辑模板.pptx`、`学术蓝-导航栏5-可编辑模板.pptx`。JSON 保留全部逐页记录，文档仅取代表包行。

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-007-p2 | 005_三段上滚/reference | 9 | 354 | none（high） | 否 | 0 | dark | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-007-p1 | 005_三段上滚/reference | 8 | 224 | none（high） | 否 | 3 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p4 | 008_纯矢量动画/reference | 45 | 189 | suspect（low） | 否 | 0 | light | 并列卡片：45 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p5 | 008_纯矢量动画/reference | 45 | 184 | suspect（low） | 否 | 0 | light | 并列卡片：45 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p6 | 008_纯矢量动画/reference | 45 | 184 | suspect（low） | 否 | 0 | light | 并列卡片：45 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p7 | 008_纯矢量动画/reference | 45 | 184 | suspect（low） | 否 | 0 | light | 并列卡片：45 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p8 | 008_纯矢量动画/reference | 45 | 184 | suspect（low） | 否 | 0 | light | 并列卡片：45 个可替换槽位，按槽位字数控制文案长度 |
| ref-019-p1 | 016_扫光动画/reference | 17 | 122 | none（high） | 否 | 92 | light | 并列卡片：17 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p2 | 009_金边弧线/reference | 10 | 80 | none（high） | 否 | 8 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p13 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 8 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p14 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 0 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p15 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 0 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p16 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 0 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p17 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 0 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p18 | 008_纯矢量动画/reference | 8 | 75 | suspect（low） | 否 | 8 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p9 | 007_环形开场/reference | 13 | 53 | none（high） | 否 | 0 | dark | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p1 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 1 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p10 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p11 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p12 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p2 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p3 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | dark | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p4 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p5 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p6 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p7 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p8 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | light | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-005-p9 | 003_哲学字动/reference | 10 | 52 | none（high） | 否 | 0 | dark | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p20 | academic/templates/学术模板1/reference | 26 | 52 | none（high） | 否 | 0 | light | 并列卡片：26 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p21 | 007_环形开场/reference | 13 | 50 | none（high） | 否 | 0 | dark | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p22 | 007_环形开场/reference | 13 | 50 | none（high） | 否 | 0 | dark | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p10 | 007_环形开场/reference | 13 | 48 | none（high） | 否 | 0 | dark | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p14 | 009_金边弧线/reference | 6 | 48 | none（high） | 否 | 0 | light | 并列卡片：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p15 | 009_金边弧线/reference | 6 | 48 | none（high） | 否 | 0 | light | 并列卡片：6 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p11 | 014_图片轮播/reference | 18 | 48 | none（high） | 否 | 0 | dark | 并列卡片：18 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p10 | academic/templates/学术模板1/reference | 24 | 48 | none（high） | 否 | 0 | light | 并列卡片：24 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p5 | 007_环形开场/reference | 12 | 46 | none（high） | 否 | 0 | dark | 并列卡片：12 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p10 | academic/templates/学术模板1/reference（同构×5） | 23 | 46 | none（high） | 否 | 0 | light | 并列卡片：23 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p13 | academic/templates/学术模板1/reference | 22 | 44 | none（high） | 否 | 0 | light | 并列卡片：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p7 | academic/templates/学术模板1/reference（同构×5） | 21 | 42 | none（high） | 否 | 0 | light | 并列卡片：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p7 | academic/templates/学术模板1/reference | 21 | 42 | none（high） | 否 | 0 | light | 并列卡片：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p11 | academic/templates/学术模板1/reference（同构×5） | 20 | 40 | none（high） | 否 | 0 | light | 并列卡片：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p13 | academic/templates/学术模板1/reference（同构×5） | 20 | 40 | none（high） | 否 | 0 | light | 并列卡片：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p11 | 007_环形开场/reference | 11 | 38 | none（high） | 否 | 11 | dark | 并列卡片：11 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p21 | academic/templates/学术模板1/reference（同构×5） | 19 | 38 | none（high） | 否 | 0 | light | 并列卡片：19 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p11 | academic/templates/学术模板1/reference | 19 | 38 | none（high） | 否 | 0 | light | 并列卡片：19 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p21 | academic/templates/学术模板1/reference | 19 | 38 | none（high） | 否 | 0 | light | 并列卡片：19 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p18 | academic/templates/学术模板1/reference（同构×5） | 18 | 36 | none（high） | 否 | 0 | light | 并列卡片：18 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p18 | academic/templates/学术模板1/reference | 18 | 36 | none（high） | 否 | 0 | light | 并列卡片：18 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p16 | academic/templates/学术模板1/reference（同构×5） | 17 | 34 | none（high） | 否 | 0 | light | 并列卡片：17 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p12 | academic/templates/学术模板1/reference | 17 | 34 | none（high） | 否 | 0 | light | 并列卡片：17 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p16 | academic/templates/学术模板1/reference | 17 | 34 | none（high） | 否 | 0 | light | 并列卡片：17 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p17 | academic/templates/学术模板1/reference | 17 | 34 | none（high） | 否 | 0 | light | 并列卡片：17 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p19 | 007_环形开场/reference | 10 | 32 | none（high） | 否 | 3 | dark | 并列卡片：10 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p14 | academic/templates/学术模板1/reference（同构×5） | 16 | 32 | none（high） | 否 | 0 | light | 并列卡片：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p17 | academic/templates/学术模板1/reference（同构×5） | 16 | 32 | none（high） | 否 | 0 | light | 并列卡片：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p14 | academic/templates/学术模板1/reference | 16 | 32 | none（high） | 否 | 0 | light | 并列卡片：16 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p17 | 007_环形开场/reference | 8 | 30 | none（high） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p18 | 007_环形开场/reference | 8 | 30 | none（high） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p12 | academic/templates/学术模板1/reference（同构×5） | 15 | 30 | none（high） | 否 | 0 | light | 并列卡片：15 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p19 | academic/templates/学术模板1/reference（同构×5） | 15 | 30 | none（high） | 否 | 0 | light | 并列卡片：15 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p19 | academic/templates/学术模板1/reference | 15 | 30 | none（high） | 否 | 0 | light | 并列卡片：15 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p8 | academic/templates/学术模板1/reference（同构×5） | 14 | 28 | none（high） | 否 | 0 | light | 并列卡片：14 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p9 | academic/templates/学术模板1/reference（同构×5） | 14 | 28 | none（high） | 否 | 0 | light | 并列卡片：14 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p8 | academic/templates/学术模板1/reference | 14 | 28 | none（high） | 否 | 0 | light | 并列卡片：14 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p9 | academic/templates/学术模板1/reference | 14 | 28 | none（high） | 否 | 0 | light | 并列卡片：14 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p20 | academic/templates/学术模板1/reference（同构×5） | 13 | 26 | none（high） | 否 | 0 | light | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p22 | academic/templates/学术模板1/reference | 13 | 26 | none（high） | 否 | 0 | light | 并列卡片：13 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p13 | 007_环形开场/reference | 7 | 24 | none（high） | 否 | 0 | dark | 并列卡片：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p14 | 007_环形开场/reference | 7 | 24 | none（high） | 否 | 0 | dark | 并列卡片：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p15 | academic/templates/学术模板1/reference（同构×5） | 12 | 24 | none（high） | 否 | 0 | light | 并列卡片：12 个可替换槽位，按槽位字数控制文案长度 |
| ref-022-p22 | academic/templates/学术模板1/reference（同构×5） | 12 | 24 | none（high） | 否 | 0 | light | 并列卡片：12 个可替换槽位，按槽位字数控制文案长度 |
| ref-027-p15 | academic/templates/学术模板1/reference | 12 | 24 | none（high） | 否 | 0 | light | 并列卡片：12 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p2 | 008_纯矢量动画/reference | 9 | 22 | none（high） | 否 | 0 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p15 | 007_环形开场/reference | 9 | 21 | none（high） | 否 | 6 | dark | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p22 | 008_纯矢量动画/reference | 9 | 19 | none（high） | 否 | 0 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p24 | 008_纯矢量动画/reference | 9 | 19 | none（high） | 否 | 0 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p26 | 008_纯矢量动画/reference | 9 | 19 | none（high） | 否 | 0 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p28 | 008_纯矢量动画/reference | 9 | 19 | none（high） | 否 | 0 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p29 | 008_纯矢量动画/reference | 9 | 19 | none（high） | 否 | 6 | light | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p20 | 008_纯矢量动画/reference | 8 | 16 | none（high） | 否 | 0 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p21 | 008_纯矢量动画/reference | 8 | 16 | none（high） | 否 | 1 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p23 | 008_纯矢量动画/reference | 8 | 16 | none（high） | 否 | 1 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p25 | 008_纯矢量动画/reference | 8 | 16 | none（high） | 否 | 1 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-010-p27 | 008_纯矢量动画/reference | 8 | 16 | none（high） | 否 | 1 | light | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p12 | 002_玻璃碎片/reference | 9 | 15 | suspect（low） | 否 | 0 | dark | 并列卡片：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p13 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p14 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p15 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p16 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p17 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p18 | 002_玻璃碎片/reference | 8 | 13 | suspect（low） | 否 | 0 | dark | 并列卡片：8 个可替换槽位，按槽位字数控制文案长度 |
| ref-015-p1 | 012_六边形旋转/reference | 2 | 4 | none（high） | 否 | 0 | dark | 并列卡片：2 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p1 | 002_玻璃碎片/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-003-p1 | 002_玻璃碎片/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-004-p1 | 002_玻璃碎片/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p1 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 48 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p23 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p24 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p25 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p3 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p6 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-009-p7 | 007_环形开场/reference | 0 | 0 | none（high） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-010-p30 | 008_纯矢量动画/reference | 0 | 0 | none（high） | 否 | 0 | light | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-010-p9 | 008_纯矢量动画/reference | 0 | 0 | none（high） | 否 | 1 | light | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |
| ref-011-p18 | 009_金边弧线/reference | 0 | 0 | suspect（mid） | 否 | 0 | dark | 并列卡片版式骨架，无原生文本槽，仅作视觉参考 |

## 致谢页（thanks）· 1 页

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-016-p5 | 013_模糊遮罩/reference | 4 | 9 | suspect（low） | 否 | 0 | dark | 致谢页：4 个可替换槽位，按槽位字数控制文案长度 |

## 图集 / 图墙（gallery）· 83 页

- 同构簇：**2 包 × 20 页**，代表包 `017_纵横导航栏/reference/效果17-偏学术风版.pptx`；成员：`效果17-偏学术风版.pptx`、`学术导航0-纵横导航栏-偏学术.pptx`。JSON 保留全部逐页记录，文档仅取代表包行。

| id | 来源效果 | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |
|---|---|---|---|---|---|---|---|---|
| ref-029-p11 | academic/学术导航0/samples（filled_sample） | 33 | 122 | suspect（low） | 否 | 0 | light | 图集 / 图墙：33 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p2 | academic/学术导航0/samples（filled_sample） | 34 | 118 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p1 | academic/学术导航0/samples（filled_sample） | 34 | 116 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p5 | academic/学术导航0/samples（filled_sample） | 34 | 115 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p3 | academic/学术导航0/samples（filled_sample） | 34 | 111 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p7 | academic/学术导航0/samples（filled_sample） | 31 | 110 | suspect（low） | 否 | 0 | light | 图集 / 图墙：31 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p6 | academic/学术导航0/samples（filled_sample） | 34 | 109 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p14 | academic/学术导航0/samples（filled_sample） | 34 | 108 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p8 | academic/学术导航0/samples（filled_sample） | 34 | 108 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p9 | academic/学术导航0/samples（filled_sample） | 31 | 108 | suspect（low） | 否 | 0 | light | 图集 / 图墙：31 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p4 | academic/学术导航0/samples（filled_sample） | 34 | 107 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p10 | academic/学术导航0/samples（filled_sample） | 34 | 105 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p13 | academic/学术导航0/samples（filled_sample） | 34 | 103 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p12 | academic/学术导航0/samples（filled_sample） | 31 | 102 | suspect（low） | 否 | 0 | light | 图集 / 图墙：31 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p16 | academic/学术导航0/samples（filled_sample） | 34 | 101 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p15 | academic/学术导航0/samples（filled_sample） | 34 | 99 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p17 | academic/学术导航0/samples（filled_sample） | 34 | 94 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p19 | academic/学术导航0/samples（filled_sample） | 34 | 92 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p18 | academic/学术导航0/samples（filled_sample） | 34 | 91 | suspect（low） | 否 | 0 | light | 图集 / 图墙：34 个可替换槽位，按槽位字数控制文案长度 |
| ref-029-p20 | academic/学术导航0/samples（filled_sample） | 33 | 90 | suspect（low） | 否 | 0 | light | 图集 / 图墙：33 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p10 | 014_图片轮播/reference | 22 | 60 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p10 | 009_金边弧线/reference | 7 | 56 | suspect（low） | 否 | 0 | light | 图集 / 图墙：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p11 | 009_金边弧线/reference | 7 | 56 | suspect（low） | 否 | 0 | light | 图集 / 图墙：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p12 | 009_金边弧线/reference | 7 | 56 | suspect（low） | 否 | 0 | light | 图集 / 图墙：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-011-p9 | 009_金边弧线/reference | 7 | 56 | suspect（low） | 否 | 0 | light | 图集 / 图墙：7 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p3 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p4 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p5 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p6 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p7 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p8 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p9 | 014_图片轮播/reference | 22 | 56 | none（high） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-017-p2 | 014_图片轮播/reference | 18 | 44 | none（high） | 否 | 0 | dark | 图集 / 图墙：18 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p7 | 010_暗红塔影/reference | 22 | 32 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p8 | 010_暗红塔影/reference | 22 | 32 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：22 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p10 | 010_暗红塔影/reference | 21 | 30 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p5 | 010_暗红塔影/reference | 21 | 30 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p6 | 010_暗红塔影/reference | 21 | 30 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p9 | 010_暗红塔影/reference | 21 | 30 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：21 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p1 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p10 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p11 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p12 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p13 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p14 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p15 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p16 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p17 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p18 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p19 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p2 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p20 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p3 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p4 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p5 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p6 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p7 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p8 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-020-p9 | 017_纵横导航栏/reference（同构×2） | 20 | 30 | suspect（low） | 否 | 0 | light | 图集 / 图墙：20 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p3 | 010_暗红塔影/reference | 19 | 27 | suspect（low） | 否 | 2 | dark | 图集 / 图墙：19 个可替换槽位，按槽位字数控制文案长度 |
| ref-012-p4 | 010_暗红塔影/reference | 18 | 24 | suspect（low） | 否 | 1 | dark | 图集 / 图墙：18 个可替换槽位，按槽位字数控制文案长度 |
| ref-002-p19 | 002_玻璃碎片/reference | 9 | 15 | suspect（low） | 否 | 0 | dark | 图集 / 图墙：9 个可替换槽位，按槽位字数控制文案长度 |
| ref-009-p2 | 007_环形开场/reference | 0 | 0 | suspect（low） | 否 | 0 | dark | 图集 / 图墙版式骨架，无原生文本槽，仅作视觉参考 |

## 缺口角色

- `table_list`（表格清单）：暂无成品页
- `timeline`（时间线 / 学期推进）：暂无成品页
- `two_column_compare`（双栏对比）：暂无成品页
- `agenda`（目录）：暂无成品页
- `metric_band`（指标带）：暂无成品页
- `trend_chart`（趋势图）：暂无成品页
- `compare_table`（对比表）：暂无成品页
- `org_chart`（组织结构）：暂无成品页
- `risk_alert`（风险提示）：暂无成品页
- `case_card`（案例卡）：暂无成品页
- `conclusion_list`（结论清单）：暂无成品页
- `quote`（引语页）：暂无成品页
- `closing`（结尾）：暂无成品页

## 覆盖率速览

- 有 ≥1 张可换文案成品页的角色（7）：`cover`、`section_transition`、`kpi`、`process`、`card_grid`、`thanks`、`gallery`
- 有 ≥2 张可换文案成品页的角色（6）：`cover`、`section_transition`、`kpi`、`process`、`card_grid`、`gallery`
- 0 张可用的角色（13）：`table_list`、`timeline`、`two_column_compare`、`agenda`、`metric_band`、`trend_chart`、`compare_table`、`org_chart`、`risk_alert`、`case_card`、`conclusion_list`、`quote`、`closing`
- 仅可整页复用、不可换文案：**20** 页；完整阻断清单与 academic 家族覆盖见覆盖率报告。
