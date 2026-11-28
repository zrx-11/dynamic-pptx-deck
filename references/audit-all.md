# 合并校验入口 audit_all

> 本文件规定 `scripts/audit_all.py` 的 CLI、报告字段、动效配额判定、`--fast` 取舍、
> 退出码，以及它与既有校验脚本的关系。动画编排与配对矩阵细节见
> `animation-orchestration.md`；分档阻断规则见 `delivery-tiers.md`；
> 可编辑性扣分码见 `editability-audit.md`。

## 1. 定位

`audit_all.py` 是交付前校验的**唯一入口**：一次调用合并五类既有检查与三条后处理检查，
出一份合并报告与一个退出码。建稿完成后只跑这一个入口，不要把结构体检、配对检查、
可编辑性审计、路由脚本分别跑四遍。

三条后处理检查（`morph-missing` / `image-placeholder` / `letters-count`）的判定细则与
反例构造见 `layout-tools.md` 第 5 节。

## 2. CLI

```bash
python3 scripts/audit_all.py <deck.pptx> [--manifest P]
        [--tier fidelity_first|balanced|compat] [--fast]
        [--require-morph-pages 3,5,7]
        [--json OUT] [--md OUT]
python3 scripts/audit_all.py --selftest
```

| 参数 | 含义 |
|---|---|
| `deck` | 待校验 `.pptx`（位置参数） |
| `--manifest` | `deck-manifest.json` 路径；缺省取成稿同级同名文件 |
| `--tier` | 交付档位，默认 `fidelity_first`；可编辑性阻断按档位解析 |
| `--fast` | 跳过全包 schema 深扫，只做结构关键项与动效配额（取舍见第 5 节） |
| `--require-morph-pages` | 显式要求这些页必须挂 `p159:morph` 转场；与 manifest 页字段 `transition=morph` 等效，两者取并集 |
| `--json OUT` | 合并报告写机器可读 JSON |
| `--md OUT` | 人读逐页配额表写 Markdown |
| `--selftest` | 运行内联自检（达标样例退出 0；只挂 Morph、无页内动效的反例退出 1 且命中该页） |

## 3. 合并报告字段

报告顶层结构：

| 字段 | 含义 |
|---|---|
| `deck.file` / `deck.tier` / `deck.exit` / `deck.verdict` | 文件名、档位、总体退出码、`pass` / `fail` |
| `checks.structure` | 结构体检：`exit`、`issues[]`、汇总 `effects` / `ctn` |
| `checks.pairing` | 容器—文字配对：`exit`、`fail`（失败页）、`orphan`（孤儿文字）、`morph_pairs[]`、`conflict`（配对冲突数） |
| `checks.editability` | 分档可编辑性：`exit`、整包 `score`、`blocking_pages[]`、`by_code` 扣分码计数 |
| `checks.schema` | schema 深扫：`ran`（是否执行）、`issues[]` |
| `checks.morph_missing` | 应为 Morph 的页里缺 `p159:morph`（2015/09 命名空间）的页号清单 |
| `checks.image_placeholder` | 占位文案（添加图片 / 插入图片 / 图片占位）但无 `a:blip` 的页号清单 |
| `checks.letters_count` | 单字母文本对象数与页声明不符的页：`{page, expected, actual}` |
| `checks.motion` | 页内动效配额：`exit`、`pages_below_quota[]`（不达标页号） |
| `pages[]` | 逐页明细，字段见下 |
| `summary.pages` | 总页数；`summary.motion_fail_pages` 同不达标页；`summary.extra_fail_pages` 三条后处理检查命中的页并集；`summary.by_code` 全部发现码计数 |

逐页 `pages[]` 字段：

| 字段 | 含义 |
|---|---|
| `index` / `page_type` | 页码（1 起）/ manifest 登记的页型 |
| `effects` / `ctn` | 效果总数 / cTn 节点数 |
| `kinds` | `entr` / `emph` / `exit` / `path` 四类分布 |
| `perceptible` | 计入可感知运动的对象数 |
| `loops` | 持续循环动效数 |
| `quota.ok` / `quota.missing[]` | 配额是否达标；未达标项文案 |
| `findings[]` | 本页发现：`QUOTA_MISSING`（配额不足）、`MORPH_MOTION_CONFLICT`（Morph 配对对象挂了位置类页内动画） |

## 4. 动效配额判定

配额口径写死在脚本内，逐页判定：

- **可感知运动对象 ≥3**：位移 / 缩放 / 旋转 / 擦除 / 路径 / fly / zoom 计入；
  纯 fade、opacity、appear 不计；退出类效果不计入可感知运动对象计数（可存在，不充数）。
- **持续循环 ≥1**：环自转（spin + 无限重复）或光点路径巡游（path + 无限重复）。
- **效果总数 ≤12**；**cTn 节点 ≤60**。
- **Morph 配对对象**（同名同类型跨页出现）不得挂页内位置类动画（位移 / 路径 / fly / wipe），
  其页内只允许 fade / opacity；冲突记 `MORPH_MOTION_CONFLICT`。
- **不合格形态**：只靠 Morph 转场衔接、页内近乎静态（仅纯 fade 入场）的页面判不达标；
  页内动效应挂在非配对对象上（装饰环、光点、连线、数字、标题下划线、卡片内文字等）。

## 5. `--fast` 取舍

- 不加 `--fast`（默认全量）：执行全包 schema 深扫——部件被 ContentTypes 覆盖、关系目标
  齐备、媒体部件存在且非零字节、timing 内 `spTgt` 不悬空。**交付前必须全量跑一次。**
- 加 `--fast`：跳过 schema 深扫，只做结构关键项、配对、可编辑性与配额；适合建稿中途
  快速回环。`--fast` 下 schema 硬伤不纳入退出码，不能替代交付前全量校验。
- 三条后处理检查（`morph_missing` / `image_placeholder` / `letters_count`）在 `--fast`
  与全量两条路径**都执行**，不受 `--fast` 影响。

## 6. 退出码

| 退出码 | 含义 |
|---|---|
| 0 | 全部检查通过 |
| 1 | 任一检查失败：配额不足、配对 / 可编辑性按档位阻断、三条后处理检查命中、非 `--fast` 时 schema 有硬伤 |
| 2 | 输入不可读或参数错误 |

## 7. 与既有脚本的关系

`audit_all.py` 以模块方式复用既有脚本的判定能力，校验阶段不再单独调用它们；
下列脚本保留独立 CLI，**仅用于排障定位**（拆开看某一类问题的明细）：

| 脚本 | 承担的检查 | 单独调用（排障用） |
|---|---|---|
| `dynamic_deck.py` | 结构体检（spid 命中、cTn id 唯一、子元素顺序） | `python3 scripts/dynamic_deck.py check <deck.pptx>` |
| `pairing_check.py` | 容器—文字配对时序、孤儿文字、Morph 配对互斥 | `python3 scripts/pairing_check.py <deck.pptx> [--json OUT] [--md OUT]` |
| `editability_check.py` | 分档可编辑性评分与结构化源齐全性 | `python3 scripts/editability_check.py <deck.pptx> [--manifest P] [--tier 档位]` |
| `effect_router.py` | 内容识别与路线计划，属规划工具，不是校验入口 | 规划阶段取页型 / 批次 / 路线建议；校验阶段不单独跑 |

schema 深扫、页内动效配额与三条后处理检查没有独立脚本入口，只在 `audit_all.py` 内执行。
三条后处理检查源自上游成品后处理工具的缺陷自检，本仓库以 python-pptx 口径重写，
判定细则与反例见 `layout-tools.md` 第 5 节。
