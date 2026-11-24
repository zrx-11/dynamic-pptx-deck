# 装饰层预渲染管线（decor_bake）

> 本文件规定装饰层（渐变、光晕、噪点、纹理、柔化阴影、模糊、光斑、暗角）的预渲染口径：
> 支持哪些图层类型与参数、内置预设、产物清单与可复现性判定，以及成品优先档下产出图如何
> 登记进 deck-manifest。工具为 `scripts/decor_bake.py`，仅依赖 Pillow 与标准库。
>
> 分层原则与档位口径见 `editability-audit.md`；装饰层属于可整体图像化的一层，内容层与
> 结构层不在本管线范围内。

## 目录

1. 运行方式
2. spec 结构
3. 图层类型与参数表
4. 内置预设
5. 产物与清单结构
6. 可复现性判定
7. 在成品优先档中的消费方式
8. 管线不覆盖什么

---

## 1. 运行方式

```
python3 scripts/decor_bake.py --spec SPEC.json --out-dir DIR [--only ID] [--preset 名称]
python3 scripts/decor_bake.py --preset 深底光斑 --out-dir DIR
```

- `--spec SPEC.json`：参数化图层定义；与 `--preset` 二选一。
- `--out-dir DIR`：产物目录，不存在时自动创建。
- `--only ID`：只输出指定 `id` 的图层，其余图层不参与合成。
- `--preset 名称`：直接使用内置预设出图，不需要先写 spec 文件。
- 退出码：`0` 成功；`2` 用法错误、spec 不可读或输出目录不可用。

## 2. spec 结构

```json
{
  "canvas": [1920, 1080],
  "layers": [
    {"id": "bg",   "type": "gradient",    "params": {"colors": ["#0E3B36", "#071E1C"], "angle": 120, "stops": [0, 1]}},
    {"id": "glow", "type": "glow",        "params": {"center": [0.62, 0.38], "radius": 0.55, "color": "#7FE3C8", "strength": 0.55}},
    {"id": "noise","type": "noise",       "params": {"amount": 0.04, "scale": 2, "seed": 7}},
    {"id": "bn",   "type": "blur_noise",  "params": {"amount": 0.05, "scale": 8, "radius": 6, "seed": 11}},
    {"id": "vig",  "type": "vignette",    "params": {"strength": 0.35, "radius": 0.75}},
    {"id": "tex",  "type": "texture",     "params": {"kind": "grain", "amount": 0.03, "seed": 3}},
    {"id": "blur", "type": "blur",        "params": {"radius": 4}}
  ]
}
```

- `canvas`：画布宽高（像素）。
- `layers`：按数组顺序叠合；每个图层含 `id`、`type`、`params`。
- `id` 为 `bg` 的图层参与整幅合成；合成图额外输出为 `baked-<sha8>.png`。

## 3. 图层类型与参数表

| 类型 | 用途 | 参数 | 默认值 |
|---|---|---|---|
| `gradient` | 线性渐变底色 | `colors`：色标 `#RRGGBB` 列表（≥2）；`angle`：渐变角度（度）；`stops`：各色标位置 0–1，缺省均分 | `angle: 90`，`stops` 均分 |
| `glow` | 径向光斑 / 光晕 | `center`：中心归一化坐标 `[x, y]`；`radius`：半径（占画布短边比例）；`color`：光晕色；`strength`：强度 0–1 | `center: [0.5, 0.5]`，`radius: 0.5`，`strength: 0.5` |
| `noise` | 细颗粒噪点 | `amount`：噪点不透明度 0–1；`scale`：颗粒块边长（像素）；`seed`：随机种子 | `amount: 0.05`，`scale: 2`，`seed: 0` |
| `blur_noise` | 噪点后接高斯模糊 | `amount`、`scale`、`seed`：同 `noise`；`radius`：模糊半径（像素） | `radius: 6` |
| `vignette` | 边缘暗角 | `strength`：暗角强度 0–1；`radius`：从边缘向内生效的范围比例 0–1 | `strength: 0.35`，`radius: 0.75` |
| `texture` | 表面纹理 | `kind`：`grain`（细颗粒）/ `paper`（纸张纤维）/ `dots`（网点）；`amount`：纹理强度 0–1；`seed`：随机种子 | `kind: grain`，`amount: 0.03` |
| `blur` | 整幅柔化 | `radius`：模糊半径（像素）；该类型不产出独立位图，在合成阶段直接作用于当前整幅画布 | `radius: 4` |

参数约定：颜色统一为 `#RRGGBB`；归一化坐标与比例取值范围 0–1；涉及随机的图层必须带
`seed`，缺省为 0。

## 4. 内置预设

| 预设名 | 画面取向 | 图层构成 |
|---|---|---|
| `深底光斑` | 深青底 + 薄荷色中心光晕 + 细噪点 + 轻暗角 | `gradient`、`glow`、`noise`、`vignette` |
| `暖橙渐变` | 暖橙到深棕斜向渐变 + 低位柔光 + 细颗粒 | `gradient`、`glow`、`texture(grain)` |
| `靛蓝噪点` | 靛蓝深底 + 颗粒噪点 + 模糊噪点层 | `gradient`、`noise`、`blur_noise` |
| `暗角纹理` | 近黑灰底 + 强暗角 + 纸张纹理 | `gradient`、`vignette`、`texture(paper)` |

预设画布统一为 1920×1080；需要其他尺寸时改用 `--spec` 自定义。

## 5. 产物与清单结构

输出目录包含：

- 每个图层一张 `<id>.png`（RGBA；`blur` 层为透明占位图，参数记录在清单中）。
- 存在 `bg` 图层时，额外输出整幅合成图 `baked-<sha8>.png`，`<sha8>` 为合成图 sha256
  的前 8 位十六进制。
- `manifest.json`：

```json
{
  "canvas": [1920, 1080],
  "layers": [
    {"id": "bg", "type": "gradient", "params": {"colors": ["#0E3B36", "#071E1C"], "angle": 120},
     "out": "bg.png", "sha256": "<图层图 sha256>"}
  ],
  "baked": "baked-<sha8>.png",
  "baked_sha256": "<合成图 sha256>",
  "generator_version": "decor_bake/1"
}
```

## 6. 可复现性判定

- **字节级可复现**：同一份 spec、同一个 `seed` 运行两次，每个图层 PNG 与合成图的
  sha256 必须一致；清单中的 `sha256`、`baked_sha256` 随之一致。
- 随机数只来自图层参数中的 `seed`（`random.Random(seed)`），不读取系统时间、不依赖目录
  顺序或环境状态；图像滤波与缩放使用固定参数，不引入随运行变化的默认值。
- 复现性校验方式：同一 spec 连续输出到两个独立目录，逐文件比对 sha256；不一致即视为管线
  未通过复现验收。
- 修改任一颜色、角度、强度、尺寸或 `seed`，产物哈希随之变化，属正常结果。

## 7. 在成品优先档中的消费方式

成品优先档（`fidelity_first`）下，装饰层允许整体图像化：预渲染出的背景、光晕、噪点、
暗角、纹理作为贴图放进成稿，不要求在成稿内保留可调的装饰参数。消费流程：

1. 用 `decor_bake.py` 按 spec 渲染，取 `baked-<sha8>.png` 或需要的单层 PNG。
2. 在成稿同级 `deck-manifest.json` 的顶层 `assets[]` 登记该图：

```json
{
  "path": "assets/baked-decor.png",
  "generator": "decor_bake/gradient",
  "params": {"colors": ["#0E3B36", "#071E1C"], "angle": 120},
  "sha256": "<与产物清单 baked_sha256 一致>"
}
```

3. 用到该图的页面条目在 `assets` 中引用同一 `path`；页面文案、要点、数据仍登记在
   `content`，保持内容层归属可追溯。
4. 兼容优先档（`compat`）需要装饰参数可回调时，不在本管线贴成品图，改用成稿内可编辑
   效果节点；烘焙位图层在该档计入阻断。

边界：预渲染图只承担装饰层；正文、要点、数据、表格文字仍应保持 `a:t` 文本节点，页序、
页型、内容归属仍以结构化源为准，不能烤进装饰图。

## 8. 管线不覆盖什么

- **不渲染成稿版面**：本管线只出装饰位图，不排版文字、不绘制图标与分区结构。
- **不做内容层图像化**：文案、数据、表格不属于装饰层，不通过本管线烘焙。
- **不校验贴图与成稿是否贴合**：图放进哪一页、放在背景还是叠层，由排版流程决定；管线
  只保证指定 spec 下产物可复现。
- **不处理字体与文字效果**：渐变文字、发光字等涉及字形的效果不在本管线范围。
