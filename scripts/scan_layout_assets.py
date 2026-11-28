#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scan_layout_assets.py — 成品版式页资产扫描器（可重复执行）。

对参考件目录下每个 ``reference/*.pptx`` 成品页包逐页解析，产出结构化资产索引：
形状 / 图片 / 文本槽位 / 配色 / 背景依赖 / 烘焙文字判定 / 动画节点 / 转场 / 页角色。
页角色优先调用同目录 ``effect_router.py``（等价于 CLI ``--json`` 路由计划）；
router 置信度为 low 时，按形状与文案形态的确定性规则归入 20 类之一，并在记录里标明。

设计约束：
- 零第三方依赖（标准库 zipfile / xml.etree / json）；effect_router 为同目录内置脚本。
- 只读参考件目录，不复制、不打包任何 pptx / png / svg 本体；索引里只写相对路径。
- 不摘抄参考件正文、作者信息或来源仓库名；recipe.md / spec.json 仅用于结构与依赖判定。

CLI：
    python3 scripts/scan_layout_assets.py [--root DIR] [--out JSON]
        [--md MD] [--report MD] [--skill-root DIR]
不带参数时按脚本所在位置推断参考件根目录，并在技能根目录默认位置重建三份产物：
``references/layout-assets.json``、``references/layout-assets.md``、
``reports/layout-asset-coverage.md``。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import effect_router as router  # noqa: E402

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "{%s}" % P_NS
A = "{%s}" % A_NS

# 20 类页角色键，顺序与 effect_router.priority 一致。
ROLE_KEYS = [
    "cover", "section_transition", "kpi", "table_list", "timeline", "process",
    "two_column_compare", "card_grid", "agenda", "metric_band", "trend_chart",
    "compare_table", "org_chart", "risk_alert", "case_card", "conclusion_list",
    "thanks", "gallery", "quote", "closing",
]
ROLE_LABEL = dict(router.TYPE_LABEL)

# 标题字号阈值（百分之一磅）：≥28pt 视为主标题槽。
TITLE_SIZE = 2800
# 整页覆盖阈值：图片 / 形状包围盒占画布比例。
FULLBLEED_W = 0.92
FULLBLEED_H = 0.85

PAGE_NUM_RE = re.compile(r"(?:第\s*(\d+)\s*页|_p(\d+)|p(\d+))\.png$", re.I)
SLIDE_FILE_RE = re.compile(r"ppt/slides/slide\d+\.xml$")

# ----------------------------------------------------------------------
# v1.3.0 新增包人工覆盖策略
#
# 018–022 五个效果包由人工按 spec.json 核定角色 / 变体 / 硬约束；
# 覆盖记录的 role_source 标为 ``人工-覆盖``，不伪造 router 结果。
# 旧包（001–017 与 academic）仍走 router + 兜底规则，结构不变。
# ----------------------------------------------------------------------

# 上一轮（v1.2.0，29 包 406 页）逐角色基线：(成品页数, 可换文案页数)。
PREV_ROLE_BASELINE: dict[str, tuple[int, int]] = {
    "cover": (121, 115), "section_transition": (4, 4), "kpi": (7, 7),
    "table_list": (0, 0), "timeline": (0, 0), "process": (20, 20),
    "two_column_compare": (0, 0), "card_grid": (170, 157), "agenda": (0, 0),
    "metric_band": (0, 0), "trend_chart": (0, 0), "compare_table": (0, 0),
    "org_chart": (0, 0), "risk_alert": (0, 0), "case_card": (0, 0),
    "conclusion_list": (0, 0), "thanks": (1, 1), "gallery": (83, 82),
    "quote": (0, 0), "closing": (0, 0),
}
PREV_TOTAL_PAGES = 406
PREV_USABLE_PAGES = 386
PREV_PACKS = 29


def _pr(pages: list[int], role: str, default: str | None = None) -> dict:
    """生成 ``{页码: 角色}`` 映射，default 为未列出页的兜底角色。"""
    return {"pages": {p: role for p in pages}, "default": default}


PACK_POLICY: dict[str, dict] = {
    "018_摄像机转场": {
        "constraint": "normal",
        "lock_palette": False,
        "roles": {**_pr([1], "cover"), "default": "section_transition"},
        "notes": [
            "材料：1 张 16:9 横构图照片（宽高比容差 0.08）；副标题 ≤20 字，超长会缩字号",
            "三层结构：全屏底图 / 中部横幅 / 底部装饰；中部横幅可缺省，单张照片版只保留底图",
            "三层图须同一色调，否则旋转切换时接缝明显；中部横幅建议透明 PNG",
        ],
        "boundary": {
            "cover": "三层结构（全屏底图 / 中部横幅 / 底部装饰）；可承担开场镜头帧，不能承担正文内容页与多图并列展示；中部横幅可缺省，单张照片版只保留底图",
            "section_transition": "三层结构（全屏底图 / 中部横幅 / 底部装饰）；可承担章节过渡的镜头甩动帧，不能承担正文内容页与多图并列展示；中部横幅可缺省，单张照片版只保留底图",
        },
    },
    "019_情景演绎双版本": {
        "constraint": "strong",
        "lock_palette": False,
        "roles": {
            "pages": {**{p: "cover" for p in (1, 2, 3)},
                      **{p: "quote" for p in range(4, 9)},
                      **{p: "case_card" for p in range(9, 14)}},
            "default": "quote",
        },
        "notes_shared": [
            "实为情景演绎页：选页必须按用户素材情况定版本——有透明 PNG 主体 5–6 个 + 背景 2–4 张走 cutout，只有普通照片 5–6 张走 plain",
            "两版页数相同；抠图版每页比免抠版多 1–2 张图，选页时按素材情况判成本",
            "台词为中英两行；英文过长缩字号，不得换行超过两行",
            "画布外的同名对象是下一页 Morph 起点，不可删；低版本 PowerPoint 只剩淡入淡出",
        ],
        "notes_variant": {
            "cutout": "抠图版：透明 PNG 主体 5–6 个（建议 ≥1000px 宽）+ 背景 2–4 张横构图；主体可单独位移/缩放，层次最强",
            "plain": "免抠图版：普通照片 5–6 张（建议 ≥1600px 宽），主体居中偏上；零准备成本，层次弱一档",
        },
        "boundary": {
            "cover": "实为情景演绎页（开场帧）：可承担对话开场的纯文字 + 背景帧；不能承担标准封面长文案排版",
            "quote": "实为情景演绎页（对话帧）：可承担中英两行台词叙事；不能承担标准引语页或流程步骤页",
            "case_card": "实为情景演绎页（招式卡）：可承担单主体招式/案例式卡片；最接近既有 case_card 键，不能承担标准并列卡片网格",
        },
    },
    "020_照片接力转场": {
        "constraint": "strong",
        "lock_palette": False,
        "roles": {**_pr([1, 2, 3, 4], "cover"), "default": "section_transition"},
        "notes": [
            "仅作封面 / 章节过渡，不承担正文内容页",
            "机制：同一张照片两种承载（p:pic 小图 + 幻灯片背景 blipFill 整页填充）；Morph 靠对象名 图片1/2/3 与 !!1/!!2/!!3 配对",
            "背景填充页换图须同时换图片对象与对应背景；画布外待命对象不可删；照片比例与原媒体一致，否则铺满时拉伸",
        ],
        "boundary": {
            "cover": "无足量 16:9 照片不上调；照片数不足会导致重复；可承担标题结构开场帧，不能承担正文信息页",
            "section_transition": "无足量 16:9 照片不上调；照片数不足会导致重复；可承担照片接力过渡帧与 THANK YOU 收尾帧，不能承担正文信息页",
        },
    },
    "021_词云开场": {
        "constraint": "strong",
        "lock_palette": True,
        "roles": {**_pr([1], "cover"), "default": "cover"},
        "notes": [
            "白字黑底，强调色不可改（spec 实测改色像素差 0.00）",
            "槽位：标题 115pt / 短句引子 80pt / 三句 72pt / 关键词 44pt×7",
            "部分槽位初始 x 为负（画布外左侧，靠动画入场）：初始出画槽位，禁止在静态稿里判为越界",
            "全部 12 个文本框都要换，漏一个会留下占位；文本框单行，文字变长加长框，不缩字号、不换行",
        ],
        "boundary": {
            "cover": "可承担单页词云开场；不能承担多段正文与数据展示；配色与出画位置锁定",
        },
    },
    "022_园林摄影集": {
        "constraint": "normal",
        "lock_palette": False,
        "roles": {
            "pages": {
                **{p: "cover" for p in (1, 2)},
                **{p: "section_transition" for p in range(3, 20)},
                20: "agenda",
                **{p: "section_transition" for p in (21, 24, 27, 30)},
                **{p: "card_grid" for p in (22, 23, 25, 26, 28, 29, 31, 32)},
                33: "section_transition",
            },
            "default": "section_transition",
        },
        "notes": [
            "需要 16:9 照片流，最多 17 张；不足会循环重复，须先确认用户是否接受重复",
            "照片全部按 16:9 铺满，竖构图会裁掉大量内容；帧号 n/45 换图后最好连续编号",
        ],
        "boundary": {
            "cover": "摄影集封面帧；需要 16:9 照片流，最多 17 张，不足会循环重复，须先确认用户是否接受重复；不能承担非摄影类封面长文案",
            "section_transition": "摄影集照片流 / 章节扉帧；需要 16:9 照片流，最多 17 张，不足会循环重复，须先确认用户是否接受重复；不能承担非摄影类正文长文案",
            "agenda": "目录页（01–04 四章）；随摄影集风格，文字槽为 18pt 小字，长文案不适用",
            "card_grid": "正文三栏页；随摄影集风格，文字槽为 18pt 小字，每段 ≤14 字，长文案不适用",
        },
        "boundary_pages": {
            33: "结尾待续帧（未完待续 / TO BE CONTINUED）；不能承担标准致谢页，也不承担正文长文案",
        },
    },
}

# 扫描器固定附加段：槽位自适应边界（上游 tools/flex_slots.py 口径，转换为本仓库口径）。
# 写成脚本常量，重扫自动带上；不得在产物 md 里手工删改。
SLOT_ADAPT_RULES_MD = """## 附录：槽位自适应边界（flex_slots 口径，扫描器固定附加段）

> 本段由 `scripts/scan_layout_assets.py` 作为常量写入，重扫自动保留，不得手工删改。
> 坐标换算：原画布 1280×720 px（13.333in 宽幻灯片），1 px = 9525 EMU。

1. **槽位身份可互换**：文框与图形框在视觉层级不变时可互换承载；文本优先文框，整块图优先图形框。
2. **相邻槽可合并**：文案字数超单槽容量 1.5 倍以上时，允许合并同行 / 同列相邻两槽；合并宽度取两槽外框并集，不得越出内容区。
3. **可放大**：向外扩展不得超过原尺寸 25%，且不得与其它槽重叠。
4. **可下移避让**：与标题 / 装饰冲突时整体下移，下移量以冲突消除为准，上限为内容区高度的 15%。
5. **标题槽**：宽 380–1160 px（折算 3,619,500–11,047,800 EMU），高 56 px 单行 / 96 px 两行；顶部对齐，两行时不得压行距。
6. **横向导航条**：厚度 64–72 px 可变，宽度固定。
7. **内容槽一律不得移出画布。**
"""


# ----------------------------------------------------------------------
# 基础解析工具
# ----------------------------------------------------------------------

def _rel(path: str, root: str) -> str:
    return os.path.relpath(path, root).replace(os.sep, "/")


def _sld_size(zf: zipfile.ZipFile) -> tuple[int, int]:
    """画布尺寸（EMU）；缺失时回落 16:9 标准画布。"""
    try:
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        sz = pres.find(f"{P}sldSz")
        if sz is not None:
            return int(sz.get("cx", 12192000)), int(sz.get("cy", 6858000))
    except (KeyError, ET.ParseError, TypeError, ValueError):
        pass
    return 12192000, 6858000


def _luminance(hex_val: str | None) -> int:
    if not hex_val or len(hex_val) != 6:
        return 255
    try:
        r, g, b = (int(hex_val[i:i + 2], 16) for i in (0, 2, 4))
        return int(0.299 * r + 0.587 * g + 0.114 * b)
    except ValueError:
        return 255


def _hue_bucket(hex_val: str | None) -> str:
    """主色系归并：无颜色时返回 none。"""
    if not hex_val or len(hex_val) != 6:
        return "none"
    try:
        r, g, b = (int(hex_val[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return "none"
    mx, mn = max(r, g, b), min(r, g, b)
    if mx - mn < 0.08:
        if mx > 0.85:
            return "white"
        if mx < 0.18:
            return "black"
        return "gray"
    h = 0.0
    if mx == r:
        h = (60 * ((g - b) / (mx - mn)) + 360) % 360
    elif mx == g:
        h = 60 * ((b - r) / (mx - mn)) + 120
    else:
        h = 60 * ((r - g) / (mx - mn)) + 240
    if h < 15 or h >= 345:
        return "red"
    if h < 45:
        return "orange"
    if h < 70:
        return "yellow"
    if h < 160:
        return "green"
    if h < 200:
        return "cyan"
    if h < 255:
        return "blue"
    if h < 290:
        return "purple"
    return "pink"


def _box_intersect(a: tuple, b: tuple) -> float:
    """b 占 a 的面积比例（0–1）；box=(x,y,cx,cy)。"""
    if not a or not b:
        return 0.0
    ax, ay, acx, acy = a
    bx, by, bcx, bcy = b
    ix = max(0, min(ax + acx, bx + bcx) - max(ax, bx))
    iy = max(0, min(ay + acy, by + bcy) - max(ay, by))
    inter = ix * iy
    area_a = acx * acy
    return inter / area_a if area_a > 0 else 0.0


def _shape_fill_hex(sp: ET.Element) -> str | None:
    fill = sp.find(f"{P}spPr/{A}solidFill/{A}srgbClr")
    if fill is not None:
        return fill.get("val")
    fill = sp.find(f"{P}spPr/{A}solidFill/{A}schemeClr")
    if fill is not None:
        return "theme:" + (fill.get("val") or "")
    return None


def _run_colors(sp: ET.Element) -> list[str]:
    colors: list[str] = []
    for rpr in sp.iter(A + "rPr"):
        c = rpr.find(f"{A}solidFill/{A}srgbClr")
        if c is not None and c.get("val"):
            colors.append(c.get("val"))
            continue
        sc = rpr.find(f"{A}solidFill/{A}schemeClr")
        if sc is not None and sc.get("val"):
            colors.append("theme:" + sc.get("val"))
    return colors


# ----------------------------------------------------------------------
# 背景 / 图片 / 烘焙判定
# ----------------------------------------------------------------------

def _background(root: ET.Element, pics: list[dict],
                canvas: tuple[int, int]) -> dict:
    W, H = canvas
    bg = root.find(f"{P}cSld/{P}bg")
    bg_kind = "none"
    bg_color = None
    if bg is not None:
        if bg.find(f".//{A}blip") is not None:
            bg_kind = "blipfill"
        else:
            fill = bg.find(f".//{A}srgbClr")
            if fill is not None:
                bg_kind = "solid"
                bg_color = fill.get("val")
            else:
                scheme = bg.find(f".//{A}schemeClr")
                if scheme is not None:
                    bg_kind = "solid"
                    bg_color = "theme:" + (scheme.get("val") or "")

    fullbleed = False
    for pic in pics:
        box = pic.get("box")
        if not box:
            continue
        _, _, cx, cy = box
        if cx / W >= FULLBLEED_W and cy / H >= FULLBLEED_H:
            fullbleed = True
            break

    if bg_kind == "none" and fullbleed:
        bg_kind = "baked-image"
    return {"kind": bg_kind, "color": bg_color, "fullbleed_pic": fullbleed}


def _image_coverage(pics: list[dict], canvas: tuple[int, int]) -> float:
    W, H = canvas
    area = W * H
    if area <= 0:
        return 0.0
    cov = sum((b[2] * b[3]) for p in pics if (b := p.get("box"))) / area
    return min(1.0, cov)


def _baked_text(coverage: float, text_shapes: list[dict],
                pics: list[dict]) -> dict:
    """烘焙文字判定：图片覆盖率高且原生文本节点极少；文本槽与图片区域重叠加码。"""
    n_text = len(text_shapes)
    overlap = 0.0
    for ts in text_shapes:
        tb = ts.get("box")
        if not tb:
            continue
        for pic in pics:
            pb = pic.get("box")
            if pb and _box_intersect(tb, pb) >= 0.5:
                overlap = max(overlap, _box_intersect(tb, pb))
                break

    if coverage >= 0.75 and n_text <= 1 and (overlap >= 0.5 or
                                            (text_shapes and len(text_shapes[0]["text"]) <= 2)):
        state, conf = "yes", "high"
        reason = "图片覆盖 ≥75%，原生文本节点 ≤1 且与图重叠或仅剩装饰短字，大字大概率已烘焙进位图"
    elif coverage >= 0.60 and n_text <= 2 and overlap >= 0.5:
        state, conf = "yes", "mid"
        reason = "图片覆盖 ≥60%，文本槽压在图片区域 ≥50%，标题 / 文案疑似烘焙"
    elif coverage >= 0.35 and n_text <= 2:
        state, conf = "suspect", "mid"
        reason = "图片覆盖 35%–70% 且文本稀疏，存在局部烘焙文字风险"
    elif coverage >= 0.20 and n_text <= 1:
        state, conf = "suspect", "low"
        reason = "单文本槽压在图片上，需人工核对是否为位图文字"
    else:
        state, conf = "none", "high"
        reason = "原生文本节点充足，未见整页烘焙文字特征"

    if overlap >= 0.5 and state == "none":
        state, conf = "suspect", "low"
        reason = "文本槽与图片区域大面积重叠，需人工核对"
    return {"state": state, "confidence": conf, "coverage": round(coverage, 3),
            "text_nodes": n_text, "text_image_overlap": round(overlap, 3),
            "reason": reason}


# ----------------------------------------------------------------------
# 槽位 / 配色 / 动画
# ----------------------------------------------------------------------

def _estimate_capacity(box: tuple | None, size_pt: float | None) -> int:
    """按包围盒与字号估算可容纳 CJK 字数（保守上界）。"""
    if not box or not size_pt:
        return 0
    _, _, cx, cy = box
    emu_per_char = size_pt * 12700 * 1.05  # 全角字宽约等于字号
    emu_per_line = size_pt * 12700 * 1.35
    chars_per_line = max(1, int(cx / emu_per_char))
    lines = max(1, int(cy / emu_per_line))
    return chars_per_line * lines


def _slot(shape: dict, page_max_pt: float | None) -> dict:
    sizes = shape.get("sizes") or []
    size_pt = round(max(sizes) / 100, 1) if sizes else None
    text = shape.get("text") or ""
    name = shape.get("name") or ""
    box = shape.get("box")
    big_on_page = bool(size_pt and page_max_pt and
                       size_pt >= TITLE_SIZE / 100 and
                       size_pt >= 0.7 * page_max_pt)
    is_title = big_on_page or bool(re.search(r"标题|title|课题|heading", name, re.I))
    capacity = _estimate_capacity(box, size_pt)
    return {
        "name": name,
        "size": size_pt,
        "chars": len(text),
        "capacity": capacity,
        "kind": "title" if is_title else "body",
    }


def _palette(root: ET.Element, shapes: list[dict], bg_color: str | None) -> dict:
    ink_counter: Counter = Counter()
    mut_counter: Counter = Counter()
    fill_counter: Counter = Counter()
    container_fill: tuple[int, str | None] = (0, None)

    for sp in root.iter():
        if sp.tag not in (P + "sp", P + "pic"):
            continue
        colors = _run_colors(sp)
        if colors:
            ink_counter.update(colors[:1])
            for c in colors[1:]:
                mut_counter[c] += 1
        fill = _shape_fill_hex(sp)
        if fill:
            fill_counter[fill] += 1
            box = None
            xfrm = sp.find(f"{P}spPr/{A}xfrm")
            if xfrm is not None:
                off = xfrm.find(f"{A}off")
                ext = xfrm.find(f"{A}ext")
                if off is not None and ext is not None:
                    box = (int(off.get("x", 0)), int(off.get("y", 0)),
                           int(ext.get("cx", 0)), int(ext.get("cy", 0)))
            if box and box[2] * box[3] > container_fill[0]:
                container_fill = (box[2] * box[3], fill)

    ink = ink_counter.most_common(1)[0][0] if ink_counter else None
    mut = mut_counter.most_common(1)[0][0] if mut_counter else None
    card = container_fill[1]
    if card is None and fill_counter:
        card = fill_counter.most_common(1)[0][0]

    accent = None
    for color, _ in fill_counter.most_common():
        if color.startswith("theme:"):
            accent = color
            break
        if color in (bg_color, ink):
            continue
        lum = _luminance(color)
        if 40 < lum < 220:  # 排除近黑 / 近白
            accent = color
            break
    if accent is None:
        accent = next((c for c in fill_counter if c != bg_color and c != ink), None)

    return {"bg": bg_color, "card": card, "ink": ink, "mut": mut,
            "accent": accent}


def _anim(root: ET.Element) -> dict:
    timing = root.find(f"{P}timing")
    classes = {"entr": 0, "emph": 0, "exit": 0, "path": 0}
    total = 0
    loop = False
    if timing is not None:
        for ctn in timing.iter(P + "cTn"):
            preset = ctn.get("presetClass")
            if preset:
                total += 1
                if preset in classes:
                    classes[preset] += 1
            repeat = ctn.get("repeatCount")
            if repeat == "indefinite":
                loop = True
            if ctn.get("dur") == "indefinite" and \
                    ctn.get("nodeType") in ("ambient", "afterEffect"):
                loop = True
        for anim in timing.iter(A + "anim"):
            if anim.get("repeatCount") == "indefinite":
                loop = True

    transition = None
    trans = next((el for el in root.iter()
                  if el.tag.rsplit("}", 1)[-1] == "transition"), None)
    if trans is not None:
        if any(c.tag.rsplit("}", 1)[-1] == "morph" for c in trans.iter()):
            transition = "morph"
        else:
            prst = trans.find(".//{*}prstTrans")
            if prst is not None and prst.get("prst"):
                transition = prst.get("prst")
            else:
                for child in trans:
                    tag = child.tag.rsplit("}", 1)[-1]
                    if tag:
                        transition = tag
                        break
            if transition is None:
                transition = "custom"
    return {"total": total, "entr": classes["entr"], "emph": classes["emph"],
            "exit": classes["exit"], "path": classes["path"], "loop": loop,
            "transition": transition}


# ----------------------------------------------------------------------
# 角色兜底（router 低置信时的确定性人工归类）
# ----------------------------------------------------------------------

def _fallback_role(feat: dict) -> str:
    if feat["has_table"]:
        return "table_list"
    if feat["section_marks"]:
        return "section_transition"
    if feat["max_font_size"] >= 3000 and feat["body_shapes"] <= 6:
        return "cover"
    if feat["pic_count"] >= 3:
        return "gallery"
    if feat["two_column"] and feat["compare_words"]:
        return "two_column_compare"
    if feat["step_count"] >= 3 or feat["arrow_count"] >= 2:
        return "process"
    if feat["card_count"] >= 3:
        return "card_grid"
    return "card_grid"


# ----------------------------------------------------------------------
# 包发现与包级依赖判定
# ----------------------------------------------------------------------

def _discover_packs(root: str) -> list[dict]:
    """递归扫全根目录下所有 .pptx（含 reference/ 母版包与 samples/ 实填样例）。

    kind: ``master`` 为版式母版 / 成品页包；``filled_sample`` 为 samples/ 下已换文案的
    成品样例，用于核对槽位容量与替换可行性的真实边界。
    """
    packs: list[dict] = []
    for dirpath, _, filenames in os.walk(root):
        rel_dir = _rel(dirpath, root)
        for fname in sorted(filenames):
            if not fname.lower().endswith(".pptx"):
                continue
            deck = fname if rel_dir == "." else rel_dir + "/" + fname
            rel_norm = "/" + rel_dir
            kind = ("filled_sample"
                    if "/samples" in rel_norm or rel_dir.startswith("samples")
                    else "master")
            packs.append({
                "rel": rel_dir,
                "deck": deck,
                "kind": kind,
                "usebgfill": _pack_usebgfill(dirpath),
            })
    packs.sort(key=lambda p: p["deck"])
    return packs


def _pack_usebgfill(pack_dir: str) -> bool:
    """包级强依赖：spec.json / recipe.md 声明 useBgFill（幻灯片背景填充透出）。"""
    for fname in ("spec.json", "recipe.md"):
        fpath = os.path.join(pack_dir, fname)
        if not os.path.isfile(fpath):
            continue
        try:
            text = open(fpath, encoding="utf-8").read()
        except OSError:
            continue
        if "useBgFill" in text:
            return True
    return False


def _preview_map(pack_dir: str, variant: str | None = None
                 ) -> tuple[dict[int, str], list[str]]:
    """预览静帧：按文件名页号映射到页码；无页号的总览图归入包级概览。

    019 双版本预览带 `抠图版-` / `免抠图版-` 前缀；按 ``variant`` 过滤，
    避免两版同页号互相覆盖。
    """
    preview_dir = os.path.join(pack_dir, "preview")
    per_page: dict[int, str] = {}
    overview: list[str] = []
    if not os.path.isdir(preview_dir):
        return per_page, overview
    prefix = {"cutout": "抠图版", "plain": "免抠图版"}.get(variant, "")
    for fname in sorted(os.listdir(preview_dir)):
        if not fname.lower().endswith(".png"):
            continue
        if prefix and not fname.startswith(prefix + "-"):
            continue
        m = PAGE_NUM_RE.search(fname)
        rel = "preview/" + fname
        if m:
            num = next((g for g in m.groups() if g), None)
            if num:
                per_page[int(num)] = rel
                continue
        overview.append(rel)
    return per_page, overview


def _load_spec(pack_dir: str) -> dict:
    """读取包内 spec.json；先查 reference 同级包根，再查上级目录；缺失返回空 dict。"""
    for candidate in (pack_dir, os.path.dirname(pack_dir)):
        spec_path = os.path.join(candidate, "spec.json")
        if not os.path.isfile(spec_path):
            continue
        try:
            with open(spec_path, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}
    return {}


def _resolve_variant(spec: dict, deck_rel: str, fname: str) -> str | None:
    """按 spec.variants 的 reference 路径匹配变体；回落文件名关键词。"""
    variants = spec.get("variants") or {}
    for vkey, vmeta in variants.items():
        ref = (vmeta or {}).get("reference")
        if ref and (deck_rel.endswith(ref) or fname.endswith(os.path.basename(ref))):
            return vkey
    if "抠图" in fname:
        return "cutout"
    if "免抠" in fname:
        return "plain"
    return None


def _materials(spec: dict, variant: str | None) -> dict:
    """从 spec.materials 抽取素材门槛；spec 缺失时留空，不编造。"""
    m = spec.get("materials") or {}
    out: dict = {
        "min": m.get("min"),
        "max": m.get("max"),
        "aspect": m.get("aspects") or [],
        "orientation": m.get("orientation"),
        "note": m.get("requirement") or "",
    }
    if variant and isinstance(m.get("variants"), dict):
        vk = "抠图版" if variant == "cutout" else "免抠图版"
        vm = (m["variants"] or {}).get(vk) or {}
        if vm:
            out["min"] = vm.get("min", out["min"])
            out["max"] = vm.get("max", out["max"])
            out["note"] = vm.get("requirement") or out["note"]
    return out


def _policy_for(pack_rel: str) -> tuple[str, dict] | None:
    """按包目录首段命中 v1.3.0 人工覆盖策略。"""
    head = pack_rel.split("/", 1)[0]
    if head in PACK_POLICY:
        return head, PACK_POLICY[head]
    return None


# ----------------------------------------------------------------------
# 单页扫描
# ----------------------------------------------------------------------

def scan_deck(pack: dict, root: str, seq: int) -> list[dict]:
    deck_path = os.path.join(root, pack["deck"])
    records: list[dict] = []
    try:
        plans = {int(p["page"]): p for p in router.plan_pptx(deck_path, None)}
    except router.RouterError:
        plans = {}

    pack_dir = os.path.join(root, pack["rel"])
    spec = _load_spec(pack_dir)
    fname = os.path.basename(pack["deck"])
    variant = _resolve_variant(spec, pack["deck"], fname)
    materials = _materials(spec, variant)
    policy_hit = _policy_for(pack["rel"])
    policy = policy_hit[1] if policy_hit else None

    preview_map, _ = _preview_map(pack_dir, variant)
    if not preview_map:
        parent_dir = os.path.dirname(pack_dir)
        if parent_dir != root:
            preview_map, _ = _preview_map(parent_dir, variant)

    with zipfile.ZipFile(deck_path) as zf:
        parts = router._slide_order(zf)
        canvas = _sld_size(zf)
        for idx, part in enumerate(parts, start=1):
            xml = zf.read(part)
            root_el = ET.fromstring(xml)
            shapes = router._collect(root_el)
            pics = [s for s in shapes if s["is_pic"]]
            text_shapes = [s for s in shapes if s["text"]]

            bg = _background(root_el, pics, canvas)
            coverage = _image_coverage(pics, canvas)
            page_max_pt = None
            if text_shapes:
                all_sizes = [sz for s in text_shapes for sz in (s.get("sizes") or [])]
                if all_sizes:
                    page_max_pt = round(max(all_sizes) / 100, 1)
            baked = _baked_text(coverage, text_shapes, pics)
            palette = _palette(root_el, shapes, bg["color"])
            anim = _anim(root_el)
            slots = [_slot(s, page_max_pt) for s in text_shapes]
            capacity_total = sum(s["chars"] for s in slots)

            feat = router.extract_features(part, xml)
            plan = plans.get(idx, {})
            router_role = plan.get("page_type")
            confidence = plan.get("confidence")

            notes: list[str] = []
            if confidence == "high":
                role = router_role or _fallback_role(feat)
                role_source = "router"
            else:
                role = _fallback_role(feat)
                role_source = "人工"
                notes.append("router=low，人工归为 %s（router 原判 %s）"
                             % (role, router_role or "unknown"))

            # v1.3.0 人工覆盖策略：角色 / 硬约束 / 边界 / 素材门槛以 spec 核定为准。
            constraint = "normal"
            constraint_notes: list[str] = []
            role_boundary = ""
            lock_palette = False
            if policy is not None:
                role = policy["roles"]["pages"].get(
                    idx, policy["roles"].get("default") or role)
                role_source = "人工-覆盖"
                constraint = policy["constraint"]
                lock_palette = bool(policy["lock_palette"])
                constraint_notes = list(policy.get("notes") or [])
                if variant and policy.get("notes_variant", {}).get(variant):
                    constraint_notes.append(policy["notes_variant"][variant])
                if policy.get("notes_shared"):
                    constraint_notes = list(policy["notes_shared"]) + constraint_notes
                role_boundary = policy["boundary"].get(role, "")
                if idx in policy.get("boundary_pages", {}):
                    role_boundary = policy["boundary_pages"][idx]

            # 基调：背景亮度优先；背景为图片时按墨色深浅反推。
            if bg["kind"] == "blipfill" or bg["fullbleed_pic"]:
                tone = "dark" if _luminance(palette["ink"]) > 150 else "light"
            else:
                tone = "dark" if _luminance(bg["color"]) < 128 else "light"

            main_family = _hue_bucket(palette["accent"] or bg["color"]
                                      or palette["ink"])

            preview = preview_map.get(idx)
            record = {
                "id": "ref-%03d-p%d" % (seq, idx),
                "pack": pack["rel"],
                "deck": pack["deck"],
                "page": idx,
                "role": role,
                "role_source": role_source,
                "kind": pack["kind"],
                "tone": tone,
                "main_family": main_family,
                "palette_roles": palette,
                "bg_kind": bg["kind"],
                "usebgfill": bool(pack["usebgfill"]),
                "baked_text": baked,
                "locked_images": bool(coverage >= 0.50 or bg["fullbleed_pic"]),
                "slots": slots,
                "capacity_total": capacity_total,
                "anim": anim,
                "preview": preview,
                "variant": variant,
                "constraint": constraint,
                "constraint_notes": constraint_notes,
                "materials": materials,
                "role_boundary": role_boundary,
                "lock_palette": lock_palette,
                "stats": {
                    "shapes": len(shapes),
                    "images": len(pics),
                    "text_nodes": len(text_shapes),
                    "chars": feat["char_count"],
                    "size": {"w": canvas[0], "h": canvas[1]},
                },
                "notes": notes,
            }
            records.append(record)
    return records


# ----------------------------------------------------------------------
# 产物生成
# ----------------------------------------------------------------------

def _usable(rec: dict) -> bool:
    """可整页导入换文案：非烘焙文字、不依赖背景填充透出、至少有一个文本槽。"""
    return (rec["baked_text"]["state"] != "yes"
            and not rec["usebgfill"]
            and rec["capacity_total"] > 0)


def build_index(packs: list[dict], root: str) -> dict:
    pages: list[dict] = []
    pack_meta: list[dict] = []
    for seq, pack in enumerate(packs, start=1):
        records = scan_deck(pack, root, seq)
        pages.extend(records)
        _, overview = _preview_map(os.path.join(root, pack["rel"]))
        variant = records[0]["variant"] if records else None
        pack_meta.append({
            "id": "ref-%03d" % seq,
            "pack": pack["rel"],
            "deck": pack["deck"],
            "kind": pack["kind"],
            "variant": variant,
            "usebgfill": pack["usebgfill"],
            "pages": len(records),
            "overview_previews": overview,
        })
    physical = len(pages)
    # canonical：同一效果只计一版；019 免抠图版（plain）为重复 variant，不计入。
    canonical = sum(1 for p in pages if p["variant"] != "plain")
    return {"schema": "layout-assets/v1.3",
            "generated_by": "scripts/scan_layout_assets.py",
            "reference_root_label": "ext-research/ppt-effects/effects",
            "kind_legend": {"master": "版式母版 / 成品页包",
                             "filled_sample": "samples/ 下已换文案的成品样例"},
            "role_keys": ROLE_KEYS,
            "counts": {"physical": physical,
                       "canonical": canonical,
                       "variant_duplicate": physical - canonical,
                       "caliber_note": "physical 按每个 pptx 文件计页；canonical 同一效果只计一版，"
                                       "019 双版本中 plain 免抠图版为重复 variant，差值即重复页数"},
            "packs": pack_meta,
            "pages": pages}


def _role_summary(pages: list[dict]) -> dict:
    by_role: dict[str, list[dict]] = {r: [] for r in ROLE_KEYS}
    for rec in pages:
        by_role.setdefault(rec["role"], []).append(rec)
    return by_role


def _page_sig(p: dict) -> tuple:
    """包内逐页结构指纹（忽略容量档 / 烘焙状态等实例级噪声，槽位数精确比对）。"""
    return (p["role"], p["bg_kind"], len(p["slots"]), p["tone"])


def _pack_clusters(pages: list[dict]) -> tuple[dict, dict]:
    """按整包指纹聚合同构包；页数相同且逐页粗指纹一致才并为一簇。

    返回 cluster[deck] = {rep, size, members} 与 by_deck。
    """
    by_deck: dict[str, list[dict]] = {}
    for p in pages:
        by_deck.setdefault(p["deck"], []).append(p)
    sig_map: dict[tuple, list[str]] = {}
    for deck, recs in by_deck.items():
        sig = (len(recs),) + tuple(
            _page_sig(r) for r in sorted(recs, key=lambda x: x["page"]))
        sig_map.setdefault(sig, []).append(deck)
    cluster: dict[str, dict] = {}
    for decks in sig_map.values():
        decks = sorted(decks)
        rep = decks[0]
        for d in decks:
            cluster[d] = {"rep": rep, "size": len(decks), "members": decks}
    return cluster, by_deck


def _one_liner(role: str, rec: dict) -> str:
    label = ROLE_LABEL.get(role, role)
    n_slots = len(rec["slots"])
    if n_slots == 0:
        return "%s版式骨架，无原生文本槽，仅作视觉参考" % label
    return "%s：%d 个可替换槽位，按槽位字数控制文案长度" % (label, n_slots)


def render_md(index: dict) -> str:
    pages = index["pages"]
    by_role = _role_summary(pages)
    lines: list[str] = []
    lines.append("# 成品页资产目录（layout-assets）")
    lines.append("")
    lines.append("> 本目录是「成品页资产索引」：资产本体（pptx / 静帧 png / svg）存放在")
    lines.append("> 参考件目录内，**不随技能分发**；本文与配套 JSON 只记录相对路径与结构特征，")
    lines.append("> 供「选页 → 整页导入 → 换文案」流水线按页角色、槽位与动画特征检索。")
    lines.append("")
    lines.append("- 索引由 `scripts/scan_layout_assets.py` 一键重扫重建，不要手工改表。")
    lines.append("- 页角色共 20 类；`router` 表示高置信自动路由，`人工` 表示 router 低置信后按规则兜底归类。")
    lines.append("- 标记「仅可整页复用、不可换文案」的页面：大字已烘焙进位图，或强依赖幻灯片背景填充透出。")
    lines.append("- `samples/` 下的实填样例以 `filled_sample` 标记纳入索引，用于核对槽位容量与替换边界；其余为 `master` 母版包。")
    lines.append("")

    total = len(pages)
    counts = index.get("counts", {})
    canonical = counts.get("canonical", total)
    covered = [r for r in ROLE_KEYS if by_role.get(r)]
    lines.append("## 统计概览")
    lines.append("")
    lines.append("- 总页数（physical，按每个 pptx 文件计）：**%d**" % total)
    lines.append("- 总页数（canonical，同一效果只计一版）：**%d**；差值 %d 页来自 019 双版本的重复 variant（plain 免抠图版）。"
                 % (canonical, total - canonical))
    lines.append("- 覆盖页角色数：**%d / 20**" % len(covered))
    usable = [r for r in pages if _usable(r)]
    lines.append("- 可换文案页数（非烘焙文字、非背景填充依赖、有文本槽）：**%d**" % len(usable))
    lines.append("")

    cluster, by_deck = _pack_clusters(pages)

    def _sort_key(rec: dict) -> tuple:
        return (0 if _usable(rec) else 1, -rec["capacity_total"], rec["id"])

    for role in ROLE_KEYS:
        recs = by_role.get(role) or []
        if not recs:
            continue
        # 同构簇说明：仅在本角色出现多包同构时列出，代表包取指纹簇内首个 deck。
        role_clusters: dict[str, dict] = {}
        for rec in recs:
            cl = cluster[rec["deck"]]
            if cl["size"] > 1:
                role_clusters[cl["rep"]] = cl
        visible = [r for r in recs if cluster[r["deck"]]["rep"] == r["deck"]]
        visible.sort(key=_sort_key)

        lines.append("## %s（%s）· %d 页" % (ROLE_LABEL.get(role, role), role, len(recs)))
        lines.append("")
        for rep, cl in sorted(role_clusters.items()):
            rep_pages = len(by_deck[rep])
            members = "、".join("`%s`" % m.rsplit("/", 1)[-1] for m in cl["members"])
            lines.append("- 同构簇：**%d 包 × %d 页**，代表包 `%s`；成员：%s。"
                         "JSON 保留全部逐页记录，文档仅取代表包行。"
                         % (cl["size"], rep_pages, rep, members))
        if role_clusters:
            lines.append("")
        lines.append("| id | 来源效果 | variant | 槽位数 | 可替换字数 | 烘焙文字 | bgFill | 动画节点 | 基调 | 适用说明 |")
        lines.append("|---|---|---|---|---|---|---|---|---|---|")
        for rec in visible:
            baked = rec["baked_text"]["state"]
            baked_conf = rec["baked_text"]["confidence"]
            baked_cell = "%s（%s）" % (baked, baked_conf)
            bgfill = "是" if rec["usebgfill"] else "否"
            lock_note = ""
            if rec["baked_text"]["state"] == "yes" or rec["usebgfill"]:
                lock_note = "；仅可整页复用、不可换文案"
            source = rec["pack"]
            if rec["kind"] == "filled_sample":
                source += "（filled_sample）"
            cl = cluster[rec["deck"]]
            if cl["size"] > 1:
                source += "（同构×%d）" % cl["size"]
            variant_cell = rec.get("variant") or "—"
            boundary = rec.get("role_boundary") or ""
            explain = _one_liner(role, rec)
            if boundary:
                explain += "；边界：" + boundary
            lines.append("| %s | %s | %s | %d | %d | %s | %s | %d | %s | %s%s |" % (
                rec["id"], source, variant_cell, len(rec["slots"]),
                rec["capacity_total"], baked_cell, bgfill,
                rec["anim"]["total"], rec["tone"], explain, lock_note))
        lines.append("")

    # 缺口角色
    missing = [r for r in ROLE_KEYS if not by_role.get(r)]
    lines.append("## 缺口角色")
    lines.append("")
    if missing:
        for r in missing:
            lines.append("- `%s`（%s）：暂无成品页" % (r, ROLE_LABEL.get(r, r)))
    else:
        lines.append("- 无缺口，20 类均有覆盖。")
    lines.append("")

    # 覆盖率速览（明细见 reports/layout-asset-coverage.md）
    usable_by_role: dict[str, list[dict]] = {}
    for rec in usable:
        usable_by_role.setdefault(rec["role"], []).append(rec)
    ge1 = [r for r in ROLE_KEYS if usable_by_role.get(r)]
    ge2 = [r for r in ROLE_KEYS if len(usable_by_role.get(r, [])) >= 2]
    zero = [r for r in ROLE_KEYS if not usable_by_role.get(r)]
    lines.append("## 覆盖率速览")
    lines.append("")
    lines.append("- 有 ≥1 张可换文案成品页的角色（%d）：%s"
                 % (len(ge1), "、".join("`%s`" % r for r in ge1) or "无"))
    lines.append("- 有 ≥2 张可换文案成品页的角色（%d）：%s"
                 % (len(ge2), "、".join("`%s`" % r for r in ge2) or "无"))
    lines.append("- 0 张可用的角色（%d）：%s"
                 % (len(zero), "、".join("`%s`" % r for r in zero) or "无"))
    lines.append("- 仅可整页复用、不可换文案：**%d** 页；完整阻断清单与 academic 家族覆盖见覆盖率报告。"
                 % (total - len(usable)))
    lines.append("")
    lines.append(SLOT_ADAPT_RULES_MD.rstrip("\n"))
    lines.append("")
    return "\n".join(lines)


def render_report(index: dict) -> str:
    pages = index["pages"]
    by_role = _role_summary(pages)
    usable = [r for r in pages if _usable(r)]
    usable_by_role: dict[str, list[dict]] = {}
    for rec in usable:
        usable_by_role.setdefault(rec["role"], []).append(rec)

    counts = index.get("counts", {})
    physical = len(pages)
    canonical = counts.get("canonical", physical)
    lines: list[str] = []
    lines.append("# 成品版式页覆盖率报告")
    lines.append("")
    lines.append("- 参考根目录（相对标签）：`%s`" % index["reference_root_label"])
    lines.append("- 扫描包数：**%d**（master %d · filled_sample %d）"
                 % (len(index["packs"]),
                    sum(1 for p in index["packs"] if p["kind"] == "master"),
                    sum(1 for p in index["packs"] if p["kind"] == "filled_sample")))
    lines.append("- 总页数（physical，每个 pptx 文件都计）：**%d**；"
                 "总页数（canonical，同一效果只计一版）：**%d**；差值 **%d** 页来自 019 双版本"
                 "（plain 免抠图版为重复 variant，cutout 抠图版计入 canonical）。"
                 % (physical, canonical, physical - canonical))
    lines.append("- 可换文案页数：**%d**；不可直接换文案页数：**%d**"
                 % (len(usable), physical - len(usable)))
    lines.append("")

    ge1 = [r for r in ROLE_KEYS if usable_by_role.get(r)]
    ge2 = [r for r in ROLE_KEYS if len(usable_by_role.get(r, [])) >= 2]
    zero = [r for r in ROLE_KEYS if not usable_by_role.get(r)]

    lines.append("## 一、角色覆盖：本轮 vs 上一轮（v1.2.0 → v1.3.0）")
    lines.append("")
    lines.append("上一轮基线：%d 包 %d 页，可换文案 %d 页（基线取自 v1.2.0 报告）。"
                 % (PREV_PACKS, PREV_TOTAL_PAGES, PREV_USABLE_PAGES))
    lines.append("")
    lines.append("| 角色键 | 角色名 | 上一轮成品页 | 上一轮可用 | 本轮成品页 | 本轮可用 | 变化 |")
    lines.append("|---|---|---|---|---|---|---|")
    for role in ROLE_KEYS:
        prev_all, prev_ok = PREV_ROLE_BASELINE.get(role, (0, 0))
        all_n = len(by_role.get(role, []))
        ok_n = len(usable_by_role.get(role, []))
        phys_delta = all_n - prev_all
        ok_delta = ok_n - prev_ok
        parts = []
        if phys_delta:
            parts.append("成品 %+d" % phys_delta)
        if ok_delta:
            parts.append("可用 %+d" % ok_delta)
        change = "；".join(parts) or "—"
        if prev_ok == 0 and ok_n >= 1:
            change = "新覆盖（可用 +%d）" % ok_n
        elif prev_ok == 0 and all_n >= 1 and ok_n == 0:
            change = "新出成品页 %d，暂不可换文案" % all_n
        lines.append("| `%s` | %s | %d | %d | %d | %d | %s |"
                     % (role, ROLE_LABEL.get(role, role), prev_all, prev_ok,
                        all_n, ok_n, change))
    lines.append("")
    lines.append("- 有 ≥1 张可用成品页的角色（%d）：%s"
                 % (len(ge1), "、".join("`%s`" % r for r in ge1) or "无"))
    lines.append("- 有 ≥2 张可用成品页的角色（%d）：%s"
                 % (len(ge2), "、".join("`%s`" % r for r in ge2) or "无"))
    lines.append("- 0 张可用的角色（%d）：%s"
                 % (len(zero), "、".join("`%s`" % r for r in zero) or "无"))
    lines.append("")

    # 上一轮 13 个缺口角色逐条核对：本轮是否补上、靠哪个包、补了几页。
    prev_zero = [r for r in ROLE_KEYS if PREV_ROLE_BASELINE.get(r, (0, 0))[1] == 0]
    lines.append("## 二、上一轮缺口角色逐条核对")
    lines.append("")
    lines.append("状态分三档：**已补上**＝本轮有可换文案成品页；**已有成品页、暂不可换文案**＝"
                 "页已入索引但被烘焙 / 背景填充等阻断条件拦下；**仍为 0**＝本轮无任何成品页。")
    lines.append("")
    lines.append("| 角色键 | 角色名 | 本轮成品页 | 本轮可用 | 状态 | 来源包 | 页数（physical / canonical） |")
    lines.append("|---|---|---|---|---|---|---|")
    newly_physical: list[str] = []
    newly_usable: list[str] = []
    still_empty: list[str] = []
    for role in prev_zero:
        ok_recs = usable_by_role.get(role, [])
        all_recs = by_role.get(role, [])
        packs_hit = sorted({r["pack"].split("/", 1)[0] for r in all_recs})
        phys = len(all_recs)
        canon = sum(1 for r in all_recs if r["variant"] != "plain")
        src = "、".join("`%s`" % p for p in packs_hit) or "—"
        if ok_recs:
            status = "已补上"
            newly_usable.append(role)
            newly_physical.append(role)
        elif all_recs:
            status = "已有成品页，暂不可换文案"
            newly_physical.append(role)
        else:
            status = "仍为 0，本轮无包补上"
            still_empty.append(role)
        lines.append("| `%s` | %s | %d | %d | %s | %s | %d / %d |"
                     % (role, ROLE_LABEL.get(role, role), phys, len(ok_recs),
                        status, src, phys, canon))
    lines.append("")
    lines.append("- 本轮新出成品页的角色（%d）：%s"
                 % (len(newly_physical),
                    "、".join("`%s`" % r for r in newly_physical) or "无"))
    lines.append("- 其中已可换文案、完成新覆盖的角色（%d）：%s"
                 % (len(newly_usable),
                    "、".join("`%s`" % r for r in newly_usable) or "无"))
    lines.append("- 仍无成品页的角色（%d）：%s"
                 % (len(still_empty),
                    "、".join("`%s`" % r for r in still_empty) or "无"))
    lines.append("- 按可换文案口径仍为 0 的角色（%d）：%s"
                 % (len(zero), "、".join("`%s`" % r for r in zero) or "无"))
    lines.append("")

    acad = [r for r in pages if r["pack"].startswith("academic/")]
    acad_usable = [r for r in acad if _usable(r)]
    acad_by_role: dict[str, list[dict]] = {}
    for rec in acad:
        acad_by_role.setdefault(rec["role"], []).append(rec)
    acad_roles = sorted(acad_by_role.keys(), key=lambda r: ROLE_KEYS.index(r))
    lines.append("## 三、academic（学术蓝）家族角色覆盖")
    lines.append("")
    lines.append("academic/ 子树是最可能直接复用的母版家族（含 6 个导航栏模板、")
    lines.append("1 个纵横导航母版与 1 个实填样例）。下表只统计该家族内部的角色分布。")
    lines.append("")
    if acad:
        lines.append("| 角色键 | 角色名 | 家族页数 | 其中可换文案 |")
        lines.append("|---|---|---|---|")
        for role in acad_roles:
            n_all = len(acad_by_role[role])
            n_ok = sum(1 for r in acad_by_role[role] if _usable(r))
            lines.append("| `%s` | %s | %d | %d |"
                         % (role, ROLE_LABEL.get(role, role), n_all, n_ok))
        lines.append("")
        lines.append("- 家族覆盖角色（%d）：%s"
                     % (len(acad_roles),
                        "、".join("`%s`" % r for r in acad_roles) or "无"))
        acad_gap = [r for r in ROLE_KEYS if r not in acad_by_role]
        lines.append("- 家族未覆盖角色（%d）：%s"
                     % (len(acad_gap),
                        "、".join("`%s`" % r for r in acad_gap) or "无"))
    else:
        lines.append("- 无 academic 家族页面。")
    lines.append("")

    blocked = [r for r in pages if not _usable(r)]
    lines.append("## 四、仅可整页复用、不可换文案的页面")
    lines.append("")
    lines.append("以下页面满足任一阻断条件：大字已烘焙进位图（baked_text=yes），")
    lines.append("或强依赖 `useBgFill` 幻灯片背景填充透出。整页视觉可复用，换文案前需重建文本层。")
    lines.append("")
    if blocked:
        lines.append("| id | 来源效果 | 阻断原因 | 依据 |")
        lines.append("|---|---|---|---|")
        for rec in sorted(blocked, key=lambda r: r["id"]):
            reasons = []
            if rec["baked_text"]["state"] == "yes":
                reasons.append("烘焙文字=yes（%s）" % rec["baked_text"]["confidence"])
            if rec["usebgfill"]:
                reasons.append("useBgFill 强依赖")
            if rec["capacity_total"] == 0 and not reasons:
                reasons.append("无原生文本槽")
            if rec["baked_text"]["state"] == "yes":
                basis = rec["baked_text"]["reason"]
            elif rec["capacity_total"] == 0:
                basis = "无原生可替换文本槽（动画过渡帧 / 纯装饰页）"
            else:
                basis = "依赖幻灯片背景填充透出，换文案前需重建文本层"
            lines.append("| %s | %s | %s | %s |" % (
                rec["id"], rec["pack"], "；".join(reasons), basis))
    else:
        lines.append("- 无阻断页面。")
    lines.append("")

    # 优先复用：按槽位数与容量排序，只在可用页里选。
    ranked = sorted(usable, key=lambda r: (r["capacity_total"], len(r["slots"])),
                    reverse=True)
    top = ranked[:5]
    lines.append("## 五、优先复用建议（Top 5，按槽位容量与槽位数排序）")
    lines.append("")
    if top:
        lines.append("| 排名 | id | 角色 | 槽位数 | 可替换字数 | 基调 | 理由 |")
        lines.append("|---|---|---|---|---|---|---|")
        for i, rec in enumerate(top, start=1):
            reason = "%d 个槽位共 %d 字容量，动画节点 %d，%s 底；槽位结构可直接套文案" % (
                len(rec["slots"]), rec["capacity_total"], rec["anim"]["total"],
                rec["tone"])
            lines.append("| %d | %s | %s | %d | %d | %s | %s |" % (
                i, rec["id"], rec["role"], len(rec["slots"]),
                rec["capacity_total"], rec["tone"], reason))
    else:
        lines.append("- 暂无可换文案页面。")
    lines.append("")
    return "\n".join(lines)


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------

def _default_root() -> str:
    skill_root = os.path.dirname(SCRIPT_DIR)
    return os.path.normpath(os.path.join(
        skill_root, "..", "ext-research", "ppt-effects", "effects"))


def _materials_line(m: dict) -> str:
    if not m or m.get("min") is None and not m.get("note"):
        return "无图片材料门槛（spec 未声明）"
    parts = []
    if m.get("min") is not None or m.get("max") is not None:
        parts.append("数量 %s–%s" % (m.get("min"), m.get("max")))
    if m.get("aspect"):
        parts.append("比例 %s" % "/".join(str(a) for a in m["aspect"]))
    if m.get("orientation"):
        parts.append("朝向：%s" % m["orientation"])
    if m.get("note"):
        parts.append(m["note"])
    return "；".join(parts)


def render_new_packs(index: dict) -> str:
    """v1.3.0 新增效果包清单：一包一节，角色 / 材料门槛 / 硬约束 / 边界。"""
    pages = index["pages"]
    by_head: dict[str, list[dict]] = {}
    for rec in pages:
        head = rec["pack"].split("/", 1)[0]
        if head in PACK_POLICY:
            by_head.setdefault(head, []).append(rec)

    lines: list[str] = []
    lines.append("# v1.3.0 新增效果包清单（018–022）")
    lines.append("")
    phys_new = sum(len(v) for v in by_head.values())
    canon_new = sum(1 for v in by_head.values() for r in v
                    if r["variant"] != "plain")
    lines.append("- 本轮新增包：**%d** 个；页数 physical **%d** / canonical **%d**"
                 % (len(by_head), phys_new, canon_new))
    lines.append("- 口径：physical 按每个 pptx 文件计；canonical 同一效果只计一版，"
                 "019 的 plain 免抠图版为重复 variant。")
    lines.append("- 角色 / 约束以各包 `spec.json` 为准人工核定，记录 `role_source=人工-覆盖`。")
    lines.append("")

    for head in sorted(by_head):
        recs = sorted(by_head[head], key=lambda r: (r["deck"], r["page"]))
        policy = PACK_POLICY[head]
        decks = sorted({r["deck"] for r in recs})
        variants = sorted({r["variant"] for r in recs if r["variant"]})
        role_count: dict[str, int] = {}
        for r in recs:
            role_count[r["role"]] = role_count.get(r["role"], 0) + 1
        canon_n = sum(1 for r in recs if r["variant"] != "plain")
        materials = recs[0]["materials"]

        lines.append("## %s" % head)
        lines.append("")
        lines.append("- 参考件：%s" % "；".join("`%s`" % d for d in decks))
        lines.append("- 页数：physical **%d** / canonical **%d**；variant：%s"
                     % (len(recs), canon_n,
                        "、".join(variants) if variants else "null（单版本）"))
        role_txt = "；".join("%s %d 页" % (ROLE_LABEL.get(k, k), n)
                            for k, n in sorted(role_count.items()))
        lines.append("- 角色分布：%s" % role_txt)
        mat_by_variant: dict = {}
        for r in recs:
            mat_by_variant.setdefault(r["variant"], r["materials"])
        if len(mat_by_variant) > 1:
            lines.append("- 材料门槛（按 variant 分列）：")
            for vkey, vmat in sorted(mat_by_variant.items(), key=lambda kv: str(kv[0])):
                lines.append("  - %s：%s" % (vkey, _materials_line(vmat)))
        else:
            lines.append("- 材料门槛：%s" % _materials_line(materials))
        lines.append("- 硬约束档位：**%s**" % policy["constraint"])
        for note in recs[0]["constraint_notes"]:
            lines.append("  - %s" % note)
        boundaries = sorted({r["role_boundary"] for r in recs if r["role_boundary"]})
        for b in boundaries:
            lines.append("- 边界：%s" % b)
        lines.append("")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="成品版式页资产扫描器")
    ap.add_argument("--root", help="参考件根目录（默认按脚本位置推断）")
    ap.add_argument("--out", help="JSON 索引输出路径")
    ap.add_argument("--md", help="人读 Markdown 索引输出路径")
    ap.add_argument("--report", help="覆盖率报告输出路径")
    ap.add_argument("--new-packs", help="新增效果包清单输出路径")
    ap.add_argument("--skill-root", help="技能根目录（用于推断默认产物位置）")
    args = ap.parse_args(argv)

    skill_root = args.skill_root or os.path.dirname(SCRIPT_DIR)
    root = args.root or _default_root()
    out_json = args.out or os.path.join(skill_root, "references", "layout-assets.json")
    out_md = args.md or os.path.join(skill_root, "references", "layout-assets.md")
    out_report = args.report or os.path.join(skill_root, "reports",
                                             "layout-asset-coverage.md")
    out_new = args.new_packs or os.path.join(skill_root, "reports",
                                            "new-packs-v1.3.0.md")

    if not os.path.isdir(root):
        print("错误：参考件根目录不存在：%s" % root, file=sys.stderr)
        return 2

    packs = _discover_packs(root)
    if not packs:
        print("错误：参考件根目录下未发现 reference/*.pptx 包", file=sys.stderr)
        return 2

    index = build_index(packs, root)

    for out in (out_json, out_md, out_report, out_new):
        os.makedirs(os.path.dirname(out), exist_ok=True)

    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    with open(out_md, "w", encoding="utf-8") as f:
        f.write(render_md(index))
    with open(out_report, "w", encoding="utf-8") as f:
        f.write(render_report(index))
    with open(out_new, "w", encoding="utf-8") as f:
        f.write(render_new_packs(index))

    pages = index["pages"]
    usable = [r for r in pages if _usable(r)]
    print("扫描完成：%d 包，%d 页（canonical %d）；可换文案 %d 页"
          % (len(packs), len(pages), index["counts"]["canonical"], len(usable)))
    print("JSON  : %s" % out_json)
    print("MD    : %s" % out_md)
    print("报告  : %s" % out_report)
    print("新包  : %s" % out_new)
    return 0


if __name__ == "__main__":
    sys.exit(main())
