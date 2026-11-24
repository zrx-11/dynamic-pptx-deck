# -*- coding: utf-8 -*-
"""
effect_router.py — 内容识别与效果路由（仅标准库：zipfile + xml.etree + json）。

输入任一 .pptx（逐页解析形状 / 文本 / 数字 / 层级 / 时间词 / 对齐模式），或一份
JSON 页描述；输出 EffectPlan：页型、信息特征、批次拆法、预设族、常驻环境动效、
强调对象、转场建议与路线建议（原生重建 / 整页导入 / 阻断），每个推荐项带命中的
判定条件编号（basis），识别不确定时给 confidence=low 与保守档，不猜。

CLI：
    python3 effect_router.py <deck.pptx> [--page N] [--tier fidelity_first|balanced|compat]
        [--engine effect_first|compat] [--json OUT] [--md OUT]
    python3 effect_router.py <page-plan.json> [--tier fidelity_first|balanced|compat]
        [--engine effect_first|compat] [--json OUT] [--md OUT]

交付档位：fidelity_first（默认，成品优先；装饰层可预渲染烘焙，内容层保原生文本）；
balanced（效果与可编辑兼顾）；compat（兼容优先，只给标准 DrawingML 安全族）。
`--engine` 为既有引擎族覆盖项；未显式给定时按档位映射。balanced / compat 输出与
旧版逐字段一致，新增的 `decor` 字段与烘焙候选只出现在 fidelity_first。
退出码：0=正常出计划；2=输入不可读。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

P = "{%s}" % P_NS
A = "{%s}" % A_NS

CONTAINER_PRST = {"rect", "roundRect", "roundRect2", "snip1Rect", "snip2Rect",
                  "snip2SameRect", "chamferedRect"}
# 卡片几何：圆角族；普通 rect 是标签 / 行底 / 文本框，不计入卡片计数
CARD_PRST = {"roundRect", "roundRect2", "snip1Rect", "snip2Rect",
             "snip2SameRect", "chamferedRect"}

TIME_WORDS = ["学期", "周次", "阶段", "步骤", "季度"]
HEADER_WORDS = ["课程", "学分", "序号", "问题", "方案", "科目", "周次", "模块", "对比", "类别"]
QUOTE_MARKS = ["“", "”", "「", "」", "‘", "’"]
# 收束结论词（结尾页）与感谢 / 联系词（致谢页）分开：结尾页讲结论，致谢页讲感谢与联络。
CLOSING_WORDS = ["总结", "三年之后", "结尾", "展望"]
THANKS_WORDS = ["谢谢", "感谢", "致谢", "THANKS", "Thanks", "敬请指正",
                "联系方式", "邮箱", "电话", "微信", "QQ"]
SECTION_MARKS = ["PART", "第", "章", "章节", "01", "02", "03", "04"]
COMPARE_WORDS = ["对照", "优劣", "正反", "vs", "VS", "两组", "对比"]
AGENDA_WORDS = ["目录", "大纲", "议程", "CONTENTS", "Contents", "概览"]
RISK_WORDS = ["风险", "隐患", "对策", "警示", "矛盾", "不足", "挑战", "风险提示"]
CASE_WORDS = ["案例", "实例", "例如", "背景介绍", "做法", "成效"]
CONCL_WORDS = ["结论", "综上", "小结", "要点", "建议清单", "清单"]
ORG_WORDS = ["部门", "岗位", "职责", "主管", "经理", "团队", "层级", "主任", "架构"]
TREND_WORDS = ["趋势", "增长", "下降", "同比", "环比", "走势", "折线", "柱状"]
# 对比表两侧列头词：与双栏对比（两个自由容器）区分，对比表是成行成对的列结构。
COMPARE_SIDE_WORDS = ["优势", "劣势", "问题", "对策", "原方案", "核对动作", "方案 A", "方案 B"]

NUMBER_RE = re.compile(r"\d+(?:\.\d+)?%?")
STEP_RE = re.compile(r"^(?:0?[1-9]|[一二三四五六])\s*[、.．]")
ARROW_RE = re.compile(r"→|➜|->|▶")
# 纯编号圆牌（01–12 / 一–六）：是卡片序号徽章，不是被强调的统计数字，不计入数字块。
BADGE_RE = re.compile(r"^(?:0?[1-9]|1[0-2]|[一二三四五六])$")
# KPI 与指标带的字号分界：主角数字 run ≥ 28pt 且数字块 ≤3 判 KPI；
# 多个 16–27pt 中小数字横排（≥4 块）判指标带。
HERO_SIZE = 2800
MEDIUM_SIZE = 1600
KPI_MAX_BLOCKS = 3
BAND_MIN_BLOCKS = 4

# 规模上限（保守档）：超过即提示拆分或阻断核对
MAX_CHARS_PER_SHAPE = 60
MAX_SHAPES_PER_PAGE = 40
# 单页节点预算：效果节点 ≤12，cTn ≤60（沿用技能既有配额）
MAX_EFFECTS = 12
MAX_CTN = 60
CTN_PER_EFFECT = 3.8
# 各批量角色每点击批的元素配额
PER_BATCH = {"cards": 3, "rows": 4, "nodes": 3, "kpi_value": 3, "steps": 3,
             "summary": 4, "band_value": 4, "chart_series": 3, "cmp_rows": 4,
             "org_node": 3, "risk_row": 3, "concl_item": 4, "gallery_pic": 3}
# 每个元素占用的效果节点数（badge 随卡片 zoom，不单独占节点）
EFFECTS_PER = {
    "ambient": 1, "title": 1, "subtitle": 1, "hero": 1, "kicker": 1,
    "chapter_title": 1, "kpi_value": 1, "label": 1, "action": 1, "header": 1,
    "rows": 1, "axis": 1, "nodes": 1, "connectors": 1, "text": 1, "steps": 1,
    "left": 1, "right": 1, "cards": 1, "card_text": 1, "quote_line": 1,
    "attribution": 1, "summary": 1, "closing": 1, "body": 1,
    "agenda_row": 1, "agenda_badge": 1, "band_value": 1, "band_label": 1,
    "chart_axis": 1, "chart_series": 1, "cmp_rows": 1, "org_root": 1,
    "org_node": 1, "org_link": 1, "risk_tag": 1, "risk_row": 1,
    "case_card": 1, "case_body": 1, "concl_item": 1, "thanks_line": 1,
    "gallery_pic": 1, "gallery_label": 1,
}

# 引擎档位：effect_first 为默认（365 / 新版 WPS）；compat 仅在明确要求兼容
# 2019 / 旧 WPS 时启用。高版本预设在弱引擎上降级但不损坏文件。
ENGINE_FIRST = "effect_first"
ENGINE_COMPAT = "compat"
ENGINES = (ENGINE_FIRST, ENGINE_COMPAT)

# 交付档位（交付取向维度）：fidelity_first 成品优先（默认）；balanced 效果与可编辑
# 兼顾；compat 兼容优先。引擎族由档位映射，--engine 仅作既有覆盖项。
TIER_FIDELITY = "fidelity_first"
TIER_BALANCED = "balanced"
TIER_COMPAT = "compat"
TIERS = (TIER_FIDELITY, TIER_BALANCED, TIER_COMPAT)
TIER_ENGINE = {
    TIER_FIDELITY: ENGINE_FIRST,
    TIER_BALANCED: ENGINE_FIRST,
    TIER_COMPAT: ENGINE_COMPAT,
}

# fidelity_first 下装饰层预烘焙配方映射：策略 / 观感强度 / 配方文档节号
# 第 3 节为图层类型与参数表；策略字段标明具体图层类型。
DECOR_RECIPES = {
    "cover": ("baked_gradient", "high", "decor-bake.md#3"),
    "section_transition": ("baked_glow", "high", "decor-bake.md#3"),
    "closing": ("baked_gradient", "high", "decor-bake.md#3"),
    "thanks": ("baked_gradient", "high", "decor-bake.md#3"),
    "kpi": ("baked_glow", "medium", "decor-bake.md#3"),
    "metric_band": ("baked_glow", "medium", "decor-bake.md#3"),
    "trend_chart": ("baked_noise", "medium", "decor-bake.md#3"),
    "quote": ("baked_vignette", "medium", "decor-bake.md#3"),
    "gallery": ("baked_noise", "medium", "decor-bake.md#3"),
}
DECOR_DEFAULT = ("baked_noise", "medium", "decor-bake.md#3")

# 高版本预设族：compat 档不下发；effect_first 档下发并附 2019 降级提示。
HIGH_TIER = {"zoom_rot", "bounce_up", "drop_in", "grow_once"}
DEGRADE_NOTE = {
    "zoom_rot": "2019 映射不稳时退回出现 / zoom",
    "bounce_up": "旧版本差异，可能回退直接出现",
    "drop_in": "旧版本差异，可能回退直接出现",
    "grow_once": "旧版 WPS 识别不稳，退回出现",
}
SAFE_NOTE = "标准 DrawingML 预设，2019 / 365 / WPS 均可解析，无降级"

# 各角色在两档下的预设族；未列出的角色两档共用默认族。
FAMILIES = {
    "kpi_value": {ENGINE_FIRST: ["zoom", "zoom_rot"], ENGINE_COMPAT: ["zoom"]},
    "cards": {ENGINE_FIRST: ["bounce_up", "fly_up", "wipe_l"],
              ENGINE_COMPAT: ["fly_up", "wipe_l"]},
    "closing": {ENGINE_FIRST: ["grow_once", "fade"], ENGINE_COMPAT: ["fade"]},
}

# compat 档替换掉的 p14 / 现代转场备选（主转场保持 ECMA 基础组）。
COMPAT_ALT_TRANSITION = {"process": "pull", "quote": "fade",
                          "gallery": "cover_l"}


def _load_asset_index() -> dict:
    """读取双库索引（references/asset-index.json）；缺失 / 不可读时返回空表，
    路由器退化为内置保守候选，不报错。"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "..", "references", "asset-index.json")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


ASSET_INDEX = _load_asset_index()


def _families(role: str, engine: str, default: list[str]) -> list[str]:
    return FAMILIES.get(role, {}).get(engine, default)


def _degrade(presets: list[str]) -> str:
    notes = [DEGRADE_NOTE[p] for p in presets if p in DEGRADE_NOTE]
    return "；".join(notes) if notes else SAFE_NOTE


class RouterError(RuntimeError):
    """输入不可读或结构无法解析。"""


# ----------------------------------------------------------------------
# 特征提取
# ----------------------------------------------------------------------
def _box(node: ET.Element):
    xfrm = node.find(f"{P}spPr/{A}xfrm")
    if xfrm is None:
        xfrm = node.find(f"{P}grpSpPr/{A}xfrm")
    off = xfrm.find(f"{A}off") if xfrm is not None else None
    ext = xfrm.find(f"{A}ext") if xfrm is not None else None
    if off is None or ext is None:
        return None
    try:
        return (int(off.get("x")), int(off.get("y")),
                int(ext.get("cx")), int(ext.get("cy")))
    except (TypeError, ValueError):
        return None


def _number_block(text: str) -> bool:
    """数字块判定：短文本（≤12 字）且以数字 / 比较符号开头；叙述句 / 标题里
    夹带的数字不计入 KPI / 指标带数字块。"""
    t = text.strip()
    if len(t) > 12 or not NUMBER_RE.search(t):
        return False
    return t[0] in "0123456789≥≤±"


def _font_sizes(node: ET.Element) -> list[int]:
    sizes = []
    for rpr in node.iter(A + "rPr"):
        if rpr.get("sz"):
            try:
                sizes.append(int(rpr.get("sz")))
            except ValueError:
                pass
    return sizes


def _collect(root: ET.Element) -> list[dict]:
    sp_tree = root.find(f".//{P}spTree")
    shapes: list[dict] = []

    def walk(node: ET.Element) -> None:
        for child in node:
            if child.tag not in (P + "sp", P + "grpSp", P + "pic", P + "graphicFrame"):
                continue
            nv = child.find(f".//{P}cNvPr")
            sid = nv.get("id") if nv is not None else ""
            name = nv.get("name") if nv is not None else ""
            geom = child.find(f"{P}spPr/{A}prstGeom")
            prst = geom.get("prst") if geom is not None else ""
            box = _box(child)
            text = "" if child.tag == P + "grpSp" else \
                "".join(t.text or "" for t in child.iter(A + "t")).strip()
            sizes = _font_sizes(child)
            trs = child.findall(f".//{A}tr")
            shapes.append({
                "id": sid, "name": name, "prst": prst, "box": box,
                "text": text, "sizes": sizes,
                "is_table": child.tag == P + "graphicFrame" or prst == "tbl",
                "is_pic": child.tag == P + "pic",
                "table_rows": len(trs),
            })
            if child.tag == P + "grpSp":
                walk(child)

    if sp_tree is not None:
        walk(sp_tree)
    return shapes


def _row_clusters(shapes: list[dict], tol: int = 150_000) -> int:
    """按 y 坐标聚行：同一 y 带（容差内）的多段文本归为一行，不按单元格计数。

    通栏标题 / 页脚备注（cx ≥ 9_000_000）不参与行计数；行内左右两栏 y 略有
    偏移时落在同一 y 带内。
    """
    ys = sorted(
        s["box"][1] for s in shapes if s["box"] and s["box"][2] < 9_000_000)
    if not ys:
        return 0
    clusters = 1
    anchor = ys[0]
    for y in ys[1:]:
        if abs(y - anchor) > tol:
            clusters += 1
            anchor = y
    return clusters


def _bg_info(root: ET.Element) -> dict:
    bg = root.find(f"{P}cSld/{P}bg")
    info = {"type": "none", "color": None, "dark": None, "media": False}
    if bg is None:
        return info
    blip = bg.find(f".//{A}blip")
    if blip is not None:
        info["type"] = "image"
        info["media"] = True
        return info
    fill = bg.find(f".//{A}srgbClr")
    if fill is not None:
        info["type"] = "solid"
        info["color"] = fill.get("val")
        info["dark"] = _luminance(fill.get("val")) < 128
    return info


def _luminance(hex_val: str | None) -> int:
    if not hex_val or len(hex_val) != 6:
        return 255
    try:
        r, g, b = (int(hex_val[i:i + 2], 16) for i in (0, 2, 4))
        return int(0.299 * r + 0.587 * g + 0.114 * b)
    except ValueError:
        return 255


def _boilerplate(name: str, text: str) -> bool:
    """页脚页码、课程名页脚、装饰线等模板文本，不参与页型规则匹配。"""
    if name.startswith(("hd-foot", "hd-kicker", "deco-foot", "deco-tick")):
        return True
    if "action" in name or name.startswith(("note", "g-action")):
        return True
    if re.match(r"^\d+\s*/\s*\d+$", text):
        return True
    if "级 ·" in text and ("规划" in text or "报告" in text or "调研" in text):
        return True
    return False


def extract_features(part: str, xml: bytes) -> dict:
    root = ET.fromstring(xml)
    shapes = _collect(root)
    texts = [s for s in shapes if s["text"]]
    body = [s for s in texts if not _boilerplate(s["name"], s["text"])]
    all_text = " ".join(s["text"] for s in body)
    numbers = NUMBER_RE.findall(all_text)
    time_hits = [w for w in TIME_WORDS if w in all_text]
    header_hits = [w for w in HEADER_WORDS if w in all_text]
    quote = any(m in all_text for m in QUOTE_MARKS)
    closing = any(w in all_text for w in CLOSING_WORDS)
    thanks = any(w in all_text for w in THANKS_WORDS)
    section = bool(re.search(r"PART\s*\d|第[一二三四五六七八九十0-9]+章|章节", all_text))
    compare = any(w in all_text for w in COMPARE_WORDS)
    # 语义页型词只在短标签（≤10 字）中命中；长叙述句 / 备注里夹带的同类词不计入。
    short_text = " ".join(s["text"] for s in body if len(s["text"]) <= 10)
    agenda = [w for w in AGENDA_WORDS if w in all_text]
    risk = [w for w in RISK_WORDS if w in short_text]
    case = [w for w in CASE_WORDS if w in short_text]
    concl = [w for w in CONCL_WORDS if w in short_text]
    org = [w for w in ORG_WORDS if w in short_text]
    trend = [w for w in TREND_WORDS if w in short_text]
    side_hits = [w for w in COMPARE_SIDE_WORDS
                 if any(w in s["text"] and len(s["text"]) <= 10 for s in body)]
    steps = sum(1 for s in body if STEP_RE.match(s["text"]))
    arrows = len(ARROW_RE.findall(all_text))

    containers = [
        s for s in body if s["prst"] in CONTAINER_PRST and s["box"]
        and s["box"][2] >= 1_500_000 and s["box"][3] >= 400_000
        and s["box"][2] < 11_000_000
    ]
    sizes = sorted({sz for s in shapes for sz in s["sizes"]})
    max_size = max(sizes) if sizes else 0
    # 数字块拆分：hero ≥28pt（KPI 主角）、medium 16–27pt（指标带中小数字）；
    # 纯编号圆牌（01–12）不计入任何数字块。big_number_shapes 为兼容旧字段，
    # 口径 = hero + ≥20pt 的 medium 块。
    hero_numbers = 0
    medium_numbers = 0
    badge_numbers = 0
    big_numbers = 0
    for s in body:
        if not _number_block(s["text"]) or not s["sizes"]:
            continue
        if BADGE_RE.match(s["text"].strip()):
            badge_numbers += 1
            continue
        own = max(s["sizes"])
        if own >= HERO_SIZE:
            hero_numbers += 1
        elif own >= MEDIUM_SIZE:
            medium_numbers += 1
        if own >= 2000:
            big_numbers += 1

    # 并列双栏：左右两个等高容器
    two_column = False
    if len(containers) >= 2:
        left = [c for c in containers if c["box"][0] < 6_000_000]
        right = [c for c in containers if c["box"][0] >= 6_000_000]
        if left and right:
            two_column = True

    bg = _bg_info(root)
    char_count = len(all_text)
    short_labels = sum(1 for s in body if len(s["text"]) <= 8)
    header_cells = [w for w in HEADER_WORDS
                    if any(w in s["text"] and len(s["text"]) <= 8 for s in body)]

    # 时间轴通栏细长矩形；小圆 / 椭圆节点；编号圆牌
    oval_prst = {"ellipse", "oval", "donut"}
    axis_shapes = [
        s for s in shapes if s["prst"] in CONTAINER_PRST and s["box"]
        and s["box"][2] >= 8_000_000 and s["box"][3] <= 350_000
    ]
    node_shapes = [
        s for s in shapes if s["prst"] in oval_prst and s["box"]
        and s["box"][2] <= 2_000_000 and s["box"][3] <= 2_000_000
    ]
    # 图表：c:chart 元素或 graphicFrame 图表数据；内容图片：p:pic 且面积不过小
    # （卡片侧边阴影等窄长位图不计入图集计数）。
    has_chart = any(el.tag.endswith("}chart") for el in root.iter())
    pic_shapes = [s for s in shapes if s.get("is_pic") and s["box"]
                  and s["box"][3] >= 900_000 and s["box"][2] >= 900_000]
    pic_count = len(pic_shapes)
    table_tr = sum(s.get("table_rows", 0) for s in shapes)
    if table_tr > 0:
        row_count = table_tr
    else:
        row_count = _row_clusters(body)
        if header_cells:
            row_count -= 1  # 表头单独成行，不进 rows 逐条批次
    row_count = max(row_count, 1)
    return {
        "part": part,
        "shape_count": len(shapes),
        "text_shapes": len(texts),
        "body_shapes": len(body),
        "char_count": char_count,
        "number_tokens": numbers,
        "number_density": round(len(numbers) / char_count, 3) if char_count else 0,
        "big_number_shapes": big_numbers,
        "hero_number_shapes": hero_numbers,
        "medium_number_shapes": medium_numbers,
        "badge_number_shapes": badge_numbers,
        "font_sizes": sizes,
        "font_levels": len(sizes),
        "max_font_size": max_size,
        "has_table": any(s["is_table"] for s in shapes),
        "has_chart": has_chart,
        "pic_count": pic_count,
        "time_words": time_hits,
        "header_words": header_hits,
        "header_cells": header_cells,
        "has_quote": quote,
        "closing_words": [w for w in CLOSING_WORDS if w in all_text],
        "thanks_words": [w for w in THANKS_WORDS if w in all_text],
        "section_marks": section,
        "compare_words": [w for w in COMPARE_WORDS if w in all_text],
        "agenda_words": agenda,
        "risk_words": risk,
        "case_words": case,
        "concl_words": concl,
        "org_words": org,
        "trend_words": trend,
        "side_words": side_hits,
        "step_count": steps,
        "arrow_count": arrows,
        "card_count": len(containers),
        "short_labels": short_labels,
        "axis_count": len(axis_shapes),
        "node_count": len(node_shapes),
        "row_count": row_count,
        "two_column": two_column,
        "background": bg,
        "texts": [s["text"] for s in texts],
    }


# ----------------------------------------------------------------------
# 页型识别（规则编号与 effect-routing.md 矩阵一致）
# ----------------------------------------------------------------------
def classify(feat: dict) -> dict:
    rules: list[tuple[str, str, bool]] = []

    def hit(code: str, label: str, ok: bool) -> None:
        if ok:
            rules.append((code, label, True))

    body_n = feat["body_shapes"]
    hero = feat.get("hero_number_shapes", 0)
    medium = feat.get("medium_number_shapes", 0)
    hit("END-1", "收束词命中且无数字",
        bool(feat["closing_words"]) and not feat["number_tokens"])
    hit("THX-1", "感谢 / 联系词命中，无收束结论词与主角数字",
        bool(feat["thanks_words"]) and not feat["closing_words"] and hero == 0)
    hit("COVER-1", "大标题字号且文本块少、无卡片网格、无大字号数字、非章节标记页",
        feat["max_font_size"] >= 3000 and body_n <= 14
        and feat["card_count"] < 5 and feat["big_number_shapes"] < 2
        and not feat["section_marks"])
    hit("SEC-1", "章节标记且正文稀疏",
        feat["section_marks"] and body_n <= 4
        and feat["card_count"] < 3)
    hit("QUO-1", "引号命中、无数字、长句少量块",
        feat["has_quote"] and not feat["number_tokens"]
        and body_n <= 4 and feat["char_count"] >= 30)
    hit("AGD-1", "目录词命中、成行条目 ≥3、无主角数字",
        bool(feat["agenda_words"]) and 3 <= body_n <= 12
        and feat["row_count"] >= 3 and hero == 0 and feat["card_count"] < 7)
    hit("KPI-1", "主角数字 ≤3 且 ≥1；或大字号数字 ≥2 / 密度达标",
        (1 <= hero <= KPI_MAX_BLOCKS) or
        feat["big_number_shapes"] >= 2 or
        (feat["big_number_shapes"] >= 1 and len(feat["number_tokens"]) >= 3
         and feat["number_density"] >= 0.08))
    hit("MBD-1", "无主角数字，中小数字块 ≥4 横排",
        hero == 0 and medium >= BAND_MIN_BLOCKS
        and len(feat["number_tokens"]) >= 4)
    hit("TRC-1", "图表存在，或趋势词命中且带轴 / 节点",
        bool(feat["has_chart"]) or
        (bool(feat["trend_words"]) and
         (feat.get("axis_count", 0) >= 1 or feat.get("node_count", 0) >= 2)))
    hit("RSK-1", "风险词命中且风险—对策成对成行 ≥2",
        bool(feat["risk_words"]) and feat["row_count"] >= 2 and body_n >= 5)
    hit("CMPT-1", "真实表格带对比信号，或双侧列头成对且成行 ≥2（非左右自由双栏）",
        (feat["has_table"] and (feat["compare_words"] or len(feat["side_words"]) >= 2)) or
        (len(feat["side_words"]) >= 2 and feat["row_count"] >= 2
         and feat["short_labels"] >= 4 and not feat["risk_words"]
         and not feat["two_column"]))
    hit("TBL-1", "表头词成格且短标签 ≥4",
        feat["has_table"] or
        (feat["header_cells"] and feat["short_labels"] >= 4 and body_n >= 7))
    hit("ORG-1", "组织词命中且节点形 ≥3",
        bool(feat["org_words"]) and feat.get("node_count", 0) >= 3)
    hit("TIME-1", "时间词命中且有序条目成组",
        bool(feat["time_words"]) and (feat["step_count"] >= 2 or body_n >= 5))
    hit("PROC-1", "步骤编号或箭头成链，条目 ≥3",
        (feat["step_count"] >= 3 or feat["arrow_count"] >= 2) and body_n >= 5)
    hit("CMP-1", "左右双栏且有对比词或双大容器",
        feat["two_column"] and (feat["compare_words"] or feat["card_count"] == 2))
    hit("CASE-1", "案例词命中且单 / 双特征卡",
        bool(feat["case_words"]) and feat["card_count"] <= 2 and body_n >= 3)
    hit("CONC-1", "结论词命中、条目 ≥3、无卡片网格与表格",
        bool(feat["concl_words"]) and feat["card_count"] < 3
        and not feat["has_table"] and not feat["time_words"]
        and (feat["step_count"] >= 2 or feat["short_labels"] >= 3) and body_n >= 3)
    hit("GAL-1", "内容图片 ≥3 成墙，非整页背景媒体",
        feat.get("pic_count", 0) >= 3 and not feat["background"]["media"]
        and body_n >= 2)
    hit("CARD-1", "并列卡片 ≥3 且各带文字",
        feat["card_count"] >= 3 and body_n >= 4)

    priority = ["closing", "thanks", "cover", "section_transition", "quote",
                "agenda", "kpi", "metric_band", "trend_chart", "risk_alert",
                "compare_table", "table_list", "org_chart", "timeline", "process",
                "two_column_compare", "case_card", "conclusion_list", "gallery",
                "card_grid"]
    code_map = {
        "END-1": "closing", "THX-1": "thanks", "COVER-1": "cover",
        "SEC-1": "section_transition", "QUO-1": "quote", "AGD-1": "agenda",
        "KPI-1": "kpi", "MBD-1": "metric_band", "TRC-1": "trend_chart",
        "RSK-1": "risk_alert", "CMPT-1": "compare_table", "TBL-1": "table_list",
        "ORG-1": "org_chart", "TIME-1": "timeline", "PROC-1": "process",
        "CMP-1": "two_column_compare", "CASE-1": "case_card",
        "CONC-1": "conclusion_list", "GAL-1": "gallery", "CARD-1": "card_grid",
    }
    matched = {code_map[c]: c for c, _, _ in rules}
    page_type = next((t for t in priority if t in matched), "card_grid")

    # 强特征：命中即可独立支撑该页型，不依赖密度 / 弱命中凑数
    strong = {
        "closing": bool(feat["closing_words"]),
        "thanks": bool(feat["thanks_words"]) and not feat["closing_words"],
        "cover": feat["max_font_size"] >= 3000 and feat["card_count"] < 3
        and not feat["section_marks"],
        "section_transition": feat["section_marks"] and body_n <= 4,
        "quote": feat["has_quote"] and feat["char_count"] >= 30 and body_n <= 3,
        "agenda": bool(feat["agenda_words"]) and feat["row_count"] >= 3,
        "kpi": hero >= 1 or feat["big_number_shapes"] >= 2,
        "metric_band": hero == 0 and medium >= BAND_MIN_BLOCKS,
        "trend_chart": bool(feat["has_chart"]),
        "risk_alert": bool(feat["risk_words"]) and feat["row_count"] >= 2,
        "compare_table": feat["has_table"] and
        (bool(feat["compare_words"]) or len(feat["side_words"]) >= 2),
        "table_list": feat["has_table"],
        "org_chart": bool(feat["org_words"]) and feat.get("node_count", 0) >= 3,
        "timeline": bool(feat["time_words"]) and
        (feat["step_count"] >= 2 or feat.get("axis_count", 0) >= 1
         or feat.get("node_count", 0) >= 2),
        "process": feat["step_count"] >= 3 or feat["arrow_count"] >= 2,
        "two_column_compare": feat["two_column"] and bool(feat["compare_words"]),
        "case_card": bool(feat["case_words"]) and feat["card_count"] <= 2,
        "conclusion_list": bool(feat["concl_words"]) and feat["card_count"] < 3,
        "gallery": feat.get("pic_count", 0) >= 3 and not feat["background"]["media"],
        "card_grid": feat["card_count"] >= 4 and feat["big_number_shapes"] < 2
        and not feat["header_cells"] and feat.get("node_count", 0) < 2
        and not feat["time_words"],
    }

    # 只有同为强特征的页型才算竞争；弱命中（密度凑数、默认卡片兜底）只进 basis，不压低置信
    competing = [c for c in matched if c != page_type and strong.get(c, False)]
    weak = [c for c in matched if c != page_type and not strong.get(c, False)]
    sparse = body_n < 3
    # 仅当主规则为强特征、无强竞争页型、特征不稀疏时判 high；其余给 low 保守档
    primary_strong = strong.get(page_type, False)
    confidence = "high" if primary_strong and not competing and not sparse else "low"
    basis = ([matched[page_type]] if page_type in matched else ["CARD-1"]) + \
            [matched[c] for c in competing + weak]
    return {"page_type": page_type, "confidence": confidence,
            "basis": basis, "competing": competing, "strong": primary_strong}


# ----------------------------------------------------------------------
# 计划生成
# ----------------------------------------------------------------------
TRANSITION_BY_TYPE = {
    "cover": ("circle", "wedge", "COVER-1"),
    "section_transition": ("wheel4", "wedge", "SEC-1"),
    "kpi": ("split_h", "blinds_h", "KPI-1"),
    "table_list": ("fade", "wipe_u", "TBL-1"),
    "timeline": ("wipe_l", "pull", "TIME-1"),
    "process": ("pull", "conveyor", "PROC-1"),
    "two_column_compare": ("cover_l", "split_h", "CMP-1"),
    "card_grid": ("cover_l", "split_h", "CARD-1"),
    "quote": ("fade", "reveal", "QUO-1"),
    "closing": ("zoom_t", "fade", "END-1"),
    "agenda": ("push_l", "comb", "AGD-1"),
    "metric_band": ("split_h", "blinds_h", "MBD-1"),
    "trend_chart": ("wipe_u", "checker", "TRC-1"),
    "compare_table": ("fade", "wipe_u", "CMPT-1"),
    "org_chart": ("wheel4", "wedge", "ORG-1"),
    "risk_alert": ("cover_l", "split_h", "RSK-1"),
    "case_card": ("cover_l", "split_h", "CASE-1"),
    "conclusion_list": ("fade", "wipe_u", "CONC-1"),
    "thanks": ("zoom_t", "fade", "THX-1"),
    "gallery": ("gallery", "cover_l", "GAL-1"),
}

TYPE_LABEL = {
    "cover": "封面", "section_transition": "章节过渡", "kpi": "KPI 数字",
    "table_list": "表格清单", "timeline": "时间线 / 学期推进", "process": "流程步骤",
    "two_column_compare": "双栏对比", "card_grid": "并列卡片", "quote": "引语页",
    "closing": "结尾",
    "agenda": "目录", "metric_band": "指标带", "trend_chart": "趋势图",
    "compare_table": "对比表", "org_chart": "组织结构", "risk_alert": "风险提示",
    "case_card": "案例卡", "conclusion_list": "结论清单", "thanks": "致谢页",
    "gallery": "图集 / 图墙",
}


def _b(role: str, count: int, presets: list[str], delay: list[int],
       note: str = "") -> dict:
    count = max(count, 1)
    per_batch = PER_BATCH.get(role)
    batch_count = (count + per_batch - 1) // per_batch if per_batch else 1
    entry = {
        "role": role, "count": count, "preset_family": presets,
        "delay_range": delay, "note": note,
        "degrade": _degrade(presets),
        "effects": count * EFFECTS_PER.get(role, 1),
    }
    if per_batch:
        entry["per_batch"] = per_batch
        entry["batch_count"] = batch_count
    return entry


def _batches(page_type: str, feat: dict, engine: str = ENGINE_FIRST) -> list[dict]:
    """按页型角色矩阵生成批次；数量尽量从形状特征推导，推不出时取该页型默认保守档。"""
    cards = max(feat["card_count"], 1)
    values = max(feat["big_number_shapes"], 1)
    nodes = max(feat.get("node_count", 0), 2)
    rows = feat.get("row_count", max(feat["text_shapes"] - 2, 1))
    steps = max(feat["step_count"], feat["arrow_count"], 3)
    ambient = _b("ambient", 2, ["spin", "path"], [0, 0],
                 "环自转 / 光点往返，indefinite 常驻")

    if page_type == "cover":
        return [
            ambient,
            _b("title", 1, ["fade", "wipe_l"], [0, 120]),
            _b("subtitle", 1, ["fade"], [150, 300]),
        ]
    if page_type == "section_transition":
        hero_note = ("高冲击母题；稀疏页可整页导入复用，或按 Morph 跨页配对衔接"
                     if engine == ENGINE_FIRST else
                     "高冲击母题；稀疏页可整页导入复用")
        return [
            _b("hero", 1, ["spin", "path"], [0, 0], hero_note),
            _b("kicker", 1, ["fade"], [120, 250]),
            _b("chapter_title", 1, ["fade", "wipe_l"], [200, 400]),
        ]
    if page_type == "kpi":
        return [
            ambient,
            _b("kpi_value", values, _families("kpi_value", engine, ["zoom"]), [0, 200],
               "数字载体 zoom 入场，随后 grow 回弹一次；不用 fade 平铺数字"),
            _b("label", values, ["fade"], [150, 450], "标签文字晚于数字 150—450ms"),
            _b("action", 1, ["fade"], [300, 500]),
        ]
    if page_type == "table_list":
        return [
            _b("header", 1, ["fade"], [0, 120], "表头一次 fade，不参与逐条滚动"),
            _b("rows", rows, ["wipe_u"], [120, 300],
               "按行入场；每批 4 行，批间错开 120—300ms，不整表同帧"),
        ]
    if page_type == "timeline":
        return [
            ambient,
            _b("axis", 1, ["wipe_l"], [0, 150], "主轴生长"),
            _b("nodes", nodes, ["fly_l", "fly_up"], [150, 300], "节点错开 150—300ms"),
            _b("connectors", nodes - 1, ["path"], [150, 300], "连线跟随节点"),
            _b("text", nodes, ["fade"], [200, 400]),
        ]
    if page_type == "process":
        return [
            ambient,
            _b("steps", steps, ["fly_up", "wipe_l"], [150, 300], "步骤沿箭头成链依次入场"),
            _b("connectors", steps - 1, ["path"], [150, 300]),
        ]
    if page_type == "two_column_compare":
        return [
            _b("left", 1, ["fly_l"], [0, 200], "左栏先入"),
            _b("right", 1, ["fly_r"], [200, 350], "右栏晚入；左右不同向、不同速"),
        ]
    if page_type == "card_grid":
        return [
            ambient,
            _b("cards", cards, _families("cards", engine, ["fly_up", "wipe_l"]), [150, 300],
               "第 i 张卡与其文字同批成对；编号圆牌随卡片 zoom，不另占效果节点；每批 ≤3 张"),
            _b("card_text", cards, ["fade"], [150, 450],
               "卡内文字与同序号卡片同批，晚 150—450ms"),
        ]
    if page_type == "agenda":
        rows = max(feat["row_count"], 3)
        return [
            _b("title", 1, ["fade", "wipe_l"], [0, 120]),
            _b("agenda_badge", rows, ["zoom"], [120, 260], "编号圆环逐条 zoom，与章名同批"),
            _b("agenda_row", rows, ["fade"], [150, 300], "章名 / 副标随圆环同批，晚 60—260ms"),
        ]
    if page_type == "metric_band":
        values = max(feat["medium_number_shapes"], feat["big_number_shapes"], 3)
        return [
            ambient,
            _b("band_value", values, ["zoom", "fade"], [0, 200],
               "中小数字成排 zoom / fade；不挂 grow 回弹，主角感让位于标题"),
            _b("band_label", values, ["fade"], [150, 350], "标签晚于数字 150—350ms"),
        ]
    if page_type == "trend_chart":
        series = max(feat.get("node_count", 0), 2)
        return [
            _b("chart_axis", 1, ["wipe_l"], [0, 150], "坐标轴 / 基线生长"),
            _b("chart_series", series, ["fade", "wipe_u"], [150, 300],
               "系列逐组入场；当前系列 grow 一次"),
            _b("label", series, ["fade"], [200, 400], "数值 / 图例文字晚于系列"),
        ]
    if page_type == "compare_table":
        return [
            _b("header", 1, ["fade"], [0, 120], "双侧列头一次 fade"),
            _b("cmp_rows", rows, ["wipe_u"], [120, 300],
               "成对行逐条入场；左右单元格同批，不整表同帧"),
        ]
    if page_type == "org_chart":
        nodes = max(feat.get("node_count", 0), 3)
        return [
            _b("org_root", 1, ["zoom"], [0, 150], "根节点先落位"),
            _b("org_link", nodes - 1, ["path"], [150, 300], "层级连线跟随节点"),
            _b("org_node", nodes - 1, ["fly_up", "fade"], [150, 300],
               "下级节点按层错开入场"),
        ]
    if page_type == "risk_alert":
        risk_rows = max(feat["row_count"], 2)
        return [
            _b("risk_tag", 1, ["grow_once", "fade"], [0, 150],
               "风险标记单次强调；红 / 警示色只落在风险标记上"),
            _b("risk_row", risk_rows, ["fly_up", "fade"], [150, 300],
               "风险项与对策成对同批入场，对策文字 fade 晚 150—300ms"),
        ]
    if page_type == "case_card":
        return [
            _b("case_card", 1, ["fly_up", "cover_l"], [0, 200], "案例卡整体入场"),
            _b("case_body", max(feat["text_shapes"] - 2, 1), ["fade"], [150, 350],
               "案例标题 / 正文随卡同批，文字只挂 fade"),
        ]
    if page_type == "conclusion_list":
        items = max(feat["row_count"], 3)
        return [
            _b("header", 1, ["fade"], [0, 120], "结论标题一次 fade"),
            _b("concl_item", items, ["fade", "wipe_l"], [120, 300],
               "结论条目分批入场，每批 ≤4 条；不挂弹跳类"),
        ]
    if page_type == "thanks":
        return [
            _b("title", 1, ["grow_once", "fade"], [0, 200], "致谢语单次放大淡入"),
            _b("thanks_line", max(feat["text_shapes"] - 1, 1), ["fade"], [200, 400],
               "联络 / 署名行逐条 fade，不起新数据批次"),
        ]
    if page_type == "gallery":
        pics = max(feat.get("pic_count", 0), 3)
        return [
            _b("gallery_pic", pics, ["fly_up", "wipe_l"], [150, 300],
               "图片按网格分批入场，每批 ≤3 张"),
            _b("gallery_label", pics, ["fade"], [200, 400], "图注随同序号图片同批"),
        ]
    if page_type == "quote":
        return [
            _b("quote_line", 1, ["fade"], [0, 200], "逐句 fade，不挂弹跳类"),
            _b("attribution", 1, ["fade"], [150, 300], "署名晚 150—300ms"),
        ]
    if page_type == "closing":
        return [
            _b("summary", max(feat["text_shapes"] - 1, 1), ["fade"], [120, 300],
               "结论分批 fade"),
            _b("closing", 1, _families("closing", engine, ["fade"]), [300, 450],
               "收束语最后入场，不堆强调"),
        ]
    # 保守默认档
    return [
        ambient,
        _b("body", max(feat["text_shapes"] - 2, 1), ["fade"], [150, 400]),
    ]


def _budget(batches: list[dict]) -> dict:
    """按批次估算效果节点与 cTn，超预算给出同角色拆批方案。"""
    effects = sum(b["effects"] for b in batches)
    ctn = round(effects * CTN_PER_EFFECT)
    fits = effects <= MAX_EFFECTS and ctn <= MAX_CTN
    split_plan = [
        {"role": b["role"], "total": b["count"],
         "per_batch": b["per_batch"], "batch_count": b["batch_count"]}
        for b in batches if b.get("batch_count", 1) > 1
    ]
    return {
        "estimated_effects": effects,
        "estimated_ctn": ctn,
        "max_effects": MAX_EFFECTS,
        "max_ctn": MAX_CTN,
        "fits": fits,
        "split_plan": split_plan,
    }


def _route(page_type: str, feat: dict, batches: list[dict] | None = None,
           engine: str = ENGINE_FIRST) -> dict:
    cautions: list[str] = []
    missing: list[str] = []

    if feat["background"]["media"]:
        return {
            "decision": "block",
            "reason": "整页背景为照片媒体，需先声明烘焙像素处置（replace/keep/block）"
                      "与 recolor 路径，否则不得整页导入；无显式声明前阻断。",
            "basis": ["R3"],
            "check_slide": {
                "target": {"tone": "dark" if feat["background"]["dark"] else "light"},
                "background": "需先核对 p:bg 媒体与烘焙文字",
                "recolor": "主题色按目标基线角色重着色，冲突无路径时阻断",
            },
            "missing": ["烘焙像素声明", "背景处置口径", "基调 recolor 路径"],
        }

    if feat["char_count"] > MAX_CHARS_PER_SHAPE * max(feat["text_shapes"], 1):
        cautions.append("单块文本超长，先按句拆文本位再编排")
    if feat["shape_count"] > MAX_SHAPES_PER_PAGE:
        cautions.append("形状数超出单页保守上限，考虑拆页")

    # 装饰性过渡 / 封面母题可走整页导入复用；内容结构页走原生重建
    if page_type in ("section_transition", "closing", "thanks") \
            and feat["text_shapes"] <= 5:
        return {
            "decision": "import",
            "reason": "稀疏装饰页，可复用参考件现成环境动效母题，导入后只换 kicker "
                      "与章名 / 致谢文本。",
            "basis": ["SEC-1", "END-1", "THX-1"],
            "check_slide": {
                "target": {"tone": "dark" if feat["background"]["dark"] else "light"},
                "background": "随目标包 p:bg 统一",
                "recolor": "按角色映射重着色，保留颜色变换节点",
            },
            "batches_note": "导入页不改时序，只做文字替换与页序编排",
        }

    return {
        "decision": "native_rebuild",
        "reason": "内容结构可由封装函数原生构建；批次结构直接喂给 build 脚本。",
        "basis": ["BUILD-1"],
        "build": {
            "batches": batches or _batches(page_type, feat, engine),
            "preset_tier": ("effect_first 引擎族：高版本预设按观感选用，2019 降级但不崩"
                            if engine == ENGINE_FIRST else
                            "compat 引擎族：标准 DrawingML 安全族打底，高版本预设不下发"),
        },
        "cautions": cautions,
        "missing": missing,
    }


def _build_alternatives(page_type: str, budget: dict, engine: str,
                        index: dict, tier: str = TIER_FIDELITY) -> list[dict]:
    """多方案候选：至少 1 个锁结构模板式方案 + 1 个定制方案；索引有第三条候选时追加。

    每条候选：scheme_id / scheme(template|custom|effect|layout) / summary /
    cost（节点数估算 + 实现复杂度 + 字数要求）/ engine / tradeoff / ref。
    既有 plan 字段一律保留，alternatives 只增不改。
    """
    candidates = (index.get("page_types") or {}).get(page_type, [])
    layouts = [c for c in candidates if c.get("kind") == "layout"]
    effects = [c for c in candidates if c.get("kind") == "effect"]

    def _cost(cand: dict, complexity: str) -> dict:
        c = cand.get("cost", {}) if isinstance(cand.get("cost"), dict) else {}
        return {
            "effects": budget["estimated_effects"],
            "ctn": budget["estimated_ctn"],
            "complexity": c.get("complexity", complexity),
            "nodes": c.get("nodes", "≤%d 效果节点" % budget["estimated_effects"]),
            "copy_fit": c.get("copy_fit", "按骨架槽位控制字数"),
        }

    alts: list[dict] = []
    if layouts:
        t = layouts[0]
        alts.append({
            "scheme_id": "%s-tpl-1" % page_type,
            "scheme": "template",
            "summary": t.get("fit", "复用索引版式骨架，只换文案与配色"),
            "cost": _cost(t, "low"),
            "engine": t.get("engine", "both"),
            "tradeoff": "骨架槽位与批次顺序固定，长文案 / 非常规结构需压缩进槽位；"
                        + t.get("forbidden", ""),
            "ref": t.get("id", ""),
        })
    else:
        alts.append({
            "scheme_id": "%s-tpl-1" % page_type,
            "scheme": "template",
            "summary": "按页型保守骨架落位（标题 + 条目分批 fade），只换文案",
            "cost": {"effects": budget["estimated_effects"],
                     "ctn": budget["estimated_ctn"], "complexity": "low",
                     "nodes": "≤%d 效果节点" % budget["estimated_effects"],
                     "copy_fit": "单块 ≤60 字"},
            "engine": "both",
            "tradeoff": "无索引骨架可套，槽位按保守档排，观感弹性小。",
            "ref": "",
        })
    if effects:
        e = effects[0]
        alts.append({
            "scheme_id": "%s-cust-1" % page_type,
            "scheme": "custom",
            "summary": e.get("fit", "按页角色重新编排效果族与批次时序"),
            "cost": _cost(e, "medium"),
            "engine": e.get("engine", engine),
            "tradeoff": "需逐对象配对、重排时序并在目标机 F5 复核；"
                        + e.get("forbidden", ""),
            "ref": e.get("id", ""),
        })
    else:
        alts.append({
            "scheme_id": "%s-cust-1" % page_type,
            "scheme": "custom",
            "summary": "按页角色定制批次方向与强调，不走现成骨架",
            "cost": {"effects": budget["estimated_effects"],
                     "ctn": budget["estimated_ctn"], "complexity": "medium",
                     "nodes": "≤%d 效果节点" % budget["estimated_effects"],
                     "copy_fit": "按句拆文本位"},
            "engine": engine,
            "tradeoff": "重建工作量高于模板式；弱引擎上高版本项需逐条降级核对。",
            "ref": "",
        })
    for extra in (layouts[1:] + effects[1:])[:2]:
        alts.append({
            "scheme_id": "%s-alt-%d" % (page_type, len(alts) + 1),
            "scheme": extra.get("kind", "effect"),
            "summary": extra.get("fit", ""),
            "cost": _cost(extra, "medium"),
            "engine": extra.get("engine", engine),
            "tradeoff": extra.get("forbidden", ""),
            "ref": extra.get("id", ""),
        })
    if tier == TIER_FIDELITY:
        strategy, impact, ref = DECOR_RECIPES.get(page_type, DECOR_DEFAULT)
        alts.append({
            "scheme_id": "%s-baked-decor-1" % page_type,
            "scheme": "effect",
            "kind": "baked_decor_native_anim",
            "summary": "装饰层预渲染铺底（%s，观感 %s），内容层保留原生文本与入场动画"
                       % (strategy, impact),
            "cost": {"effects": budget["estimated_effects"],
                     "ctn": budget["estimated_ctn"], "complexity": "medium",
                     "nodes": "≤%d 效果节点" % budget["estimated_effects"],
                     "copy_fit": "结构化源须登记素材、生成参数与 sha256"},
            "engine": engine,
            "tradeoff": "装饰层可编辑性让位于观感；正文、数据与结构层保持原生，"
                        "结构化源随稿交付可重生成",
            "ref": ref,
        })
    return alts


def _decor(page_type: str) -> dict:
    """fidelity_first 下每页装饰层预烘焙配方描述；balanced / compat 不输出此字段。"""
    strategy, impact, ref = DECOR_RECIPES.get(page_type, DECOR_DEFAULT)
    return {"strategy": strategy, "engine": "decor_bake", "recipe_ref": ref,
            "impact": impact, "editable_cost": "high"}


def plan_page(feat: dict, engine: str = ENGINE_FIRST,
              tier: str = TIER_FIDELITY) -> dict:
    cls = classify(feat)
    page_type = cls["page_type"]
    trans_main, trans_alt, trans_basis = TRANSITION_BY_TYPE[page_type]
    if engine == ENGINE_COMPAT:
        trans_alt = COMPAT_ALT_TRANSITION.get(page_type, trans_alt)

    emphasis = []
    if page_type == "kpi":
        emphasis = ["大字号数字：grow 回弹一次"]
    elif page_type in ("timeline", "process"):
        emphasis = ["当前阶段 / 当前步骤：blink 一次或 grow"]
    elif page_type == "section_transition":
        emphasis = ["章名：fade 入场，环簇自转常驻"]
    elif page_type == "metric_band":
        emphasis = ["重点指标：单次 grow，不长闪；其余数字不强调"]
    elif page_type == "trend_chart":
        emphasis = ["当前系列 / 峰值点：grow 一次"]
    elif page_type == "risk_alert":
        emphasis = ["风险标记：grow_once 单次；对策文字不强调"]
    elif page_type == "org_chart":
        emphasis = ["当前层级节点：grow 一次"]
    else:
        emphasis = ["当前章节环 / 核心结论：单次强调，不长闪"]

    batches = _batches(page_type, feat, engine)
    route = _route(page_type, feat, batches, engine)
    budget = _budget(batches)
    alternatives = _build_alternatives(page_type, budget, engine, ASSET_INDEX, tier)
    ambient = page_type not in ("table_list",)

    engine_basis = ["ENGINE:%s" % engine]
    if engine == ENGINE_FIRST:
        engine_note = ("effect_first 引擎族（目标 365 / 新版 WPS）；2019 上高版本项降级"
                       "但不损坏文件，逐批降级提示见 batches[].degrade")
    else:
        engine_note = ("compat 引擎族（明确要求兼容 2019 / 旧 WPS）；仅下发标准"
                       "DrawingML 安全族，高版本预设不下发")

    result = {
        "page": feat.get("index"),
        "page_type": page_type,
        "page_type_label": TYPE_LABEL[page_type],
        "confidence": cls["confidence"],
        "engine_profile": engine,
        "engine_basis": engine_basis,
        "engine_note": engine_note,
        "basis": cls["basis"] + engine_basis,
        "traits": {
            "text_shapes": feat["text_shapes"],
            "char_count": feat["char_count"],
            "number_tokens": len(feat["number_tokens"]),
            "number_density": feat["number_density"],
            "font_levels": feat["font_levels"],
            "max_font_size": feat["max_font_size"],
            "has_table": feat["has_table"],
            "time_words": feat["time_words"],
            "has_quote": feat["has_quote"],
            "two_column": feat["two_column"],
            "card_count": feat["card_count"],
            "background": feat["background"],
            "competing": cls["competing"],
        },
        "batches": batches,
        "budget": budget,
        "presets": {
            "entrance": ["fade", "wipe_l", "fly_up"],
            "emphasis": ["grow", "blink", "spin", "path"],
            "exit": ["disappear", "zoom_out"],
        },
        "ambient_motion": ambient,
        "emphasis_targets": emphasis,
        "transition": {"suggested": trans_main, "alternative": trans_alt,
                       "basis": trans_basis},
        "route": route,
        "alternatives": alternatives,
        "cautions": route.get("cautions", []),
    }
    if tier == TIER_FIDELITY:
        result["decor"] = _decor(page_type)
    return result


# ----------------------------------------------------------------------
# 输入装载
# ----------------------------------------------------------------------
def _slide_order(zf: zipfile.ZipFile) -> list[str]:
    names = set(zf.namelist())
    pres = "ppt/presentation.xml"
    if pres not in names:
        return sorted(
            (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda s: int(re.search(r"\d+", s).group()))
    rels_name = "ppt/_rels/presentation.xml.rels"
    rels = {}
    if rels_name in names:
        rel_root = ET.fromstring(zf.read(rels_name))
        for rel in rel_root:
            rels[rel.get("Id")] = rel.get("Target")
    pres_root = ET.fromstring(zf.read(pres))
    ordered = []
    for sld in pres_root.find(f"{P}sldIdLst"):
        rid = sld.get(f"{{{R_NS}}}id")
        target = rels.get(rid, "")
        if target:
            part = re.sub(r"ppt/ppt/", "ppt/", "ppt/" + target.lstrip("/"))
            if part in names:
                ordered.append(part)
    return ordered or sorted(
        (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
        key=lambda s: int(re.search(r"\d+", s).group()))


def plan_pptx(path: str, page_only: int | None,
              engine: str = ENGINE_FIRST,
              tier: str = TIER_FIDELITY) -> list[dict]:
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        raise RouterError("无法打开包：%s" % e)
    with zf:
        parts = _slide_order(zf)
        if not parts:
            raise RouterError("包内没有幻灯片")
        plans = []
        for i, part in enumerate(parts, start=1):
            if page_only and i != page_only:
                continue
            feat = extract_features(part, zf.read(part))
            feat["index"] = i
            plans.append(plan_page(feat, engine, tier))
    return plans


def plan_json(path: str, page_only: int | None,
              engine: str = ENGINE_FIRST,
              tier: str = TIER_FIDELITY) -> list[dict]:
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        raise RouterError("JSON 页描述不可读：%s" % e)
    raw_pages = data.get("pages", [data]) if isinstance(data, dict) else data
    plans = []
    for i, raw in enumerate(raw_pages, start=1):
        if page_only and i != page_only:
            continue
        feat = _normalize_page(raw, i)
        plans.append(plan_page(feat, engine, tier))
    return plans


def _normalize_page(raw: dict, index: int) -> dict:
    texts = raw.get("texts") or []
    all_text = " ".join(texts)
    numbers = NUMBER_RE.findall(all_text)
    sizes = raw.get("font_sizes") or []
    hero = raw.get("hero_number_shapes")
    medium = raw.get("medium_number_shapes")
    if hero is None or medium is None:
        big = raw.get("big_number_shapes", 0)
        hero = raw.get("hero_number_shapes", big if big <= KPI_MAX_BLOCKS else 0)
        medium = raw.get("medium_number_shapes", max(big - hero, 0))
    return {
        "part": "page-%d" % index, "index": index,
        "shape_count": raw.get("shape_count", len(texts)),
        "text_shapes": len(texts), "body_shapes": len(texts),
        "char_count": len(all_text),
        "number_tokens": numbers,
        "number_density": round(len(numbers) / len(all_text), 3) if all_text else 0,
        "big_number_shapes": raw.get("big_number_shapes", hero + medium),
        "hero_number_shapes": hero,
        "medium_number_shapes": medium,
        "badge_number_shapes": raw.get("badge_number_shapes", 0),
        "font_sizes": sizes, "font_levels": len(set(sizes)),
        "max_font_size": max(sizes) if sizes else raw.get("max_font_size", 0),
        "has_table": bool(raw.get("has_table")),
        "has_chart": bool(raw.get("has_chart")),
        "pic_count": raw.get("pic_count", 0),
        "time_words": [w for w in TIME_WORDS if w in all_text],
        "header_words": [w for w in HEADER_WORDS if w in all_text],
        "header_cells": raw.get("header_cells", []),
        "has_quote": any(m in all_text for m in QUOTE_MARKS),
        "closing_words": [w for w in CLOSING_WORDS if w in all_text],
        "thanks_words": [w for w in THANKS_WORDS if w in all_text],
        "section_marks": bool(re.search(
            r"PART\s*\d|第[一二三四五六七八九十0-9]+章|章节", all_text)),
        "compare_words": [w for w in COMPARE_WORDS if w in all_text],
        "agenda_words": [w for w in AGENDA_WORDS if w in all_text],
        "risk_words": [w for w in RISK_WORDS if w in all_text],
        "case_words": [w for w in CASE_WORDS if w in all_text],
        "concl_words": [w for w in CONCL_WORDS if w in all_text],
        "org_words": [w for w in ORG_WORDS if w in all_text],
        "trend_words": [w for w in TREND_WORDS if w in all_text],
        "side_words": raw.get("side_words",
                              [w for w in COMPARE_SIDE_WORDS if w in all_text]),
        "step_count": sum(1 for t in texts if STEP_RE.match(t)),
        "arrow_count": len(ARROW_RE.findall(all_text)),
        "card_count": raw.get("card_count", 0),
        "short_labels": raw.get("short_labels", sum(1 for t in texts if len(t) <= 8)),
        "axis_count": raw.get("axis_count", 0),
        "node_count": raw.get("node_count", 0),
        "row_count": raw.get("row_count", max(len(texts) - 2, 1)),
        "two_column": bool(raw.get("two_column")),
        "background": raw.get("background", {"type": "none", "dark": None, "media": False}),
        "texts": texts,
    }


# ----------------------------------------------------------------------
# 输出
# ----------------------------------------------------------------------
def _batch_summary(batches: list[dict]) -> str:
    parts = []
    for b in batches:
        seg = "%s:%s×%d" % (b["role"], "/".join(b["preset_family"]), b["count"])
        if b.get("batch_count", 1) > 1:
            seg += "（%d 批·每批 %d）" % (b["batch_count"], b["per_batch"])
        parts.append(seg)
    return " → ".join(parts)


def _budget_summary(budget: dict) -> str:
    head = "预算 效果 %d/%d · cTn %d/%d" % (
        budget["estimated_effects"], budget["max_effects"],
        budget["estimated_ctn"], budget["max_ctn"])
    if budget["fits"]:
        return head + " ✓"
    k = max((s["batch_count"] for s in budget["split_plan"]), default=2)
    return head + " 超限→拆 %d 批" % k


def _md(plans: list[dict]) -> str:
    lines = ["# 效果路由计划", ""]
    if plans:
        lines.append("- 引擎档位：**%s**；%s" % (
            plans[0].get("engine_profile", ENGINE_FIRST), plans[0].get("engine_note", "")))
        lines.append("")
    lines.append("| 页 | 页型 | 置信度 | 批次角色 | 转场 | 路线 | 依据 |")
    lines.append("|---|---|---|---|---|---|---|")
    for p in plans:
        roles = " / ".join(b["role"] for b in p["batches"])
        lines.append("| %s | %s | %s | %s | %s | %s | %s |" % (
            p["page"], p["page_type_label"], p["confidence"], roles,
            p["transition"]["suggested"], p["route"]["decision"],
            ", ".join(p["basis"])))
    lines.append("")
    for p in plans:
        lines.append("## 第 %s 页 · %s（%s）" % (p["page"], p["page_type_label"],
                                                  p["confidence"]))
        lines.append("- 特征：文本块 %d，字数 %d，数字 %d，字号档 %d，卡片 %d" % (
            p["traits"]["text_shapes"], p["traits"]["char_count"],
            p["traits"]["number_tokens"], p["traits"]["font_levels"],
            p["traits"]["card_count"]))
        lines.append("- 路线：%s —— %s" % (p["route"]["decision"], p["route"]["reason"]))
        if p["route"].get("missing"):
            lines.append("- 待补：%s" % "、".join(p["route"]["missing"]))
        for b in p["batches"]:
            seg = "- 批次 %s：%d 个，预设 %s，延迟 %sms" % (
                b["role"], b["count"], "/".join(b["preset_family"]),
                "%d—%d" % tuple(b["delay_range"]))
            if b.get("degrade"):
                seg += "；降级提示：%s" % b["degrade"]
            if b.get("batch_count", 1) > 1:
                seg += "；每批 %d 个，拆 %d 批" % (b["per_batch"], b["batch_count"])
            lines.append(seg)
        lines.append("- 预算：效果 %d/%d，cTn %d/%d，fits=%s" % (
            p["budget"]["estimated_effects"], p["budget"]["max_ctn"],
            p["budget"]["estimated_ctn"], p["budget"]["max_ctn"],
            p["budget"]["fits"]))
        if p.get("decor"):
            d = p["decor"]
            lines.append("- 装饰层预烘焙：%s（引擎 %s，配方 %s，观感 %s，可编辑代价 %s）" % (
                d["strategy"], d["engine"], d["recipe_ref"], d["impact"],
                d["editable_cost"]))
        lines.append("- 候选方案（%d）：" % len(p["alternatives"]))
        for alt in p["alternatives"]:
            lines.append("  - `%s`（%s）%s｜复杂度 %s｜引擎 %s｜取舍：%s" % (
                alt["scheme_id"], alt["scheme"], alt["summary"],
                alt["cost"]["complexity"], alt["engine"], alt["tradeoff"]))
        if p["cautions"]:
            lines.append("- 注意：%s" % "；".join(p["cautions"]))
        lines.append("")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="内容识别与效果路由")
    ap.add_argument("input", help=".pptx 或 JSON 页描述")
    ap.add_argument("--page", type=int, default=None, help="只分析第 N 页（1 起）")
    ap.add_argument("--tier", choices=list(TIERS), default=TIER_FIDELITY,
                    help="交付档位：fidelity_first（默认，成品优先）/ balanced（兼顾）/ compat（兼容优先）")
    ap.add_argument("--engine", choices=list(ENGINES), default=None,
                    help="既有引擎族覆盖项：effect_first 或 compat；未给定时按 --tier 映射")
    ap.add_argument("--json", dest="json_out", help="机器可读计划输出路径")
    ap.add_argument("--md", dest="md_out", help="人读表格输出路径")
    args = ap.parse_args(argv)

    tier = args.tier
    engine = args.engine if args.engine else TIER_ENGINE[tier]
    # 显式 --engine 是既有调用路径：按对应旧档口径抑制 fidelity_first 新增字段，
    # 保证旧调用方输出与改动前逐字段一致。
    if args.engine:
        plan_tier = TIER_COMPAT if engine == ENGINE_COMPAT else TIER_BALANCED
    else:
        plan_tier = tier

    try:
        if args.input.lower().endswith(".json"):
            plans = plan_json(args.input, args.page, engine, plan_tier)
        else:
            plans = plan_pptx(args.input, args.page, engine, plan_tier)
    except RouterError as e:
        print("错误：%s" % e)
        return 2

    report = {"engine_profile": engine,
              "plans": plans,
              "totals": {"pages": len(plans),
                         "high_confidence": sum(1 for p in plans if p["confidence"] == "high"),
                         "low_confidence": sum(1 for p in plans if p["confidence"] == "low"),
                         "routes": {r: sum(1 for p in plans if p["route"]["decision"] == r)
                                    for r in ("native_rebuild", "import", "block")}}}
    if plan_tier == TIER_FIDELITY:
        report["tier"] = tier

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    if args.md_out:
        with open(args.md_out, "w", encoding="utf-8") as f:
            f.write(_md(plans))

    for p in plans:
        print("第 %s 页 [%s/%s] %s | 转场 %s | 批次 %s | %s | 依据 %s" % (
            p["page"], p["page_type_label"], p["confidence"], p["route"]["decision"],
            p["transition"]["suggested"], _batch_summary(p["batches"]),
            _budget_summary(p["budget"]), ", ".join(p["basis"])))
        alt_sum = "；".join("%s: %s" % (a["scheme_id"], a["summary"])
                            for a in p["alternatives"])
        print("    候选 %d：%s" % (len(p["alternatives"]), alt_sum))
    print("合计 %d 页，档位 %s，高置信 %d，低置信 %d，路线分布 %s" % (
        report["totals"]["pages"], engine, report["totals"]["high_confidence"],
        report["totals"]["low_confidence"], report["totals"]["routes"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
