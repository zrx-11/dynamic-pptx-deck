# -*- coding: utf-8 -*-
"""
effect_router.py — 内容识别与效果路由（仅标准库：zipfile + xml.etree + json）。

输入任一 .pptx（逐页解析形状 / 文本 / 数字 / 层级 / 时间词 / 对齐模式），或一份
JSON 页描述；输出 EffectPlan：页型、信息特征、批次拆法、预设族、常驻环境动效、
强调对象、转场建议与路线建议（原生重建 / 整页导入 / 阻断），每个推荐项带命中的
判定条件编号（basis），识别不确定时给 confidence=low 与保守档，不猜。

CLI：
    python3 effect_router.py <deck.pptx> [--page N] [--json OUT] [--md OUT]
    python3 effect_router.py <page-plan.json> [--json OUT] [--md OUT]

退出码：0=正常出计划；2=输入不可读。
"""
from __future__ import annotations

import argparse
import json
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
CLOSING_WORDS = ["谢谢", "总结", "三年之后", "THANKS", "Thanks", "致谢", "结尾", "展望"]
SECTION_MARKS = ["PART", "第", "章", "章节", "01", "02", "03", "04"]
COMPARE_WORDS = ["对照", "优劣", "正反", "vs", "VS", "两组", "对比"]

NUMBER_RE = re.compile(r"\d+(?:\.\d+)?%?")
STEP_RE = re.compile(r"^(?:0?[1-9]|[一二三四五六])\s*[、.．]")
ARROW_RE = re.compile(r"→|➜|->|▶")

# 规模上限（保守档）：超过即提示拆分或阻断核对
MAX_CHARS_PER_SHAPE = 60
MAX_SHAPES_PER_PAGE = 40
# 单页节点预算：效果节点 ≤12，cTn ≤60（沿用技能既有配额）
MAX_EFFECTS = 12
MAX_CTN = 60
CTN_PER_EFFECT = 3.8
# 各批量角色每点击批的元素配额
PER_BATCH = {"cards": 3, "rows": 4, "nodes": 3, "kpi_value": 3, "steps": 3, "summary": 4}
# 每个元素占用的效果节点数（badge 随卡片 zoom，不单独占节点）
EFFECTS_PER = {
    "ambient": 1, "title": 1, "subtitle": 1, "hero": 1, "kicker": 1,
    "chapter_title": 1, "kpi_value": 1, "label": 1, "action": 1, "header": 1,
    "rows": 1, "axis": 1, "nodes": 1, "connectors": 1, "text": 1, "steps": 1,
    "left": 1, "right": 1, "cards": 1, "card_text": 1, "quote_line": 1,
    "attribution": 1, "summary": 1, "closing": 1, "body": 1,
}


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
    section = bool(re.search(r"PART\s*\d|第[一二三四五六七八九十0-9]+章|章节", all_text))
    compare = any(w in all_text for w in COMPARE_WORDS)
    steps = sum(1 for s in body if STEP_RE.match(s["text"]))
    arrows = len(ARROW_RE.findall(all_text))

    containers = [
        s for s in body if s["prst"] in CONTAINER_PRST and s["box"]
        and s["box"][2] >= 1_500_000 and s["box"][3] >= 400_000
        and s["box"][2] < 11_000_000
    ]
    sizes = sorted({sz for s in shapes for sz in s["sizes"]})
    max_size = max(sizes) if sizes else 0
    # 大字号数字：数字所在 run 字号 ≥ 2000（20pt）
    big_numbers = 0
    for s in body:
        if NUMBER_RE.search(s["text"]) and s["sizes"] and max(s["sizes"]) >= 2000:
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
        "font_sizes": sizes,
        "font_levels": len(sizes),
        "max_font_size": max_size,
        "has_table": any(s["is_table"] for s in shapes),
        "time_words": time_hits,
        "header_words": header_hits,
        "header_cells": header_cells,
        "has_quote": quote,
        "closing_words": [w for w in CLOSING_WORDS if w in all_text],
        "section_marks": section,
        "compare_words": [w for w in COMPARE_WORDS if w in all_text],
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
    hit("END-1", "收束词命中且无数字",
        bool(feat["closing_words"]) and not feat["number_tokens"])
    hit("COVER-1", "大标题字号且文本块少、无卡片网格、无大字号数字",
        feat["max_font_size"] >= 3000 and body_n <= 14
        and feat["card_count"] < 5 and feat["big_number_shapes"] < 2)
    hit("SEC-1", "章节标记且正文稀疏",
        feat["section_marks"] and body_n <= 4
        and feat["card_count"] < 3)
    hit("QUO-1", "引号命中、无数字、长句少量块",
        feat["has_quote"] and not feat["number_tokens"]
        and body_n <= 4 and feat["char_count"] >= 30)
    hit("KPI-1", "大字号数字 ≥2，或有大字号数字且占比达标",
        feat["big_number_shapes"] >= 2 or
        (feat["big_number_shapes"] >= 1 and len(feat["number_tokens"]) >= 3
         and feat["number_density"] >= 0.08))
    hit("TBL-1", "表头词成格且短标签 ≥4",
        feat["has_table"] or
        (feat["header_cells"] and feat["short_labels"] >= 4 and body_n >= 7))
    hit("TIME-1", "时间词命中且有序条目成组",
        bool(feat["time_words"]) and (feat["step_count"] >= 2 or body_n >= 5))
    hit("PROC-1", "步骤编号或箭头成链，条目 ≥3",
        (feat["step_count"] >= 3 or feat["arrow_count"] >= 2) and body_n >= 5)
    hit("CMP-1", "左右双栏且有对比词或双大容器",
        feat["two_column"] and (feat["compare_words"] or feat["card_count"] == 2))
    hit("CARD-1", "并列卡片 ≥3 且各带文字",
        feat["card_count"] >= 3 and body_n >= 4)

    priority = ["closing", "cover", "section_transition", "quote", "kpi",
                "table_list", "timeline", "process", "two_column_compare",
                "card_grid"]
    code_map = {
        "END-1": "closing", "COVER-1": "cover", "SEC-1": "section_transition",
        "QUO-1": "quote", "KPI-1": "kpi", "TBL-1": "table_list",
        "TIME-1": "timeline", "PROC-1": "process", "CMP-1": "two_column_compare",
        "CARD-1": "card_grid",
    }
    matched = {code_map[c]: c for c, _, _ in rules}
    page_type = next((t for t in priority if t in matched), "card_grid")

    # 强特征：命中即可独立支撑该页型，不依赖密度 / 弱命中凑数
    strong = {
        "closing": bool(feat["closing_words"]),
        "cover": feat["max_font_size"] >= 3000 and feat["card_count"] < 3,
        "section_transition": feat["section_marks"] and body_n <= 4,
        "quote": feat["has_quote"] and feat["char_count"] >= 30 and body_n <= 3,
        "kpi": feat["big_number_shapes"] >= 2,
        "table_list": feat["has_table"],
        "timeline": bool(feat["time_words"]) and
        (feat["step_count"] >= 2 or feat.get("axis_count", 0) >= 1
         or feat.get("node_count", 0) >= 2),
        "process": feat["step_count"] >= 3 or feat["arrow_count"] >= 2,
        "two_column_compare": feat["two_column"] and bool(feat["compare_words"]),
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
}

TYPE_LABEL = {
    "cover": "封面", "section_transition": "章节过渡", "kpi": "KPI 数字",
    "table_list": "表格清单", "timeline": "时间线 / 学期推进", "process": "流程步骤",
    "two_column_compare": "双栏对比", "card_grid": "并列卡片", "quote": "引语页",
    "closing": "结尾",
}


def _b(role: str, count: int, presets: list[str], delay: list[int],
       note: str = "") -> dict:
    count = max(count, 1)
    per_batch = PER_BATCH.get(role)
    batch_count = (count + per_batch - 1) // per_batch if per_batch else 1
    entry = {
        "role": role, "count": count, "preset_family": presets,
        "delay_range": delay, "note": note,
        "effects": count * EFFECTS_PER.get(role, 1),
    }
    if per_batch:
        entry["per_batch"] = per_batch
        entry["batch_count"] = batch_count
    return entry


def _batches(page_type: str, feat: dict) -> list[dict]:
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
        return [
            _b("hero", 1, ["spin", "path"], [0, 0], "高冲击母题；稀疏页可整页导入复用"),
            _b("kicker", 1, ["fade"], [120, 250]),
            _b("chapter_title", 1, ["fade", "wipe_l"], [200, 400]),
        ]
    if page_type == "kpi":
        return [
            ambient,
            _b("kpi_value", values, ["zoom", "zoom_rot"], [0, 200],
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
            _b("cards", cards, ["fly_up", "wipe_l"], [150, 300],
               "第 i 张卡与其文字同批成对；编号圆牌随卡片 zoom，不另占效果节点；每批 ≤3 张"),
            _b("card_text", cards, ["fade"], [150, 450],
               "卡内文字与同序号卡片同批，晚 150—450ms"),
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
            _b("closing", 1, ["fade"], [300, 450], "收束语最后入场，不堆强调"),
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


def _route(page_type: str, feat: dict, batches: list[dict] | None = None) -> dict:
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
    if page_type in ("section_transition", "closing") and feat["text_shapes"] <= 5:
        return {
            "decision": "import",
            "reason": "稀疏装饰页，可复用参考件现成环境动效母题，导入后只换 kicker "
                      "与章名文本。",
            "basis": ["SEC-1", "END-1"],
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
            "batches": batches or _batches(page_type, feat),
            "preset_tier": "L0 安全层打底，L1 仅作同组点缀",
        },
        "cautions": cautions,
        "missing": missing,
    }


def plan_page(feat: dict) -> dict:
    cls = classify(feat)
    page_type = cls["page_type"]
    trans_main, trans_alt, trans_basis = TRANSITION_BY_TYPE[page_type]

    emphasis = []
    if page_type == "kpi":
        emphasis = ["大字号数字：grow 回弹一次"]
    elif page_type in ("timeline", "process"):
        emphasis = ["当前阶段 / 当前步骤：blink 一次或 grow"]
    elif page_type == "section_transition":
        emphasis = ["章名：fade 入场，环簇自转常驻"]
    else:
        emphasis = ["当前章节环 / 核心结论：单次强调，不长闪"]

    batches = _batches(page_type, feat)
    route = _route(page_type, feat, batches)
    budget = _budget(batches)
    ambient = page_type not in ("table_list",)

    return {
        "page": feat.get("index"),
        "page_type": page_type,
        "page_type_label": TYPE_LABEL[page_type],
        "confidence": cls["confidence"],
        "basis": cls["basis"],
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
        "cautions": route.get("cautions", []),
    }


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


def plan_pptx(path: str, page_only: int | None) -> list[dict]:
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
            plans.append(plan_page(feat))
    return plans


def plan_json(path: str, page_only: int | None) -> list[dict]:
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
        plans.append(plan_page(feat))
    return plans


def _normalize_page(raw: dict, index: int) -> dict:
    texts = raw.get("texts") or []
    all_text = " ".join(texts)
    numbers = NUMBER_RE.findall(all_text)
    sizes = raw.get("font_sizes") or []
    return {
        "part": "page-%d" % index, "index": index,
        "shape_count": raw.get("shape_count", len(texts)),
        "text_shapes": len(texts), "body_shapes": len(texts),
        "char_count": len(all_text),
        "number_tokens": numbers,
        "number_density": round(len(numbers) / len(all_text), 3) if all_text else 0,
        "big_number_shapes": raw.get("big_number_shapes", 0),
        "font_sizes": sizes, "font_levels": len(set(sizes)),
        "max_font_size": max(sizes) if sizes else raw.get("max_font_size", 0),
        "has_table": bool(raw.get("has_table")),
        "time_words": [w for w in TIME_WORDS if w in all_text],
        "header_words": [w for w in HEADER_WORDS if w in all_text],
        "header_cells": raw.get("header_cells", []),
        "has_quote": any(m in all_text for m in QUOTE_MARKS),
        "closing_words": [w for w in CLOSING_WORDS if w in all_text],
        "section_marks": any(w in all_text for w in SECTION_MARKS),
        "compare_words": [w for w in COMPARE_WORDS if w in all_text],
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
            if b.get("batch_count", 1) > 1:
                seg += "；每批 %d 个，拆 %d 批" % (b["per_batch"], b["batch_count"])
            lines.append(seg)
        lines.append("- 预算：效果 %d/%d，cTn %d/%d，fits=%s" % (
            p["budget"]["estimated_effects"], p["budget"]["max_effects"],
            p["budget"]["estimated_ctn"], p["budget"]["max_ctn"],
            p["budget"]["fits"]))
        if p["cautions"]:
            lines.append("- 注意：%s" % "；".join(p["cautions"]))
        lines.append("")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="内容识别与效果路由")
    ap.add_argument("input", help=".pptx 或 JSON 页描述")
    ap.add_argument("--page", type=int, default=None, help="只分析第 N 页（1 起）")
    ap.add_argument("--json", dest="json_out", help="机器可读计划输出路径")
    ap.add_argument("--md", dest="md_out", help="人读表格输出路径")
    args = ap.parse_args(argv)

    try:
        if args.input.lower().endswith(".json"):
            plans = plan_json(args.input, args.page)
        else:
            plans = plan_pptx(args.input, args.page)
    except RouterError as e:
        print("错误：%s" % e)
        return 2

    report = {"plans": plans,
              "totals": {"pages": len(plans),
                         "high_confidence": sum(1 for p in plans if p["confidence"] == "high"),
                         "low_confidence": sum(1 for p in plans if p["confidence"] == "low"),
                         "routes": {r: sum(1 for p in plans if p["route"]["decision"] == r)
                                    for r in ("native_rebuild", "import", "block")}}}

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
    print("合计 %d 页，高置信 %d，低置信 %d，路线分布 %s" % (
        report["totals"]["pages"], report["totals"]["high_confidence"],
        report["totals"]["low_confidence"], report["totals"]["routes"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
