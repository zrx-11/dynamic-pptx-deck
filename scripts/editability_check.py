# -*- coding: utf-8 -*-
"""
editability_check.py — 成稿可编辑性与结构化源审计（仅标准库：zipfile + xml.etree）。

按交付档位分档出报告：逐页评分、不可编辑元素清单（类型 / 位置 / 原因）、结构化源
齐全性。判定口径详见 references/editability-audit.md。

扣分码与分值（不随档位改变，只增不改）：
    IMAGE_PAGE       40  整页大图（覆盖率 ≥ 80%）且文字节点 ≤ 3，整页图片式幻灯片；
    TEXT_BITMAP      25  文本槽位存在但 a:t 为空或极少，文字疑似位图化；
    PIC_FLATTENED    10  仅 1–2 张全幅图片承载内容，形状被扁平化；
    TABLE_SHAPE      10  无原生 a:tbl，页面呈多列多行短文本网格，表格疑似形状拼装；
    CHART_SCREENSHOT 10  无原生图表部件，却有占页 ≥ 35% 的图片，图表疑似截图；
    BAKED_EFFECT      5  a14 烘焙位图层，效果参数不可再编辑；
    EMPTY_CONTENT     3  无文字、无图片、无表格图表，页面近乎空白；
    BLIPFILL_HEAVY     1  形状图片填充较多（可编辑，仅提示核对用途）。

结构化源码：
    SOURCE_MISSING     无 deck-manifest.json，或页面在 pages 中无对应 index；
    SOURCE_INCOMPLETE  页面条目存在但 content 为空，或图片填充引用未在 assets[] 登记。

CLI：
    python3 editability_check.py <deck.pptx> [--manifest PATH]
        [--tier fidelity_first|balanced|compat]
        [--report-json OUT.json] [--report-md OUT.md] [--threshold N]
    python3 editability_check.py --selftest

退出码：0 = 通过（当前档位无阻断页）；1 = 存在阻断页；2 = 用法或读取错误。
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"

P = "{%s}" % P_NS
A = "{%s}" % A_NS
R = "{%s}" % R_NS

# 扣分权重：与 references/editability-audit.md 评分表一一对应。
WEIGHTS = {
    "IMAGE_PAGE": 40,
    "TEXT_BITMAP": 25,
    "PIC_FLATTENED": 10,
    "TABLE_SHAPE": 10,
    "CHART_SCREENSHOT": 10,
    "BAKED_EFFECT": 5,
    "EMPTY_CONTENT": 3,
    "BLIPFILL_HEAVY": 1,
}

# 覆盖率阈值（占画布面积比例）。
FULL_PAGE_COVERAGE = 0.80
CHART_PIC_COVERAGE = 0.35
# 表格拼装网格阈值。
GRID_MIN_COLS = 3
GRID_MIN_ROWS = 4
GRID_MIN_CELLS = 20
GRID_SHORT_RATIO = 0.65
GRID_SHORT_LEN = 10

TEXT_NAME_RE = re.compile(
    r"title|kicker|head|label|text|txt|lb|desc|note|body|caption|sub|nav|tab|cell|headline|para",
    re.I)
GRID_X_GAP = 800000
GRID_Y_GAP = 500000

TIERS = ("fidelity_first", "balanced", "compat")
DEFAULT_THRESHOLD = 70

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
DEFAULT_CANVAS = (12192000, 6858000)


class CheckError(RuntimeError):
    """输入不可读或不是有效 pptx。"""


# ----------------------------------------------------------------------
# 包读取
# ----------------------------------------------------------------------
def _slide_order(zf: zipfile.ZipFile) -> list[str]:
    """按 presentation.xml 的 sldIdLst 顺序取 slide part；缺关系时回退文件名排序。"""
    names = set(zf.namelist())
    pres_name = "ppt/presentation.xml"
    if pres_name not in names:
        return sorted(
            (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda s: int(re.search(r"\d+", s).group()),
        )
    rels_name = "ppt/_rels/presentation.xml.rels"
    rels: dict[str, str] = {}
    if rels_name in names:
        rel_root = ET.fromstring(zf.read(rels_name))
        for rel in rel_root:
            rels[rel.get("Id")] = rel.get("Target")
    pres = ET.fromstring(zf.read(pres_name))
    ordered: list[str] = []
    for sld in pres.find(f"{P}sldIdLst"):
        rid = sld.get(R + "id")
        target = rels.get(rid, "")
        if target:
            part = "ppt/" + target.lstrip("/")
            part = re.sub(r"ppt/ppt/", "ppt/", part)
            if part in names:
                ordered.append(part)
    if not ordered:
        ordered = sorted(
            (n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda s: int(re.search(r"\d+", s).group()),
        )
    return ordered


def _canvas(zf: zipfile.ZipFile) -> tuple[int, int]:
    if "ppt/presentation.xml" not in zf.namelist():
        return DEFAULT_CANVAS
    pres = ET.fromstring(zf.read("ppt/presentation.xml"))
    sz = pres.find(f"{P}sldSz")
    if sz is None:
        return DEFAULT_CANVAS
    try:
        return int(sz.get("cx")), int(sz.get("cy"))
    except (TypeError, ValueError):
        return DEFAULT_CANVAS


class Package:
    """持有打开的包，按需读取 slide XML 与 slide 级关系。"""

    def __init__(self, path: str):
        try:
            self.zf = zipfile.ZipFile(path)
        except (OSError, zipfile.BadZipFile) as e:
            raise CheckError("无法打开包：%s" % e)
        self.names = set(self.zf.namelist())
        self.slide_parts = _slide_order(self.zf)
        if not self.slide_parts:
            raise CheckError("包内没有幻灯片")
        self.canvas = _canvas(self.zf)
        self.chart_parts = {n for n in self.names if re.match(r"ppt/charts/chart\d+\.xml$", n)}
        self._rels_cache: dict[str, dict[str, str]] = {}

    def read(self, part: str) -> bytes:
        return self.zf.read(part)

    def slide_rels(self, part: str) -> dict[str, str]:
        if part in self._rels_cache:
            return self._rels_cache[part]
        base = part.rsplit("/", 1)[-1]
        rels_name = "ppt/slides/_rels/%s.rels" % base
        rels: dict[str, str] = {}
        if rels_name in self.names:
            root = ET.fromstring(self.zf.read(rels_name))
            for rel in root:
                rels[rel.get("Id")] = rel.get("Target") or ""
        self._rels_cache[part] = rels
        return rels

    def close(self) -> None:
        self.zf.close()


# ----------------------------------------------------------------------
# 形状收集
# ----------------------------------------------------------------------
def _box(node: ET.Element) -> tuple[int, int, int, int] | None:
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


def collect_shapes(root: ET.Element) -> list[dict]:
    """递归收集 sp / grpSp / pic / graphicFrame，返回形状记录列表。"""
    out: list[dict] = []
    sp_tree = root.find(f".//{P}spTree")

    def walk(node: ET.Element) -> None:
        for child in node:
            if child.tag not in (P + "sp", P + "grpSp", P + "pic", P + "graphicFrame"):
                continue
            nv = child.find(f".//{P}cNvPr")
            sid = nv.get("id") if nv is not None else None
            name = nv.get("name") if nv is not None else ""
            text = "".join(t.text or "" for t in child.iter(A + "t")).strip()
            out.append({
                "id": sid, "name": name,
                "tag": child.tag.replace(P, ""),
                "box": _box(child), "text": text,
            })
            if child.tag == P + "grpSp":
                walk(child)

    if sp_tree is not None:
        walk(sp_tree)
    return out


def _cluster(values: list[int], gap: int) -> list[list[int]]:
    if not values:
        return []
    ordered = sorted(values)
    clusters: list[list[int]] = [[ordered[0]]]
    for v in ordered[1:]:
        if v - clusters[-1][-1] <= gap:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return clusters


def _grid(shapes: list[dict]) -> dict:
    cells = [s for s in shapes if s["text"] and s["box"] is not None]
    if len(cells) < GRID_MIN_CELLS:
        return {"cols": 0, "rows": 0, "cells": 0, "short_cells": 0, "short_ratio": 0}
    xs = [s["box"][0] for s in cells]
    ys = [s["box"][1] for s in cells]
    col_clusters = _cluster(xs, GRID_X_GAP)
    row_clusters = _cluster(ys, GRID_Y_GAP)
    if len(col_clusters) < GRID_MIN_COLS or len(row_clusters) < GRID_MIN_ROWS:
        return {"cols": len(col_clusters), "rows": len(row_clusters),
                "cells": len(cells), "short_cells": 0, "short_ratio": 0}

    def cluster_idx(v: int, clusters: list[list[int]]) -> int:
        for i, cl in enumerate(clusters):
            if min(cl) - GRID_X_GAP <= v <= max(cl) + GRID_X_GAP:
                return i
        return -1

    grid_cells: list[dict] = []
    for s in cells:
        ci = cluster_idx(s["box"][0], col_clusters)
        ri = cluster_idx(s["box"][1], row_clusters)
        if ci >= 0 and ri >= 0:
            grid_cells.append({"col": ci, "row": ri, "shape": s})
    short = [c for c in grid_cells if len(c["shape"]["text"]) <= GRID_SHORT_LEN]
    ratio = len(short) / len(grid_cells) if grid_cells else 0
    return {
        "cols": len(col_clusters), "rows": len(row_clusters),
        "cells": len(grid_cells), "short_cells": len(short),
        "short_ratio": round(ratio, 2),
    }


def _para_has_fmt(sp_el: ET.Element) -> bool:
    for p in sp_el.iter(A + "p"):
        if p.find(A + "rPr") is not None or p.find(A + "pPr") is not None:
            return True
    return False


def _attr(root: ET.Element, shape_id: str | None) -> str | None:
    if shape_id is None:
        return None
    for sp_el in root.iter(P + "sp"):
        nv = sp_el.find(f"{P}nvSpPr/{P}cNvPr")
        if nv is not None and nv.get("id") == shape_id:
            return sp_el.get("useBgFill")
    return None


# ----------------------------------------------------------------------
# 单页结构分析（只产扣分码与统计，档位阻断在后面统一解析）
# ----------------------------------------------------------------------
def analyze_slide(part: str, xml: bytes, pkg: Package) -> dict:
    root = ET.fromstring(xml)
    shapes = collect_shapes(root)
    canvas_area = pkg.canvas[0] * pkg.canvas[1]

    sp = [s for s in shapes if s["tag"] == "sp"]
    pics = [s for s in shapes if s["tag"] == "pic"]
    frames = [s for s in shapes if s["tag"] == "graphicFrame"]
    groups = [s for s in shapes if s["tag"] == "grpSp"]

    text_nodes = len(list(root.iter(A + "t")))
    text_slots: list[str] = []
    for sp_el in root.iter(P + "sp"):
        nv = sp_el.find(f"{P}nvSpPr/{P}cNvPr")
        name = nv.get("name") if nv is not None else ""
        txt = "".join(t.text or "" for t in sp_el.iter(A + "t")).strip()
        if txt or TEXT_NAME_RE.search(name or "") or _para_has_fmt(sp_el):
            text_slots.append(name)
    text_shapes = [s for s in shapes if s["text"]]

    native_tables = len(root.findall(f".//{A}tbl"))

    rels = pkg.slide_rels(part)
    native_charts = 0
    for chart in root.iter(C_NS + "chart"):
        rid = chart.get(R + "id")
        if rid or rels.get(rid, ""):
            native_charts += 1
    if native_charts == 0 and pkg.chart_parts:
        native_charts = sum(1 for target in rels.values() if "charts/chart" in target)

    shape_blipfills = 0
    for sp_el in root.iter(P + "sp"):
        if sp_el.find(f"{P}spPr/{A}blipFill") is not None:
            shape_blipfills += 1
    pic_blips = len(root.findall(f".//{P}pic/{P}blipFill"))

    bg = root.find(f".//{P}bg")
    bg_blipfill = bg is not None and bg.find(f".//{A}blipFill") is not None
    use_bg_fill = sum(1 for s in sp if _attr(root, s["id"]) == "1")

    editable_effects = 0
    for eff in root.iter(A + "effectLst"):
        editable_effects += sum(1 for c in eff
                                if c.tag in (A + "blur", A + "glow", A + "prstShdw"))
    baked_effects = 0
    for el in root.iter():
        if el.tag.startswith("{%s}" % A14_NS):
            local = el.tag.split("}", 1)[1]
            if local in ("imgProps", "imgLayer"):
                baked_effects += 1

    pic_ratios: list[dict] = []
    max_coverage = 0.0
    for s in pics:
        if s["box"] is None:
            continue
        area = s["box"][2] * s["box"][3]
        ratio = area / canvas_area if canvas_area else 0.0
        pic_ratios.append({"id": s["id"], "name": s["name"], "coverage": round(ratio, 3)})
        max_coverage = max(max_coverage, ratio)

    grid = _grid(shapes)
    findings: list[dict] = []

    def add(code: str, region: str, count: int, reason: str) -> None:
        findings.append({"code": code, "region": region, "count": count, "reason": reason})

    pic_names = ", ".join(s["name"] or s["id"] or "pic" for s in pics) or "slide"

    if max_coverage >= FULL_PAGE_COVERAGE and text_nodes <= 3:
        add("IMAGE_PAGE", pic_names, len(pics),
            "整页图片式幻灯片：图片覆盖率 %.0f%%，文字节点仅 %d 个，内容无法逐元素修改。"
            % (max_coverage * 100, text_nodes))

    if text_slots:
        ratio_text = len(text_shapes) / len(text_slots)
        if (len(text_slots) >= 2 and text_nodes == 0) or \
           (len(text_slots) >= 3 and ratio_text < 0.5):
            add("TEXT_BITMAP", "slide", len(text_slots),
                "文字疑似位图化：%d 个文本槽位中仅 %d 个含真实文字节点（a:t 共 %d 个）。"
                % (len(text_slots), len(text_shapes), text_nodes))

    if 1 <= len(pics) <= 2 and max_coverage >= FULL_PAGE_COVERAGE and text_nodes > 3:
        add("PIC_FLATTENED", pic_names, len(pics),
            "内容扁平化为单张图：仅 %d 张图片且覆盖率 %.0f%%，其余元素可选中性差。"
            % (len(pics), max_coverage * 100))

    if native_tables == 0 and grid["cols"] >= GRID_MIN_COLS and \
       grid["rows"] >= GRID_MIN_ROWS and grid["cells"] >= GRID_MIN_CELLS and \
       grid.get("short_ratio", 0) >= GRID_SHORT_RATIO:
        add("TABLE_SHAPE", "slide", grid["cells"],
            "表格疑似形状拼装：无原生 a:tbl，呈 %d 列 × %d 行、%d 个短文本单元格（占比 %.0f%%）。"
            % (grid["cols"], grid["rows"], grid["cells"], grid.get("short_ratio", 0) * 100))

    if native_charts == 0:
        big_pics = [p for p in pic_ratios if p["coverage"] >= CHART_PIC_COVERAGE]
        if big_pics:
            add("CHART_SCREENSHOT", ", ".join(p["name"] for p in big_pics), len(big_pics),
                "图表疑似截图：无原生图表部件，%d 张图片占页 ≥ %.0f%%。"
                % (len(big_pics), CHART_PIC_COVERAGE * 100))

    if baked_effects > 0:
        add("BAKED_EFFECT", "slide", baked_effects,
            "艺术效果已烘焙为位图：%d 个 a14 烘焙层节点，参数不可再编辑。" % baked_effects)

    if not shapes and text_nodes == 0:
        add("EMPTY_CONTENT", "slide", 1, "页面无任何形状、文字或图片，近乎空白。")

    if shape_blipfills >= 3:
        add("BLIPFILL_HEAVY", "slide", shape_blipfills,
            "%d 个形状使用图片填充；填充本身可编辑，需核对是否为背景纹理而非内容截图。"
            % shape_blipfills)

    score = 100
    for f in findings:
        score -= WEIGHTS[f["code"]]
    score = max(0, score)

    has_image_fill = bool(pics) or shape_blipfills > 0 or bg_blipfill

    stats = {
        "shapes": len(shapes), "sp": len(sp), "pic": len(pics),
        "graphic_frame": len(frames), "group": len(groups),
        "text_shapes": len(text_shapes), "text_slots": len(text_slots),
        "text_nodes": text_nodes, "native_tables": native_tables,
        "native_charts": native_charts, "shape_blipfills": shape_blipfills,
        "pic_blips": pic_blips, "bg_blipfill": bg_blipfill,
        "use_bg_fill": use_bg_fill, "editable_effects": editable_effects,
        "baked_effects": baked_effects,
        "max_pic_coverage": round(max_coverage, 3),
        "pic_ratios": pic_ratios, "grid": grid,
        "has_image_fill": has_image_fill,
    }
    return {"part": part, "score": score, "findings": findings, "stats": stats}


# ----------------------------------------------------------------------
# 结构化源（deck-manifest.json）
# ----------------------------------------------------------------------
def load_manifest(path: str | None) -> dict | None:
    """读取 manifest；路径缺省或文件不存在时返回 None（按 SOURCE_MISSING 处理）。"""
    if not path:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def check_source(manifest: dict | None, index: int, has_image_fill: bool) -> dict:
    """按冻结稿第三节判定结构化源齐全性，返回 complete 与 missing 码列表。"""
    missing: list[str] = []
    if manifest is None:
        missing.append("SOURCE_MISSING")
        return {"complete": False, "missing": missing}

    entries = {}
    for entry in manifest.get("pages", []) or []:
        if isinstance(entry, dict) and "index" in entry:
            try:
                entries[int(entry["index"])] = entry
            except (TypeError, ValueError):
                continue
    entry = entries.get(index)
    if entry is None:
        missing.append("SOURCE_MISSING")
        return {"complete": False, "missing": missing}

    content = entry.get("content") or []
    if not content:
        missing.append("SOURCE_INCOMPLETE")

    if has_image_fill:
        registry = {a.get("path") for a in manifest.get("assets", [])
                    if isinstance(a, dict) and a.get("path")}
        refs = entry.get("assets") or []
        if not refs or any(r not in registry for r in refs):
            missing.append("SOURCE_INCOMPLETE")

    return {"complete": not missing, "missing": missing}


# ----------------------------------------------------------------------
# 分档阻断解析
# ----------------------------------------------------------------------
def _resolve_severity(code: str, tier: str) -> str:
    """档位感知的级别解析。返回 block / warn。"""
    if code == "SOURCE_MISSING":
        return "block"
    if code == "SOURCE_INCOMPLETE":
        # 成品优先档：结构化源缺失即阻断；兼顾 / 兼容档仅告警。
        return "block" if tier == "fidelity_first" else "warn"
    if code in ("IMAGE_PAGE", "TEXT_BITMAP"):
        # 成品优先档允许低分与图片化，不阻断；兼顾 / 兼容档阻断。
        return "warn" if tier == "fidelity_first" else "block"
    if code == "BAKED_EFFECT":
        return "block" if tier == "compat" else "warn"
    return "warn"


def build_report(slides: list[tuple[str, bytes]], pkg: Package, manifest: dict | None,
                 tier: str, threshold: int) -> dict:
    pages: list[dict] = []
    blocking_pages: list[int] = []
    by_code: dict[str, int] = {}

    for i, (part, xml) in enumerate(slides, start=1):
        base = analyze_slide(part, xml, pkg)
        source = check_source(manifest, i, base["stats"]["has_image_fill"])

        findings: list[dict] = []
        for f in base["findings"]:
            findings.append({**f, "severity": _resolve_severity(f["code"], tier)})
        for code in source["missing"]:
            findings.append({
                "code": code, "region": "source", "count": 1,
                "severity": _resolve_severity(code, tier),
                "reason": {
                    "SOURCE_MISSING": "结构化源缺失：manifest 无本页条目。",
                    "SOURCE_INCOMPLETE": "结构化源不完整：content 为空或图片资产未在 assets[] 登记。",
                }[code],
            })

        # 档位阻断判定：
        # fidelity_first：仅结构化源阻断，低分 / 图片页不阻断；
        # balanced / compat：阻断级 findings，或单页评分低于阈值；compat 另计 BAKED_EFFECT。
        score_low = base["score"] < threshold
        blocked = any(f["severity"] == "block" for f in findings)
        if tier in ("balanced", "compat") and score_low:
            blocked = True
        if blocked:
            blocking_pages.append(i)

        verdict = "block" if blocked else ("warn" if any(f["severity"] == "warn" for f in findings) else "ok")

        for f in findings:
            by_code[f["code"]] = by_code.get(f["code"], 0) + 1

        pages.append({
            "index": i,
            "score": base["score"],
            "verdict": verdict,
            "findings": findings,
            "source": source,
            "stats": base["stats"],
        })

    deck_score = round(sum(p["score"] for p in pages) / len(pages)) if pages else 0
    has_image = any(f["code"] == "IMAGE_PAGE" for p in pages for f in p["findings"])
    has_finding = any(p["findings"] for p in pages)
    deck_verdict = "image_based" if has_image else ("mixed" if has_finding else "editable")

    return {
        "deck": {"file": "", "tier": tier, "score": deck_score,
                 "verdict": deck_verdict, "blocking_pages": blocking_pages},
        "pages": pages,
        "summary": {"by_code": by_code, "threshold": threshold, "pages": len(pages)},
    }


# ----------------------------------------------------------------------
# 输出
# ----------------------------------------------------------------------
def _md_table(report: dict) -> str:
    d = report["deck"]
    lines = ["# 成稿可编辑性与结构化源审计报告", ""]
    lines.append("- 档位：**%s**；总体判定：**%s**；整包分数 **%d**；阈值 %d 分；阻断页 %s"
                 % (d["tier"], d["verdict"], d["score"], report["summary"]["threshold"],
                    ", ".join(str(x) for x in d["blocking_pages"]) or "无"))
    lines.append("")
    lines.append("| 页 | 分数 | 判定 | 图片 | 原生表格 | 原生图表 | 结构化源 | 问题 |")
    lines.append("|---|---|---|---|---|---|---|---|")
    for p in report["pages"]:
        s = p["stats"]
        src = "齐全" if p["source"]["complete"] else "、".join(p["source"]["missing"])
        brief = "；".join("%s(%s)" % (f["code"], f["severity"]) for f in p["findings"]) or "—"
        lines.append("| %d | %d | %s | %d | %d | %d | %s | %s |"
                     % (p["index"], p["score"], p["verdict"], s["pic"],
                        s["native_tables"], s["native_charts"], src, brief))
    lines.append("")
    for p in report["pages"]:
        if not p["findings"]:
            continue
        lines.append("## 第 %d 页" % p["index"])
        for f in p["findings"]:
            lines.append("- **[%s] %s**（%s，×%d）%s"
                         % (f["severity"], f["code"], f["region"], f["count"], f["reason"]))
        lines.append("")
    return "\n".join(lines)


def print_console(report: dict) -> None:
    d = report["deck"]
    for p in report["pages"]:
        s = p["stats"]
        src = "source-ok" if p["source"]["complete"] else ",".join(p["source"]["missing"])
        print("第 %2d 页 [%s] 分数 %3d 图 %d 表格 %d 图表 %d 源 %s"
              % (p["index"], p["verdict"], p["score"], s["pic"],
                 s["native_tables"], s["native_charts"], src))
        for f in p["findings"]:
            print("  %s %s（%s）：%s" % (f["severity"].upper(), f["code"], f["region"], f["reason"]))
    print("合计：%d 页，档位 %s，阻断页 %s，整包分数 %d —— %s"
          % (report["summary"]["pages"], d["tier"], d["blocking_pages"],
             d["score"], d["verdict"]))


# ----------------------------------------------------------------------
# 自检 fixture
# ----------------------------------------------------------------------
def _fixture_zip(slide_xmls: list[bytes]) -> io.BytesIO:
    buf = io.BytesIO()
    n = len(slide_xmls)
    overrides = []
    for i in range(1, n + 1):
        overrides.append(
            '<Override PartName="/ppt/slides/slide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            % i)
    ct = ('<Types xmlns="%s"><Default Extension="rels" '
          'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>%s</Types>'
          % (REL_NS, "".join(overrides)))
    rels = ('<Relationships xmlns="%s"><Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="ppt/presentation.xml"/></Relationships>' % REL_NS)
    sld_ids = "".join('<p:sldId id="%d" r:id="rId%d"/>' % (255 + i, i + 1)
                      for i in range(1, n + 1))
    pres = ('<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
            '<p:sldIdLst>%s</p:sldIdLst></p:presentation>'
            % (A_NS, R_NS, P_NS, sld_ids))
    pres_rels = "".join(
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide%d.xml"/>' % (i + 1, i)
        for i in range(1, n + 1))
    pres_rels = '<Relationships xmlns="%s">%s</Relationships>' % (REL_NS, pres_rels)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", XML_DECL + ct.encode())
        zf.writestr("_rels/.rels", XML_DECL + rels.encode())
        zf.writestr("ppt/presentation.xml", XML_DECL + pres.encode())
        zf.writestr("ppt/_rels/presentation.xml.rels", XML_DECL + pres_rels.encode())
        for i, xml in enumerate(slide_xmls, start=1):
            zf.writestr("ppt/slides/slide%d.xml" % i, xml)
    buf.seek(0)
    return buf


def _native_slide() -> bytes:
    return XML_DECL + (
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>原生标题</a:t></a:r></a:p></p:txBody></p:sp>'
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="正文"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="500000" y="1600000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>原生正文文字</a:t></a:r></a:p></p:txBody></p:sp>'
        '</p:spTree></p:cSld></p:sld>'
        % (A_NS, R_NS, P_NS)).encode()


def _image_slide() -> bytes:
    return XML_DECL + (
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>'
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="整页图"/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:spPr>'
        '</p:pic>'
        '</p:spTree></p:cSld></p:sld>'
        % (A_NS, R_NS, P_NS)).encode()


def _baked_effect_slide() -> bytes:
    return XML_DECL + (
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s" xmlns:a14="%s">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>'
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="效果图"/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId1"/></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="3000000" cy="2000000"/></a:xfrm></p:spPr>'
        '<p:extLst><a14:imgProps><a14:imgLayer name="layer1"/></a14:imgProps></p:extLst>'
        '</p:pic>'
        '</p:spTree></p:cSld></p:sld>'
        % (A_NS, R_NS, P_NS, A14_NS)).encode()


def _package_from_buffer(buf: io.BytesIO) -> Package:
    pkg = Package.__new__(Package)
    pkg.zf = zipfile.ZipFile(buf)
    pkg.names = set(pkg.zf.namelist())
    pkg.slide_parts = _slide_order(pkg.zf)
    if not pkg.slide_parts:
        raise CheckError("包内没有幻灯片")
    pkg.canvas = _canvas(pkg.zf)
    pkg.chart_parts = {n for n in pkg.names if re.match(r"ppt/charts/chart\d+\.xml$", n)}
    pkg._rels_cache = {}
    return pkg


def _run_fixture(slides: list[bytes], manifest: dict | None = None,
                 tier: str = "fidelity_first", threshold: int = DEFAULT_THRESHOLD) -> dict:
    buf = _fixture_zip(slides)
    pkg = _package_from_buffer(buf)
    loaded = [(part, pkg.read(part)) for part in pkg.slide_parts]
    report = build_report(loaded, pkg, manifest, tier, threshold)
    pkg.close()
    return report


def _selftest() -> int:
    ok_manifest = {"pages": [{"index": 1, "content": ["原生标题", "原生正文文字"],
                              "assets": ["assets/decor.png"]}],
                    "assets": [{"path": "assets/decor.png"}]}

    # 1. 原生页 + 齐全源，fidelity_first —— 通过，editable。
    report = _run_fixture([_native_slide()], ok_manifest, "fidelity_first")
    assert report["deck"]["verdict"] == "editable"
    assert report["deck"]["blocking_pages"] == []
    assert report["pages"][0]["source"]["complete"] is True

    # 2. 整页图：fidelity_first 下图片化不阻断（源齐全）；balanced 下 IMAGE_PAGE 阻断。
    report = _run_fixture([_image_slide()], ok_manifest, "fidelity_first")
    assert report["deck"]["blocking_pages"] == [], "成品优先档图片页不应阻断"
    report = _run_fixture([_image_slide()], ok_manifest, "balanced")
    assert report["deck"]["blocking_pages"] == [1], "兼顾档整页图应阻断"

    # 3. 无 manifest —— 两档都 SOURCE_MISSING 阻断。
    for tier in ("fidelity_first", "balanced", "compat"):
        report = _run_fixture([_native_slide()], None, tier)
        assert report["deck"]["blocking_pages"] == [1], "%s 档缺源应阻断" % tier
        assert any(f["code"] == "SOURCE_MISSING" for f in report["pages"][0]["findings"])

    # 4. SOURCE_INCOMPLETE：fidelity_first 阻断，balanced 仅告警。
    incomplete = {"pages": [{"index": 1, "content": [], "assets": []}], "assets": []}
    report = _run_fixture([_native_slide()], incomplete, "fidelity_first")
    assert report["deck"]["blocking_pages"] == [1]
    report = _run_fixture([_native_slide()], incomplete, "balanced")
    assert report["deck"]["blocking_pages"] == [], "兼顾档 INCOMPLETE 只告警"

    # 5. BAKED_EFFECT：compat 阻断，其余档告警。
    report = _run_fixture([_baked_effect_slide()], ok_manifest, "compat")
    assert report["deck"]["blocking_pages"] == [1]
    report = _run_fixture([_baked_effect_slide()], ok_manifest, "balanced")
    assert report["deck"]["blocking_pages"] == []

    # 6. 扣分权重不变：整页图 IMAGE_PAGE 扣 40（另叠图表截图提示扣 10），烘焙扣 5。
    report = _run_fixture([_image_slide()], ok_manifest, "fidelity_first")
    assert report["pages"][0]["score"] == 50
    report = _run_fixture([_baked_effect_slide()], ok_manifest, "fidelity_first")
    assert report["pages"][0]["score"] == 95

    print("selftest OK：三档阻断 / 结构化源齐全性 / 低分口径 / 扣分权重 均通过")
    return 0


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="成稿可编辑性与结构化源审计")
    ap.add_argument("deck", nargs="?", help="待检查 .pptx")
    ap.add_argument("--manifest", help="deck-manifest.json 路径；缺省取成稿同级 deck-manifest.json")
    ap.add_argument("--tier", choices=TIERS, default="fidelity_first",
                    help="交付档位（默认 fidelity_first）")
    ap.add_argument("--report-json", dest="json_out", help="机器可读报告输出路径")
    ap.add_argument("--report-md", dest="md_out", help="人读表格输出路径")
    ap.add_argument("--threshold", type=int, default=DEFAULT_THRESHOLD,
                    help="单页达标分数线（默认 70）")
    ap.add_argument("--selftest", action="store_true", help="运行内联自检")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()

    if not args.deck:
        ap.print_help()
        return 2

    manifest_path = args.manifest
    if not manifest_path:
        cand = os.path.join(os.path.dirname(os.path.abspath(args.deck)), "deck-manifest.json")
        if os.path.exists(cand):
            manifest_path = cand

    try:
        pkg = Package(args.deck)
    except CheckError as e:
        print("错误：%s" % e)
        return 2

    manifest = load_manifest(manifest_path)
    slides = [(part, pkg.read(part)) for part in pkg.slide_parts]
    report = build_report(slides, pkg, manifest, args.tier, args.threshold)
    report["deck"]["file"] = os.path.basename(args.deck)
    pkg.close()

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    if args.md_out:
        with open(args.md_out, "w", encoding="utf-8") as f:
            f.write(_md_table(report))

    print_console(report)
    return 0 if not report["deck"]["blocking_pages"] else 1


if __name__ == "__main__":
    sys.exit(main())
