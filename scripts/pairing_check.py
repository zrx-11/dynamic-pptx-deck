# -*- coding: utf-8 -*-
"""
pairing_check.py — 容器—文字配对检查器（仅标准库：zipfile + xml.etree）。

用途：逐页核对卡片 / 行底 / 胶囊等容器与其内部文字的动画编排是否成对成立：
同一点击批次、文字晚于容器 150—450ms 入场、文字只挂 fade 不跟位移、退出时
文字先消失容器后走 100—150ms；同时统计孤儿文字、节点预算、转场与悬空目标。

判定口径（详见 references/animation-orchestration.md）：
    1. 容器（有入场动画的 rect / roundRect 等自动形状）与内部文字同批绑定；
    2. 文字入场延迟 − 容器入场延迟 ∈ [min-gap, max-gap]，负值判 FAIL；
    3. 文字入场只挂 fade，出现位移类行为节点判 FAIL；
    4. 两侧都有退出时，容器退出晚于文字 100—150ms；
    5. 容器有动画而内部文字零节点，计入孤儿文字；
    6. 逐页汇总效果数、cTn 数、转场预设、advTm，并核对每个 spTgt 命中真实形状 id。

CLI：
    python3 pairing_check.py <deck.pptx> [--min-gap 150] [--max-gap 450]
            [--max-ctn 60] [--json OUT] [--md OUT]
    python3 pairing_check.py --selftest

退出码：0=全部 PASS；1=存在 FAIL；2=输入错误或包不可读。
"""
from __future__ import annotations

import argparse
import io
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

P = "{%s}" % P_NS
A = "{%s}" % A_NS

# 视为“容器”的自动形状几何：rect / roundRect / 圆角变体。椭圆、环、自由曲线不计。
CONTAINER_PRST = {
    "rect", "roundRect", "roundRect2", "snip1Rect", "snip2Rect",
    "snip2SameRect", "snip1SingleCornerRect", "snip2SingleCornerRect",
    "chamferedRect",
}

# 退出顺序：容器退出 − 文字退出的允许区间（ms）
EXIT_LEAD_MIN = 100
EXIT_LEAD_MAX = 150

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


# ----------------------------------------------------------------------
# 包读取
# ----------------------------------------------------------------------
class CheckError(RuntimeError):
    """输入不可读或不是有效 pptx。"""


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
        rid = sld.get(f"{{{R_NS}}}id")
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


def load_slides(path: str) -> list[tuple[str, bytes]]:
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        raise CheckError("无法打开包：%s" % e)
    with zf:
        parts = _slide_order(zf)
        if not parts:
            raise CheckError("包内没有幻灯片")
        return [(n, zf.read(n)) for n in parts]


# ----------------------------------------------------------------------
# 形状清单
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


def _center(box: tuple[int, int, int, int]) -> tuple[float, float]:
    x, y, cx, cy = box
    return (x + cx / 2.0, y + cy / 2.0)


def _inside(inner: tuple[int, int, int, int], outer: tuple[int, int, int, int],
            tol: int = 20000) -> bool:
    ix, iy, icx, icy = inner
    ox, oy, ocx, ocy = outer
    cx, cy = ix + icx / 2.0, iy + icy / 2.0
    return (ox - tol <= cx <= ox + ocx + tol and
            oy - tol <= cy <= oy + ocy + tol)


def collect_shapes(root: ET.Element) -> dict[str, dict]:
    """递归收集 sp / grpSp / pic / graphicFrame，返回 id -> 形状记录。"""
    sp_tree = root.find(f".//{P}spTree")
    shapes: dict[str, dict] = {}

    def walk(node: ET.Element, parent_group: str | None) -> None:
        for child in node:
            if child.tag not in (P + "sp", P + "grpSp", P + "pic", P + "graphicFrame"):
                continue
            nv = child.find(f".//{P}cNvPr")
            sid = nv.get("id") if nv is not None else None
            name = nv.get("name") if nv is not None else ""
            prst = ""
            if child.tag != P + "grpSp":
                geom = child.find(f"{P}spPr/{A}prstGeom")
                if geom is not None:
                    prst = geom.get("prst") or ""
            box = _box(child)
            text = "".join(t.text or "" for t in child.iter(A + "t")).strip()
            entry = {
                "id": sid, "name": name, "tag": child.tag.replace(P, ""),
                "prst": prst, "box": box, "text": text,
                "group": parent_group, "children": [],
            }
            shapes[sid] = entry
            if child.tag == P + "grpSp":
                entry["children"] = [
                    c.get("id") for c in child.iter() if c.tag == P + "cNvPr"
                ]
                walk(child, sid)
            else:
                walk(child, parent_group)

    if sp_tree is not None:
        walk(sp_tree, None)
    return shapes


def resolve_text_ids(shape: dict, shapes: dict[str, dict]) -> list[str]:
    """动画目标命中的文字形状 id：组节点展开到内部子形状，自身原样返回。"""
    if shape["tag"] == "grpSp" and shape["children"]:
        ids = [i for i in shape["children"] if i in shapes]
        return ids or [shape["id"]]
    return [shape["id"]]


# ----------------------------------------------------------------------
# timing 解析
# ----------------------------------------------------------------------
def _cond_delay(ctn: ET.Element) -> int | None:
    cond = ctn.find(f"{P}stCondLst/{P}cond")
    if cond is None:
        return None
    raw = cond.get("delay")
    if raw is None or raw == "indefinite":
        return None
    try:
        return int(float(raw))
    except ValueError:
        return None


def _is_displacement(ctn: ET.Element) -> bool:
    """位移类判定：ppt_x/ppt_y 位移、animMotion 路径、非 fade 的 animEffect。"""
    for anim in ctn.iter(P + "anim"):
        attr = anim.find(f".//{P}attrName")
        if attr is not None and attr.text in ("ppt_x", "ppt_y"):
            return True
    if ctn.find(f".//{P}animMotion") is not None:
        return True
    for eff in ctn.iter(P + "animEffect"):
        flt = eff.get("filter")
        if flt and flt != "fade":
            return True
    return False


def parse_timing(root: ET.Element) -> dict:
    """返回效果列表、cTn 计数、转场信息。"""
    effects: list[dict] = []
    ctn_count = len(root.findall(f".//{P}cTn"))
    timing = root.find(P + "timing")

    transition = None
    tr = root.find(P + "transition")
    if tr is not None:
        adv = tr.get("advTm")
        child = next((c for c in tr if c.tag != P + "stCondLst"), None)
        transition = {
            "kind": child.tag.split("}")[1] if child is not None else "fade",
            "advTm": int(adv) if adv and adv.isdigit() else adv,
        }

    if timing is None:
        return {"effects": effects, "ctn": ctn_count, "transition": transition}

    main_seq = timing.find(f".//{P}cTn[@nodeType='mainSeq']")
    if main_seq is None:
        return {"effects": effects, "ctn": ctn_count, "transition": transition}

    batch_lst = main_seq.find(f"{P}childTnLst")
    if batch_lst is None:
        return {"effects": effects, "ctn": ctn_count, "transition": transition}

    for batch_idx, batch_par in enumerate(batch_lst.findall(P + "par"), start=1):
        # 批次与效果之间可能套 1—2 层 par/cTn 包装；直接遍历批次内所有带
        # presetClass 的效果节点，批次号取外层 par 的序号。
        for ctn in batch_par.iter(P + "cTn"):
            cls = ctn.get("presetClass") or ""
            if cls not in ("entr", "emph", "exit", "path"):
                continue
            tgt = ctn.find(f".//{P}spTgt")
            effects.append({
                "batch": batch_idx,
                "cls": cls,
                "presetID": ctn.get("presetID") or "",
                "subtype": ctn.get("presetSubtype") or "",
                "nodeType": ctn.get("nodeType") or "",
                "spid": tgt.get("spid") if tgt is not None else None,
                "delay": _cond_delay(ctn),
                "displacement": _is_displacement(ctn),
            })
    return {"effects": effects, "ctn": ctn_count, "transition": transition}


# ----------------------------------------------------------------------
# 逐页配对判定
# ----------------------------------------------------------------------
def _effects_by_target(effects: list[dict]) -> dict[str, list[dict]]:
    by: dict[str, list[dict]] = {}
    for eff in effects:
        by.setdefault(eff["spid"], []).append(eff)
    return by


def _first(effects: list[dict], cls: str) -> dict | None:
    for eff in effects:
        if eff["cls"] == cls:
            return eff
    return None


def analyze_slide(part: str, xml: bytes, opts: dict) -> dict:
    root = ET.fromstring(xml)
    shapes = collect_shapes(root)
    timing = parse_timing(root)
    effects = timing["effects"]
    by_target = _effects_by_target(effects)

    dangling = sorted(
        {e["spid"] for e in effects if e["spid"] not in shapes and e["spid"]},
        key=lambda x: int(x) if x.isdigit() else 0,
    )

    # 轮换条：同一目标退出后又在另一批重新入场，属于“旧退出 + 新入场”的独立生命周期，
    # 不挂在卡片容器上，配对时排除，避免几何包含把条带误归为卡片文字。
    rotating = {
        spid for spid, effs in by_target.items()
        if _first(effs, "exit") is not None and
        len({e["batch"] for e in effs if e["cls"] == "entr"}) >= 2
    }

    # 容器：几何在容器集合内、且存在入场动画；极细矩形是分隔线，不计容器。
    def _is_card(shape: dict) -> bool:
        if shape["prst"] not in CONTAINER_PRST or shape["box"] is None:
            return False
        _, _, cx, cy = shape["box"]
        return cx >= 300000 and cy >= 150000

    containers = [
        s for s in shapes.values()
        if _is_card(s) and _first(by_target.get(s["id"], []), "entr") is not None
    ]

    pairs: list[dict] = []
    orphans: list[dict] = []

    for ctn_shape in containers:
        c_id = ctn_shape["id"]
        # 自身承载文字的标题 / 条带不是卡片容器，不参与配对与孤儿统计。
        if ctn_shape["text"].strip():
            continue
        c_ent = _first(by_target.get(c_id, []), "entr")
        c_exit = _first(by_target.get(c_id, []), "exit")

        # 容器内部文字形状：几何中心落在容器盒内的非空文本形状；轮换条除外。
        inner_text = [
            s for s in shapes.values()
            if s["text"] and s["box"] is not None and s["id"] != c_id
            and s["id"] not in rotating
            and _inside(s["box"], ctn_shape["box"])
        ]

        # 候选文字目标：入场动画目标，且展开后命中内部文字形状。
        text_targets: list[dict] = []
        for spid, effs in by_target.items():
            if spid in rotating or _first(effs, "entr") is None:
                continue
            target = shapes.get(spid)
            if target is None:
                continue
            resolved = set(resolve_text_ids(target, shapes))
            if any(t["id"] in resolved for t in inner_text):
                text_targets.append(target)

        if not text_targets:
            orphans.append({
                "container": ctn_shape["name"], "container_id": c_id,
                "text_shapes": [t["name"] for t in inner_text],
            })
            continue

        for tgt in text_targets:
            t_id = tgt["id"]
            t_effs = by_target.get(t_id, [])
            t_ent = _first(t_effs, "entr")
            t_exit = _first(t_effs, "exit")

            checks: dict[str, str] = {}
            issues: list[str] = []

            # ① 同批
            same_batch = c_ent["batch"] == t_ent["batch"]
            checks["same_batch"] = "PASS" if same_batch else "FAIL"
            if not same_batch:
                issues.append("文字与容器不在同一点击批次（容器批 %d / 文字批 %d）"
                               % (c_ent["batch"], t_ent["batch"]))

            # ② 延迟区间
            gap = None
            if c_ent["delay"] is not None and t_ent["delay"] is not None:
                gap = t_ent["delay"] - c_ent["delay"]
                if gap < 0:
                    checks["gap"] = "FAIL"
                    issues.append("文字入场早于容器 %dms（负差）" % -gap)
                elif opts["min_gap"] <= gap <= opts["max_gap"]:
                    checks["gap"] = "PASS"
                else:
                    checks["gap"] = "FAIL"
                    issues.append("文字晚于容器 %dms，超出 [%d, %d]ms"
                                  % (gap, opts["min_gap"], opts["max_gap"]))
            else:
                checks["gap"] = "UNKNOWN"

            # ③ 文字只挂 fade
            fade_only = not t_ent["displacement"]
            checks["fade_only"] = "PASS" if fade_only else "FAIL"
            if not fade_only:
                issues.append("文字入场含位移类行为节点，不应跟容器一起位移")

            # ④ 退出顺序
            if c_exit is not None and t_exit is not None:
                lead = (c_exit["delay"] if c_exit["delay"] is not None else 0) - \
                       (t_exit["delay"] if t_exit["delay"] is not None else 0)
                if EXIT_LEAD_MIN <= lead <= EXIT_LEAD_MAX:
                    checks["exit_order"] = "PASS"
                else:
                    checks["exit_order"] = "FAIL"
                    issues.append("容器退出未晚于文字 100—150ms（差值 %sms）" % lead)
            elif c_exit is not None or t_exit is not None:
                checks["exit_order"] = "WARN"
                issues.append("退出只挂在单侧（容器退出 %s / 文字退出 %s）"
                              % ("有" if c_exit else "无", "有" if t_exit else "无"))
            else:
                checks["exit_order"] = "PASS"

            verdict = "PASS" if "FAIL" not in checks.values() else "FAIL"
            pairs.append({
                "container": ctn_shape["name"], "container_id": c_id,
                "text": tgt["name"], "text_id": t_id,
                "container_batch": c_ent["batch"],
                "text_batch": t_ent["batch"],
                "container_delay": c_ent["delay"],
                "text_delay": t_ent["delay"],
                "gap": gap,
                "container_exit_delay": c_exit["delay"] if c_exit else None,
                "text_exit_delay": t_exit["delay"] if t_exit else None,
                "checks": checks, "issues": issues, "verdict": verdict,
            })

    # 组节点跨批：同一文字组覆盖的容器落在不同批次，单点延迟无法逐张错开。
    group_batches: dict[str, set[int]] = {}
    for pair in pairs:
        group_batches.setdefault(pair["text_id"], set()).add(pair["container_batch"])
    for gid, batches in group_batches.items():
        if len(batches) > 1:
            for pair in pairs:
                if pair["text_id"] == gid:
                    pair["checks"]["shared_group"] = "FAIL"
                    pair["verdict"] = "FAIL"
                    pair["issues"].append("同一文字组覆盖不同批次的容器，应按批拆组")

    # ⑥ 节点预算
    ctn_over = timing["ctn"] > opts["max_ctn"]
    if ctn_over:
        issues_budget = "cTn 节点 %d 超出预算 %d" % (timing["ctn"], opts["max_ctn"])
    else:
        issues_budget = ""

    fail_pairs = [p for p in pairs if p["verdict"] == "FAIL"]
    slide_fail = bool(fail_pairs or dangling or ctn_over)
    if orphans:
        # 孤儿文字计为提示性缺陷：影响观感，不单独判 FAIL，计入汇总。
        pass

    return {
        "part": part,
        "shapes": len(shapes),
        "effects": len(effects),
        "ctn": timing["ctn"],
        "ctn_budget": opts["max_ctn"],
        "transition": timing["transition"],
        "dangling_targets": dangling,
        "pairs": pairs,
        "orphan_containers": orphans,
        "orphan_count": len(orphans),
        "issues": ([issues_budget] if ctn_over else []) +
                  ["悬空目标 id：%s" % d for d in dangling],
        "verdict": "FAIL" if slide_fail else "PASS",
    }


# ----------------------------------------------------------------------
# 报告输出
# ----------------------------------------------------------------------
def build_report(slides: list[tuple[str, bytes]], opts: dict) -> dict:
    page_reports = [analyze_slide(part, xml, opts) for part, xml in slides]
    totals = {
        "pairs": sum(len(p["pairs"]) for p in page_reports),
        "orphans": sum(p["orphan_count"] for p in page_reports),
        "fail_pages": sum(1 for p in page_reports if p["verdict"] == "FAIL"),
        "effects": sum(p["effects"] for p in page_reports),
        "ctn": sum(p["ctn"] for p in page_reports),
    }
    verdict = "PASS" if totals["fail_pages"] == 0 else "FAIL"
    return {"options": opts, "slides": page_reports, "totals": totals,
            "verdict": verdict}


def _md_table(report: dict) -> str:
    lines = ["# 容器—文字配对检查报告", ""]
    lines.append("- 总体判定：**%s**" % report["verdict"])
    t = report["totals"]
    lines.append("- 配对 %d 对；孤儿容器 %d；FAIL 页 %d；效果节点 %d；cTn %d"
                 % (t["pairs"], t["orphans"], t["fail_pages"], t["effects"], t["ctn"]))
    lines.append("")
    for page in report["slides"]:
        lines.append("## 第 %d 页（%s）" % (report["slides"].index(page) + 1, page["part"]))
        tr = page["transition"]
        tr_desc = "无" if tr is None else "%s（advTm=%s）" % (
            tr["kind"], tr["advTm"] if tr["advTm"] is not None else "点击")
        lines.append("- 形状 %d；效果 %d；cTn %d；转场 %s"
                     % (page["shapes"], page["effects"], page["ctn"], tr_desc))
        if page["dangling_targets"]:
            lines.append("- 悬空目标：%s" % ", ".join(page["dangling_targets"]))
        if not page["pairs"]:
            lines.append("- 本页无容器—文字配对。")
        for pair in page["pairs"]:
            lines.append("")
            lines.append("| 容器 | 文字 | 批次(容/文) | 延迟(容/文) | 差值 ms | 判定 |")
            lines.append("|---|---|---|---|---|---|")
            lines.append("| %s | %s | %s / %s | %s / %s | %s | %s |" % (
                pair["container"], pair["text"], pair["container_batch"],
                pair["text_batch"], pair["container_delay"], pair["text_delay"],
                pair["gap"] if pair["gap"] is not None else "—", pair["verdict"]))
            for issue in pair["issues"]:
                lines.append("  - %s" % issue)
        if page["orphan_containers"]:
            lines.append("- 孤儿容器：%s" % ", ".join(
                o["container"] for o in page["orphan_containers"]))
        lines.append("")
    return "\n".join(lines)


def print_console(report: dict) -> None:
    for i, page in enumerate(report["slides"], start=1):
        tr = page["transition"]
        tr_desc = "-" if tr is None else tr["kind"]
        print("第 %2d 页 [%s] 形状 %d 效果 %d cTn %d 转场 %s 配对 %d 孤儿 %d"
              % (i, page["verdict"], page["shapes"], page["effects"], page["ctn"],
                 tr_desc, len(page["pairs"]), page["orphan_count"]))
        for pair in page["pairs"]:
            if pair["verdict"] == "FAIL":
                print("  FAIL %s ↔ %s：%s" % (
                    pair["container"], pair["text"], "；".join(pair["issues"])))
        for issue in page["issues"]:
            print("  问题：%s" % issue)
    t = report["totals"]
    print("合计：配对 %d，孤儿 %d，FAIL 页 %d，效果 %d，cTn %d —— %s"
          % (t["pairs"], t["orphans"], t["fail_pages"], t["effects"], t["ctn"],
             report["verdict"]))


# ----------------------------------------------------------------------
# 自检 fixture（内联最小包，不依赖外部 pptx）
# ----------------------------------------------------------------------
def _fixture_slide(shape_xml: str, timing_xml: str) -> bytes:
    return XML_DECL + (
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm/></p:grpSpPr>%s'
        '</p:spTree></p:cSld>%s</p:sld>'
        % (A_NS, R_NS, P_NS, shape_xml, timing_xml)
    ).encode("utf-8")


def _shape_xml(sid: str, name: str, prst: str, box: tuple[int, int, int, int],
               text: str) -> str:
    x, y, cx, cy = box
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>' % (sid, name, x, y, cx, cy, prst, text)
    )


def _effect_xml(ctn_id: int, spid: str, cls: str, delay: int | None,
                node_type: str, displacement: bool = False,
                preset: str = "10", subtype: str = "0") -> str:
    delay_raw = "indefinite" if delay is None else str(delay)
    eff_body = [
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>' % (ctn_id, spid),
    ]
    if displacement:
        eff_body.append(
            '<p:anim><p:cBhvr><p:cTn id="%d" dur="500"/>'
            '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
            '</p:cBhvr></p:anim>' % (ctn_id + 1, spid))
    else:
        eff_body.append(
            '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
            '<p:cTn id="%d" dur="500"/><p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
            '</p:cBhvr></p:animEffect>' % (ctn_id + 1, spid))
    return (
        '<p:par><p:cTn id="%d" presetID="%s" presetClass="%s" presetSubtype="%s" '
        'fill="hold" grpId="0" nodeType="%s"><p:stCondLst>'
        '<p:cond delay="%s"/></p:stCondLst><p:childTnLst>%s'
        '</p:childTnLst></p:cTn></p:par>'
        % (ctn_id + 2, preset, cls, subtype, node_type, delay_raw, "".join(eff_body))
    )


def _timing_xml(batches: list[list[dict]]) -> str:
    """batches: 每批一组 effect 描述；effect: spid/cls/delay/nodeType/disp/preset/subtype。"""
    ctn = 3
    batch_xml: list[str] = []
    for batch in batches:
        eff_xml: list[str] = []
        for eff in batch:
            eff_xml.append(_effect_xml(
                ctn, eff["spid"], eff["cls"], eff["delay"], eff["nodeType"],
                displacement=eff.get("disp", False), preset=eff.get("preset", "10"),
                subtype=eff.get("subtype", "0")))
            ctn += 4
        batch_xml.append(
            '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
            '<p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>%s'
            '</p:childTnLst></p:cTn></p:par>' % (ctn, "".join(eff_xml)))
        ctn += 1
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>%s'
        '</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        '</p:timing>' % "".join(batch_xml)
    )


def _fixture_zip(slide_xml: bytes) -> io.BytesIO:
    buf = io.BytesIO()
    ct = (
        '<Types xmlns="%s"><Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '</Types>' % REL_NS)
    rels = (
        '<Relationships xmlns="%s"><Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/></Relationships>' % REL_NS)
    pres = (
        '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
        '</p:presentation>' % (A_NS, R_NS, P_NS))
    pres_rels = (
        '<Relationships xmlns="%s"><Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide1.xml"/></Relationships>' % REL_NS)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", XML_DECL + ct.encode())
        zf.writestr("_rels/.rels", XML_DECL + rels.encode())
        zf.writestr("ppt/presentation.xml", XML_DECL + pres.encode())
        zf.writestr("ppt/_rels/presentation.xml.rels", XML_DECL + pres_rels.encode())
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
    buf.seek(0)
    return buf


def _run_fixture(shape_xml: str, batches: list[list[dict]],
                 opts: dict | None = None) -> dict:
    slide = _fixture_slide(shape_xml, _timing_xml(batches))
    buf = _fixture_zip(slide)
    with zipfile.ZipFile(buf) as zf:
        slides = [("ppt/slides/slide1.xml",
                    zf.read("ppt/slides/slide1.xml"))]
    return build_report(slides, opts or {"min_gap": 150, "max_gap": 450, "max_ctn": 60})


def _selftest() -> int:
    box_c = (1000000, 2000000, 3000000, 2000000)
    box_t = (1100000, 2100000, 2800000, 600000)
    shapes_ok = (
        _shape_xml("10", "card0", "roundRect", box_c, "") +
        _shape_xml("11", "label0", "rect", box_t, "文字")
    )

    # 用例 1：同批、fade、延迟差 300ms —— 应 PASS
    report = _run_fixture(shapes_ok, [[
        {"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
         "preset": "53", "subtype": "16"},
        {"spid": "11", "cls": "entr", "delay": 300, "nodeType": "withEffect"},
    ]])
    page = report["slides"][0]
    assert report["verdict"] == "PASS", "合格配对不应判 FAIL：%s" % page["issues"]
    assert len(page["pairs"]) == 1 and page["pairs"][0]["verdict"] == "PASS"
    assert page["pairs"][0]["gap"] == 300

    # 用例 2：文字落到第二批 —— 同批判定 FAIL
    report = _run_fixture(shapes_ok, [
        [{"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
          "preset": "53", "subtype": "16"}],
        [{"spid": "11", "cls": "entr", "delay": 0, "nodeType": "clickEffect"}],
    ])
    pair = report["slides"][0]["pairs"][0]
    assert pair["checks"]["same_batch"] == "FAIL", "跨批应判同批 FAIL"
    assert report["verdict"] == "FAIL"

    # 用例 3：负延迟差 —— gap 判定 FAIL
    report = _run_fixture(shapes_ok, [[
        {"spid": "10", "cls": "entr", "delay": 300, "nodeType": "clickEffect",
         "preset": "53", "subtype": "16"},
        {"spid": "11", "cls": "entr", "delay": 100, "nodeType": "withEffect"},
    ]])
    pair = report["slides"][0]["pairs"][0]
    assert pair["checks"]["gap"] == "FAIL" and pair["gap"] == -200

    # 用例 4：文字跟位移 —— fade_only 判定 FAIL
    report = _run_fixture(shapes_ok, [[
        {"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
         "preset": "53", "subtype": "16"},
        {"spid": "11", "cls": "entr", "delay": 300, "nodeType": "withEffect",
         "disp": True},
    ]])
    pair = report["slides"][0]["pairs"][0]
    assert pair["checks"]["fade_only"] == "FAIL"

    # 用例 5：容器有动画、文字零节点 —— 计入孤儿
    report = _run_fixture(shapes_ok, [[
        {"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
         "preset": "53", "subtype": "16"},
    ]])
    assert report["slides"][0]["orphan_count"] == 1, "应识别孤儿容器"

    # 用例 6：悬空 spTgt —— 计入 dangling_targets 并判 FAIL
    report = _run_fixture(shapes_ok, [[
        {"spid": "999", "cls": "entr", "delay": 0, "nodeType": "clickEffect"},
    ]])
    assert report["slides"][0]["dangling_targets"] == ["999"]
    assert report["verdict"] == "FAIL"

    # 用例 7：退出顺序 —— 文字晚于容器应 FAIL；容器晚 120ms 应 PASS
    shapes_exit = shapes_ok
    report = _run_fixture(shapes_exit, [
        [{"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
          "preset": "53", "subtype": "16"},
         {"spid": "11", "cls": "entr", "delay": 300, "nodeType": "withEffect"}],
        [{"spid": "10", "cls": "exit", "delay": 1000, "nodeType": "clickEffect",
          "preset": "10", "subtype": "0"},
         {"spid": "11", "cls": "exit", "delay": 1200, "nodeType": "withEffect"}],
    ])
    pair = report["slides"][0]["pairs"][0]
    assert pair["checks"]["exit_order"] == "FAIL", "文字先消失才对，反向应 FAIL"

    report = _run_fixture(shapes_exit, [
        [{"spid": "10", "cls": "entr", "delay": 0, "nodeType": "clickEffect",
          "preset": "53", "subtype": "16"},
         {"spid": "11", "cls": "entr", "delay": 300, "nodeType": "withEffect"}],
        [{"spid": "10", "cls": "exit", "delay": 1200, "nodeType": "clickEffect"},
         {"spid": "11", "cls": "exit", "delay": 1080, "nodeType": "withEffect"}],
    ])
    pair = report["slides"][0]["pairs"][0]
    assert pair["checks"]["exit_order"] == "PASS"

    print("selftest OK：同批 / 延迟差 / fade-only / 退出顺序 / 孤儿容器 / 悬空目标 六类判定均通过")
    return 0


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="容器—文字配对检查器")
    ap.add_argument("deck", nargs="?", help="待检查 .pptx")
    ap.add_argument("--min-gap", type=int, default=150, help="文字晚于容器的最小 ms")
    ap.add_argument("--max-gap", type=int, default=450, help="文字晚于容器的最大 ms")
    ap.add_argument("--max-ctn", type=int, default=60, help="单页 cTn 节点预算")
    ap.add_argument("--json", dest="json_out", help="机器可读报告输出路径")
    ap.add_argument("--md", dest="md_out", help="人读表格输出路径")
    ap.add_argument("--selftest", action="store_true", help="运行内联自检")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()

    if not args.deck:
        ap.print_help()
        return 2

    opts = {"min_gap": args.min_gap, "max_gap": args.max_gap,
            "max_ctn": args.max_ctn}
    try:
        slides = load_slides(args.deck)
    except CheckError as e:
        print("错误：%s" % e)
        return 2

    report = build_report(slides, opts)

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    if args.md_out:
        with open(args.md_out, "w", encoding="utf-8") as f:
            f.write(_md_table(report))

    print_console(report)
    return 0 if report["verdict"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
