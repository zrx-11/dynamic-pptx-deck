# -*- coding: utf-8 -*-
"""
audit_all.py — 合并校验入口：结构体检、容器—文字配对、分档可编辑性、
schema 深扫与页内动效可见性配额，一次调用出一份合并报告与一个退出码。

既有能力以模块方式复用：
    - 结构体检：``dynamic_deck.validate``
    - 配对判定：``pairing_check.load_slides`` / ``build_report``
    - 可编辑性：``editability_check.Package`` / ``build_report``

CLI：
    python3 audit_all.py <deck.pptx> [--manifest P]
            [--tier fidelity_first|balanced|compat] [--fast]
            [--json OUT] [--md OUT]
    python3 audit_all.py --selftest

选项：
    --manifest   结构化源 deck-manifest.json 路径；缺省取成稿同级同名文件。
    --tier       交付档位，默认 fidelity_first；可编辑性阻断按档位解析。
    --fast       跳过全包 schema 深扫（rels / ContentTypes / 媒体与部件
                 完整性），只做结构关键项与动效配额。
    --json OUT   合并报告写机器可读 JSON。
    --md OUT     人读逐页配额表写 Markdown。
    --require-morph-pages  逗号分隔页号：这些页必须挂 2015/09 morph
                           （manifest 声明 transition=morph 的页同样计入）。

后处理三条检查（fast / 全量均跑）：morph-missing（缺平滑切换）、
image-placeholder（图片占位未填）、letters-count（单字母数与页声明不符）。

退出码：
    0  全部检查通过；
    1  任一检查失败（含配额不足、配对 / 可编辑性按档位阻断、非 fast 时
       schema 有硬伤）；
    2  输入不可读或参数错误。
"""
from __future__ import annotations

import argparse
import io
import json
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import dynamic_deck as dd          # noqa: E402
import pairing_check as pc         # noqa: E402
import editability_check as ec     # noqa: E402

P_NS = pc.P_NS
A_NS = pc.A_NS
R_NS = pc.R_NS
REL_NS = pc.REL_NS
P = "{%s}" % P_NS
A = "{%s}" % A_NS
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# 配额口径（写死）：可感知运动对象 ≥3、持续循环 ≥1、效果总数 ≤12、cTn ≤60。
QUOTA = {"min_perceptible": 3, "min_loops": 1, "max_effects": 12, "max_ctn": 60}

# 位置 / 形态类预设族：fly(2) / wipe(22) / path(42) / zoom(53,31) /
# bounce(37) / drop(47) / spin(8) / grow(26,6) / zoom_out(23)。
MOTION_PRESET_IDS = {"2", "22", "42", "53", "31", "37", "47", "8",
                     "26", "6", "23"}
# 纯透明度 / 出现类预设：fade(10) / appear(1) / blink(35)。
NON_MOTION_PRESET_IDS = {"10", "1", "35"}
# 行为节点属性名：位移 / 缩放 / 旋转计入可感知；visibility / opacity 不计。
MOTION_ATTRS = {"ppt_x", "ppt_y", "ppt_w", "ppt_h", "style.rotation"}
NON_MOTION_ATTRS = {"style.visibility", "style.opacity"}

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


class AuditInputError(RuntimeError):
    """输入不可读或不是有效 pptx。"""


# ----------------------------------------------------------------------
# 结构体检（复用 dynamic_deck.validate）
# ----------------------------------------------------------------------
def _run_structure(deck) -> dict:
    try:
        rows = dd.validate(deck, verbose=False)
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError,
            ValueError, dd.etree.XMLSyntaxError) as e:
        return {"exit": 2, "issues": ["结构体检读取失败: %s" % e],
                "effects": 0, "ctn": 0, "pages": {}}
    issues: list[str] = []
    effects = 0
    ctn = 0
    page_info: dict[int, dict] = {}
    for name, info, problems in rows:
        effects += info["effects"]
        ctn += info["ctn"]
        m = re.search(r"\d+", name)
        idx = int(m.group()) if m else 0
        page_info[idx] = {"effects": info["effects"], "ctn": info["ctn"]}
        for prob in problems:
            issues.append("第 %d 页: %s" % (idx, prob))
    return {"exit": 0 if not issues else 1, "issues": issues,
            "effects": effects, "ctn": ctn, "pages": page_info}


# ----------------------------------------------------------------------
# 配对检查（复用 pairing_check）
# ----------------------------------------------------------------------
def _run_pairing(slides: list[tuple[str, bytes]]) -> dict:
    opts = {"min_gap": 150, "max_gap": 450, "max_ctn": QUOTA["max_ctn"]}
    report = pc.build_report(slides, opts)
    totals = report["totals"]
    morph_pairs = [
        {"name": m["name"], "type": m["type"], "pages": m["pages"]}
        for m in report["paired_objects"]
    ]
    return {
        "exit": 0 if report["verdict"] == "PASS" else 1,
        "fail": totals["fail_pages"],
        "orphan": totals["orphans"],
        "morph_pairs": morph_pairs,
        "conflict": totals["paired_conflicts"],
        "report": report,
    }


# ----------------------------------------------------------------------
# 可编辑性（复用 editability_check）
# ----------------------------------------------------------------------
def _run_editability(slides: list[tuple[str, bytes]], pkg: ec.Package,
                     manifest: dict | None, tier: str) -> dict:
    report = ec.build_report(slides, pkg, manifest, tier,
                             ec.DEFAULT_THRESHOLD)
    blocking = report["deck"]["blocking_pages"]
    return {
        "exit": 0 if not blocking else 1,
        "score": report["deck"]["score"],
        "blocking_pages": blocking,
        "by_code": report["summary"]["by_code"],
        "report": report,
    }


# ----------------------------------------------------------------------
# schema 深扫（非 fast）：部件与关系齐备、ContentTypes 覆盖、slide rels
# 指向存在、媒体非零字节、timing 内 spTgt 不悬空。
# ----------------------------------------------------------------------
def _rels_base(part: str) -> str:
    """关系部件所在基准目录：``ppt/slides/_rels/x.rels`` -> ``ppt/slides/``；
    根级 ``_rels/.rels`` -> ``""``。"""
    directory = posixpath.dirname(part)
    if posixpath.basename(directory) == "_rels":
        directory = posixpath.dirname(directory)
    return directory + "/" if directory else ""


def _resolve_target(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(base + target).lstrip("/")


def _schema_scan(zf: zipfile.ZipFile) -> list[str]:
    issues: list[str] = []
    names = set(zf.namelist())

    if "[Content_Types].xml" not in names:
        return ["缺少 [Content_Types].xml"]
    try:
        ct = ET.fromstring(zf.read("[Content_Types].xml"))
    except ET.ParseError as e:
        return ["ContentTypes 解析失败: %s" % e]

    defaults = {d.get("Extension", "").lower()
                for d in ct.findall("{%s}Default" % CT_NS)}
    overrides = {o.get("PartName", "").lstrip("/")
                 for o in ct.findall("{%s}Override" % CT_NS)}

    # ① 部件被 ContentTypes 覆盖。
    for part in names:
        if part.endswith("/") or part.endswith(".rels"):
            continue
        if part == "[Content_Types].xml":
            continue
        ext = part.rsplit(".", 1)[-1].lower() if "." in part else ""
        if part not in overrides and ext not in defaults:
            issues.append("部件未被 ContentTypes 覆盖: %s" % part)

    # ② 关系目标齐备；③ 媒体部件存在且非零字节。
    for part in sorted(names):
        if not part.endswith(".rels"):
            continue
        try:
            rel_root = ET.fromstring(zf.read(part))
        except ET.ParseError as e:
            issues.append("关系部件解析失败 %s: %s" % (part, e))
            continue
        base = _rels_base(part)
        for rel in rel_root:
            mode = rel.get("TargetMode")
            target = rel.get("Target") or ""
            if mode == "External" or re.match(r"^(https?|mailto|ftp):", target):
                continue
            resolved = _resolve_target(base, target)
            if resolved not in names:
                issues.append("关系指向缺失部件: %s -> %s" % (part, target))
                continue
            if resolved.startswith("ppt/media/"):
                if zf.getinfo(resolved).file_size == 0:
                    issues.append("媒体部件为空: %s" % resolved)

    # ④ 全包媒体部件非零字节（含未被关系引用的残留媒体）。
    for part in sorted(names):
        if part.startswith("ppt/media/") and zf.getinfo(part).file_size == 0:
            issues.append("媒体部件为空: %s" % part)

    # ⑤ timing 树内 spTgt 不悬空。
    for part in sorted(names):
        if not re.match(r"ppt/slides/slide\d+\.xml$", part):
            continue
        try:
            root = ET.fromstring(zf.read(part))
        except ET.ParseError as e:
            issues.append("幻灯片解析失败 %s: %s" % (part, e))
            continue
        shape_ids = {n.get("id") for n in root.iter(P + "cNvPr") if n.get("id")}
        for tgt in root.iter(P + "spTgt"):
            spid = tgt.get("spid")
            if spid and spid not in shape_ids:
                issues.append("spTgt 悬空: %s -> 形状 %s" % (part, spid))

    return issues


# ----------------------------------------------------------------------
# 动效可见性与配额（新能力）
# ----------------------------------------------------------------------
def _behavior_is_motion(node: ET.Element) -> bool:
    """单个行为节点是否属于位移 / 缩放 / 旋转 / 擦除 / 路径 / fly / zoom。"""
    tag = node.tag.split("}")[1]
    if tag in ("animMotion", "animRot", "animScale"):
        return True
    if tag == "animEffect":
        filt = (node.get("filter") or "").lower()
        if filt.startswith("fade") or filt.startswith("appear"):
            return False
        return True
    if tag == "anim":
        attr = node.find(".//" + P + "attrName")
        name = (attr.text or "") if attr is not None else ""
        return name in MOTION_ATTRS
    return False


def _effect_loop(ctn: ET.Element) -> bool:
    """效果子树内是否存在无限重复：repeatCount=indefinite，或行为级
    cond delay=indefinite，或 <p:anim> 无限迭代。外层批次点击门控的
    indefinite cond 不在效果子树内，不计入。"""
    if ctn.get("repeatCount") == "indefinite":
        return True
    for inner in ctn.iter(P + "cTn"):
        if inner.get("repeatCount") == "indefinite":
            return True
    for cond in ctn.iter(P + "cond"):
        if cond.get("delay") == "indefinite":
            return True
    return False


def _effect_target(ctn: ET.Element) -> str | None:
    tgt = ctn.find(".//" + P + "spTgt")
    return tgt.get("spid") if tgt is not None else None


def _analyze_motion(part: str, xml: bytes) -> dict:
    root = ET.fromstring(xml)
    effects = [c for c in root.iter(P + "cTn") if c.get("presetClass")]

    kinds = {"entr": 0, "emph": 0, "exit": 0, "path": 0}
    perceptible_targets: set[str] = set()
    loops = 0

    for ctn in effects:
        cls = ctn.get("presetClass") or ""
        pid = ctn.get("presetID") or ""
        spid = _effect_target(ctn)
        behaviors = [
            b for b in ctn.iter()
            if b.tag.split("}")[1] in
            ("animEffect", "anim", "animMotion", "animRot", "animScale", "set")
        ]
        behavior_motion = any(_behavior_is_motion(b) for b in behaviors)
        preset_motion = pid in MOTION_PRESET_IDS and pid not in NON_MOTION_PRESET_IDS
        motion = behavior_motion or preset_motion

        if cls == "exit":
            kinds["exit"] += 1
        elif cls == "path":
            kinds["path"] += 1
        elif cls == "emph":
            kinds["emph"] += 1
        else:
            kinds["entr"] += 1
        if any(b.tag.split("}")[1] == "animMotion" for b in behaviors):
            kinds["path"] += 1 if cls != "path" else 0

        # 退出类不计入可感知运动对象。
        if cls != "exit" and motion and spid:
            perceptible_targets.add(spid)
        if _effect_loop(ctn):
            loops += 1

    ctn_count = len(list(root.iter(P + "cTn")))
    effects_count = len(effects)
    perceptible = len(perceptible_targets)

    missing: list[str] = []
    if perceptible < QUOTA["min_perceptible"]:
        missing.append("可感知运动对象仅 %d 个（需 ≥%d）"
                       % (perceptible, QUOTA["min_perceptible"]))
    if loops < QUOTA["min_loops"]:
        missing.append("缺持续循环（需 ≥%d 个无限重复动效）" % QUOTA["min_loops"])
    if effects_count > QUOTA["max_effects"]:
        missing.append("效果总数 %d 超出预算 %d"
                       % (effects_count, QUOTA["max_effects"]))
    if ctn_count > QUOTA["max_ctn"]:
        missing.append("cTn 节点 %d 超出预算 %d" % (ctn_count, QUOTA["max_ctn"]))

    return {
        "effects": effects_count,
        "ctn": ctn_count,
        "kinds": kinds,
        "perceptible": perceptible,
        "loops": loops,
        "quota": {"ok": not missing, "missing": missing},
    }


def _manifest_page_type(manifest: dict | None, index: int) -> str:
    if not manifest:
        return "unknown"
    for entry in manifest.get("pages", []) or []:
        if isinstance(entry, dict):
            try:
                if int(entry.get("index", -1)) == index:
                    return entry.get("page_type") or "unknown"
            except (TypeError, ValueError):
                continue
    return "unknown"


def _collect_motion_pages(slides: list[tuple[str, bytes]],
                          pairing_report: dict,
                          manifest: dict | None) -> tuple[list[dict], list[int]]:
    pair_pages = {i + 1: p for i, p in enumerate(pairing_report["slides"])}
    pages: list[dict] = []
    below: list[int] = []

    for idx, (part, xml) in enumerate(slides, start=1):
        motion = _analyze_motion(part, xml)
        findings: list[dict] = []

        # Morph 配对对象挂页内位置类动画：复用配对检查的跨页判定结果。
        pair_page = pair_pages.get(idx, {})
        for paired in pair_page.get("paired_objects", []):
            if paired.get("status") == "fail":
                findings.append({
                    "code": "MORPH_MOTION_CONFLICT",
                    "object": paired.get("name"),
                    "issues": paired.get("issues", []),
                })

        if not motion["quota"]["ok"]:
            for msg in motion["quota"]["missing"]:
                findings.append({"code": "QUOTA_MISSING", "message": msg})
            below.append(idx)

        page = {
            "index": idx,
            "page_type": _manifest_page_type(manifest, idx),
            "effects": motion["effects"],
            "ctn": motion["ctn"],
            "kinds": motion["kinds"],
            "perceptible": motion["perceptible"],
            "loops": motion["loops"],
            "quota": motion["quota"],
            "findings": findings,
        }
        pages.append(page)

    return pages, below


# ----------------------------------------------------------------------
# 成品后处理三条检查（移植自 postfix_deck.py，python-pptx/EMU 路线重写）
#   morph-missing      指定页缺 2015/09 命名空间的 p159:morph
#   image-placeholder  文本命中图片占位词且形状 / 所在组合内无 a:blip
#   letters-count      单字母对象数与页声明 letters_count 不符（出画则跳过）
# ----------------------------------------------------------------------
P159_NS = pc.P159_NS
PLACEHOLDER_WORDS = ("添加图片", "插入图片", "图片占位", "点击添加图片")


def _manifest_page_field(manifest: dict | None, index: int, field: str,
                         default=None):
    if not manifest:
        return default
    for entry in manifest.get("pages", []) or []:
        if isinstance(entry, dict):
            try:
                if int(entry.get("index", -1)) == index:
                    return entry.get(field, default)
            except (TypeError, ValueError):
                continue
    return default


def _localname(tag: str) -> str:
    return tag.split("}")[1] if "}" in tag else tag


def _group_state(xfrm):
    off = xfrm.find(A + "off")
    ext = xfrm.find(A + "ext")
    if off is None or ext is None:
        return None
    chOff = xfrm.find(A + "chOff")
    chExt = xfrm.find(A + "chExt")
    return (int(off.get("y") or 0), int(ext.get("cy") or 1),
            int(chOff.get("y") or 0) if chOff is not None else 0,
            int(chExt.get("cy") or 1) if chExt is not None else 1)


def _abs_top(el, grp_state) -> int | None:
    xfrm = el.find(".//" + A + "xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(A + "off")
    if off is None:
        return None
    y = int(off.get("y") or 0)
    if grp_state is None:
        return y
    off_y, ext_cy, ch_off_y, ch_ext_cy = grp_state
    if ch_ext_cy:
        return int(off_y + (y - ch_off_y) * ext_cy / ch_ext_cy)
    return off_y + y


def _walk_shapes(root, grp_state=None, group_blip: bool = False):
    """递归产出 (元素, 绝对 top EMU, 所在组合是否含 blip)，含组合内部。"""
    out: list[tuple[ET.Element, int | None, bool]] = []
    for child in list(root):
        tag = _localname(child.tag)
        if tag in ("cSld", "spTree", "grpSp"):
            xfrm = child.find(A + "xfrm")
            state = _group_state(xfrm) if xfrm is not None else grp_state
            gblip = group_blip or child.find(".//" + A + "blip") is not None
            out.extend(_walk_shapes(child, state, gblip))
        elif tag in ("sp", "pic", "graphicFrame"):
            out.append((child, _abs_top(child, grp_state), group_blip))
    return out


def _shape_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(A + "t"))


def _check_image_placeholder(root: ET.Element) -> list[str]:
    hits: list[str] = []
    for el, _top, group_blip in _walk_shapes(root):
        if _localname(el.tag) != "sp":
            continue
        text = _shape_text(el).strip()
        if not text or not any(w in text for w in PLACEHOLDER_WORDS):
            continue
        if el.find(".//" + A + "blip") is None and not group_blip:
            hits.append(text[:20])
    return hits


def _single_letters(root: ET.Element) -> list[tuple[str, int | None]]:
    letters: list[tuple[str, int | None]] = []
    for el, top, _ in _walk_shapes(root):
        if _localname(el.tag) != "sp":
            continue
        text = _shape_text(el).strip()
        if len(text) == 1 and text.isascii() and text.isalpha():
            letters.append((text, top))
    return letters


def _run_postfix(slides: list[tuple[str, bytes]], manifest: dict | None,
                 require_morph_pages: list[int] | None) -> dict:
    morph_missing: list[int] = []
    image_placeholder: list[int] = []
    letters_bad: list[dict] = []
    letters_skipped: list[int] = []
    required = set(require_morph_pages or [])

    for idx, (_part, xml) in enumerate(slides, start=1):
        root = ET.fromstring(xml)

        # ① morph-missing：manifest 声明 transition=morph 或显式指定的页，
        #    必须挂 2015/09 命名空间的 p159:morph；fade / 旧命名空间算缺失。
        declared = _manifest_page_field(manifest, idx, "transition")
        if declared == "morph" or idx in required:
            if root.find(".//{%s}morph" % P159_NS) is None:
                morph_missing.append(idx)

        # ② image-placeholder：占位词文本且形状 / 组合内无 blip。
        if _check_image_placeholder(root):
            image_placeholder.append(idx)

        # ③ letters-count：无声明记 skipped；全部出画（top<0）视为布局，跳过。
        expected = _manifest_page_field(manifest, idx, "letters_count")
        letters = _single_letters(root)
        if expected is None:
            letters_skipped.append(idx)
            continue
        # 所有单字母对象都移出画布（top<0）视为布局本身，跳过；部分出画不豁免。
        if letters and all(top is not None and top < 0
                           for _ch, top in letters):
            letters_skipped.append(idx)
            continue
        if len(letters) != int(expected):
            letters_bad.append({"page": idx, "expected": int(expected),
                                "actual": len(letters)})

    exit_code = 0 if not (morph_missing or image_placeholder or letters_bad) else 1
    return {"exit": exit_code, "morph_missing": morph_missing,
            "image_placeholder": image_placeholder,
            "letters_count": letters_bad, "letters_skipped": letters_skipped}


# ----------------------------------------------------------------------
# 合并报告
# ----------------------------------------------------------------------
def run_audit(deck, manifest=None, tier: str = "fidelity_first",
              fast: bool = False,
              require_morph_pages: list[int] | None = None) -> tuple[dict, int]:
    """执行全部检查。``deck`` 可为路径字符串或二进制流；``manifest`` 可为
    路径 / 字典 / None。返回 (报告, 退出码)。"""
    deck_name = os.path.basename(deck) if isinstance(deck, str) else "<memory>"

    # 输入可读性：配对加载同时负责解包与取页序。
    try:
        slides = pc.load_slides(deck)
    except pc.CheckError as e:
        report = {
            "deck": {"file": deck_name, "tier": tier, "exit": 2,
                     "verdict": "fail"},
            "checks": {},
            "pages": [],
            "summary": {"pages": 0, "motion_fail_pages": [], "by_code": {}},
            "error": str(e),
        }
        return report, 2

    structure = _run_structure(deck)
    pairing = _run_pairing(slides)

    if isinstance(manifest, dict) or manifest is None:
        manifest_data = manifest
    else:
        manifest_data = ec.load_manifest(manifest)

    try:
        pkg = ec.Package(deck)
    except ec.CheckError as e:
        report = {
            "deck": {"file": deck_name, "tier": tier, "exit": 2,
                     "verdict": "fail"},
            "checks": {"structure": structure, "pairing": pairing},
            "pages": [],
            "summary": {"pages": 0, "motion_fail_pages": [], "by_code": {}},
            "error": str(e),
        }
        return report, 2

    editability = _run_editability(slides, pkg, manifest_data, tier)
    pkg.close()

    # schema 深扫：fast 跳过。
    if fast:
        schema = {"ran": False, "exit": 0,
                  "issues": ["fast 模式：跳过全包 schema 深扫"]}
    else:
        try:
            with zipfile.ZipFile(deck) as zf:
                schema_issues = _schema_scan(zf)
        except (OSError, zipfile.BadZipFile) as e:
            schema_issues = ["schema 深扫读取失败: %s" % e]
        schema = {"ran": True, "exit": 0 if not schema_issues else 1,
                  "issues": schema_issues}

    pages, motion_fail_pages = _collect_motion_pages(slides, pairing["report"],
                                                     manifest_data)
    motion_exit = 0 if not motion_fail_pages else 1

    # 成品后处理三条检查：fast / 全量两条路径都跑。
    postfix = _run_postfix(slides, manifest_data, require_morph_pages)
    extra_fail_pages = sorted(set(
        postfix["morph_missing"] + postfix["image_placeholder"]
        + [b["page"] for b in postfix["letters_count"]]))

    # 汇总 by_code：可编辑性码 + 动效发现码 + schema / 结构问题计数。
    by_code: dict[str, int] = dict(editability["by_code"])
    for page in pages:
        for finding in page["findings"]:
            code = finding["code"]
            by_code[code] = by_code.get(code, 0) + 1
    for issue in schema["issues"]:
        by_code["SCHEMA_ISSUE"] = by_code.get("SCHEMA_ISSUE", 0) + 1
    for issue in structure["issues"]:
        by_code["STRUCTURE_ISSUE"] = by_code.get("STRUCTURE_ISSUE", 0) + 1
    for _p in postfix["morph_missing"]:
        by_code["MORPH_MISSING"] = by_code.get("MORPH_MISSING", 0) + 1
    for _p in postfix["image_placeholder"]:
        by_code["IMAGE_PLACEHOLDER"] = by_code.get("IMAGE_PLACEHOLDER", 0) + 1
    for bad in postfix["letters_count"]:
        by_code["LETTERS_COUNT"] = by_code.get("LETTERS_COUNT", 0) + 1

    checks = {
        "structure": {"exit": structure["exit"], "issues": structure["issues"],
                      "effects": structure["effects"], "ctn": structure["ctn"]},
        "pairing": {"exit": pairing["exit"], "fail": pairing["fail"],
                    "orphan": pairing["orphan"],
                    "morph_pairs": pairing["morph_pairs"],
                    "conflict": pairing["conflict"]},
        "editability": {"exit": editability["exit"],
                        "score": editability["score"],
                        "blocking_pages": editability["blocking_pages"],
                        "by_code": editability["by_code"]},
        "schema": schema,
        "motion": {"exit": motion_exit, "pages_below_quota": motion_fail_pages},
        "morph_missing": postfix["morph_missing"],
        "image_placeholder": postfix["image_placeholder"],
        "letters_count": postfix["letters_count"],
    }

    exits = [checks[k]["exit"] for k in
             ("structure", "pairing", "editability", "schema", "motion")]
    exits.append(postfix["exit"])
    overall = 0 if all(e == 0 for e in exits) else 1

    report = {
        "deck": {"file": deck_name, "tier": tier, "exit": overall,
                 "verdict": "pass" if overall == 0 else "fail"},
        "checks": checks,
        "pages": pages,
        "summary": {"pages": len(pages),
                    "motion_fail_pages": motion_fail_pages,
                    "extra_fail_pages": extra_fail_pages,
                    "by_code": by_code},
    }
    return report, overall


# ----------------------------------------------------------------------
# 人读 Markdown 表
# ----------------------------------------------------------------------
def _render_md(report: dict) -> str:
    lines = ["# 合并校验报告", ""]
    d = report["deck"]
    lines.append("- 文件：**%s**；档位：**%s**；总体判定：**%s**；退出码：%d"
                 % (d["file"], d["tier"], d["verdict"], d["exit"]))
    lines.append("")
    c = report["checks"]
    lines.append("| 检查项 | 退出 | 要点 |")
    lines.append("|---|---|---|")
    lines.append("| 结构 | %d | 效果 %d / cTn %d，问题 %d 条 |"
                 % (c["structure"]["exit"], c["structure"]["effects"],
                    c["structure"]["ctn"], len(c["structure"]["issues"])))
    lines.append("| 配对 | %d | 失败页 %d，孤儿 %d，Morph 配对 %d，冲突 %d |"
                 % (c["pairing"]["exit"], c["pairing"]["fail"],
                    c["pairing"]["orphan"], len(c["pairing"]["morph_pairs"]),
                    c["pairing"]["conflict"]))
    lines.append("| 可编辑性 | %d | 整包分数 %d，阻断页 %s |"
                 % (c["editability"]["exit"], c["editability"]["score"],
                    ", ".join(str(x) for x in c["editability"]["blocking_pages"])
                    or "无"))
    lines.append("| schema | %d | 深扫 %s，问题 %d 条 |"
                 % (c["schema"]["exit"], "已跑" if c["schema"]["ran"] else "跳过",
                    len(c["schema"]["issues"])))
    lines.append("| 动效配额 | %d | 不达标页 %s |"
                 % (c["motion"]["exit"],
                    ", ".join(str(x) for x in c["motion"]["pages_below_quota"])
                    or "无"))
    lines.append("| Morph 缺失 | %d | 缺平滑页 %s |"
                 % (1 if c["morph_missing"] else 0,
                    ", ".join(str(x) for x in c["morph_missing"]) or "无"))
    lines.append("| 图片占位 | %d | 未填占位页 %s |"
                 % (1 if c["image_placeholder"] else 0,
                    ", ".join(str(x) for x in c["image_placeholder"]) or "无"))
    lines.append("| 字母数 | %d | 不符页 %s |"
                 % (1 if c["letters_count"] else 0,
                    ", ".join(str(b["page"]) for b in c["letters_count"])
                    or "无"))
    lines.append("")
    lines.append("## 逐页配额")
    lines.append("| 页 | 页型 | 效果 | entr | emph | exit | path | 可感知 | 循环 | 达标 | 缺失项 |")
    lines.append("|---|---|---|---|---|---|---|---|---|---|---|")
    for p in report["pages"]:
        k = p["kinds"]
        missing = "；".join(p["quota"]["missing"]) or "—"
        lines.append("| %d | %s | %d | %d | %d | %d | %d | %d | %d | %s | %s |"
                     % (p["index"], p["page_type"], p["effects"], k["entr"],
                        k["emph"], k["exit"], k["path"], p["perceptible"],
                        p["loops"], "是" if p["quota"]["ok"] else "否", missing))
    lines.append("")
    for p in report["pages"]:
        if p["findings"]:
            lines.append("### 第 %d 页发现" % p["index"])
            for f in p["findings"]:
                lines.append("- **[%s]** %s" % (f["code"], f.get("message")
                             or "; ".join(f.get("issues", [])) or f.get("object", "")))
            lines.append("")
    return "\n".join(lines)


def _print_console(report: dict) -> None:
    d = report["deck"]
    c = report["checks"]
    print("总体判定：%s（退出码 %d）" % (d["verdict"], d["exit"]))
    print("  结构 exit=%d  效果=%d  cTn=%d  问题=%d"
          % (c["structure"]["exit"], c["structure"]["effects"],
             c["structure"]["ctn"], len(c["structure"]["issues"])))
    print("  配对 exit=%d  失败页=%d  孤儿=%d  冲突=%d"
          % (c["pairing"]["exit"], c["pairing"]["fail"],
             c["pairing"]["orphan"], c["pairing"]["conflict"]))
    print("  可编辑性 exit=%d  分数=%d  阻断页=%s"
          % (c["editability"]["exit"], c["editability"]["score"],
             c["editability"]["blocking_pages"] or "无"))
    print("  schema exit=%d  深扫=%s  问题=%d"
          % (c["schema"]["exit"], "已跑" if c["schema"]["ran"] else "跳过",
             len(c["schema"]["issues"])))
    print("  动效 exit=%d  不达标页=%s"
          % (c["motion"]["exit"], c["motion"]["pages_below_quota"] or "无"))
    print("  Morph 缺失=%s  图片占位=%s  字母数不符=%s"
          % (c["morph_missing"] or "无", c["image_placeholder"] or "无",
             [b["page"] for b in c["letters_count"]] or "无"))
    for p in report["pages"]:
        k = p["kinds"]
        flag = "OK" if p["quota"]["ok"] else "缺"
        print("  第 %2d 页 [%s] 效果=%d 可感知=%d 循环=%d 分布=%s"
              % (p["index"], flag, p["effects"], p["perceptible"], p["loops"],
                 k))
        for msg in p["quota"]["missing"]:
            print("      ! %s" % msg)


# ----------------------------------------------------------------------
# 自检夹具（内存构造，不依赖外部文件）
# ----------------------------------------------------------------------
def _fixture_shape(sid: str, name: str, text: str) -> str:
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="1000000" y="1000000"/>'
        '<a:ext cx="2000000" cy="1000000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>' % (sid, name, text)
    )


def _fixture_shape_at(sid: str, name: str, text: str, x: int, y: int,
                      w: int = 2000000, h: int = 1000000) -> str:
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
        '<a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>' % (sid, name, x, y, w, h, text)
    )


def _fixture_timing(effect_pars: list[str]) -> str:
    return (
        '<p:timing xmlns:p="%s" xmlns:a="%s"><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" '
        'nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>%s'
        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>'
        '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        % (P_NS, A_NS, "".join(effect_pars))
    )


def _fx_fly(spid: str, ctn: int, delay: int) -> str:
    """入场 fly：ppt_y 位移，可感知运动。"""
    return (
        '<p:par><p:cTn id="%d" presetID="2" presetClass="entr" presetSubtype="4" '
        'fill="hold" grpId="0" nodeType="withEffect">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
        '<p:cTn id="%d" dur="500" fill="hold"/>'
        '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y-.1"/></p:val></p:tav>'
        '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>'
        '</p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>'
        % (ctn, delay, ctn + 1, spid, ctn + 2, spid)
    )


def _fx_spin_loop(spid: str, ctn: int) -> str:
    """强调自转：animRot + repeatCount=indefinite，持续循环。"""
    return (
        '<p:par><p:cTn id="%d" presetID="8" presetClass="emph" presetSubtype="0" '
        'fill="hold" grpId="0" nodeType="withEffect" repeatCount="indefinite">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        '<p:animRot by="21600000"><p:cBhvr>'
        '<p:cTn id="%d" dur="2000" fill="hold" repeatCount="indefinite"/>'
        '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>'
        '</p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>'
        % (ctn, ctn + 1, spid)
    )


def _fx_fade(spid: str, ctn: int, delay: int) -> str:
    """纯淡入：不计入可感知运动。"""
    return (
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" presetSubtype="0" '
        'fill="hold" grpId="0" nodeType="withEffect">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
        '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
        '<p:cTn id="%d" dur="500"/><p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>'
        '</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>'
        % (ctn, delay, ctn + 1, spid, ctn + 2, spid)
    )


MORPH_TRANSITION = (
    '<p:transition xmlns:mc="%s" xmlns:p159="%s" spd="med">'
    '<mc:AlternateContent><mc:Choice Requires="p159">'
    '<p159:morph option="byObject"/></mc:Choice>'
    '<mc:Fallback><p:fade/></mc:Fallback></mc:AlternateContent></p:transition>'
    % (pc.MC_NS if hasattr(pc, "MC_NS") else
       "http://schemas.openxmlformats.org/markup-compatibility/2006",
       pc.P159_NS)
)

FADE_TRANSITION = (
    '<p:transition xmlns:p="%s" spd="med"><p:fade/></p:transition>' % P_NS
)


def _fixture_package(slide_xml: bytes) -> io.BytesIO:
    buf = io.BytesIO()
    ct = (
        '<Types xmlns="%s"><Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '</Types>' % CT_NS)
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
        zf.writestr("[Content_Types].xml", XML_DECL + ct.encode("utf-8"))
        zf.writestr("_rels/.rels", XML_DECL + rels.encode("utf-8"))
        zf.writestr("ppt/presentation.xml", XML_DECL + pres.encode("utf-8"))
        zf.writestr("ppt/_rels/presentation.xml.rels",
                     XML_DECL + pres_rels.encode("utf-8"))
        zf.writestr("ppt/slides/slide1.xml", slide_xml)
    buf.seek(0)
    return buf


def _fixture_slide(shape_xml: str, timing_xml: str,
                   transition_xml: str = "") -> bytes:
    body = (
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm/></p:grpSpPr>%s'
        '</p:spTree></p:cSld>%s%s</p:sld>'
        % (A_NS, R_NS, P_NS, shape_xml, transition_xml, timing_xml)
    )
    return XML_DECL + body.encode("utf-8")


def _selftest() -> int:
    manifest = {"pages": [{"index": 1, "page_type": "cover",
                           "content": ["自检占位"], "assets": []}],
                "assets": []}

    # (a) 达标样例：3 个可感知运动对象 + 1 个无限循环。
    shapes_a = "".join([
        _fixture_shape("2", "dot-a", "A"),
        _fixture_shape("3", "dot-b", "B"),
        _fixture_shape("4", "ring", "R"),
    ])
    timing_a = _fixture_timing([
        _fx_fly("2", 5, 0),
        _fx_fly("3", 10, 120),
        _fx_spin_loop("4", 15),
    ])
    slide_a = _fixture_slide(shapes_a, timing_a)
    buf_a = _fixture_package(slide_a)
    report_a, code_a = run_audit(buf_a, manifest=manifest, fast=False)

    # (b) 反例：只挂 Morph 转场，页内仅纯 fade，无任何非 fade 动效。
    shapes_b = _fixture_shape("2", "hero", "H")
    timing_b = _fixture_timing([_fx_fade("2", 5, 0)])
    slide_b = _fixture_slide(shapes_b, timing_b, MORPH_TRANSITION)
    buf_b = _fixture_package(slide_b)
    report_b, code_b = run_audit(buf_b, manifest=manifest, fast=False)

    assert code_a == 0, "达标样例应退出 0，实际 %d；报告：%s" % (
        code_a, json.dumps(report_a, ensure_ascii=False))
    assert code_b == 1, "反例应退出 1，实际 %d" % code_b
    assert report_b["summary"]["motion_fail_pages"] == [1], \
        "反例不达标页应精确命中第 1 页，实际 %s" % \
        report_b["summary"]["motion_fail_pages"]
    missing_b = " ".join(report_b["pages"][0]["quota"]["missing"])
    assert "持续循环" in missing_b, "反例 missing 应包含「持续循环」，实际：%s" % missing_b
    assert "可感知运动对象" in missing_b, \
        "反例 missing 应包含「可感知运动对象」，实际：%s" % missing_b

    # (c) 反例：manifest 声明 transition=morph，但页内只挂 fade → morph-missing。
    manifest_c = {"pages": [{"index": 1, "transition": "morph"}], "assets": []}
    slide_c = _fixture_slide(_fixture_shape("2", "hero", "H"),
                             _fixture_timing([]), FADE_TRANSITION)
    buf_c = _fixture_package(slide_c)
    report_c, code_c = run_audit(buf_c, manifest=manifest_c, fast=False)
    assert code_c == 1, "morph 缺失反例应退出 1，实际 %d" % code_c
    assert report_c["checks"]["morph_missing"] == [1], \
        "morph-missing 应精确命中第 1 页，实际 %s" % \
        report_c["checks"]["morph_missing"]
    assert 1 in report_c["summary"]["extra_fail_pages"]

    # (d) 反例：文本命中占位词且形状 / 组合内无 blip → image-placeholder。
    shape_d = _fixture_shape("2", "ph", "点击添加图片")
    slide_d = _fixture_slide(shape_d, _fixture_timing([]))
    buf_d = _fixture_package(slide_d)
    report_d, code_d = run_audit(buf_d, manifest=manifest, fast=False)
    assert code_d == 1, "图片占位反例应退出 1，实际 %d" % code_d
    assert report_d["checks"]["image_placeholder"] == [1], \
        "image-placeholder 应精确命中第 1 页，实际 %s" % \
        report_d["checks"]["image_placeholder"]

    # (e) 反例：声明 letters_count=3，实际 2 个单字母 → letters-count。
    shapes_e = (_fixture_shape("2", "a", "A")
                + _fixture_shape("3", "b", "B"))
    manifest_e = {"pages": [{"index": 1, "letters_count": 3}], "assets": []}
    slide_e = _fixture_slide(shapes_e, _fixture_timing([]))
    buf_e = _fixture_package(slide_e)
    report_e, code_e = run_audit(buf_e, manifest=manifest_e, fast=False)
    assert code_e == 1, "字母数不符反例应退出 1，实际 %d" % code_e
    assert report_e["checks"]["letters_count"] == [
        {"page": 1, "expected": 3, "actual": 2}], \
        "letters-count 应命中第 1 页（3 应 2），实际 %s" % \
        report_e["checks"]["letters_count"]

    # (e2) 无 letters_count 声明 → skipped，不算 FAIL。
    report_e2, _ = run_audit(buf_e, manifest=manifest, fast=False)
    assert report_e2["checks"]["letters_count"] == [], \
        "无声明页不应判字母数不符，实际 %s" % \
        report_e2["checks"]["letters_count"]

    # (e3) 单字母全部出画（top<0）→ 视为布局，跳过不判 FAIL。
    shape_z = _fixture_shape_at("2", "z", "Z", 1000000, -1000000)
    slide_z = _fixture_slide(shape_z, _fixture_timing([]))
    buf_z = _fixture_package(slide_z)
    report_z, _ = run_audit(buf_z, manifest=manifest_e, fast=False)
    assert report_z["checks"]["letters_count"] == [], \
        "出画字母应跳过，实际 %s" % report_z["checks"]["letters_count"]

    # (e4) 部分出画不豁免：一个字母出画、一个在画布内，仍按实际数量核对。
    mixed = (_fixture_shape_at("2", "z", "Z", 1000000, -1000000)
             + _fixture_shape("3", "a", "A"))
    slide_m = _fixture_slide(mixed, _fixture_timing([]))
    buf_m = _fixture_package(slide_m)
    manifest_m = {"pages": [{"index": 1, "letters_count": 3}], "assets": []}
    report_m, code_m = run_audit(buf_m, manifest=manifest_m, fast=False)
    assert code_m == 1, "部分出画不应整体跳过，应退出 1，实际 %d" % code_m
    assert report_m["checks"]["letters_count"] == [
        {"page": 1, "expected": 3, "actual": 2}], \
        "部分出画仍应核对字母数，实际 %s" % \
        report_m["checks"]["letters_count"]

    print("selftest OK：达标样例退出 0；配额反例退出 1 命中 [1]；"
          "morph 缺失 / 图片占位 / 字母数不符三条反例均退出 1 且命中第 1 页；"
          "无声明与全部出画字母记 skipped 不判 FAIL，部分出画仍核对。")
    return 0


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="合并校验入口：结构 / 配对 / 可编辑性 / schema / 动效配额")
    ap.add_argument("deck", nargs="?", help="待校验 .pptx")
    ap.add_argument("--manifest", help="deck-manifest.json 路径")
    ap.add_argument("--tier", choices=ec.TIERS, default="fidelity_first",
                    help="交付档位（默认 fidelity_first）")
    ap.add_argument("--fast", action="store_true",
                    help="跳过全包 schema 深扫，只做结构关键项与配额")
    ap.add_argument("--require-morph-pages", default="",
                    help="逗号分隔的页号：这些页必须挂平滑切换（manifest 声明 "
                         "transition=morph 的页同样计入）")
    ap.add_argument("--json", dest="json_out", help="机器可读报告输出路径")
    ap.add_argument("--md", dest="md_out", help="人读 Markdown 表输出路径")
    ap.add_argument("--selftest", action="store_true", help="运行内联自检")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()
    if not args.deck:
        ap.print_help()
        return 2
    if not os.path.exists(args.deck):
        print("输入不可读: %s" % args.deck)
        return 2

    manifest_path = args.manifest
    if not manifest_path:
        cand = os.path.join(os.path.dirname(os.path.abspath(args.deck)),
                            "deck-manifest.json")
        if os.path.exists(cand):
            manifest_path = cand

    require_morph = [int(v) for v in args.require_morph_pages.split(",") if v.strip()]
    try:
        report, code = run_audit(args.deck, manifest=manifest_path,
                                 tier=args.tier, fast=args.fast,
                                 require_morph_pages=require_morph)
    except (OSError, zipfile.BadZipFile, ET.ParseError) as e:
        print("输入不可读: %s" % e)
        return 2

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
    if args.md_out:
        with open(args.md_out, "w", encoding="utf-8") as f:
            f.write(_render_md(report))

    if code != 2:
        _print_console(report)
    else:
        print("输入不可读: %s" % report.get("error", ""))
    return code


if __name__ == "__main__":
    sys.exit(main())
