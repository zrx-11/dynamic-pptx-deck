# -*- coding: utf-8 -*-
"""
page_regen.py — 页级局部重生成回填器（仅标准库：zipfile + xml.etree）。

用途：把目标稿（deck）的第 N 页替换为来源包（--from）的指定页；总页数、页序与
其余页保持不变，其余 slide 部件的 XML 字节在回填前后必须逐字节一致。被替换页的
图片 / 媒体 / 图表 / 嵌入对象等资源按传递闭包复制进目标包并完成 rId 重映射；版式
关系改挂目标稿原有版式；命名冲突自动加后缀；`<p:sldIdLst>` 与 presentation 关系
保持原样。

CLI：
    python3 page_regen.py --deck DECK.pptx --page N \\
        --from NEW.pptx [--from-page M] --out OUT.pptx \\
        [--report-json OUT.json] [--check-only]

    页码均为 1 基、按演示顺序；--from-page 默认 1。
    --check-only 只做可行性校验，不写出包文件。

退出码：0=成功；1=校验失败（页码越界、来源页数不足、关系悬空 / 不支持、页数变化、
其余页被改）；2=用法错误或文件读取错误。
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

RELS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
T_LAYOUT = RELS_TYPE + "slideLayout"
T_MASTER = RELS_TYPE + "slideMaster"
T_NOTES = RELS_TYPE + "notesSlide"
T_NOTES_MASTER = RELS_TYPE + "notesMaster"
T_THEME = RELS_TYPE + "theme"
T_SLIDE = RELS_TYPE + "slide"
T_HLINK = RELS_TYPE + "hyperlink"

# 作为内部资源复制进目标包的关系类型（含嵌套部件的传递闭包）。
COPY_TYPES = {
    RELS_TYPE + "image",
    RELS_TYPE + "media",
    RELS_TYPE + "audio",
    RELS_TYPE + "video",
    RELS_TYPE + "chart",
    RELS_TYPE + "oleObject",
    RELS_TYPE + "embeddedPackage",
    RELS_TYPE + "package",
    RELS_TYPE + "diagramData",
    RELS_TYPE + "customXml",
}

# 显式不搬运、由目标稿侧保留的关系类型。
DROP_TYPES = {T_MASTER, T_NOTES, T_NOTES_MASTER, T_THEME, T_SLIDE}

MEDIA_CT = {
    "jpeg": "image/jpeg", "jpg": "image/jpeg", "png": "image/png",
    "gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff",
    "wmf": "image/x-wmf", "emf": "image/x-emf", "svg": "image/svg+xml",
    "mp4": "video/mp4", "mp3": "audio/mpeg", "wav": "audio/wav",
    "webm": "video/webm",
}

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_RID_RE = re.compile(r'(r:(?:id|embed|link)=")(rId\d+)(")')


class RegenError(RuntimeError):
    """校验失败：退出码 1。"""


class ReadError(RuntimeError):
    """用法 / 读取错误：退出码 2。"""


def _norm_part(base_dir: str, target: str) -> str:
    """把 rels Target 相对路径归一化为包内绝对 partname。"""
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_dir + "/" + target).split("/")
    stack = []
    for p in parts:
        if p in ("", "."):
            continue
        if p == "..":
            stack.pop()
        else:
            stack.append(p)
    return "/".join(stack)


def _rel_target_rel(src_partname: str, dest_partname: str) -> str:
    """从源 slide / 部件所在目录到目标 partname 的相对 Target。"""
    src_dir = src_partname.rsplit("/", 1)[0]
    dest_dir, dest_base = dest_partname.rsplit("/", 1)
    src_parts = src_dir.split("/")
    dest_parts = dest_dir.split("/")
    common = 0
    while common < len(src_parts) and common < len(dest_parts) \
            and src_parts[common] == dest_parts[common]:
        common += 1
    up = [".."] * (len(src_parts) - common)
    return "/".join(up + dest_parts[common:] + [dest_base])


def _nested_rels_name(partname: str) -> str:
    d, b = partname.rsplit("/", 1)
    return d + "/_rels/" + b + ".rels"


def _load_zip(path: str) -> dict[str, bytes]:
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad is not None:
                raise ReadError("压缩包损坏：%s" % bad)
            return {n: z.read(n) for n in z.namelist()}
    except (OSError, zipfile.BadZipFile) as e:
        raise ReadError("包不可读：%s（%s）" % (path, e))


def _parse_xml(data: bytes, what: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise ReadError("XML 解析失败：%s（%s）" % (what, e))


def _slide_order(parts: dict[str, bytes], label: str) -> list[str]:
    pres_name = "ppt/presentation.xml"
    rels_name = "ppt/_rels/presentation.xml.rels"
    if pres_name not in parts or rels_name not in parts:
        raise ReadError("%s 缺少 presentation.xml 或其关系文件" % label)
    pres = _parse_xml(parts[pres_name], label + "/presentation.xml")
    rels = _parse_xml(parts[rels_name], label + "/presentation.xml.rels")
    rel_map = {r.get("Id"): r.get("Target")
               for r in rels.findall("{%s}Relationship" % REL_NS)}
    lst = pres.find("{%s}sldIdLst" % P_NS)
    if lst is None:
        raise ReadError("%s presentation.xml 缺少 sldIdLst" % label)
    order = []
    for sld in lst:
        rid = sld.get("{%s}id" % R_NS)
        if rid not in rel_map:
            raise ReadError("%s sldIdLst 存在悬空 rId：%s" % (label, rid))
        order.append(_norm_part("ppt", rel_map[rid]))
    return order


def _slide_rels_name(slide_part: str) -> str:
    return "ppt/slides/_rels/" + slide_part.rsplit("/", 1)[-1] + ".rels"


def _read_rels(parts: dict[str, bytes], rels_name: str) -> list[dict]:
    if rels_name not in parts:
        return []
    root = _parse_xml(parts[rels_name], rels_name)
    out = []
    for rel in root.findall("{%s}Relationship" % REL_NS):
        out.append({
            "id": rel.get("Id"),
            "type": rel.get("Type"),
            "target": rel.get("Target"),
            "mode": rel.get("TargetMode"),
        })
    return out


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _next_free_partname(dest_parts: set[str], source_part: str) -> str:
    if source_part not in dest_parts:
        return source_part
    head, ext = os.path.splitext(source_part)
    n = 2
    while True:
        candidate = "%s-%d%s" % (head, n, ext)
        if candidate not in dest_parts:
            return candidate
        n += 1


def _ensure_content_types(ct_root: ET.Element, copied: dict[str, bytes],
                          dest_to_src: dict[str, str],
                          src_ct_root: ET.Element) -> None:
    """为新复制部件补 Default 扩展名声明与 Override 条目。"""
    defaults = {d.get("Extension").lower(): d.get("ContentType")
                for d in ct_root.findall("{%s}Default" % CT_NS)}
    src_defaults = {d.get("Extension").lower(): d.get("ContentType")
                    for d in src_ct_root.findall("{%s}Default" % CT_NS)}
    overrides = {o.get("PartName").lstrip("/")
                  for o in ct_root.findall("{%s}Override" % CT_NS)}
    src_overrides = {o.get("PartName").lstrip("/"): o
                     for o in src_ct_root.findall("{%s}Override" % CT_NS)}

    for dest_part in copied:
        if dest_part in overrides:
            continue
        src_part = dest_to_src[dest_part]
        ext = dest_part.rsplit(".", 1)[-1].lower() if "." in dest_part else ""
        if src_part in src_overrides:
            src_o = src_overrides[src_part]
            ET.SubElement(ct_root, "{%s}Override" % CT_NS,
                          {"PartName": "/" + dest_part,
                           "ContentType": src_o.get("ContentType")})
            overrides.add(dest_part)
        elif ext in MEDIA_CT and ext not in defaults:
            ET.SubElement(ct_root, "{%s}Default" % CT_NS,
                          {"Extension": ext, "ContentType": MEDIA_CT[ext]})
            defaults[ext] = MEDIA_CT[ext]
        elif ext in src_defaults and ext not in defaults:
            ET.SubElement(ct_root, "{%s}Default" % CT_NS,
                          {"Extension": ext, "ContentType": src_defaults[ext]})
            defaults[ext] = src_defaults[ext]
        elif ext == "bin" and ext not in defaults:
            ET.SubElement(ct_root, "{%s}Default" % CT_NS,
                          {"Extension": ext,
                           "ContentType": "application/vnd.openxmlformats-officedocument.oleObject"})
            defaults[ext] = "application/vnd.openxmlformats-officedocument.oleObject"


def _build(args: argparse.Namespace) -> dict:
    deck_parts = _load_zip(args.deck)
    from_parts = _load_zip(args.from_)

    deck_order = _slide_order(deck_parts, "deck")
    from_order = _slide_order(from_parts, "from")
    pages_before = len(deck_order)

    if args.page < 1 or args.page > pages_before:
        raise RegenError("目标页码越界：第 %d 页，共 %d 页" % (args.page, pages_before))
    if args.from_page < 1 or args.from_page > len(from_order):
        raise RegenError("来源包页数不足：取第 %d 页，共 %d 页"
                         % (args.from_page, len(from_order)))

    dst_slide = deck_order[args.page - 1]
    src_slide = from_order[args.from_page - 1]
    dst_rels_name = _slide_rels_name(dst_slide)
    src_rels_name = _slide_rels_name(src_slide)

    # 回填前记录全部 slide XML 哈希（按演示顺序）。
    before_hashes = {p: _sha256(deck_parts[p]) for p in deck_order}

    # 目标稿该页现有版式 partname：替换页改挂同一版式，保持母版 / 版式侧不变。
    dst_layout = None
    for rel in _read_rels(deck_parts, dst_rels_name):
        if rel["type"] == T_LAYOUT:
            dst_layout = _norm_part("ppt/slides", rel["target"])
            break

    src_rels = _read_rels(from_parts, src_rels_name)
    src_slide_bytes = from_parts[src_slide]
    try:
        src_slide_text = src_slide_bytes.decode("utf-8")
    except UnicodeDecodeError as e:
        raise ReadError("来源页 XML 非 UTF-8：%s" % e)
    used_src_rids = {m[1] for m in _RID_RE.finditer(src_slide_text)}

    copied_parts: dict[str, bytes] = {}
    dest_to_src: dict[str, str] = {}
    # 嵌套部件 rels：dest_part -> 重写后的关系条目清单。
    nested_rels: dict[str, list[dict]] = {}
    src_to_dest: dict[str, str] = {}
    visited_src: set[str] = set()
    dest_part_set = set(deck_parts.keys())

    src_ct_root = _parse_xml(from_parts["[Content_Types].xml"],
                             "from/[Content_Types].xml")

    def copy_resource(src_part: str) -> str:
        """传递闭包复制来源资源部件，返回目标包内 partname。"""
        if src_part in src_to_dest:
            return src_to_dest[src_part]
        if src_part not in from_parts:
            raise RegenError("来源关系悬空：%s" % src_part)
        visited_src.add(src_part)
        data = from_parts[src_part]
        reused = src_part in deck_parts and \
            _sha256(deck_parts[src_part]) == _sha256(data)
        if reused:
            dest_part = src_part
        else:
            dest_part = _next_free_partname(dest_part_set, src_part)
            copied_parts[dest_part] = data
            dest_part_set.add(dest_part)
        dest_to_src[dest_part] = src_part
        src_to_dest[src_part] = dest_part

        # 复用目标包已有部件时保留其原有 rels，不重写嵌套关系。
        rels_name = _nested_rels_name(src_part)
        out_nested: list[dict] = []
        if reused:
            nested_rels[dest_part] = out_nested
            return dest_part
        if rels_name in from_parts:
            next_rid = 1
            for rel in _read_rels(from_parts, rels_name):
                rtype = rel["type"]
                if rtype in DROP_TYPES:
                    continue
                if rtype == T_HLINK:
                    if rel["mode"] != "External":
                        raise RegenError("内部超链接关系暂不支持：%s" % rel["id"])
                    out_nested.append({"id": "rId%d" % next_rid, "type": rtype,
                                       "target": rel["target"], "mode": "External"})
                    next_rid += 1
                    continue
                if rtype not in COPY_TYPES:
                    raise RegenError("不支持的嵌套关系类型：%s（%s）"
                                     % (rtype, rel["id"]))
                nsrc = _norm_part(src_part.rsplit("/", 1)[0], rel["target"])
                ndest = copy_resource(nsrc)
                out_nested.append({"id": "rId%d" % next_rid, "type": rtype,
                                   "target": _rel_target_rel(dest_part, ndest),
                                   "mode": None})
                next_rid += 1
        nested_rels[dest_part] = out_nested
        return dest_part

    # 解析来源页关系：分类并解析内部目标。
    remap: dict[str, str] = {}          # 来源页 rId -> 目标页 rId
    out_rels: list[dict] = []
    next_rid = 1

    for rel in src_rels:
        rid = rel["id"]
        rtype = rel["type"]
        if rtype == T_LAYOUT:
            if dst_layout is not None:
                remap[rid] = "rId%d" % next_rid
                out_rels.append({"id": remap[rid], "type": rtype,
                                 "target": _rel_target_rel(dst_slide, dst_layout),
                                 "mode": None})
                next_rid += 1
            # 无目标版式可挂时丢弃版式关系，不引入来源母版侧部件。
            continue
        if rtype in DROP_TYPES:
            if rid in used_src_rids:
                raise RegenError("来源页正文引用了不搬运的关系：%s（%s）" % (rid, rtype))
            continue
        if rtype == T_HLINK:
            if rel["mode"] != "External":
                raise RegenError("内部超链接关系暂不支持：%s" % rid)
            remap[rid] = "rId%d" % next_rid
            out_rels.append({"id": remap[rid], "type": rtype,
                             "target": rel["target"], "mode": "External"})
            next_rid += 1
            continue
        if rtype not in COPY_TYPES:
            raise RegenError("不支持的关系类型：%s（%s，目标 %s）"
                             % (rtype, rid, rel["target"]))

        src_part = _norm_part("ppt/slides", rel["target"])
        dest_part = copy_resource(src_part)
        remap[rid] = "rId%d" % next_rid
        out_rels.append({"id": remap[rid], "type": rtype,
                         "target": _rel_target_rel(dst_slide, dest_part),
                         "mode": None})
        next_rid += 1

    # 正文 rId 引用校验与重写。
    def _sub(m: re.Match) -> str:
        old = m.group(2)
        if old not in remap:
            raise RegenError("页面正文引用了未解析的 rId：%s" % old)
        return m.group(1) + remap[old] + m.group(3)

    rewritten_bytes = _RID_RE.sub(_sub, src_slide_text).encode("utf-8")

    # 组装目标包部件。
    new_parts = dict(deck_parts)
    new_parts[dst_slide] = rewritten_bytes

    # 重写被替换页的 rels。
    ET.register_namespace("", REL_NS)
    rels_root = ET.Element("{%s}Relationships" % REL_NS)
    for rel in out_rels:
        attrs = {"Id": rel["id"], "Type": rel["type"], "Target": rel["target"]}
        if rel["mode"]:
            attrs["TargetMode"] = rel["mode"]
        ET.SubElement(rels_root, "{%s}Relationship" % REL_NS, attrs)
    new_parts[dst_rels_name] = XML_DECL + ET.tostring(rels_root, encoding="UTF-8")

    # 落盘复制部件与其嵌套 rels。
    for partname, data in copied_parts.items():
        new_parts[partname] = data
    for dest_part, rel_entries in nested_rels.items():
        if not rel_entries:
            continue
        nroot = ET.Element("{%s}Relationships" % REL_NS)
        for rel in rel_entries:
            attrs = {"Id": rel["id"], "Type": rel["type"], "Target": rel["target"]}
            if rel["mode"]:
                attrs["TargetMode"] = rel["mode"]
            ET.SubElement(nroot, "{%s}Relationship" % REL_NS, attrs)
        new_parts[_nested_rels_name(dest_part)] = \
            XML_DECL + ET.tostring(nroot, encoding="UTF-8")

    # Content-Types 补声明。
    ct_root = _parse_xml(new_parts["[Content_Types].xml"], "deck/[Content_Types].xml")
    _ensure_content_types(ct_root, copied_parts, dest_to_src, src_ct_root)
    ET.register_namespace("", CT_NS)
    new_parts["[Content_Types].xml"] = XML_DECL + ET.tostring(ct_root, encoding="UTF-8")

    # 回填后校验：页数不变；除目标页外其余 slide XML 哈希逐字节一致。
    after_order = _slide_order(new_parts, "deck(回填后)")
    pages_after = len(after_order)
    unchanged_others = True
    if pages_after != pages_before:
        unchanged_others = False
    else:
        for part in deck_order:
            if part == dst_slide:
                continue
            if _sha256(new_parts[part]) != before_hashes[part]:
                unchanged_others = False
                break

    if pages_after != pages_before or not unchanged_others:
        raise RegenError("回填后校验失败：页数 %d -> %d，其余页一致=%s"
                         % (pages_before, pages_after, unchanged_others))

    report = {
        "deck": os.path.abspath(args.deck),
        "page": args.page,
        "out": os.path.abspath(args.out) if args.out else None,
        "pages_before": pages_before,
        "pages_after": pages_after,
        "unchanged_others": unchanged_others,
        "copied_parts": sorted(copied_parts.keys()),
    }

    if not args.check_only:
        if not args.out:
            raise ReadError("非 --check-only 模式必须提供 --out")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in new_parts.items():
                z.writestr(name, data)
        with open(args.out, "wb") as f:
            f.write(buffer.getvalue())

    return report


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="页级局部重生成回填器：替换目标稿指定页，其余页逐字节不变")
    ap.add_argument("--deck", required=True, help="目标稿 .pptx")
    ap.add_argument("--page", type=int, required=True, help="目标页码（1 基，按演示顺序）")
    ap.add_argument("--from", dest="from_", required=True, help="来源单页包 .pptx")
    ap.add_argument("--from-page", type=int, default=1, help="来源包页码（1 基，默认 1）")
    ap.add_argument("--out", help="回填后输出 .pptx（--check-only 时可省略）")
    ap.add_argument("--report-json", dest="report_json", help="机器可读报告输出路径")
    ap.add_argument("--check-only", action="store_true", help="只做可行性校验，不写包")
    args = ap.parse_args(argv)

    try:
        report = _build(args)
    except ReadError as e:
        print("读取错误：%s" % e)
        return 2
    except RegenError as e:
        print("校验失败：%s" % e)
        return 1
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, ValueError) as e:
        print("包不可读：%s" % e)
        return 2

    if args.report_json and not args.check_only:
        with open(args.report_json, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    print("页级重生成完成：第 %d 页（%d -> %d 页，其余页一致=%s）"
          % (report["page"], report["pages_before"], report["pages_after"],
             report["unchanged_others"]))
    if report["copied_parts"]:
        print("复制部件：%s" % ", ".join(report["copied_parts"]))
    if args.check_only:
        print("（--check-only：未写出包文件）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
