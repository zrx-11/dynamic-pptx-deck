# -*- coding: utf-8 -*-
"""
slide_importer.py — 跨包 .pptx 幻灯片导入工具（仅标准库，不依赖 python-pptx）。

用途：把参考 pptx 中现成的高级效果页（形状 + 媒体 + timing/transition 动画节点）
整体手术式复制进目标 deck，重指版式、按内容哈希去重媒体、解析主题色，再做文本替换
与页序调整。timing 与 transition 节点除主题色解析与 rId 改写外原样保留。

公开 API：
    SlideImporter(dst_pptx: str)
        .import_slide(src_pptx, src_index=0, text_map=None, by_name=False) -> int
        .replace_text(slide_index, mapping, by_name=False) -> None
        .reorder(order: list[int]) -> None
        .save(path=None) -> None
    inspect(src_pptx, src_index=0) -> dict

CLI：
    python3 slide_importer.py inspect --src X.pptx [--index N]
    python3 slide_importer.py import --src X.pptx --index N --dst D.pptx
            [--out O.pptx] [--map map.json] [--by-name]
    python3 slide_importer.py replace --dst D.pptx --index N --map map.json
            [--by-name] [--out O.pptx]
    python3 slide_importer.py reorder --dst D.pptx --order 0,3,1,2 [--out O.pptx]
    python3 slide_importer.py selftest

text_map 语义：
    默认（by_name=False）：按键为 a:t 旧全文、值为新文本，按文档顺序整串精确匹配替换；
    by_name=True：键为形状名，保留首个 run 的 rPr，其余 run 文本清空。
    所有替换只动 a:t 文本节点，不碰 rPr 与段落结构。
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

# ----------------------------------------------------------------------
# 命名空间与关系类型
# ----------------------------------------------------------------------
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
EXT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

NS = {"p": P_NS, "a": A_NS, "r": R_NS, "rel": REL_NS, "ct": CT_NS, "vt": VT_NS,
      "ext": EXT_NS}

RELS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
T_LAYOUT = RELS_TYPE + "slideLayout"
T_MASTER = RELS_TYPE + "slideMaster"
T_IMAGE = RELS_TYPE + "image"
T_NOTES = RELS_TYPE + "notesSlide"
T_HLINK = RELS_TYPE + "hyperlink"
T_SLIDE = RELS_TYPE + "slide"

# 显式不搬运、直接抛错的关系类型（L2 / 未实现）
UNSUPPORTED_TYPES = {
    RELS_TYPE + "audio": "音频",
    RELS_TYPE + "video": "视频",
    RELS_TYPE + "oleObject": "嵌入对象",
    RELS_TYPE + "embeddedPackage": "嵌入包",
    RELS_TYPE + "chart": "图表",
    RELS_TYPE + "diagramData": "SmartArt 数据",
    RELS_TYPE + "notesMaster": "备注母版",
}

# 媒体扩展名 -> Content-Type（缺什么补什么）
MEDIA_CT = {
    "jpeg": "image/jpeg", "jpg": "image/jpeg", "png": "image/png",
    "gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff",
    "wmf": "image/x-wmf", "emf": "image/x-emf",
}

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_REGISTERED = False


def _register_namespaces():
    global _REGISTERED
    if _REGISTERED:
        return
    for prefix, uri in [
        ("p", P_NS), ("a", A_NS), ("r", R_NS),
        ("vt", VT_NS),
        ("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main"),
        ("p159", "http://schemas.microsoft.com/office/powerpoint/2015/12/main"),
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main"),
    ]:
        ET.register_namespace(prefix, uri)
    _REGISTERED = True


class ImporterError(RuntimeError):
    """导入过程中的可读错误。"""


# ----------------------------------------------------------------------
# 基础 XML 工具
# ----------------------------------------------------------------------
def _parse(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _serialize(root: ET.Element, default_ns: str | None = None) -> bytes:
    if default_ns:
        ET.register_namespace("", default_ns)
    return XML_DECL + ET.tostring(root, encoding="UTF-8")


def _q(ns: str, tag: str) -> str:
    return "{%s}%s" % (NS[ns], tag)


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


def _rel_dir(partname: str) -> str:
    return partname.rsplit("/", 1)[0]


def _rels_base_dir(rels_name: str) -> str:
    """rels 部件所在目录的上一级（关系 Target 相对该目录解析）。"""
    d = rels_name.rsplit("/", 1)[0]
    if d == "_rels":
        return ""
    if d.endswith("/_rels"):
        return d[: -len("/_rels")]
    return d


# ----------------------------------------------------------------------
# 主题色解析
# ----------------------------------------------------------------------
def _parse_theme_colors(theme_bytes: bytes) -> dict:
    """读 theme1.xml 的 clrScheme，返回 scheme 名 -> 十六进制色值（小写 6 位）。"""
    root = _parse(theme_bytes)
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        raise ImporterError("主题中未找到 a:clrScheme")
    colors = {}
    aliases = {
        "dk1": "dk1", "tx1": "dk1",
        "lt1": "lt1", "bg1": "lt1",
        "dk2": "dk2", "tx2": "dk2",
        "lt2": "lt2", "bg2": "lt2",
    }
    for child in scheme:
        name = etree_local(child)
        srgb = child.find("a:srgbClr", NS)
        sysc = child.find("a:sysClr", NS)
        if srgb is not None:
            val = srgb.get("val")
        elif sysc is not None:
            val = sysc.get("lastClr")
        else:
            continue
        if val:
            colors[name] = val.lower()
    resolved = {}
    for alias, canonical in aliases.items():
        if canonical in colors:
            resolved[alias] = colors[canonical]
    for k, v in colors.items():
        resolved[k] = v
    return resolved


def etree_local(el: ET.Element) -> str:
    return etree_qname(el).split("}", 1)[-1]


def etree_qname(el: ET.Element) -> str:
    return el.tag if "}" in el.tag else "{" + A_NS + "}" + el.tag


# ----------------------------------------------------------------------
# 主类
# ----------------------------------------------------------------------
class SlideImporter:
    """把源 pptx 的指定幻灯片导入目标 pptx 包。"""

    def __init__(self, dst_pptx: str):
        _register_namespaces()
        self.dst_path = dst_pptx
        self.parts: dict[str, bytes] = {}
        with zipfile.ZipFile(dst_pptx) as z:
            for name in z.namelist():
                self.parts[name] = z.read(name)
        self._media_hash: dict[str, str] = {}
        self._scan_existing_media()
        self._blank_layout: str | None = self._detect_blank_layout()

    # ---------- 内部：包状态 ----------
    def _scan_existing_media(self):
        for name, data in self.parts.items():
            if name.startswith("ppt/media/"):
                self._media_hash[hashlib.sha1(data).hexdigest()] = name

    def _detect_blank_layout(self) -> str | None:
        """从目标包现有原生幻灯片的 rels 识别其版式 partname，供导入页重指。"""
        slide_rels = sorted(
            n for n in self.parts
            if re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", n)
        )
        for rels_name in slide_rels:
            root = _parse(self.parts[rels_name])
            for rel in root.findall("rel:Relationship", NS):
                if rel.get("Type") == T_LAYOUT:
                    base = _rels_base_dir(rels_name)
                    return _norm_part(base, rel.get("Target"))
        return None

    def _pres_root(self) -> ET.Element:
        return _parse(self.parts["ppt/presentation.xml"])

    def _pres_rels_root(self) -> ET.Element:
        return _parse(self.parts["ppt/_rels/presentation.xml.rels"])

    def _slide_parts_in_order(self) -> list[str]:
        pres = self._pres_root()
        rels = self._pres_rels_root()
        rel_map = {r.get("Id"): r.get("Target")
                   for r in rels.findall("rel:Relationship", NS)}
        out = []
        for sld in pres.find("p:sldIdLst", NS):
            rid = sld.get(_q("r", "id"))
            out.append(_norm_part("ppt", rel_map[rid]))
        return out

    def _next_slide_name(self) -> str:
        nums = [int(re.search(r"\d+", n).group()) for n in self.parts
                if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        return "ppt/slides/slide%d.xml" % (max(nums) + 1 if nums else 1)

    def _next_rel_id(self, rels_root: ET.Element) -> str:
        ids = [int(r.get("Id")[3:]) for r in rels_root.findall("rel:Relationship", NS)
               if r.get("Id", "").startswith("rId")]
        return "rId%d" % (max(ids) + 1 if ids else 1)

    # ---------- 文本替换 ----------
    @staticmethod
    def _text_frames(slide_root: ET.Element) -> dict[str, list[ET.Element]]:
        """形状名 -> 该形状内按文档顺序的 a:t 节点列表（含组内嵌套）。"""
        frames: dict[str, list[ET.Element]] = {}
        for shape in slide_root.iter():
            local = etree_local(shape)
            if local not in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame"):
                continue
            cnv = shape.find(".//p:cNvPr", NS)
            if cnv is None:
                continue
            name = cnv.get("name") or ""
            frames.setdefault(name, []).extend(shape.findall(".//a:t", NS))
        return frames

    def replace_text(self, slide_index: int, mapping: dict, by_name: bool = False) -> None:
        slide_name = self._slide_parts_in_order()[slide_index]
        root = _parse(self.parts[slide_name])
        misses: list[str] = []

        if by_name:
            frames = self._text_frames(root)
            for shape_name, new_text in mapping.items():
                runs = frames.get(shape_name)
                if not runs:
                    misses.append(shape_name)
                    continue
                runs[0].text = new_text
                for extra in runs[1:]:
                    extra.text = ""
        else:
            all_runs = root.findall(".//a:t", NS)
            for old_text, new_text in mapping.items():
                hits = [i for i, t in enumerate(all_runs) if (t.text or "") == old_text]
                if not hits:
                    misses.append(old_text)
                    continue
                for i in hits:
                    all_runs[i].text = new_text

        if misses:
            raise ImporterError("文本替换未命中：%s" % "；".join(str(m) for m in misses))
        self.parts[slide_name] = _serialize(root)

    # ---------- 导入 ----------
    def import_slide(self, src_pptx: str, src_index: int = 0,
                     text_map: dict | None = None, by_name: bool = False) -> int:
        with zipfile.ZipFile(src_pptx) as z:
            names = set(z.namelist())
            src_pres = _parse(z.read("ppt/presentation.xml"))
            src_pres_rels = _parse(z.read("ppt/_rels/presentation.xml.rels"))
            rel_map = {r.get("Id"): r.get("Target")
                       for r in src_pres_rels.findall("rel:Relationship", NS)}
            src_slides = []
            for sld in src_pres.find("p:sldIdLst", NS):
                rid = sld.get(_q("r", "id"))
                src_slides.append(_norm_part("ppt", rel_map[rid]))
            if not (0 <= src_index < len(src_slides)):
                raise ImporterError("源幻灯片序号 %d 越界（共 %d 页）" % (src_index, len(src_slides)))
            src_slide = src_slides[src_index]
            src_slide_bytes = z.read(src_slide)

            src_rels_name = "ppt/slides/_rels/" + src_slide.rsplit("/", 1)[-1] + ".rels"
            src_rels_root = _parse(z.read(src_rels_name)) if src_rels_name in names else None

            theme_name = "ppt/theme/theme1.xml"
            theme_colors = _parse_theme_colors(z.read(theme_name)) if theme_name in names else {}

            # 目标新 slide part
            dst_slide = self._next_slide_name()
            dst_rels_name = "ppt/slides/_rels/" + dst_slide.rsplit("/", 1)[-1] + ".rels"

            # 1) 重写 slide rels：rId 重编号、layout 重指、媒体去重、notes 丢弃
            dst_rels_root = ET.Element(_q("rel", "Relationships"))
            old_to_new: dict[str, str] = {}
            new_rn = 1
            media_added: list[str] = []

            if src_rels_root is not None:
                for rel in src_rels_root.findall("rel:Relationship", NS):
                    rtype = rel.get("Type")
                    target = rel.get("Target")
                    old_id = rel.get("Id")
                    target_mode = rel.get("TargetMode")

                    if rtype == T_LAYOUT:
                        if self._blank_layout is None:
                            raise ImporterError("目标包未识别到原生版式，无法重指 layout")
                        new_id = "rId%d" % new_rn
                        new_rn += 1
                        ET.SubElement(dst_rels_root, _q("rel", "Relationship"), {
                            "Id": new_id, "Type": T_LAYOUT,
                            "Target": "../slideLayouts/" + self._blank_layout.rsplit("/", 1)[-1],
                        })
                        old_to_new[old_id] = new_id

                    elif rtype == T_NOTES:
                        # 备注关系与备注 part 一律丢弃
                        continue

                    elif rtype == T_IMAGE:
                        src_media = _norm_part("ppt/slides", target)
                        data = z.read(src_media)
                        digest = hashlib.sha1(data).hexdigest()
                        if digest in self._media_hash:
                            dst_media = self._media_hash[digest]
                        else:
                            dst_media = self._next_media_name(src_media)
                            self.parts[dst_media] = data
                            self._media_hash[digest] = dst_media
                            media_added.append(dst_media)
                            self._register_media_content_type(dst_media)
                        new_id = "rId%d" % new_rn
                        new_rn += 1
                        ET.SubElement(dst_rels_root, _q("rel", "Relationship"), {
                            "Id": new_id, "Type": T_IMAGE,
                            "Target": "../media/" + dst_media.rsplit("/", 1)[-1],
                        })
                        old_to_new[old_id] = new_id

                    elif rtype == T_HLINK:
                        new_id = "rId%d" % new_rn
                        new_rn += 1
                        attrs = {"Id": new_id, "Type": T_HLINK, "Target": target}
                        if target_mode:
                            attrs["TargetMode"] = target_mode
                        ET.SubElement(dst_rels_root, _q("rel", "Relationship"), attrs)
                        old_to_new[old_id] = new_id

                    elif rtype in UNSUPPORTED_TYPES:
                        raise ImporterError(
                            "源页含不支持的%s关系（%s），已中止导入；未生成坏包"
                            % (UNSUPPORTED_TYPES[rtype], target))
                    else:
                        raise ImporterError(
                            "未识别的关系类型 %s（target=%s），已中止导入" % (rtype, target))

            # 2) 改写 slide xml：rId 引用 + schemeClr 解析
            slide_root = _parse(src_slide_bytes)
            self._rewrite_rids(slide_root, old_to_new)
            self._resolve_scheme_colors(slide_root, theme_colors)

            self.parts[dst_slide] = _serialize(slide_root)
            self.parts[dst_rels_name] = _serialize(dst_rels_root, REL_NS)

            # 3) Content_Types 补 slide Override
            self._register_slide_content_type(dst_slide)

            # 4) presentation rels + sldIdLst 注册
            pres_root = self._pres_root()
            pres_rels = self._pres_rels_root()
            new_pres_rid = self._next_rel_id(pres_rels)
            ET.SubElement(pres_rels, _q("rel", "Relationship"), {
                "Id": new_pres_rid, "Type": T_SLIDE,
                "Target": "slides/" + dst_slide.rsplit("/", 1)[-1],
            })
            sld_id_lst = pres_root.find("p:sldIdLst", NS)
            existing_ids = [int(s.get("id")) for s in sld_id_lst]
            new_sld_id = max(existing_ids) + 1 if existing_ids else 256
            ET.SubElement(sld_id_lst, _q("p", "sldId"), {
                "id": str(new_sld_id), _q("r", "id"): new_pres_rid,
            })
            self.parts["ppt/presentation.xml"] = _serialize(pres_root)
            self.parts["ppt/_rels/presentation.xml.rels"] = _serialize(pres_rels, REL_NS)

            # 5) app.xml 页数尽力同步
            self._sync_app_props(len(self._slide_parts_in_order()))

        # 6) 文本替换
        new_index = len(self._slide_parts_in_order()) - 1
        if text_map:
            self.replace_text(new_index, text_map, by_name=by_name)
        return new_index

    # ---------- 内部：XML 改写 ----------
    @staticmethod
    def _rewrite_rids(root: ET.Element, old_to_new: dict[str, str]):
        rid_attrs = [_q("r", "embed"), _q("r", "link"), _q("r", "id")]
        for el in root.iter():
            for attr in rid_attrs:
                val = el.get(attr)
                if val in old_to_new:
                    el.set(attr, old_to_new[val])

    def _resolve_scheme_colors(self, root: ET.Element, theme_colors: dict):
        for clr in list(root.iter(_q("a", "schemeClr"))):
            name = clr.get("val")
            if name not in theme_colors:
                raise ImporterError("主题色 %s 未在源主题 clrScheme 中解析到" % name)
            hex_val = theme_colors[name]
            srgb = ET.Element(_q("a", "srgbClr"), {"val": hex_val})
            for child in list(clr):
                srgb.append(child)
            parent = self._parent(root, clr)
            idx = list(parent).index(clr)
            parent.remove(clr)
            parent.insert(idx, srgb)

    @staticmethod
    def _parent(root: ET.Element, child: ET.Element) -> ET.Element:
        for p in root.iter():
            if child in list(p):
                return p
        raise ImporterError("内部错误：找不到元素父节点")

    # ---------- 内部：媒体与 Content_Types ----------
    def _next_media_name(self, src_media: str) -> str:
        ext = src_media.rsplit(".", 1)[-1].lower()
        nums = [int(re.search(r"\d+", n).group()) for n in self.parts
                if re.match(r"ppt/media/image\d+\.", n)]
        return "ppt/media/image%d.%s" % (max(nums) + 1 if nums else 1, ext)

    def _content_types_root(self) -> ET.Element:
        return _parse(self.parts["[Content_Types].xml"])

    def _register_media_content_type(self, media_name: str):
        ct = self._content_types_root()
        ext = media_name.rsplit(".", 1)[-1].lower()
        existing_defaults = {d.get("Extension") for d in ct.findall("ct:Default", NS)}
        if ext not in existing_defaults:
            ET.SubElement(ct, _q("ct", "Default"), {
                "Extension": ext, "ContentType": MEDIA_CT.get(ext, "application/octet-stream")})
            self.parts["[Content_Types].xml"] = _serialize(ct, CT_NS)

    def _register_slide_content_type(self, slide_name: str):
        ct = self._content_types_root()
        overrides = {o.get("PartName") for o in ct.findall("ct:Override", NS)}
        part_name = "/" + slide_name
        if part_name not in overrides:
            ET.SubElement(ct, _q("ct", "Override"), {
                "PartName": part_name,
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"})
            self.parts["[Content_Types].xml"] = _serialize(ct, CT_NS)

    def _sync_app_props(self, slide_count: int):
        if "docProps/app.xml" not in self.parts:
            return
        root = _parse(self.parts["docProps/app.xml"])
        slides_el = root.find("{%s}Slides" % EXT_NS)
        if slides_el is not None:
            slides_el.text = str(slide_count)
        # HeadingPairs 中“幻灯片标题 / Slide Titles”计数同步
        hp = root.find("vt:HeadingPairs", NS)
        if hp is not None:
            vec = hp.find("vt:vector", NS)
            if vec is not None:
                pairs = vec.findall("vt:variant", NS)
                for i in range(0, len(pairs) - 1, 2):
                    label = pairs[i].find("vt:lpstr", NS)
                    count = pairs[i + 1].find("vt:i4", NS)
                    if label is not None and count is not None and label.text in (
                            "幻灯片标题", "Slide Titles"):
                        count.text = str(slide_count)
        self.parts["docProps/app.xml"] = _serialize(root, EXT_NS)

    # ---------- 页序 ----------
    def reorder(self, order: list[int]) -> None:
        pres_root = self._pres_root()
        lst = pres_root.find("p:sldIdLst", NS)
        children = list(lst)
        if sorted(order) != list(range(len(children))):
            raise ImporterError("reorder 必须给出当前页的完整排列（长度 %d）" % len(children))
        for child in children:
            lst.remove(child)
        for idx in order:
            lst.append(children[idx])
        self.parts["ppt/presentation.xml"] = _serialize(pres_root)

    # ---------- 保存 ----------
    def save(self, path: str | None = None) -> None:
        out = path or self.dst_path
        ordered = self._ordered_part_names()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for name in ordered:
                zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(zi, self.parts[name])

    def _ordered_part_names(self) -> list[str]:
        head = ["[Content_Types].xml", "_rels/.rels"]
        rest = [n for n in self.parts if n not in head]
        return head + sorted(rest)


# ----------------------------------------------------------------------
# inspect
# ----------------------------------------------------------------------
def inspect(src_pptx: str, src_index: int = 0) -> dict:
    _register_namespaces()
    with zipfile.ZipFile(src_pptx) as z:
        pres = _parse(z.read("ppt/presentation.xml"))
        pres_rels = _parse(z.read("ppt/_rels/presentation.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target")
                   for r in pres_rels.findall("rel:Relationship", NS)}
        slides = [_norm_part("ppt", rel_map[s.get(_q("r", "id"))])
                  for s in pres.find("p:sldIdLst", NS)]
        if not (0 <= src_index < len(slides)):
            raise ImporterError("源幻灯片序号 %d 越界（共 %d 页）" % (src_index, len(slides)))
        slide = slides[src_index]
        root = _parse(z.read(slide))

        shape_text: dict[str, str] = {}
        for shape in root.iter():
            local = etree_local(shape)
            if local not in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame"):
                continue
            cnv = shape.find(".//p:cNvPr", NS)
            if cnv is None:
                continue
            runs = [t.text or "" for t in shape.findall(".//a:t", NS)]
            if runs:
                shape_text[cnv.get("name") or ""] = "\n".join(runs)

        runs = [t.text or "" for t in root.findall(".//a:t", NS)]

        rels_name = "ppt/slides/_rels/" + slide.rsplit("/", 1)[-1] + ".rels"
        media = []
        if rels_name in z.namelist():
            rels = _parse(z.read(rels_name))
            for rel in rels.findall("rel:Relationship", NS):
                if rel.get("Type") == T_IMAGE:
                    mp = _norm_part("ppt/slides", rel.get("Target"))
                    data = z.read(mp)
                    media.append({
                        "part": mp,
                        "sha1": hashlib.sha1(data).hexdigest(),
                        "size": len(data),
                    })

        scheme_count = len(root.findall(".//a:schemeClr", NS))
        return {
            "slide": slide,
            "shape_text": shape_text,
            "runs": runs,
            "media": media,
            "scheme_count": scheme_count,
        }


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def _load_map(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="跨包幻灯片导入工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_in = sub.add_parser("inspect")
    p_in.add_argument("--src", required=True)
    p_in.add_argument("--index", type=int, default=0)

    p_imp = sub.add_parser("import")
    p_imp.add_argument("--src", required=True)
    p_imp.add_argument("--index", type=int, default=0)
    p_imp.add_argument("--dst", required=True)
    p_imp.add_argument("--out")
    p_imp.add_argument("--map")
    p_imp.add_argument("--by-name", action="store_true")

    p_rep = sub.add_parser("replace")
    p_rep.add_argument("--dst", required=True)
    p_rep.add_argument("--index", type=int, required=True)
    p_rep.add_argument("--map", required=True)
    p_rep.add_argument("--by-name", action="store_true")
    p_rep.add_argument("--out")

    p_reo = sub.add_parser("reorder")
    p_reo.add_argument("--dst", required=True)
    p_reo.add_argument("--order", required=True)
    p_reo.add_argument("--out")

    sub.add_parser("selftest")

    args = ap.parse_args(argv)

    if args.cmd == "inspect":
        print(json.dumps(inspect(args.src, args.index), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "import":
        imp = SlideImporter(args.dst)
        mapping = _load_map(args.map) if args.map else None
        idx = imp.import_slide(args.src, args.index, text_map=mapping, by_name=args.by_name)
        imp.save(args.out)
        print("imported slide index:", idx)
        return 0

    if args.cmd == "replace":
        imp = SlideImporter(args.dst)
        imp.replace_text(args.index, _load_map(args.map), by_name=args.by_name)
        imp.save(args.out)
        print("replaced slide", args.index)
        return 0

    if args.cmd == "reorder":
        order = [int(x) for x in args.order.split(",") if x.strip()]
        imp = SlideImporter(args.dst)
        imp.reorder(order)
        imp.save(args.out)
        print("reordered:", order)
        return 0

    if args.cmd == "selftest":
        return _selftest()

    return 2


# ----------------------------------------------------------------------
# 自测：最小构造包导入往返
# ----------------------------------------------------------------------
def _minimal_parts(seed_text: str, with_image: bool) -> dict[str, bytes]:
    parts: dict[str, bytes] = {}
    parts["[Content_Types].xml"] = XML_DECL + (
        '<Types xmlns="%s">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '</Types>' % CT_NS).encode()
    parts["_rels/.rels"] = XML_DECL + (
        '<Relationships xmlns="%s"><Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/></Relationships>' % REL_NS).encode()
    parts["docProps/app.xml"] = XML_DECL + (
        '<Properties xmlns="%s" xmlns:vt="%s"><Slides>1</Slides></Properties>'
        % (EXT_NS, VT_NS)).encode()
    parts["ppt/presentation.xml"] = XML_DECL + (
        '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>'
        % (A_NS, R_NS, P_NS)).encode()
    parts["ppt/_rels/presentation.xml.rels"] = XML_DECL + (
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
        '</Relationships>' % REL_NS).encode()
    parts["ppt/slideMasters/slideMaster1.xml"] = XML_DECL + (
        '<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld/><p:clrMap/></p:sldMaster>'
        % (A_NS, R_NS, P_NS)).encode()
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = XML_DECL + (
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>' % REL_NS).encode()
    parts["ppt/slideLayouts/slideLayout1.xml"] = XML_DECL + (
        '<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld/></p:sldLayout>'
        % (A_NS, R_NS, P_NS)).encode()
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = XML_DECL + (
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>' % REL_NS).encode()
    parts["ppt/theme/theme1.xml"] = XML_DECL + (
        '<a:theme xmlns:a="%s"><a:themeElements><a:clrScheme name="t">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
        '</a:clrScheme></a:themeElements></a:theme>' % A_NS).encode()

    blip = ""
    rels_extra = ""
    if with_image:
        # 1x1 蓝色 jpeg（最小合法字节，自测只校验哈希与引用）
        parts["ppt/media/image1.jpeg"] = bytes(
            [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46])
        blip = '<a:blip r:embed="rId2"/>'
        rels_extra = (
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>')

    slide_xml = (
        '<p:sld xmlns:a="%(a)s" xmlns:r="%(r)s" xmlns:p="%(p)s">'
        '<p:cSld><p:bg><p:bgPr><a:blipFill>%(blip)s<a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>'
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="标题文本框"/></p:nvSpPr>'
        '<p:spPr/><p:txBody><a:p><a:r><a:rPr/><a:t>%(text)s</a:t></a:r></a:p></p:txBody></p:sp>'
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="强调文本框"/></p:nvSpPr>'
        '<p:spPr/><p:txBody><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr>'
        '<a:t>%(color)s</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>'
        '<p:timing><p:tnLst><p:par><p:cTn id="1" presetID="10" presetClass="entr" presetSubtype="0">'
        '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr>'
        '<p:cTn id="2" dur="indefinite" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '<p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        '</p:sld>' % {"a": A_NS, "r": R_NS, "p": P_NS, "blip": blip,
                       "text": seed_text, "color": "色值"})
    parts["ppt/slides/slide1.xml"] = XML_DECL + slide_xml.encode()
    parts["ppt/slides/_rels/slide1.xml.rels"] = XML_DECL + (
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '%s</Relationships>' % (REL_NS, rels_extra)).encode()
    return parts


def _write_pkg(path: str, parts: dict[str, bytes]):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(parts):
            z.writestr(name, parts[name])


def _selftest() -> int:
    import tempfile
    import glob

    tmp = tempfile.mkdtemp(prefix="importer-selftest-")
    dst = os.path.join(tmp, "dst.pptx")
    src = os.path.join(tmp, "src.pptx")
    _write_pkg(dst, _minimal_parts("原生页", False))
    _write_pkg(src, _minimal_parts("源页标题", True))

    imp = SlideImporter(dst)
    idx = imp.import_slide(src, 0, text_map={"源页标题": "导入页标题", "色值": "色值替换"})
    assert idx == 1, idx
    imp.save(os.path.join(tmp, "out.pptx"))

    # 重新打开校验
    out = os.path.join(tmp, "out.pptx")
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        # 1) rels 全部可解析、无孤儿 rId
        for rels_name in [n for n in names if n.endswith(".rels")]:
            root = _parse(z.read(rels_name))
            base = _rels_base_dir(rels_name)
            for rel in root.findall("rel:Relationship", NS):
                if rel.get("TargetMode") == "External":
                    continue
                target = _norm_part(base, rel.get("Target"))
                assert target in names, "孤儿关系: %s -> %s" % (rels_name, target)
        # 2) 导入页无 schemeClr，媒体已复制
        slide2 = z.read("ppt/slides/slide2.xml").decode()
        assert "schemeClr" not in slide2
        assert "srgbClr" in slide2
        media = [n for n in names if n.startswith("ppt/media/")]
        assert len(media) == 1, media
        # 3) 文本替换命中
        assert "导入页标题" in slide2 and "源页标题" not in slide2
        # 4) presentation 注册一致
        pres = _parse(z.read("ppt/presentation.xml"))
        rels = _parse(z.read("ppt/_rels/presentation.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
        slds = pres.find("p:sldIdLst", NS)
        assert len(list(slds)) == 2
        for sld in slds:
            assert _norm_part("ppt", rel_map[sld.get(_q("r", "id"))]) in names
        # 5) app.xml 页数
        app = _parse(z.read("docProps/app.xml"))
        assert app.find("{%s}Slides" % EXT_NS).text == "2"

    # reorder 校验
    imp2 = SlideImporter(out)
    imp2.reorder([1, 0])
    imp2.save(os.path.join(tmp, "out2.pptx"))
    with zipfile.ZipFile(os.path.join(tmp, "out2.pptx")) as z:
        pres = _parse(z.read("ppt/presentation.xml"))
        order = [sld.get(_q("r", "id")) for sld in pres.find("p:sldIdLst", NS)]
        assert len(order) == 2 and order[0] != order[1]

    # 确定性：两次保存字节一致
    imp3 = SlideImporter(dst)
    imp3.import_slide(src, 0)
    imp3.save(os.path.join(tmp, "a.pptx"))
    imp4 = SlideImporter(dst)
    imp4.import_slide(src, 0)
    imp4.save(os.path.join(tmp, "b.pptx"))
    assert open(os.path.join(tmp, "a.pptx"), "rb").read() == open(os.path.join(tmp, "b.pptx"), "rb").read()

    print("selftest OK：rels 无孤儿 rId，媒体去重复制，schemeClr 已解析，文本替换/reorder/确定性保存通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
