# -*- coding: utf-8 -*-
"""
slide_importer.py — 跨包 .pptx 幻灯片导入工具（仅标准库，不依赖 python-pptx）。

用途：把参考 pptx 中现成的高级效果页（形状 + 媒体 + timing/transition 动画节点）
整体手术式复制进目标 deck，重指版式、按内容哈希去重媒体、解析主题色，再做文本替换
与页序调整。timing 与 transition 节点除主题色解析与 rId 改写外原样保留。

公开 API：
    TargetProfile(tone, bg_color, palette, fonts=None, baked_declarations=None, strict=False)
        .to_dict() -> dict；.from_dict(data) -> TargetProfile
    SlideImporter(dst_pptx: str)
        .import_slide(src_pptx, src_index=0, text_map=None, by_name=False) -> int
        .check_slide(src_pptx, src_index=0, target=None, target_texts=None) -> dict
        .adapt_slide(src_pptx, src_index=0, target, text_map=None, recolor_map=None,
                     media_map=None, background=None, strict=None) -> int
        .replace_text(slide_index, mapping, by_name=False) -> None
        .reorder(order: list[int]) -> None
        .save(path=None) -> None
    inspect(src_pptx, src_index=0) -> dict
    工具函数：recolor_srgb(root, mapping)、luminance(hex_color)、hue_family(hex_color)、
        estimate_slot(shape)、center_crop(src_aspect, slot_aspect)、
        detect_background(root)、find_bg_fill_shapes(root)

CLI：
    python3 slide_importer.py inspect --src X.pptx [--index N]
    python3 slide_importer.py check --src X.pptx [--index N] [--profile profile.json]
            [--target-texts map.json]
    python3 slide_importer.py adapt --src X.pptx [--index N] --dst D.pptx
            --profile profile.json [--out O.pptx] [--map map.json]
            [--recolor map.json] [--media map.json] [--strict]
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
import math
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

# 显式不搬运、直接抛错的关系类型（排除层 / 未实现）
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
# 颜色与基调工具
# ----------------------------------------------------------------------
def _norm_hex(color: str) -> str:
    """归一化为 6 位小写十六进制；非法输入返回空串。"""
    if not color:
        return ""
    c = color.strip().lstrip("#")
    if re.fullmatch(r"[0-9A-Fa-f]{6}", c):
        return c.lower()
    if re.fullmatch(r"[0-9A-Fa-f]{3}", c):
        c = "".join(ch * 2 for ch in c)
        return c.lower()
    return ""


def _rgb(hex_color: str) -> tuple[int, int, int]:
    c = _norm_hex(hex_color)
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)) if c else (0, 0, 0)


def luminance(hex_color: str) -> float:
    """相对亮度（Rec.709 线性），返回 0..1；无法解析时返回 0。"""
    r, g, b = _rgb(hex_color)

    def lin(v: float) -> float:
        c = v / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return round(0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b), 6)


def _hsv(r: int, g: int, b: int) -> tuple[float, float, float]:
    rn, gn, bn = r / 255.0, g / 255.0, b / 255.0
    mx, mn = max(rn, gn, bn), min(rn, gn, bn)
    v = mx
    s = 0.0 if mx == 0 else (mx - mn) / mx
    d = mx - mn
    if d == 0:
        h = 0.0
    elif mx == rn:
        h = 60.0 * (((gn - bn) / d) % 6)
    elif mx == gn:
        h = 60.0 * ((bn - rn) / d + 2)
    else:
        h = 60.0 * ((rn - gn) / d + 4)
    return h, s, v


def hue_family(hex_color: str) -> str:
    """归并色相族：red/orange/yellow/green/cyan/blue/purple/neutral。"""
    r, g, b = _rgb(hex_color)
    h, s, v = _hsv(r, g, b)
    if s < 0.18 or v < 0.20:
        return "neutral"
    if h < 15 or h >= 345:
        return "red"
    if h < 45:
        return "orange"
    if h < 75:
        return "yellow"
    if h < 165:
        return "green"
    if h < 200:
        return "cyan"
    if h < 250:
        return "blue"
    return "purple"


def tone_of(hex_color: str) -> str:
    """按相对亮度判 dark / light；中间档返回 unknown。"""
    y = luminance(hex_color)
    if y < 0.35:
        return "dark"
    if y > 0.65:
        return "light"
    return "unknown"


def recolor_srgb(root: ET.Element, mapping: dict) -> int:
    """按 {src_hex: target_hex} 重写 srgbClr 取值；保留全部变换子节点。

    匹配大小写不敏感；返回改写的节点数。未命中的颜色原样保留。
    """
    count = 0
    table = {_norm_hex(k): _norm_hex(v) for k, v in mapping.items()}
    for clr in list(root.iter(_q("a", "srgbClr"))):
        src = _norm_hex(clr.get("val") or "")
        dst = table.get(src)
        if dst:
            clr.set("val", dst)
            count += 1
    return count


# ----------------------------------------------------------------------
# 无第三方依赖的图像尺寸解析（JPEG / PNG / GIF）
# ----------------------------------------------------------------------
def _jpeg_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 4 or data[0] != 0xFF or data[1] != 0xD8:
        return None
    i = 2
    n = len(data)
    while i + 1 < n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        if i + 3 >= n:
            return None
        seg_len = (data[i + 2] << 8) | data[i + 3]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                      0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            if i + 8 < n:
                h = (data[i + 5] << 8) | data[i + 6]
                w = (data[i + 7] << 8) | data[i + 8]
                if w and h:
                    return w, h
            return None
        i += 2 + seg_len
    return None


def _png_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19]
    h = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23]
    return (w, h) if w and h else None


def _gif_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 10 or data[:6] not in (b"GIF87a", b"GIF89a"):
        return None
    w = data[6] | (data[7] << 8)
    h = data[8] | (data[9] << 8)
    return (w, h) if w and h else None


def image_size(data: bytes) -> tuple[int, int] | None:
    """解析 JPEG/PNG/GIF 像素宽高；无法识别返回 None。"""
    return _png_size(data) or _jpeg_size(data) or _gif_size(data)


# ----------------------------------------------------------------------
# 目标画像
# ----------------------------------------------------------------------
class TargetProfile:
    """目标对象画像：基调、背景色、角色配色、字体与烘焙像素声明。

    palette 角色至少支持 bg/card/ink/mut/accent1/accent2/accent3/line；
    未知角色容忍保留。baked_declarations 条目：
    {"media": partname 或 sha1, "baked_text": str, "action": replace|keep|block}。
    strict 为真时未消解的警告升级为阻断。
    """

    ROLES = ("bg", "card", "ink", "mut", "accent1", "accent2", "accent3", "line")

    def __init__(self, tone: str, bg_color: str, palette: dict,
                 fonts: dict | None = None,
                 baked_declarations: list | None = None,
                 strict: bool = False):
        self.tone = tone if tone in ("dark", "light") else "unknown"
        self.bg_color = _norm_hex(bg_color)
        self.palette = {str(k): _norm_hex(v) for k, v in (palette or {}).items()}
        self.fonts = dict(fonts or {})
        self.baked_declarations = list(baked_declarations or [])
        self.strict = bool(strict)

    def declaration_for(self, partname: str, sha1: str) -> dict | None:
        """按 partname 或 sha1 匹配烘焙声明；未命中返回 None。"""
        for decl in self.baked_declarations:
            media = str(decl.get("media") or "")
            if media and (media == partname or media == sha1):
                return decl
        return None

    def to_dict(self) -> dict:
        return {
            "tone": self.tone,
            "bg_color": self.bg_color,
            "palette": self.palette,
            "fonts": self.fonts,
            "baked_declarations": self.baked_declarations,
            "strict": self.strict,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TargetProfile":
        return cls(
            tone=str(data.get("tone") or "unknown"),
            bg_color=str(data.get("bg_color") or ""),
            palette=dict(data.get("palette") or {}),
            fonts=data.get("fonts"),
            baked_declarations=data.get("baked_declarations"),
            strict=bool(data.get("strict", False)),
        )


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
# 槽位 / 裁切 / 背景结构工具
# ----------------------------------------------------------------------
_SHAPE_LOCAL = ("sp", "pic", "grpSp", "cxnSp", "graphicFrame")


def _shape_name(shape: ET.Element) -> str:
    cnv = shape.find(".//p:cNvPr", NS)
    return cnv.get("name") if cnv is not None else ""


def _shape_xfrm(shape: ET.Element) -> tuple[int, int, int, int] | None:
    xfrm = shape.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = shape.find(".//a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    try:
        return (int(off.get("x") or 0), int(off.get("y") or 0),
                int(ext.get("cx") or 0), int(ext.get("cy") or 0))
    except ValueError:
        return None


def _char_weight(ch: str) -> float:
    return 1.0 if ord(ch) > 0x2E7F else 0.55


def _font_pt(shape: ET.Element, body: ET.Element) -> float:
    sizes: list[int] = []
    for rpr in body.iter(_q("a", "rPr")):
        sz = rpr.get("sz")
        if sz and sz.isdigit() and int(sz) >= 400:
            sizes.append(int(sz))
    for rpr in body.iter(_q("a", "endParaRPr")):
        sz = rpr.get("sz")
        if sz and sz.isdigit() and int(sz) >= 400:
            sizes.append(int(sz))
    for rpr in body.iter(_q("a", "defRPr")):
        sz = rpr.get("sz")
        if sz and sz.isdigit() and int(sz) >= 400:
            sizes.append(int(sz))
    if sizes:
        sizes.sort()
        return sizes[len(sizes) // 2] / 100.0
    return 18.0


def estimate_slot(shape: ET.Element) -> dict:
    """估算文本槽位容量。

    返回字体磅值、EMU 框宽高、CJK 容量、估算行数与当前文本；
    容量为离线估算，最终以渲染复核为准。
    """
    body = shape.find("p:txBody", NS)
    if body is None:
        body = shape.find(".//p:txBody", NS)
    xfrm = _shape_xfrm(shape)
    box = list(xfrm[2:4]) if xfrm else [0, 0]
    texts = [t.text or "" for t in shape.findall(".//a:t", NS)]
    text = "".join(texts)
    font_pt = _font_pt(shape, body) if body is not None else 18.0

    body_pr = body.find("a:bodyPr", NS) if body is not None else None
    wrap_none = body_pr is not None and body_pr.get("wrap") == "none"
    line_pct = None
    if body_pr is not None:
        spc = body_pr.find("a:spcPct", NS)
        if spc is not None and spc.get("val"):
            try:
                line_pct = int(spc.get("val")) / 100000.0
            except ValueError:
                line_pct = None
    line_height = font_pt * (line_pct if line_pct else 1.2)

    width_pt = box[0] / 12700.0
    height_pt = box[1] / 12700.0
    weights = sum(_char_weight(ch) for ch in text)
    cpl = max(1, int(width_pt / font_pt)) if font_pt > 0 else 1
    capacity_lines = 1 if wrap_none else max(1, int(height_pt / line_height))
    capacity_cjk = cpl * capacity_lines
    lines = 1 if wrap_none else max(1, math.ceil(weights / cpl)) if cpl else 1

    return {
        "shape": _shape_name(shape),
        "font_pt": round(font_pt, 2),
        "box_emu": box,
        "capacity_cjk": capacity_cjk,
        "lines": lines,
        "wrap_none": wrap_none,
        "text": text,
        "_weights": weights,
        "_cpl": cpl,
        "_capacity_lines": capacity_lines,
    }


def _fits(slot: dict, target_text: str, pt: float) -> bool:
    weights = sum(_char_weight(ch) for ch in target_text)
    width_pt = slot["box_emu"][0] / 12700.0
    cpl = max(1, int(width_pt / pt)) if pt > 0 else 1
    if slot["wrap_none"]:
        return weights <= cpl
    height_pt = slot["box_emu"][1] / 12700.0
    lines_needed = math.ceil(weights / cpl) if cpl else 1
    capacity_lines = max(1, int(height_pt / (pt * 1.2)))
    return lines_needed <= capacity_lines


def _suggested_pt(slot: dict, target_text: str) -> float:
    """返回可容纳目标文本的最大字号（6..48 pt 内二分）；无可行字号返回 0。"""
    lo, hi = 6.0, 48.0
    best = 0.0
    while lo <= hi:
        mid = (lo + hi) / 2.0
        if _fits(slot, target_text, mid):
            best = mid
            lo = mid + 0.5
        else:
            hi = mid - 0.5
    return round(best, 1)


def center_crop(src_aspect: float, slot_aspect: float) -> dict | None:
    """按槽位比例生成居中裁切 srcRect（0..1 分数）；比例接近时返回 None。

    比例容差 5%；srcRect 四边裁切量，写回 OOXML 时换算为千分比。
    """
    if not src_aspect or not slot_aspect:
        return None
    ratio = src_aspect / slot_aspect
    if 0.95 <= ratio <= 1.05:
        return None
    if ratio > 1.0:
        keep_w = slot_aspect / src_aspect
        side = (1.0 - keep_w) / 2.0
        return {"l": round(side, 4), "t": 0.0, "r": round(side, 4), "b": 0.0}
    keep_h = src_aspect / slot_aspect
    top = (1.0 - keep_h) / 2.0
    return {"l": 0.0, "t": round(top, 4), "r": 0.0, "b": round(top, 4)}


def detect_background(root: ET.Element) -> dict:
    """识别页面背景：整页图 / 纯色 / 形状级 blip / 无显式背景。"""
    c_sld = root.find("p:cSld", NS)
    if c_sld is None:
        return {"type": "none", "rid": None, "color": None}
    bg = c_sld.find("p:bg", NS)
    if bg is not None:
        blip_fill = bg.find(".//a:blipFill", NS)
        if blip_fill is not None:
            blip = blip_fill.find("a:blip", NS)
            rid = blip.get(_q("r", "embed")) if blip is not None else None
            return {"type": "slide-bg-blip", "rid": rid, "color": None}
        solid = bg.find(".//a:solidFill", NS)
        if solid is not None:
            srgb = solid.find("a:srgbClr", NS)
            scheme = solid.find("a:schemeClr", NS)
            color = srgb.get("val") if srgb is not None else (
                "scheme:" + scheme.get("val") if scheme is not None else None)
            return {"type": "slide-bg-solid", "rid": None, "color": color}
    # 形状级整页 blip
    for shape in root.iter():
        if etree_local(shape) not in _SHAPE_LOCAL:
            continue
        blip_fill = shape.find(".//a:blipFill", NS)
        if blip_fill is None:
            continue
        xfrm = _shape_xfrm(shape)
        if xfrm and _shape_coverage(xfrm) >= 0.95:
            blip = blip_fill.find("a:blip", NS)
            rid = blip.get(_q("r", "embed")) if blip is not None else None
            return {"type": "shape-blip", "rid": rid, "color": None}
    return {"type": "none", "rid": None, "color": None}


def _shape_coverage(xfrm: tuple[int, int, int, int],
                    slide_w: int = 12192000, slide_h: int = 6858000) -> float:
    _, _, cx, cy = xfrm
    return (cx * cy) / float(slide_w * slide_h)


def find_bg_fill_shapes(root: ET.Element) -> list[str]:
    """列出使用幻灯片背景填充（bgFillStyle / blipFill 内 bg / bgRef）的形状名。"""
    names: list[str] = []
    for shape in root.iter():
        if etree_local(shape) not in _SHAPE_LOCAL:
            continue
        hit = False
        for node in shape.iter():
            local = etree_local(node)
            if local in ("bgFillStyle", "bgFillStyleLst", "bgRef"):
                hit = True
                break
            if local == "blipFill" and node.find("a:bg", NS) is not None:
                hit = True
                break
        if hit:
            names.append(_shape_name(shape))
    return names


def _parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def _ancestor_shape(root: ET.Element, el: ET.Element) -> ET.Element | None:
    pm = _parent_map(root)
    cur: ET.Element | None = el
    while cur in pm:
        cur = pm[cur]
        if cur is not None and etree_local(cur) in _SHAPE_LOCAL:
            return cur
    return None


def _source_context(src_pptx: str, src_index: int) -> dict:
    """读取源包指定页的只读上下文：slide 根、rels、主题色、媒体与不支持关系。"""
    with zipfile.ZipFile(src_pptx) as z:
        names = set(z.namelist())
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

        rels_name = "ppt/slides/_rels/" + slide.rsplit("/", 1)[-1] + ".rels"
        rels_root = _parse(z.read(rels_name)) if rels_name in names else None

        theme_name = "ppt/theme/theme1.xml"
        theme_colors = (_parse_theme_colors(z.read(theme_name))
                        if theme_name in names else {})

        media: dict[str, dict] = {}
        unsupported: list[dict] = []
        if rels_root is not None:
            for rel in rels_root.findall("rel:Relationship", NS):
                rtype = rel.get("Type")
                target = rel.get("Target")
                old_id = rel.get("Id")
                if rtype == T_IMAGE:
                    mp = _norm_part("ppt/slides", target)
                    data = z.read(mp)
                    media[old_id] = {
                        "part": mp,
                        "sha1": hashlib.sha1(data).hexdigest(),
                        "data": data,
                    }
                elif rtype in UNSUPPORTED_TYPES:
                    unsupported.append({"type": UNSUPPORTED_TYPES[rtype], "target": target})
        return {
            "root": root,
            "slide": slide,
            "rels_root": rels_root,
            "theme_colors": theme_colors,
            "media": media,
            "unsupported": unsupported,
        }


def _collect_colors(root: ET.Element, theme_colors: dict) -> tuple[dict, list, list]:
    """统计页面取色（schemeClr 先按源主题解析），返回频次、comp 色、未解析主题色。"""
    counts: dict[str, int] = {}
    comp_colors: list[str] = []
    unresolved: list[str] = []
    for clr in root.iter():
        local = etree_local(clr)
        if local == "srgbClr":
            hexv = _norm_hex(clr.get("val") or "")
        elif local == "schemeClr":
            name = clr.get("val") or ""
            hexv = _norm_hex(theme_colors.get(name, ""))
            if not hexv:
                unresolved.append(name)
        else:
            continue
        if not hexv:
            continue
        counts[hexv] = counts.get(hexv, 0) + 1
        if clr.find("a:comp", NS) is not None:
            comp_colors.append(hexv)
    return counts, comp_colors, unresolved


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

    # ---------- 适配检查 ----------
    def check_slide(self, src_pptx: str, src_index: int = 0,
                    target: TargetProfile | None = None,
                    target_texts: dict | None = None) -> dict:
        """对照目标画像出适配报告；只读源包，不写入目标包。"""
        ctx = _source_context(src_pptx, src_index)
        root = ctx["root"]
        errors: list[str] = []
        warnings: list[str] = []

        for u in ctx["unsupported"]:
            errors.append("不支持的%s关系：%s" % (u["type"], u["target"]))

        text_entries = self._check_text(root, target_texts or {}, errors, warnings)
        counts, comp_colors, unresolved = _collect_colors(root, ctx["theme_colors"])
        background = detect_background(root)
        use_bg_fill = find_bg_fill_shapes(root)
        color_section = self._check_colors(
            counts, comp_colors, unresolved, background, target, errors, warnings)
        media_entries = self._check_media(root, ctx["media"], warnings)
        baked = self._check_baked(root, ctx["media"], background, target, errors, warnings)

        if use_bg_fill:
            warnings.append(
                "形状 %s 使用幻灯片背景填充，替换背景后将继承目标背景" % "、".join(use_bg_fill))

        if background["type"] in ("slide-bg-blip", "shape-blip") \
                and color_section["source_tone"] == "unknown":
            warnings.append("照片背景无法离线判色，基调需转人工核对")

        verdict = self._verdict(errors, warnings, color_section, target)
        return {
            "verdict": verdict,
            "text": text_entries,
            "color": color_section,
            "media": media_entries,
            "background": {"type": background["type"], "useBgFill": use_bg_fill},
            "baked_pixels": baked,
            "errors": errors,
            "warnings": warnings,
        }

    def _check_text(self, root: ET.Element, target_texts: dict,
                    errors: list[str], warnings: list[str]) -> list[dict]:
        entries: list[dict] = []
        for shape in root.iter():
            if etree_local(shape) not in _SHAPE_LOCAL:
                continue
            if shape.find(".//a:t", NS) is None:
                continue
            slot = estimate_slot(shape)
            old = slot["text"]
            new = target_texts.get(old)
            if new is None:
                new = target_texts.get(slot["shape"])
            target_text = new if new is not None else old
            has_box = slot["box_emu"][0] > 0 and slot["box_emu"][1] > 0
            fit = _fits(slot, target_text, slot["font_pt"]) if has_box else True
            suggested = _suggested_pt(slot, target_text) if has_box else slot["font_pt"]
            entries.append({
                "shape": slot["shape"],
                "font_pt": slot["font_pt"],
                "box_emu": slot["box_emu"],
                "capacity_cjk": slot["capacity_cjk"],
                "lines": slot["lines"],
                "target_len": len(target_text),
                "fit": "fit" if fit else "overflow",
                "suggested_pt": suggested,
            })
            if not fit:
                if suggested < 6:
                    errors.append("文本槽位「%s」在 6pt 以上无可行字号" % slot["shape"])
                else:
                    warnings.append(
                        "文本槽位「%s」溢出，建议字号 %.1fpt" % (slot["shape"], suggested))
        return entries

    def _check_colors(self, counts: dict, comp_colors: list, unresolved: list,
                      background: dict, target: TargetProfile | None,
                      errors: list[str], warnings: list[str]) -> dict:
        dominant = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))[:5]
        dominant_list = [{"hex": h, "count": c, "luminance": luminance(h)}
                          for h, c in dominant]

        source_tone = "unknown"
        if background["type"] == "slide-bg-solid" and background["color"]:
            if not background["color"].startswith("scheme:"):
                source_tone = tone_of(background["color"])
        target_tone = target.tone if target else None

        conflicts: list[dict] = []
        if target and source_tone in ("dark", "light") and target_tone in ("dark", "light") \
                and source_tone != target_tone:
            conflicts.append({"role": "bg", "source": background["color"],
                              "target": target.bg_color})

        recolor_map = self._auto_recolor(counts, target)
        if conflicts and not recolor_map:
            errors.append("基调冲突且无同色相 recolor 路径")
        elif conflicts:
            warnings.append("基调冲突，已生成 recolor 建议映射")

        for h in sorted(set(comp_colors)):
            warnings.append("颜色 %s 含 a:comp 补色子节点，未改写其变换链" % h)
        for name in sorted(set(unresolved)):
            warnings.append("主题色 %s 未在源主题中解析到" % name)

        return {
            "source_tone": source_tone,
            "target_tone": target_tone,
            "dominant": dominant_list,
            "conflicts": conflicts,
            "recolor_map": recolor_map,
        }

    @staticmethod
    def _auto_recolor(counts: dict, target: TargetProfile | None) -> dict:
        if not target or not target.palette:
            return {}
        roles = [(role, hexv, hue_family(hexv), luminance(hexv))
                 for role, hexv in target.palette.items() if hexv]
        recolor_map: dict[str, str] = {}
        for src_hex, _ in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
            fam = hue_family(src_hex)
            src_lum = luminance(src_hex)
            candidates = [r for r in roles if r[2] == fam]
            if not candidates:
                continue
            candidates.sort(key=lambda r: (abs(r[3] - src_lum), r[0], r[1]))
            recolor_map[src_hex] = candidates[0][1]
        return recolor_map

    def _check_media(self, root: ET.Element, media: dict,
                     warnings: list[str]) -> list[dict]:
        entries: list[dict] = []
        seen: set[str] = set()
        for blip in root.iter(_q("a", "blip")):
            rid = blip.get(_q("r", "embed"))
            if rid not in media or rid in seen:
                continue
            seen.add(rid)
            info = media[rid]
            size = image_size(info["data"])
            src_aspect = (size[0] / size[1]) if size and size[1] else None
            shape = _ancestor_shape(root, blip)
            xfrm = _shape_xfrm(shape) if shape is not None else None
            slot_aspect = (xfrm[2] / xfrm[3]) if xfrm and xfrm[3] else None
            if size is None:
                warnings.append("媒体 %s 尺寸无法解析，跳过画幅核对" % info["part"])
                entries.append({"part": info["part"], "sha1": info["sha1"],
                                "image_aspect": None, "slot_aspect": slot_aspect,
                                "verdict": "unknown", "srcRect": None})
                continue
            crop = center_crop(src_aspect, slot_aspect)
            entries.append({
                "part": info["part"],
                "sha1": info["sha1"],
                "image_aspect": round(src_aspect, 4),
                "slot_aspect": round(slot_aspect, 4) if slot_aspect else None,
                "verdict": "match" if crop is None else "crop",
                "srcRect": crop,
            })
        return entries

    def _check_baked(self, root: ET.Element, media: dict, background: dict,
                     target: TargetProfile | None, errors: list[str],
                     warnings: list[str]) -> list[dict]:
        baked: list[dict] = []
        seen: set[str] = set()
        if background["type"] == "slide-bg-blip" and background["rid"] in media:
            info = media[background["rid"]]
            baked.append(self._baked_entry(info, "整页背景图", target,
                                            seen, errors, warnings))
        for shape in root.iter():
            if etree_local(shape) not in _SHAPE_LOCAL:
                continue
            blip_fill = shape.find(".//a:blipFill", NS)
            if blip_fill is None:
                continue
            blip = blip_fill.find("a:blip", NS)
            rid = blip.get(_q("r", "embed")) if blip is not None else None
            if rid not in media:
                continue
            xfrm = _shape_xfrm(shape)
            if xfrm is None or _shape_coverage(xfrm) < 0.95:
                continue
            baked.append(self._baked_entry(media[rid], "覆盖整页的形状图片", target,
                                           seen, errors, warnings))
        return baked

    def _baked_entry(self, info: dict, reason: str, target: TargetProfile | None,
                     seen: set[str], errors: list[str], warnings: list[str]) -> dict:
        decl = target.declaration_for(info["part"], info["sha1"]) if target else None
        declared = decl is not None
        action = decl.get("action") if decl else None
        entry = {"media": info["part"], "reason": reason,
                 "declared": declared, "action": action}
        if not declared:
            if info["sha1"] not in seen:
                warnings.append("未声明的烘焙像素：%s（%s）" % (info["part"], reason))
                errors.append("未声明的烘焙像素 %s，已阻断" % info["part"])
            seen.add(info["sha1"])
        elif action == "block":
            errors.append("烘焙像素 %s 声明为 block，已阻断" % info["part"])
        return entry

    @staticmethod
    def _verdict(errors: list[str], warnings: list[str], color_section: dict,
                 target: TargetProfile | None) -> str:
        if errors:
            return "block"
        strict = target.strict if target else False
        if strict and warnings:
            return "block"
        if warnings or color_section["conflicts"]:
            return "adapt"
        return "ok"

    # ---------- 适配写入 ----------
    def adapt_slide(self, src_pptx: str, src_index: int = 0,
                    target: TargetProfile | None = None,
                    text_map: dict | None = None, recolor_map: dict | None = None,
                    media_map: dict | None = None, background: str | None = None,
                    strict: bool | None = None) -> int:
        """先跑检查；未消解冲突在任何写入前阻断，否则导入并依次应用适配动作。"""
        if target is None:
            raise ImporterError("adapt_slide 需要 TargetProfile")
        report = self.check_slide(src_pptx, src_index, target=target,
                                  target_texts=text_map)
        eff_strict = target.strict if strict is None else strict
        if report["verdict"] == "block" or (eff_strict and report["errors"]):
            raise ImporterError(
                "适配未通过检查，已中止写入：%s"
                % "；".join(report["errors"] + report["warnings"]))

        new_index = self.import_slide(src_pptx, src_index)
        slide_name = self._slide_parts_in_order()[new_index]
        root = _parse(self.parts[slide_name])

        rmap = recolor_map or report["color"]["recolor_map"]
        if rmap:
            recolor_srgb(root, rmap)
        if media_map:
            self._apply_media_map(root, slide_name, media_map, report)
        if background is not None:
            self._apply_background(root, slide_name, background)

        self.parts[slide_name] = _serialize(root)
        if text_map:
            self.replace_text(new_index, text_map)
        return new_index

    def _slide_rels_root(self, slide_name: str) -> ET.Element:
        rels_name = "ppt/slides/_rels/" + slide_name.rsplit("/", 1)[-1] + ".rels"
        if rels_name not in self.parts:
            return ET.Element(_q("rel", "Relationships"))
        return _parse(self.parts[rels_name])

    def _rid_to_sha1(self, rels_root: ET.Element) -> dict[str, str]:
        rid_sha1: dict[str, str] = {}
        for rel in rels_root.findall("rel:Relationship", NS):
            if rel.get("Type") != T_IMAGE:
                continue
            part = _norm_part("ppt/slides", rel.get("Target"))
            data = self.parts.get(part)
            if data:
                rid_sha1[rel.get("Id")] = hashlib.sha1(data).hexdigest()
        return rid_sha1

    def _apply_media_map(self, root: ET.Element, slide_name: str,
                          media_map: dict, report: dict) -> None:
        rels_name = "ppt/slides/_rels/" + slide_name.rsplit("/", 1)[-1] + ".rels"
        rels_root = self._slide_rels_root(slide_name)
        current: list[tuple[ET.Element, str, str]] = []
        for rel in rels_root.findall("rel:Relationship", NS):
            if rel.get("Type") != T_IMAGE:
                continue
            part = _norm_part("ppt/slides", rel.get("Target"))
            data = self.parts.get(part)
            digest = hashlib.sha1(data).hexdigest() if data else ""
            current.append((rel, part, digest))

        for key, repl_path in media_map.items():
            match = next((c for c in current
                          if key == c[2] or key == c[1] or key in c[1]), None)
            if match is None:
                continue
            rel, match_part, _ = match
            with open(repl_path, "rb") as f:
                repl = f.read()
            digest = hashlib.sha1(repl).hexdigest()
            if digest in self._media_hash:
                dst_media = self._media_hash[digest]
            else:
                dst_media = self._next_media_name(match_part)
                self.parts[dst_media] = repl
                self._media_hash[digest] = dst_media
                self._register_media_content_type(dst_media)
            rel.set("Target", "../media/" + dst_media.rsplit("/", 1)[-1])

        # 按报告画幅核对结果写居中裁切 srcRect（千分比）
        rid_sha1 = self._rid_to_sha1(rels_root)
        rep_by_sha1 = {m["sha1"]: m for m in report["media"]}
        for blip in root.iter(_q("a", "blip")):
            rid = blip.get(_q("r", "embed"))
            entry = rep_by_sha1.get(rid_sha1.get(rid, ""))
            if not entry or entry["verdict"] != "crop" or not entry["srcRect"]:
                continue
            fill = _ancestor_shape(root, blip)
            blip_fill = None
            for node in fill.iter() if fill is not None else []:
                if etree_local(node) == "blipFill":
                    blip_fill = node
                    break
            if blip_fill is None:
                continue
            old = blip_fill.find("a:srcRect", NS)
            if old is not None:
                blip_fill.remove(old)
            crop = entry["srcRect"]
            src_rect = ET.SubElement(blip_fill, _q("a", "srcRect"))
            for side in ("l", "t", "r", "b"):
                frac = crop.get(side, 0.0)
                src_rect.set(side, str(int(round(frac * 100000))))

        self.parts[rels_name] = _serialize(rels_root, REL_NS)

    def _apply_background(self, root: ET.Element, slide_name: str,
                           background: str) -> None:
        c_sld = root.find("p:cSld", NS)
        if c_sld is None:
            return
        old_bg = c_sld.find("p:bg", NS)
        if old_bg is not None:
            c_sld.remove(old_bg)

        hex_color = _norm_hex(background)
        bg = ET.Element(_q("p", "bg"))
        if hex_color:
            bg_pr = ET.SubElement(bg, _q("p", "bgPr"))
            solid = ET.SubElement(bg_pr, _q("a", "solidFill"))
            ET.SubElement(solid, _q("a", "srgbClr"), {"val": hex_color})
            ET.SubElement(bg_pr, _q("a", "effectLst"))
        else:
            with open(background, "rb") as f:
                data = f.read()
            digest = hashlib.sha1(data).hexdigest()
            if digest in self._media_hash:
                dst_media = self._media_hash[digest]
            else:
                ext = background.rsplit(".", 1)[-1].lower()
                dst_media = self._next_media_name("image." + ext)
                self.parts[dst_media] = data
                self._media_hash[digest] = dst_media
                self._register_media_content_type(dst_media)
            rels_name = "ppt/slides/_rels/" + slide_name.rsplit("/", 1)[-1] + ".rels"
            rels_root = self._slide_rels_root(slide_name)
            new_id = self._next_rel_id(rels_root)
            ET.SubElement(rels_root, _q("rel", "Relationship"), {
                "Id": new_id, "Type": T_IMAGE,
                "Target": "../media/" + dst_media.rsplit("/", 1)[-1],
            })
            self.parts[rels_name] = _serialize(rels_root, REL_NS)
            bg_pr = ET.SubElement(bg, _q("p", "bgPr"))
            blip_fill = ET.SubElement(bg_pr, _q("a", "blipFill"))
            ET.SubElement(blip_fill, _q("a", "blip"), {_q("r", "embed"): new_id})
            stretch = ET.SubElement(blip_fill, _q("a", "stretch"))
            ET.SubElement(stretch, _q("a", "fillRect"))
            ET.SubElement(bg_pr, _q("a", "effectLst"))

        c_sld.insert(0, bg)

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
        media_by_rid: dict[str, dict] = {}
        if rels_name in z.namelist():
            rels = _parse(z.read(rels_name))
            for rel in rels.findall("rel:Relationship", NS):
                if rel.get("Type") == T_IMAGE:
                    mp = _norm_part("ppt/slides", rel.get("Target"))
                    data = z.read(mp)
                    entry = {
                        "part": mp,
                        "sha1": hashlib.sha1(data).hexdigest(),
                        "size": len(data),
                    }
                    media.append(entry)
                    media_by_rid[rel.get("Id")] = entry

        theme_name = "ppt/theme/theme1.xml"
        theme_colors = _parse_theme_colors(z.read(theme_name)) if theme_name in z.namelist() else {}
        counts, _, _ = _collect_colors(root, theme_colors)
        colors = [{"hex": h, "count": c, "tone": tone_of(h), "family": hue_family(h)}
                  for h, c in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))]

        background = detect_background(root)
        use_bg_fill = find_bg_fill_shapes(root)

        baked_candidates: list[dict] = []
        if background["type"] == "slide-bg-blip" and background["rid"] in media_by_rid:
            info = media_by_rid[background["rid"]]
            baked_candidates.append({"part": info["part"], "sha1": info["sha1"],
                                     "reason": "整页背景图"})
        for shape in root.iter():
            if etree_local(shape) not in _SHAPE_LOCAL:
                continue
            blip_fill = shape.find(".//a:blipFill", NS)
            if blip_fill is None:
                continue
            blip = blip_fill.find("a:blip", NS)
            rid = blip.get(_q("r", "embed")) if blip is not None else None
            if rid not in media_by_rid:
                continue
            xfrm = _shape_xfrm(shape)
            if xfrm is not None and _shape_coverage(xfrm) >= 0.95:
                info = media_by_rid[rid]
                baked_candidates.append({"part": info["part"], "sha1": info["sha1"],
                                         "reason": "覆盖整页的形状图片"})

        scheme_count = len(root.findall(".//a:schemeClr", NS))
        return {
            "slide": slide,
            "shape_text": shape_text,
            "runs": runs,
            "media": media,
            "scheme_count": scheme_count,
            "background": background,
            "useBgFill": use_bg_fill,
            "baked_candidates": baked_candidates,
            "colors": colors,
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

    p_chk = sub.add_parser("check")
    p_chk.add_argument("--src", required=True)
    p_chk.add_argument("--index", type=int, default=0)
    p_chk.add_argument("--profile")
    p_chk.add_argument("--target-texts")

    p_ad = sub.add_parser("adapt")
    p_ad.add_argument("--src", required=True)
    p_ad.add_argument("--index", type=int, default=0)
    p_ad.add_argument("--dst", required=True)
    p_ad.add_argument("--profile", required=True)
    p_ad.add_argument("--out")
    p_ad.add_argument("--map")
    p_ad.add_argument("--recolor")
    p_ad.add_argument("--media")
    p_ad.add_argument("--strict", action="store_true")

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
    try:
        return _dispatch(args)
    except ImporterError as e:
        print("错误：%s" % e)
        return 1


def _dispatch(args: argparse.Namespace) -> int:
    if args.cmd == "inspect":
        print(json.dumps(inspect(args.src, args.index), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "check":
        profile = (TargetProfile.from_dict(_load_map(args.profile))
                   if args.profile else None)
        texts = _load_map(args.target_texts) if args.target_texts else None
        imp = SlideImporter(args.src)
        report = imp.check_slide(args.src, args.index, target=profile,
                                 target_texts=texts)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["verdict"] != "block" else 1

    if args.cmd == "adapt":
        profile = TargetProfile.from_dict(_load_map(args.profile))
        imp = SlideImporter(args.dst)
        idx = imp.adapt_slide(
            args.src, args.index, target=profile,
            text_map=_load_map(args.map) if args.map else None,
            recolor_map=_load_map(args.recolor) if args.recolor else None,
            media_map=_load_map(args.media) if args.media else None,
            strict=True if args.strict else None)
        imp.save(args.out)
        print("adapted slide index:", idx)
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


def _adapt_parts(seed_text: str, with_image: bool,
                 bg_color: str | None = None) -> dict[str, bytes]:
    """自测构造包：可控背景（纯色 / 整页图）与带几何框的文本槽。"""
    parts = _minimal_parts(seed_text, with_image)
    if with_image:
        bg_fill = ('<a:blipFill><a:blip r:embed="rId2"/>'
                   '<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
    else:
        bg_fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                   % (bg_color or "FFFFFF"))
    slide_xml = (
        '<p:sld xmlns:a="%(a)s" xmlns:r="%(r)s" xmlns:p="%(p)s">'
        '<p:cSld><p:bg><p:bgPr>%(bg)s<a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="457200" y="457200"/>'
        '<a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400">'
        '<a:solidFill><a:srgbClr val="4472C4"><a:lumMod val="60000"/></a:srgbClr></a:solidFill>'
        '</a:rPr><a:t>%(text)s</a:t></a:r></a:p></p:txBody></p:sp>'
        '</p:spTree></p:cSld></p:sld>'
        % {"a": A_NS, "r": R_NS, "p": P_NS, "bg": bg_fill, "text": seed_text})
    parts["ppt/slides/slide1.xml"] = XML_DECL + slide_xml.encode()
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

    _selftest_adapt(tmp)
    print("selftest OK：rels 无孤儿 rId，媒体去重复制，schemeClr 已解析，文本替换/reorder/确定性保存通过")
    return 0


def _selftest_adapt(tmp: str) -> None:
    # Profile 序列化往返
    prof = TargetProfile(
        "dark", "0B1D33", {"ink": "F2FAFF", "accent1": "3FD0F0"},
        fonts={"title": "Arial"},
        baked_declarations=[{"media": "x", "baked_text": "t", "action": "keep"}],
        strict=True)
    prof2 = TargetProfile.from_dict(prof.to_dict())
    assert prof2.tone == "dark" and prof2.bg_color == "0b1d33" and prof2.strict
    assert prof2.baked_declarations[0]["action"] == "keep"

    # 颜色工具
    assert abs(luminance("FFFFFF") - 1.0) < 1e-6
    assert hue_family("4472C4") == "blue"
    assert hue_family("000000") == "neutral"

    # 居中裁切 srcRect
    wide = center_crop(16 / 9, 4 / 3)
    assert wide["t"] == 0 and wide["b"] == 0 and wide["l"] > 0
    tall = center_crop(4 / 3, 16 / 9)
    assert tall["l"] == 0 and tall["r"] == 0 and tall["t"] > 0
    assert center_crop(1.777, 1.777) is None

    # recolor 往返且变换子节点保留
    root = ET.fromstring(
        '<p:sld xmlns:a="%s" xmlns:p="%s"><p:sp><p:spPr><a:solidFill>'
        '<a:srgbClr val="4472C4"><a:lumMod val="60000"/><a:alpha val="80000"/></a:srgbClr>'
        '</a:solidFill></p:spPr></p:sp></p:sld>' % (A_NS, P_NS))
    assert recolor_srgb(root, {"4472C4": "1F6FEB"}) == 1
    clr = root.find(".//a:srgbClr", NS)
    assert clr.get("val") == "1f6feb"
    assert clr.find("a:lumMod", NS) is not None
    assert clr.find("a:alpha", NS) is not None

    # 基调匹配：纯色浅底 + 充足槽位，结论 ok
    src_ok = os.path.join(tmp, "src_ok.pptx")
    _write_pkg(src_ok, _adapt_parts("适配标题", False, "FFFFFF"))
    prof_light = TargetProfile(
        "light", "FFFFFF",
        {"bg": "FFFFFF", "card": "F2F2F2", "ink": "1A1A1A", "accent1": "1F6FEB"})
    rep = SlideImporter(src_ok).check_slide(src_ok, 0, target=prof_light)
    assert rep["color"]["source_tone"] == "light", rep["color"]
    assert rep["verdict"] == "ok", rep

    # 未声明烘焙像素一律阻断
    src_bg = os.path.join(tmp, "src_bg.pptx")
    _write_pkg(src_bg, _adapt_parts("源页标题", True))
    rep_block = SlideImporter(src_bg).check_slide(
        src_bg, 0, target=TargetProfile("light", "FFFFFF", {}))
    assert rep_block["verdict"] == "block"
    assert any("烘焙像素" in e for e in rep_block["errors"])

    # adapt 全流程：声明 keep + 文本替换 + recolor，包内结构完整
    dst = os.path.join(tmp, "dst2.pptx")
    _write_pkg(dst, _minimal_parts("原生页", False))
    prof_keep = TargetProfile(
        "light", "FFFFFF", {"accent1": "1F6FEB"},
        baked_declarations=[{"media": "ppt/media/image1.jpeg", "baked_text": "",
                             "action": "keep"}])
    imp = SlideImporter(dst)
    idx = imp.adapt_slide(src_bg, 0, target=prof_keep,
                           text_map={"源页标题": "适配标题"})
    assert idx == 1
    out = os.path.join(tmp, "adapt_out.pptx")
    imp.save(out)
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        slide2 = z.read("ppt/slides/slide2.xml").decode()
        assert "适配标题" in slide2 and "源页标题" not in slide2
        assert "1f6feb" in slide2 and "lumMod" in slide2
        assert len([n for n in names if n.startswith("ppt/media/")]) == 1


if __name__ == "__main__":
    raise SystemExit(main())
