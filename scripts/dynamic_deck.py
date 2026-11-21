# -*- coding: utf-8 -*-
"""
dynamic_deck.py — 原生可编辑 .pptx 的对象动画 / 转场 / 形状封装模块。

路线约束（详见 references/effects-catalog.md）：
  - 动画写在每页 slide XML 的 <p:timing> 节点；python-pptx / PptxGenJS 都不封装这层。
  - 三层渐进增强：L0 安全层（2019/365/WPS 通用）、L1 高版本增强（Morph/p14 转场/移植强调，
    不认则静默回退，不丢内容）、L2 排除层（内嵌视频/mediacall、翻书页卷曲、非商业素材）。
  - Morph 用 mc:AlternateContent 包裹：mc:Choice 写 p159:morph，mc:Fallback 写 fade；
    2019 剥离 Choice 后自动走 Fallback。Morph 配对要求：前后两页对应对象同名、同类型。
  - p:animEffect 的 filter 字符串被 PowerPoint 忽略，渲染只由 presetID + presetSubtype 决定；
    不要用 LibreOffice 往返去验证 filter（它原样保留字符串，不做解释）。
  - 状态标签：【实测】= 目标机放映确认；【移植】= 带真实行为节点、未在目标机复测；
    【弃用】= 已确认不生效，不在白名单内。

CLI:
    python3 dynamic_deck.py check <deck.pptx>   # 结构体检；退出码 0=全过、1=有问题、2=包不可读
    python3 dynamic_deck.py demo  <out.pptx>   # 生成两页样例自测
    python3 dynamic_deck.py selftest           # 内联构造器自检（退出 fly / 入场 fly 回归）
"""
from __future__ import annotations

import os
import re
import sys
import zipfile

from lxml import etree
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn, nsdecls

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/12/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"p": P_NS}

# 效果分层：L0 双端安全；L1 高版本增强，不认则回退；L2 排除（不进模块）。
# 分层明细与回退行为见 references/effects-catalog.md 第 10 节。
KINDS_LAYER = {
    "fade": "L0", "fly_up": "L0", "wipe_l": "L0",
    "fly_r": "L0", "fly_l": "L0", "fly_d": "L0",
    "wipe_r": "L0", "wipe_u": "L0", "wipe_d": "L0",
    "appear": "L0", "zoom": "L0",
    "spin": "L0", "grow": "L0", "blink": "L0", "path": "L0", "swing": "L0",
    "zoom_rot": "L1", "bounce_up": "L1", "drop_in": "L1", "grow_once": "L1",
    "zoom_out": "L1",
    "disappear": "L0", "wipe_out_l": "L0", "wipe_out_r": "L0",
    "wipe_out_u": "L0", "wipe_out_d": "L0",
    "fly_out_l": "L0", "fly_out_r": "L0", "fly_out_u": "L0", "fly_out_d": "L0",
}

# kind -> (presetClass, presetID, presetSubtype, filter 字符串仅 LO 参考, 状态标签)
# 方向位标记：1=上 2=右 4=下 8=左（方向位定义）
KINDS = {
    # —— 入场【实测】 ——
    "fade":    ("entr", 10, 0, "fade", "实测"),
    "fly_up":  ("entr", 2,  4, "slide(fromBottom)", "实测"),
    "wipe_l":  ("entr", 22, 8, "wipe(left)", "实测"),
    # —— 入场【移植】 ——
    "fly_r":   ("entr", 2,  2, "slide(fromRight)", "移植"),
    "fly_l":   ("entr", 2,  8, "slide(fromLeft)", "移植"),
    "fly_d":   ("entr", 2,  1, "slide(fromTop)", "移植"),
    "wipe_r":  ("entr", 22, 2, "wipe(right)", "移植"),
    "wipe_u":  ("entr", 22, 1, "wipe(up)", "移植"),
    "wipe_d":  ("entr", 22, 4, "wipe(down)", "移植"),
    "appear":  ("entr", 1,  0, "appear", "移植"),
    "zoom":    ("entr", 53, 16, "fade", "移植"),
    "zoom_rot": ("entr", 31, 0, "fade", "移植"),
    "bounce_up": ("entr", 37, 0, "fade", "移植"),
    "drop_in": ("entr", 47, 0, "fade", "移植"),
    # —— 强调【实测】 ——
    "spin":    ("emph", 8,  0, "spin", "实测"),
    "swing":   ("emph", 8,  0, "spin", "移植"),
    # —— 强调【移植】 ——
    "grow":    ("emph", 26, 0, "grow", "移植"),
    "grow_once": ("emph", 6, 0, "grow", "移植"),
    "blink":   ("emph", 35, 0, "blink", "移植"),
    # —— 退出【移植】 ——
    "disappear": ("exit", 1, 0, "appear", "移植"),
    "zoom_out": ("exit", 23, 32, "fade", "移植"),
    "wipe_out_l": ("exit", 22, 8, "wipe(left)", "移植"),
    "wipe_out_r": ("exit", 22, 2, "wipe(right)", "移植"),
    "wipe_out_u": ("exit", 22, 1, "wipe(up)", "移植"),
    "wipe_out_d": ("exit", 22, 4, "wipe(down)", "移植"),
    "fly_out_r": ("exit", 2, 2, "slide(toRight)", "移植"),
    "fly_out_l": ("exit", 2, 8, "slide(toLeft)", "移植"),
    "fly_out_u": ("exit", 2, 1, "slide(toTop)", "移植"),
    "fly_out_d": ("exit", 2, 4, "slide(toBottom)", "移植"),
    # —— 路径【移植】 ——
    "path":    ("path", 42, 0, "path", "移植"),
}

# 导入保留预设（跨包导入真实 PowerPoint 导出页时出现）：
# 这些 presetClass/presetID/presetSubtype 属标准 ECMA 预设，随导入页 timing 整体保留，
# 不回写为本模块可 authored 的 kind；校验仅做 schema 放行，不代表建议在新稿中手写。
# 标签：imported-path 运动路径预设（带真实 animMotion）；imported-fade-exit 淡出退出；
#       imported-shell 效果目录第 4 节已判空壳/弱行为，仅为导入包校验放行。
IMPORTED_PRESETS = {
    ("exit", 10, 0): "imported-fade-exit",
    ("path", 0, 0): "imported-custom-path",
    ("path", 35, 0): "imported-path-35",
    ("path", 63, 0): "imported-path-63",
    ("path", 64, 0): "imported-path-64",
    ("emph", 19, 0): "imported-shell-19",
    ("exit", 16, 37): "imported-shell-16",
}

# 已确认弃用，不要加回：
#   box(4) / 旧版 zoom / stretch —— 缺行为节点，放映等同瞬间出现；
#   preset 6/14/16/18/21、entr 23/272 —— animEffect 空壳，filter 被忽略，放映等同出现；
#   翻书页/页面卷曲、内嵌视频/mediacall —— L2 排除层，不进模块。
# Morph 归入 L1：见 _TRANSITIONS["morph"]，mc:Choice/Fallback 包裹，弱引擎走 fade 兜底。

SH: dict = {}


# ----------------------------------------------------------------------
# 形状登记与基础封装
# ----------------------------------------------------------------------
def set_name(shape, name):
    for child in shape._element:
        cnv = child.find(qn("p:cNvPr"))
        if cnv is not None:
            cnv.set("name", name)
            SH[name] = shape
            return
    raise RuntimeError("no cNvPr on shape")


def _nostyle(sh):
    st = sh._element.find(qn("p:style"))
    if st is not None:
        sh._element.remove(st)


def _fonts(run, face):
    rpr = run._r.get_or_add_rPr()
    latin = rpr.find(qn("a:latin"))
    if latin is None:
        latin = etree.SubElement(rpr, qn("a:latin"))
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
        latin.addnext(ea)
    ea.set("typeface", face)
    latin.set("typeface", face)


def text(slide, name, x, y, w, h, paras, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.15, alpha=None):
    """paras: [(text, size, color, face), ...]；text 也可为 run 列表。"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    set_name(tb, name)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (text_, size, color, face) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        runs = text_ if isinstance(text_, list) else [(text_, size, color, face)]
        for t2, s2, c2, f2 in runs:
            r = p.add_run()
            r.text = t2
            r.font.size = Pt(s2)
            r.font.bold = (f2 == "bold")
            r.font.color.rgb = c2
            _fonts(r, f2)
            if alpha is not None:
                rpr = r._r.find(qn("a:rPr"))
                srgb = rpr.find(".//" + qn("a:srgbClr")) if rpr is not None else None
                if srgb is not None:
                    a = etree.SubElement(srgb, qn("a:alpha"))
                    a.set("val", str(int(alpha * 100000)))
    return tb


def rect(slide, name, x, y, w, h, fill=None, line=None, radius=None, lw=1.0, alpha=None):
    st = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
    _nostyle(sh)
    set_name(sh, name)
    if radius is not None:
        sh.adjustments[0] = radius
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if alpha is not None:
            srgb = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            a = etree.SubElement(srgb, qn("a:alpha"))
            a.set("val", str(int(alpha * 100000)))
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def oval(slide, name, x, y, w, h, fill=None, line=None, lw=1.0, alpha=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    _nostyle(sh)
    set_name(sh, name)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if alpha is not None:
            srgb = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            a = etree.SubElement(srgb, qn("a:alpha"))
            a.set("val", str(int(alpha * 100000)))
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def donut(slide, name, cx, cy, d, thick, color, alpha=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.DONUT, Inches(cx - d / 2), Inches(cy - d / 2),
                                Inches(d), Inches(d))
    _nostyle(sh)
    set_name(sh, name)
    sh.adjustments[0] = thick
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    if alpha is not None:
        srgb = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = etree.SubElement(srgb, qn("a:alpha"))
        a.set("val", str(int(alpha * 100000)))
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


# ----------------------------------------------------------------------
# 自转组合体（锁定旋转轴心）
#
# 陀螺旋默认绕对象自身几何中心转；短弧 / 扇面质量分布不对称，枢轴会落在弧段自身中心，
# 转起来偏离圆心。标准纠偏：把自转件包进同心组合体，组内放一个透明整圆辅助形把包围盒
# 撑成对称正方形，陀螺旋挂到组合体上，枢轴 = 组合对称中心 = 转盘圆心。
# ----------------------------------------------------------------------
def _assign_id(slide):
    maxid = 0
    for c in slide.shapes._spTree.iter(qn("p:cNvPr")):
        maxid = max(maxid, int(c.get("id")))
    return maxid + 1


def _to_local(sp, d):
    xfrm = sp._element.spPr.find(qn("a:xfrm"))
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    emu = int(d * 914400)
    off.set("x", "0")
    off.set("y", "0")
    ext.set("cx", str(emu))
    ext.set("cy", str(emu))


def _band_shape(slide, cx, cy, d, deg1, deg2, thick, color, rot,
                kind=MSO_SHAPE.BLOCK_ARC, alpha=None):
    sh = slide.shapes.add_shape(kind, Inches(cx - d / 2), Inches(cy - d / 2),
                                Inches(d), Inches(d))
    _nostyle(sh)
    sh.adjustments[0] = deg1 * 0.6
    sh.adjustments[1] = deg2 * 0.6
    if kind == MSO_SHAPE.BLOCK_ARC:
        sh.adjustments[2] = thick
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    if alpha is not None:
        srgb = sh.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        a = etree.SubElement(srgb, qn("a:alpha"))
        a.set("val", str(int(alpha * 100000)))
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.rotation = rot
    return sh


def spin_group(slide, name, cx, cy, d, band):
    helper = oval(slide, name + "_h", cx - d / 2, cy - d / 2, d, d, fill=None, line=None)
    _to_local(helper, d)
    _to_local(band, d)
    gid = _assign_id(slide)
    emu = int(d * 914400)
    offx = int((cx - d / 2) * 914400)
    offy = int((cy - d / 2) * 914400)
    g = etree.fromstring(
        '<p:grpSp %s><p:nvGrpSpPr><p:cNvPr id="%d" name="%s"/><p:cNvGrpSpPr/>'
        '<p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="%d" cy="%d"/></a:xfrm></p:grpSpPr></p:grpSp>'
        % (nsdecls("p", "a"), gid, name, offx, offy, emu, emu, emu, emu))
    g.append(helper._element)
    g.append(band._element)
    slide.shapes._spTree.append(g)

    class _W:
        pass
    w = _W()
    w.shape_id = gid
    w._element = g
    SH[name] = w
    return w


def spin_arc(slide, name, cx, cy, d, deg1, deg2, thick, color, rot=0.0, alpha=None):
    band = _band_shape(slide, cx, cy, d, deg1, deg2, thick, color, rot,
                       MSO_SHAPE.BLOCK_ARC, alpha)
    return spin_group(slide, name, cx, cy, d, band)


def spin_pie(slide, name, cx, cy, d, deg1, deg2, color, rot=0.0, alpha=None):
    band = _band_shape(slide, cx, cy, d, deg1, deg2, 0.0, color, rot,
                       MSO_SHAPE.PIE, alpha)
    return spin_group(slide, name, cx, cy, d, band)


# ----------------------------------------------------------------------
# timing 树
# ----------------------------------------------------------------------
class Timing:
    def __init__(self):
        self.n = 0

    def nid(self):
        self.n += 1
        return self.n

    # —— 行为节点 ——
    def _set(self, spid, vis, delay=0):
        st = self.nid()
        return ('<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst></p:cTn>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:to><p:strVal val="%s"/></p:to></p:set>'
                ) % (st, delay, spid, vis)

    def _anim_effect(self, spid, filt, dur, transition="in"):
        an = self.nid()
        return ('<p:animEffect transition="%s" filter="%s"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '</p:cBhvr></p:animEffect>') % (transition, filt, an, dur, spid)

    def _fly_anim(self, spid, subtype, dur):
        ax = self.nid()
        ay = self.nid()
        if subtype == 2:      # 从右飞入
            x0, y0 = "1+#ppt_w/2", "#ppt_y"
        elif subtype == 8:    # 从左飞入
            x0, y0 = "-#ppt_w/2", "#ppt_y"
        elif subtype == 1:     # 从上飞入
            x0, y0 = "#ppt_x", "-#ppt_h/2"
        else:                 # 4 从下飞入（实测）
            x0, y0 = "#ppt_x", "1+#ppt_h/2"
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="%s"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>'
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="%s"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (ax, dur, spid, x0, ay, dur, spid, y0)

    def _fly_out_anim(self, spid, subtype, dur):
        """退出 fly：起点为作者位 #ppt_x/#ppt_y，终点到画外；隐藏由 _entrance 在 dur 处闭合。"""
        ax = self.nid()
        ay = self.nid()
        if subtype == 2:      # 向右出画
            x1, y1 = "1+#ppt_w/2", "#ppt_y"
        elif subtype == 8:    # 向左出画
            x1, y1 = "-#ppt_w/2", "#ppt_y"
        elif subtype == 1:    # 向上出画
            x1, y1 = "#ppt_x", "-#ppt_h/2"
        else:                 # 4 向下出画
            x1, y1 = "#ppt_x", "1+#ppt_h/2"
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="%s"/></p:val></p:tav></p:tavLst></p:anim>'
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="%s"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (ax, dur, spid, x1, ay, dur, spid, y1)

    def _zoom_anim(self, spid, dur):
        aw = self.nid()
        ah = self.nid()
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav></p:tavLst></p:anim>'
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (aw, dur, spid, ah, dur, spid)

    def _zoom_out_anim(self, spid, dur):
        aw = self.nid()
        ah = self.nid()
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav></p:tavLst></p:anim>'
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (aw, dur, spid, ah, dur, spid)

    def _zoomrot_anim(self, spid, dur):
        rot = self.nid()
        return self._zoom_anim(spid, dur) + (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr>'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:fltVal val="90"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (rot, dur, spid)

    def _bounce_up(self, spid, dur):
        ay1 = self.nid()
        ay2 = self.nid()
        settle = max(100, dur - 120)
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" decel="100000" fill="hold"/>'
            '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y+1"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y-.03"/></p:val></p:tav></p:tavLst></p:anim>'
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="100" accel="100000" fill="hold">'
            '<p:stCondLst><p:cond delay="%d"/></p:stCondLst></p:cTn>'
            '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y-.03"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (ay1, dur, spid, ay2, settle, spid)

    def _drop_in(self, spid, dur):
        ay = self.nid()
        return (
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y-.1"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        ) % (ay, dur, spid)


    def _entrance(self, spid, kind, delay, dur, nt):
        cls, preset, subtype, filt, _ = KINDS[kind]
        eff = self.nid()
        if kind == "appear":
            body = self._set(spid, "visible")
        elif kind.startswith("fly_out"):
            body = (self._set(spid, "visible")
                    + self._fly_out_anim(spid, subtype, dur)
                    + self._set(spid, "hidden", delay=dur))
        elif kind.startswith("fly"):
            body = self._set(spid, "visible") + self._fly_anim(spid, subtype, dur)
        elif kind == "zoom":
            body = (self._set(spid, "visible") + self._zoom_anim(spid, dur)
                    + self._anim_effect(spid, filt, dur))
        elif kind == "zoom_rot":
            body = (self._set(spid, "visible") + self._zoomrot_anim(spid, dur)
                    + self._anim_effect(spid, filt, dur))
        elif kind == "bounce_up":
            body = self._set(spid, "visible") + self._bounce_up(spid, dur) \
                + self._anim_effect(spid, filt, dur)
        elif kind == "drop_in":
            body = self._set(spid, "visible") + self._drop_in(spid, dur) \
                + self._anim_effect(spid, filt, dur)
        elif kind.startswith("wipe_out"):
            body = self._set(spid, "hidden") + self._anim_effect(spid, filt, dur, "out")
        elif kind == "zoom_out":
            body = self._zoom_out_anim(spid, dur) + self._set(spid, "hidden")
        elif kind == "disappear":
            body = self._set(spid, "hidden")
        else:
            body = self._set(spid, "visible") + self._anim_effect(spid, filt, dur)
        return ('<p:par><p:cTn id="%d" presetID="%d" presetClass="%s" presetSubtype="%d" '
                'fill="hold" grpId="0" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>%s'
                '</p:childTnLst></p:cTn></p:par>'
                ) % (eff, preset, cls, subtype, nt, delay, body)

    def _spin(self, spid, dur, delay, nt, by=21600000, auto=False, auto_rev=False):
        eff = self.nid()
        inner = self.nid()
        rep = ' repeatCount="indefinite"' if auto else ''
        rev = ' autoRev="1"' if auto_rev else ''
        return ('<p:par><p:cTn id="%d" presetID="8" presetClass="emph" presetSubtype="0" '
                'fill="hold" grpId="0" nodeType="%s"%s>'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:animRot by="%d"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d" fill="hold"%s%s/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>'
                '</p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>'
                ) % (eff, nt, rep, delay, by, inner, dur, rep, rev, spid)

    def _grow(self, spid, delay, dur, nt):
        eff = self.nid()
        inner = self.nid()
        return ('<p:par><p:cTn id="%d" presetID="26" presetClass="emph" presetSubtype="0" '
                'fill="hold" grpId="0" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:animScale><p:cBhvr><p:cTn id="%d" dur="%d" autoRev="1" fill="hold"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr>'
                '<p:by x="105000" y="105000"/></p:animScale>'
                '</p:childTnLst></p:cTn></p:par>'
                ) % (eff, nt, delay, inner, dur, spid)

    def _grow_once(self, spid, delay, dur, nt):
        eff = self.nid()
        inner = self.nid()
        return ('<p:par><p:cTn id="%d" presetID="6" presetClass="emph" presetSubtype="0" '
                'fill="hold" grpId="0" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:animScale><p:cBhvr><p:cTn id="%d" dur="%d" fill="hold"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr>'
                '<p:by x="110000" y="110000"/></p:animScale>'
                '</p:childTnLst></p:cTn></p:par>'
                ) % (eff, nt, delay, inner, dur, spid)

    def _blink(self, spid, delay, dur, nt):
        eff = self.nid()
        inner = self.nid()
        return ('<p:par><p:cTn id="%d" presetID="35" presetClass="emph" presetSubtype="0" '
                'fill="hold" grpId="1" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:anim calcmode="discrete" valueType="str"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d" fill="hold"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:tavLst>'
                '<p:tav tm="0"><p:val><p:strVal val="hidden"/></p:val></p:tav>'
                '<p:tav tm="50000"><p:val><p:strVal val="visible"/></p:val></p:tav>'
                '</p:tavLst></p:anim></p:childTnLst></p:cTn></p:par>'
                ) % (eff, nt, delay, inner, dur, spid)

    def _path(self, spid, delay, dur, dx, dy, nt, path=None):
        eff = self.nid()
        inner = self.nid()
        if path:
            motion_path = path.strip() + " "
            # 点数 = 坐标对数量：M 起点 1 个，每个 L 再增加 1 个
            n_points = 1 + motion_path.count(" L ")
        else:
            motion_path = "M 0 0 L %.5f %.5f " % (dx, dy)
            n_points = 2
        pts_types = "A" * n_points
        return ('<p:par><p:cTn id="%d" presetID="42" presetClass="path" presetSubtype="0" '
                'repeatCount="indefinite" autoRev="1" fill="hold" grpId="0" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:animMotion origin="layout" path="%s" pathEditMode="relative" '
                'rAng="0" ptsTypes="%s"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d" fill="hold" repeatCount="indefinite" autoRev="1"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>'
                '</p:childTnLst></p:cTn></p:par>'
                ) % (eff, nt, delay, motion_path, pts_types, inner, dur, spid)

    def build(self, groups, auto=0):
        root = self.nid()
        seq = self.nid()
        body = ""
        bld = []
        for gi, grp in enumerate(groups):
            gid = self.nid()
            bid = self.nid()
            first_delay = "0" if gi < auto else "indefinite"
            parts = []
            for idx, item in enumerate(grp):
                nt = "clickEffect" if (idx == 0 and gi >= auto) else "withEffect"
                name = item[0]
                spid = SH[name].shape_id
                kind = item[1]
                if kind == "spin":
                    parts.append(self._spin(spid, item[3], item[2], nt,
                                            by=item[4] if len(item) > 4 else 21600000,
                                            auto=(gi < auto)))
                    bld.append('<p:bldP spid="%d" grpId="0" animBg="1"/>' % spid)
                elif kind == "swing":
                    parts.append(self._spin(spid, item[3], item[2], nt,
                                            by=item[4] if len(item) > 4 else 3600000,
                                            auto=True, auto_rev=True))
                    bld.append('<p:bldP spid="%d" grpId="0" animBg="1"/>' % spid)
                elif kind == "grow":
                    parts.append(self._grow(spid, item[2], item[3], nt))
                    bld.append('<p:bldP spid="%d" grpId="0"/>' % spid)
                elif kind == "grow_once":
                    parts.append(self._grow_once(spid, item[2], item[3], nt))
                    bld.append('<p:bldP spid="%d" grpId="0"/>' % spid)
                elif kind == "blink":
                    parts.append(self._blink(spid, item[2], item[3], nt))
                    bld.append('<p:bldP spid="%d" grpId="0"/>' % spid)
                elif kind == "path":
                    path = item[6] if len(item) > 6 else None
                    parts.append(self._path(spid, item[2], item[3], item[4], item[5], nt, path=path))
                    bld.append('<p:bldP spid="%d" grpId="0"/>' % spid)
                else:
                    _, k, d, du = item
                    parts.append(self._entrance(spid, k, d, du, nt))
                    bld.append('<p:bldP spid="%d" grpId="0"/>' % spid)
            body += ('<p:par><p:cTn id="%d" fill="hold">'
                     '<p:stCondLst><p:cond delay="%s"/></p:stCondLst>'
                     '<p:childTnLst><p:par><p:cTn id="%d" fill="hold">'
                     '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                     '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
                     '</p:childTnLst></p:cTn></p:par>') % (gid, first_delay, bid, "".join(parts))
        return ('<p:timing %s><p:tnLst><p:par>'
                '<p:cTn id="%d" dur="indefinite" restart="never" nodeType="tmRoot">'
                '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
                '<p:cTn id="%d" dur="indefinite" nodeType="mainSeq">'
                '<p:childTnLst>%s</p:childTnLst></p:cTn>'
                '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
                '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
                '<p:bldLst>%s</p:bldLst></p:timing>'
                ) % (nsdecls("p"), root, seq, body, "".join(bld))


def animate(slide, groups, auto=0):
    """groups: 点击批次列表；每批是一组 item 元组。auto=前 N 批进页即自动播。"""
    tb = Timing()
    slide._element.append(etree.fromstring(tb.build(groups, auto)))


# ----------------------------------------------------------------------
# 转场
# ----------------------------------------------------------------------
_TRANSITIONS = {
    # ECMA 基础（2019 通用）
    "fade":    '<p:transition %s spd="med"><p:fade/></p:transition>',
    "push_l":  '<p:transition %s spd="med"><p:push dir="l"/></p:transition>',
    "push_r":  '<p:transition %s spd="med"><p:push dir="r"/></p:transition>',
    "cover_l": '<p:transition %s spd="med"><p:cover dir="l"/></p:transition>',
    "cover_r": '<p:transition %s spd="med"><p:cover dir="r"/></p:transition>',
    "wipe_l":  '<p:transition %s spd="med"><p:wipe dir="l"/></p:transition>',
    "wipe_r":  '<p:transition %s spd="med"><p:wipe dir="r"/></p:transition>',
    "wipe_u":  '<p:transition %s spd="med"><p:wipe dir="u"/></p:transition>',
    "wipe_d":  '<p:transition %s spd="med"><p:wipe dir="d"/></p:transition>',
    "split_h": '<p:transition %s spd="med"><p:split orient="horz" dir="out"/></p:transition>',
    "pull":    '<p:transition %s spd="med"><p:pull/></p:transition>',
    "blinds_v": '<p:transition %s spd="med"><p:blinds dir="vert"/></p:transition>',
    "blinds_h": '<p:transition %s spd="med"><p:blinds dir="horz"/></p:transition>',
    "comb":    '<p:transition %s spd="med"><p:comb/></p:transition>',
    "random":  '<p:transition %s spd="med"><p:random/></p:transition>',
    "checker": '<p:transition %s spd="med"><p:checker/></p:transition>',
    "wheel":   '<p:transition %s spd="med"><p:wheel spokes="1"/></p:transition>',
    "wheel4":  '<p:transition %s spd="med"><p:wheel spokes="4"/></p:transition>',
    "wheel8":  '<p:transition %s spd="med"><p:wheel spokes="8"/></p:transition>',
    "wedge":   '<p:transition %s spd="med"><p:wedge/></p:transition>',
    "circle":  '<p:transition %s spd="med"><p:circle/></p:transition>',
    "zoom_t":  '<p:transition %s spd="med"><p:zoom/></p:transition>',
    "randombar_v": '<p:transition %s spd="med"><p:randomBar dir="vert"/></p:transition>',
    # p14（2010+ 预设；2019 不认则回退无转场，不伤内容）
    "doors":   ('<p:transition %s p14:dur="1400" xmlns:p14="%s">'
                '<p14:doors dir="horz"/></p:transition>'),
    "gallery": ('<p:transition %s p14:dur="1400" xmlns:p14="%s">'
                '<p14:gallery dir="l"/></p:transition>'),
    "flip":    ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:flip dir="r"/></p:transition>'),
    "switch":  ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:switch dir="r"/></p:transition>'),
    "prism":   ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:prism/></p:transition>'),
    "conveyor": ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:conveyor dir="l"/></p:transition>'),
    "window":  ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:window dir="vert"/></p:transition>'),
    "honeycomb": ('<p:transition %s p14:dur="1400" xmlns:p14="%s">'
                '<p14:honeycomb/></p:transition>'),
    "reveal":  ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:reveal/></p:transition>'),
    "ripple":  ('<p:transition %s p14:dur="1400" xmlns:p14="%s">'
                '<p14:ripple/></p:transition>'),
    "vortex":  ('<p:transition %s p14:dur="1200" xmlns:p14="%s">'
                '<p14:vortex dir="r"/></p:transition>'),
    "glitter": ('<p:transition %s p14:dur="1400" xmlns:p14="%s">'
                '<p14:glitter pattern="hexagon"/></p:transition>'),
    "ferris":  ('<p:transition %s p14:dur="1600" xmlns:p14="%s">'
                '<p14:ferris dir="l"/></p:transition>'),
}

_P14 = {"doors", "gallery", "flip", "switch", "prism", "conveyor", "window",
        "honeycomb", "reveal", "ripple", "vortex", "glitter", "ferris"}
# L1 高版本增强转场：不认则回退 Fallback 或无转场，不伤内容。
_L1_TRANS = _P14 | {"morph"}


def transition(slide, kind, adv_ms=None, on_click=True):
    """添加换页转场。L1 预设（p14 / morph）在弱引擎不认时回退，不影响内容。

    adv_ms: 自动换片间隔毫秒；给值后到点自动翻页。on_click=False 时禁用单击换片。
    """
    if kind not in _TRANSITIONS and kind != "morph":
        raise ValueError("未知转场 %r，可用: %s" % (
            kind, ", ".join(sorted(_TRANSITIONS))))
    if kind == "morph":
        xml = (
            '<mc:AlternateContent xmlns:mc="%s" xmlns:p="%s">'
            '<mc:Choice Requires="p159">'
            '<p:transition spd="med" p14:dur="1500" xmlns:p14="%s" xmlns:p159="%s">'
            '<p159:morph option="byObject"/></p:transition>'
            '</mc:Choice>'
            '<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback>'
            '</mc:AlternateContent>'
        ) % (MC, P_NS, P14, P159)
    elif kind in _P14:
        xml = _TRANSITIONS[kind] % (nsdecls("p"), P14)
    else:
        xml = _TRANSITIONS[kind] % nsdecls("p")
    frag = etree.fromstring(xml)
    for tr in frag.iter(qn("p:transition")):
        if adv_ms is not None:
            tr.set("advTm", str(int(adv_ms)))
        if not on_click:
            tr.set("advClick", "0")
    clr = slide._element.find(qn("p:clrMapOvr"))
    (clr if clr is not None else slide._element).addnext(frag)


# ----------------------------------------------------------------------
# 体检
# ----------------------------------------------------------------------
KNOWN = {(v[0], v[1], v[2]): k for k, v in KINDS.items()}
KNOWN.update(IMPORTED_PRESETS)


def validate(path, verbose=True):
    """结构体检：子元素顺序、cTn id 唯一递增、spTgt 命中真实形状、白名单预设。"""
    result = []
    with zipfile.ZipFile(path) as z:
        names = sorted(
            [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda s: int(re.findall(r"\d+", s)[-1]))
        for name in names:
            root = etree.fromstring(z.read(name))
            order = [etree.QName(c).localname for c in root]
            ctn_ids = [int(i) for i in root.xpath(".//p:cTn/@id", namespaces=NS)]
            spids = [int(i) for i in root.xpath(".//p:spTgt/@spid", namespaces=NS)]
            shape_ids = [int(i) for i in root.xpath(".//p:cNvPr/@id", namespaces=NS)]
            presets = root.xpath(".//p:cTn[@presetID]", namespaces=NS)

            problems = []
            order_norm = ["transition" if etree.QName(c).localname == "AlternateContent"
                          else etree.QName(c).localname for c in root]
            tail = [x for x in order_norm if x not in ("cSld", "clrMapOvr")]
            if tail not in (["transition", "timing"], ["timing"], ["transition"], []):
                problems.append("p:sld 子元素顺序异常: %s" % order)
            if len(set(ctn_ids)) != len(ctn_ids):
                problems.append("cTn id 重复")
            missing = sorted(set(spids) - set(shape_ids))
            if missing:
                problems.append("spTgt 指向不存在的形状 id: %s" % missing)
            for node in presets:
                cls = node.get("presetClass")
                pid = int(node.get("presetID"))
                sub = int(node.get("presetSubtype") or 0)
                if (cls, pid, sub) not in KNOWN:
                    problems.append("未在白名单的 presetClass=%s presetID=%d subtype=%d" % (cls, pid, sub))
            if root.xpath(".//p:timing", namespaces=NS) and not (
                root.xpath(".//p:animEffect", namespaces=NS)
                or root.xpath(".//p:animRot", namespaces=NS)
                or root.xpath(".//p:animScale", namespaces=NS)
                or root.xpath(".//p:animMotion", namespaces=NS)
                or root.xpath(".//p:anim", namespaces=NS)):
                problems.append("有 timing 树但没有可执行行为节点")

            info = {
                "order": order_norm,
                "transition": "transition" in order_norm,
                "effects": len(presets),
                "kinds": [KNOWN.get((n.get("presetClass"), int(n.get("presetID")),
                                    int(n.get("presetSubtype") or 0)),
                                    "preset%s" % n.get("presetID")) for n in presets],
                "ctn": len(ctn_ids),
            }
            result.append((name, info, problems))

    if verbose:
        for name, info, problems in result:
            status = "OK" if not problems else "问题"
            print("%-22s %-3s effects=%-3d ctn=%-3d %s" % (
                os.path.basename(name), status, info["effects"], info["ctn"], info["kinds"]))
            for p in problems:
                print("    ! %s" % p)
    return result


# ----------------------------------------------------------------------
# 自测样例
# ----------------------------------------------------------------------
def _selftest():
    """构造器自检：退出 fly_out_r 的预设 / 画外终值 / dur 处 hidden 闭合；入场 fly_l 回归。"""
    tb = Timing()
    frag = tb._entrance(7, "fly_out_r", 0, 600, "afterEffect")
    root = etree.fromstring('<p:timing xmlns:p="%s">%s</p:timing>' % (P_NS, frag))
    exits = root.xpath(".//p:cTn[@presetClass='exit']", namespaces=NS)
    assert len(exits) == 1, "退出 fly 应生成 1 个 exit 预设节点"
    node = exits[0]
    assert node.get("presetID") == "2" and node.get("presetSubtype") == "2", \
        "退出 fly 应为 exit 2/2，实际 %s/%s" % (node.get("presetID"), node.get("presetSubtype"))

    by_name = {}
    for a in root.xpath(".//p:anim", namespaces=NS):
        name = a.xpath(".//p:attrName/text()", namespaces=NS)[0]
        by_name[name] = a.xpath(".//p:tav/p:val/p:strVal/@val", namespaces=NS)
    assert by_name["ppt_x"] == ["#ppt_x", "1+#ppt_w/2"], by_name["ppt_x"]
    assert by_name["ppt_y"] == ["#ppt_y", "#ppt_y"], by_name["ppt_y"]

    hidden = root.xpath(
        ".//p:set[p:cBhvr/p:attrNameLst/p:attrName/text()='style.visibility']"
        "/p:to/p:strVal[@val='hidden']", namespaces=NS)
    assert len(hidden) == 1, "退出 fly 必须带结束 hidden 节点"
    delay = hidden[0].xpath(
        "ancestor::p:set/p:cBhvr/p:cTn/p:stCondLst/p:cond/@delay", namespaces=NS)
    assert delay == ["600"], "hidden 应在 dur=600 处闭合，实际 %s" % delay
    visible = root.xpath(".//p:set/p:to/p:strVal[@val='visible']", namespaces=NS)
    assert len(visible) == 1, "退出起手应先置 visible"

    # 回归：入场 fly_l 仍为 entr 2/8，起手画外、终值回作者位
    tb2 = Timing()
    frag2 = tb2._entrance(8, "fly_l", 150, 800, "withEffect")
    r2 = etree.fromstring('<p:timing xmlns:p="%s">%s</p:timing>' % (P_NS, frag2))
    e2 = r2.xpath(".//p:cTn[@presetClass='entr']", namespaces=NS)
    assert len(e2) == 1 and e2[0].get("presetSubtype") == "8"
    x2 = [a for a in r2.xpath(".//p:anim", namespaces=NS)
          if a.xpath(".//p:attrName/text()", namespaces=NS)[0] == "ppt_x"][0]
    assert x2.xpath(".//p:tav/p:val/p:strVal/@val", namespaces=NS) == ["-#ppt_w/2", "#ppt_x"]

    print("selftest OK：退出 fly_out_r 预设 / 画外终值 / dur 处 hidden 闭合 / 入场 fly_l 回归均通过")
    return 0


def _demo(out):
    from pptx import Presentation

    BG = RGBColor(0x0B, 0x1D, 0x33)
    INK = RGBColor(0xF2, 0xFA, 0xFF)
    ACC = RGBColor(0x3F, 0xD0, 0xF0)
    CARD = RGBColor(0x10, 0x2A, 0x47)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    s1.background.fill.solid()
    s1.background.fill.fore_color.rgb = BG
    t = text(s1, "title", 0.9, 2.4, 9, 1.2, [("标题先淡入", 40, INK, "bold")])
    bar = rect(s1, "bar", 0.9, 3.9, 4.6, 0.07, fill=ACC)
    d = text(s1, "desc", 0.9, 4.35, 9, 0.5, [("这条线从左擦除展开", 15, ACC, "regular")])
    transition(s1, "fade")
    animate(s1, [[("title", "fade", 120, 700),
                  ("bar", "wipe_l", 300, 600),
                  ("desc", "fade", 600, 500)]], auto=1)

    s2 = prs.slides.add_slide(blank)
    s2.background.fill.solid()
    s2.background.fill.fore_color.rgb = BG
    text(s2, "h", 0.9, 0.8, 9, 0.8, [("三张卡片依次上滑", 26, INK, "bold")])
    cards = []
    for i in range(3):
        c = rect(s2, "c%d" % i, 0.9 + i * 4.0, 2.4, 3.2, 2.6,
                 fill=CARD, line=RGBColor(0x23, 0x47, 0x6E), radius=0.05)
        cards.append(c)
    transition(s2, "push_l")
    animate(s2, [[("h", "fade", 0, 450)]] +
            [[("c%d" % i, "fly_up", 250 + i * 130, 600)] for i in range(3)], auto=1)

    prs.save(out)
    return out


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    if argv[1] == "selftest":
        return _selftest()
    if len(argv) < 3 or argv[1] not in ("check", "demo"):
        print(__doc__)
        return 2
    if argv[1] == "check":
        try:
            rows = validate(argv[2])
        except (OSError, zipfile.BadZipFile, etree.XMLSyntaxError,
                KeyError, ValueError) as e:
            print("包不可读: %s" % e)
            return 2
        bad = [r for r in rows if r[2]]
        print("\n%s" % ("全部通过" if not bad else "存在 %d 处问题" % len(bad)))
        return 0 if not bad else 1
    out = _demo(argv[2])
    print("已生成:", out)
    print("  python3 %s check %s" % (os.path.basename(__file__), out))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
