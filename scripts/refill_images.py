# -*- coding: utf-8 -*-
"""refill_images.py — 整册换图：只换图片 blob，不动几何 / 动画 / 转场 / 命名。

支持四类图片承载（含组合内部）：
    - ``bg``            幻灯片背景 ``p:bg`` 内的 ``a:blipFill``；
    - ``root_pic``      根级 ``p:pic``；
    - ``group_pic``     组合 ``p:grpSp`` 内的 ``p:pic``；
    - ``root_fill`` / ``group_fill``  ``a:blipFill`` 形状（根级 / 组合内）。

约束
    - 只替换图片关系指向的 blob：重写 ``r:embed``，不改动 ``xfrm`` 几何、
      ``a:srcRect`` 裁剪、动画节点、切换节点；
    - 原对象 ``name`` 保持不变（Morph 靠 name 配对，改名会毁掉转场）；
    - 图片不足时按顺序循环取用，报告循环轮次与重复页。

CLI
    python3 scripts/refill_images.py <pptx> --images a.jpg,b.jpg --out out.pptx
            [--page-range 5-14] [--report-json r.json]
    python3 scripts/refill_images.py --selftest

报告字段：``{page, slot, shape_name, kind, from_image, cycled, cycle}``。
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import shutil
import sys
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _in_group(el) -> bool:
    parent = el.getparent()
    while parent is not None:
        if etree.QName(parent).localname == "grpSp":
            return True
        parent = parent.getparent()
    return False


def _shape_name(el) -> str:
    cnv = el.find(".//" + qn("p:cNvPr"))
    if cnv is not None:
        return cnv.get("name") or ""
    return ""


def slots_of(slide) -> list[dict]:
    """按文档顺序返回本页全部可换图槽位（含组合内部）。"""
    root = slide._element
    slots: list[dict] = []

    bg = root.find(".//" + qn("p:bg"))
    if bg is not None:
        blip = bg.find(".//" + qn("a:blip"))
        if blip is not None:
            slots.append({"kind": "bg", "container": bg, "blip": blip,
                          "name": "(background)"})

    for pic in root.iter(qn("p:pic")):
        blip = pic.find(".//" + qn("a:blip"))
        if blip is not None:
            slots.append({
                "kind": "group_pic" if _in_group(pic) else "root_pic",
                "container": pic, "blip": blip, "name": _shape_name(pic)})

    for sp in root.iter(qn("p:sp")):
        fill = sp.find(".//" + qn("a:blipFill"))
        if fill is None:
            continue
        blip = fill.find(".//" + qn("a:blip"))
        if blip is not None:
            slots.append({
                "kind": "group_fill" if _in_group(sp) else "root_fill",
                "container": sp, "blip": blip, "name": _shape_name(sp)})

    return slots


def _parse_page_range(text: str) -> set[int]:
    text = (text or "").strip()
    if not text:
        return set()
    pages: set[int] = set()
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-", 1)
            pages.update(range(int(lo), int(hi) + 1))
        elif part:
            pages.add(int(part))
    return pages


def refill_deck(src: str, images: list[str], out: str,
                page_range: set[int] | None = None) -> list[dict]:
    """执行换图，返回逐槽位报告。``page_range`` 为空表示整册。"""
    if not images:
        raise SystemExit("至少给一张图")
    for img in images:
        if not os.path.isfile(img):
            raise SystemExit("图片不存在: %s" % img)

    shutil.copy2(src, out)
    deck = Presentation(out)
    report: list[dict] = []
    cursor = 0

    for idx, slide in enumerate(deck.slides, start=1):
        if page_range and idx not in page_range:
            continue
        for slot_no, slot in enumerate(slots_of(slide), start=1):
            image_path = images[cursor % len(images)]
            cycle = cursor // len(images)
            image_part, r_id = slide.part.get_or_add_image_part(image_path)
            slot["blip"].set(qn("r:embed"), r_id)
            report.append({
                "page": idx,
                "slot": slot_no,
                "shape_name": slot["name"],
                "kind": slot["kind"],
                "from_image": os.path.basename(image_path),
                "cycled": cycle > 0,
                "cycle": cycle + 1,
            })
            cursor += 1

    deck.save(out)
    return report


# ----------------------------------------------------------------------
# 自检夹具
# ----------------------------------------------------------------------
def _png_bytes(r: int, g: int, b: int) -> bytes:
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), (r, g, b)).save(buf, format="PNG")
    return buf.getvalue()


def _fixture_package(media: dict[str, bytes]) -> io.BytesIO:
    slide_xml = (
        '<p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
        '<p:cSld><p:bg><p:bgPr><a:blipFill>'
        '<a:blip r:embed="rIdBg"/><a:srcRect l="0" t="0" r="0" b="0"/>'
        '</a:blipFill><a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>'
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="组合"/><p:cNvGrpSpPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000000" y="1000000"/>'
        '<a:ext cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>'
        '<p:pic><p:nvPicPr><p:cNvPr id="3" name="组内图片"/><p:cNvPicPr/>'
        '</p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicGrp"/>'
        '</p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr></p:pic>'
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="组内填充形状"/><p:cNvSpPr/>'
        '</p:nvSpPr><p:spPr><a:blipFill><a:blip r:embed="rIdFill"/>'
        '</a:blipFill><a:xfrm><a:off x="0" y="1200000"/>'
        '<a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        '</p:grpSp>'
        '<p:pic><p:nvPicPr><p:cNvPr id="5" name="根级图片"/><p:cNvPicPr/>'
        '</p:nvPicPr><p:blipFill><a:blip r:embed="rIdPicRoot"/>'
        '</p:blipFill><p:spPr><a:xfrm><a:off x="4000000" y="3000000"/>'
        '<a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr></p:pic>'
        '</p:spTree></p:cSld>'
        '<p:transition spd="med"><p:fade/></p:transition>'
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
        'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"/></p:seq>'
        '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        '</p:sld>' % (P_NS, A_NS, R_NS)
    )
    rels = (
        '<Relationships xmlns="%s">'
        '<Relationship Id="rIdBg" Type="%s/image" Target="../media/image1.png"/>'
        '<Relationship Id="rIdPicGrp" Type="%s/image" Target="../media/image2.png"/>'
        '<Relationship Id="rIdFill" Type="%s/image" Target="../media/image3.png"/>'
        '<Relationship Id="rIdPicRoot" Type="%s/image" Target="../media/image4.png"/>'
        '</Relationships>' % (REL_NS, R_NS, R_NS, R_NS, R_NS))
    ct = (
        '<Types xmlns="%s"><Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '</Types>' % CT_NS)
    root_rels = (
        '<Relationships xmlns="%s"><Relationship Id="rId1" '
        'Type="%s/officeDocument" Target="ppt/presentation.xml"/></Relationships>'
        % (REL_NS, R_NS))
    pres = (
        '<p:presentation xmlns:p="%s" xmlns:r="%s"><p:sldIdLst>'
        '<p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>'
        % (P_NS, R_NS))
    pres_rels = (
        '<Relationships xmlns="%s"><Relationship Id="rId2" '
        'Type="%s/slide" Target="slides/slide1.xml"/></Relationships>'
        % (REL_NS, R_NS))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", XML_DECL + ct.encode("utf-8"))
        zf.writestr("_rels/.rels", XML_DECL + root_rels.encode("utf-8"))
        zf.writestr("ppt/presentation.xml", XML_DECL + pres.encode("utf-8"))
        zf.writestr("ppt/_rels/presentation.xml.rels",
                     XML_DECL + pres_rels.encode("utf-8"))
        zf.writestr("ppt/slides/slide1.xml", XML_DECL + slide_xml.encode("utf-8"))
        zf.writestr("ppt/slides/_rels/slide1.xml.rels",
                    XML_DECL + rels.encode("utf-8"))
        for name, blob in media.items():
            zf.writestr("ppt/media/" + name, blob)
    buf.seek(0)
    return buf


def _snapshot(zf: zipfile.ZipFile) -> dict:
    slide = etree.fromstring(zf.read("ppt/slides/slide1.xml"))
    xfrms = sorted(etree.tostring(x) for x in slide.iter(qn("a:xfrm")))
    names = sorted(c.get("name") or "" for c in slide.iter(qn("p:cNvPr")))
    timing = slide.find(qn("p:timing"))
    transition = slide.find(qn("p:transition"))
    media = {n: zf.read(n) for n in zf.namelist() if n.startswith("ppt/media/")}
    return {"xfrms": xfrms, "names": names,
            "timing": etree.tostring(timing) if timing is not None else None,
            "transition": etree.tostring(transition) if transition is not None else None,
            "media": media}


def _selftest() -> int:
    failures: list[str] = []
    media = {"image1.png": _png_bytes(255, 0, 0),
             "image2.png": _png_bytes(200, 0, 0),
             "image3.png": _png_bytes(150, 0, 0),
             "image4.png": _png_bytes(100, 0, 0)}
    src_buf = _fixture_package(media)
    src_path = os.path.join(os.path.dirname(__file__), "_refill_src.pptx")
    out_path = os.path.join(os.path.dirname(__file__), "_refill_out.pptx")
    with open(src_path, "wb") as f:
        f.write(src_buf.getvalue())

    with zipfile.ZipFile(src_path) as zf:
        before = _snapshot(zf)

    # 只给 1 张新图，4 个槽位 → 后 3 个循环取用。
    new_img = os.path.join(os.path.dirname(__file__), "_refill_new.png")
    with open(new_img, "wb") as f:
        f.write(_png_bytes(0, 200, 0))

    report = refill_deck(src_path, [new_img], out_path)

    with zipfile.ZipFile(out_path) as zf:
        after = _snapshot(zf)

    # ① blob 变了：四个原媒体 blob 不得原样保留。
    if after["media"] == before["media"]:
        failures.append("① 替换后媒体 blob 未变化")

    # ② xfrm 逐字节未变。
    if after["xfrms"] != before["xfrms"]:
        failures.append("② xfrm 几何被改动")

    # ③ shape name 未变。
    if after["names"] != before["names"]:
        failures.append("③ shape name 被改动: %s → %s"
                         % (before["names"], after["names"]))

    # ④ 动画 / 切换节点逐字节未变。
    if after["timing"] != before["timing"]:
        failures.append("④ timing 动画节点被改动")
    if after["transition"] != before["transition"]:
        failures.append("④ transition 切换节点被改动")

    # ⑤ 循环计数正确：4 槽 1 图，slot 1 不循环，slot 2–4 循环。
    if len(report) != 4:
        failures.append("⑤ 应替换 4 个槽位，实际 %d" % len(report))
    cycled = [r["slot"] for r in report if r["cycled"]]
    if cycled != [2, 3, 4]:
        failures.append("⑤ 循环槽位应为 [2,3,4]，实际 %s" % cycled)
    kinds = [r["kind"] for r in report]
    for need in ("bg", "group_pic", "group_fill", "root_pic"):
        if need not in kinds:
            failures.append("⑤ 缺少承载类型 %s，实际 %s" % (need, kinds))

    for p in (src_path, out_path, new_img):
        os.remove(p)

    if failures:
        for f in failures:
            print("FAIL:", f)
        return 1

    print("selftest OK：4 类承载全部替换；blob 已变；xfrm / name / timing / "
          "transition 逐字节未变；4 槽 1 图循环 3 次，重复槽位 [2,3,4]。")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="整册换图：只换 blob，不动几何/动画/命名")
    ap.add_argument("deck", nargs="?", help="待换图 .pptx")
    ap.add_argument("--images", help="逗号分隔的图片路径列表")
    ap.add_argument("--out", help="输出 .pptx")
    ap.add_argument("--page-range", default="", help="页码范围，如 5-14 或 5,8")
    ap.add_argument("--report-json", dest="report_json", help="报告写 JSON")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()
    if not args.deck or not args.images or not args.out:
        ap.print_help()
        return 2
    if not os.path.exists(args.deck):
        print("输入不可读: %s" % args.deck)
        return 2

    images = [p.strip() for p in args.images.split(",") if p.strip()]
    page_range = _parse_page_range(args.page_range)
    report = refill_deck(args.deck, images, args.out, page_range or None)

    if args.report_json:
        with open(args.report_json, "w", encoding="utf-8") as f:
            json.dump({"slots": report}, f, ensure_ascii=False, indent=2)

    cycled = [r["page"] for r in report if r["cycled"]]
    print("替换槽位 %d | 图数 %d | 循环重复页 %s"
          % (len(report), len(images), cycled or "无"))
    for r in report:
        print("  第%2d页 槽%d [%s] %s ← %s%s"
              % (r["page"], r["slot"], r["kind"], r["shape_name"],
                 r["from_image"], "（循环第%d轮）" % r["cycle"] if r["cycled"] else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
