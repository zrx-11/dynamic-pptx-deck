# -*- coding: utf-8 -*-
"""textfit.py — 文本框容量闸：断行量测 + 字号/加高自适应 + 空框判定。

路线：python-pptx + EMU（不依赖 SVG 画布坐标）。量测用 PIL + 真实 CJK 字体，
不做「字数 × 固定宽度」估算。

断行规则
    - CJK 按字断行；Latin 按词断行；数字与单位（``239 MB``、``1.4 GB``、
      ``17,374,215``、``24.4%``、``0.9%`` 等）整体不拆。
    - 行首禁则：``，。、；：？！）］｝」』·`` 不落在行首；
      行尾禁则：``（［｛「『`` 不落在行尾。

适配顺序（严格按序，不跳步）
    ① 缩字号至下限（默认 10pt）；
    ② 加高容器（只改高度，不改宽度）；
    ③ 报告 ``reduce_chars`` 并给出目标字数。
    禁止改宽度、禁止改字号下限之外的任何几何。

空框判定
    文本框为空，或字数撑不满一行而容器可容 ≥2 行时，判 ``empty_frame``
    （不是 ok），并给出 ``suggested_chars_min/max``、``capacity_lines``、
    ``suggested_pt``。

CLI
    python3 scripts/textfit.py <pptx> [--page N] [--json out.json]
            [--apply out.pptx]
    python3 scripts/textfit.py --selftest

``--apply`` 只在 ``shrunk`` / ``grown`` 两种动作上真改文件；
``reduce_chars`` 与 ``empty_frame`` 只报告，不改。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from PIL import ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 1 px @96 DPI = 9525 EMU；pt→px 用 96/72。
EMU_PER_PX = 9525.0
PT_TO_PX = 96.0 / 72.0
LINE_GAP = 1.35          # 行高 = 字号 px × 行距系数
PAD_PX = 12.0            # 框内上下左右留白
DEFAULT_PT = 18.0
MIN_PT = 10.0

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 真实字体候选（环境已装）；按顺序取第一个存在的。
FONT_CANDIDATES = [
    ("/home/user/.local/share/fonts/思源黑体 CN Regular.otf", 0),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 2),   # Noto Sans CJK SC
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 7),   # Noto Sans Mono CJK SC
]

NO_LINE_START = set("，。、；：？！）］｝」』·.,.;:!?)]}%”’")
NO_LINE_END = set("（［｛「『([{“‘")

# 数字 + 可选单位整体不拆：千分位逗号 / 小数点 / 百分号 / 常见单位。
UNBREAKABLE = re.compile(
    r"\d{1,3}(?:[.,]\d{3})*(?:\.\d+)?"
    r"(?:\s*(?:%|MB|GB|KB|TB|km|kg|cm|mm|pt|px|s|亿|万|元|人|页|场|天|"
    r"次|年|月|日|个|张|份))?",
    re.IGNORECASE,
)


# ----------------------------------------------------------------------
# 字体与量测
# ----------------------------------------------------------------------
def _font_path() -> tuple[str, int]:
    for path, index in FONT_CANDIDATES:
        if os.path.exists(path):
            return path, index
    raise RuntimeError("未找到可用 CJK 字体：%s" %
                       ", ".join(p for p, _ in FONT_CANDIDATES))


def load_font(pt: float) -> ImageFont.FreeTypeFont:
    path, index = _font_path()
    px = max(1, int(round(pt * PT_TO_PX)))
    return ImageFont.truetype(path, px, index=index)


def line_height_px(pt: float, gap: float = LINE_GAP) -> float:
    return pt * PT_TO_PX * gap


def is_cjk(ch: str) -> bool:
    return ("\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f"
            or "\uff00" <= ch <= "\uffef")


# ----------------------------------------------------------------------
# 分词与断行
# ----------------------------------------------------------------------
def tokenize(text: str) -> list[str]:
    """把文本切成不可再分的排版 token：数字单位整体、CJK 单字、Latin 词。"""
    tokens: list[str] = []
    i = 0
    text = text.replace("\r", "")
    while i < len(text):
        m = UNBREAKABLE.match(text, i)
        if m and m.group(0).strip():
            tokens.append(m.group(0))
            i = m.end()
            continue
        ch = text[i]
        if ch == "\n":
            tokens.append("\n")
            i += 1
        elif is_cjk(ch) or ch in NO_LINE_START or ch in NO_LINE_END or ch == " ":
            tokens.append(ch)
            i += 1
        else:
            j = i
            while j < len(text) and not is_cjk(text[j]) and text[j] not in (
                    " ，。、；：！？()（）「」『』【】《》"):
                j += 1
            tokens.append(text[i:j])
            i = j
    return tokens


def break_lines(text: str, width_px: float,
                font: ImageFont.FreeTypeFont) -> list[str]:
    """贪心断行，带行首 / 行尾禁则。返回行列表（不含空行）。"""
    lines: list[str] = []
    cur = ""
    for tk in tokenize(text):
        if tk == "\n":
            lines.append(cur.rstrip())
            cur = ""
            continue
        trial = cur + tk
        if font.getlength(trial.rstrip()) <= width_px or not cur:
            cur = trial
        else:
            # 行首禁则：标点不能顶到下一行行首，把当前行末字带下来。
            if tk and tk[0] in NO_LINE_START and len(cur) > 1:
                moved = cur[-1]
                cur = cur[:-1]
                lines.append(cur.rstrip())
                cur = moved + tk
            else:
                lines.append(cur.rstrip())
                cur = "" if tk == " " else tk
    if cur.strip():
        lines.append(cur.rstrip())

    # 行尾禁则：行尾不能挂开括号类字符，挪到下一行行首。
    fixed: list[str] = []
    for idx, line in enumerate(lines):
        if line and line[-1] in NO_LINE_END and idx + 1 < len(lines):
            fixed.append(line[:-1].rstrip())
            lines[idx + 1] = line[-1] + lines[idx + 1]
        else:
            fixed.append(line)
    return [ln for ln in fixed if ln.strip()]


# ----------------------------------------------------------------------
# 容量与适配
# ----------------------------------------------------------------------
def _avg_char_width(font: ImageFont.FreeTypeFont) -> float:
    sample = "汉Aa0数测容量框"
    return max(1.0, font.getlength(sample) / len(sample))


def capacity(text: str, w_px: float, h_px: float, pt: float) -> dict:
    """给定字号下的容量测算：行数、所需高度、每行可容字数、可容行数。"""
    font = load_font(pt)
    avail_w = max(1.0, w_px - 2 * PAD_PX)
    avail_h = max(1.0, h_px - 2 * PAD_PX)
    lines = break_lines(text, avail_w, font)
    need_h = len(lines) * line_height_px(pt)
    cpl = max(1, int(avail_w / _avg_char_width(font)))
    cap_lines = max(1, int(avail_h / line_height_px(pt)))
    return {"lines": lines, "need_h": need_h, "chars_per_line": cpl,
            "capacity_lines": cap_lines, "avail_w": avail_w,
            "avail_h": avail_h}


def analyze_empty(text: str, w_px: float, h_px: float, pt: float) -> dict:
    """空框 / 欠满判定。返回 {empty, reason, capacity}。"""
    stripped = text.strip()
    cap = capacity(stripped, w_px, h_px, pt)
    if not stripped:
        return {"empty": True, "reason": "文本框为空", "capacity": cap}
    # 只有一行且该行宽不足槽宽六成，而槽高还能放 ≥2 行 → 欠满空框。
    if (len(cap["lines"]) == 1
            and cap["capacity_lines"] >= 2
            and load_font(pt).getlength(cap["lines"][0])
            < 0.6 * cap["avail_w"]):
        return {"empty": True,
                "reason": "字数不足一行，容器可容 %d 行" % cap["capacity_lines"],
                "capacity": cap}
    return {"empty": False, "reason": "", "capacity": cap}


def fit_text(text: str, w_px: float, h_px: float,
             start_pt: float = DEFAULT_PT, min_pt: float = MIN_PT,
             grow: bool = True) -> dict:
    """对单个文本框做适配，返回结构化 dict（见模块 docstring）。"""
    stripped = text.strip()
    start_pt = max(min_pt, start_pt)
    empty = analyze_empty(stripped, w_px, h_px, start_pt)
    reasons: list[str] = []
    suggestions: dict = {}

    if empty["empty"]:
        cap = empty["capacity"]
        cpl = cap["chars_per_line"]
        cap_lines = cap["capacity_lines"]
        suggestions = {
            "suggested_pt": round(start_pt, 1),
            "capacity_lines": cap_lines,
            "chars_per_line": cpl,
            "suggested_chars_min": max(1, round(cap_lines * cpl * 0.5)),
            "suggested_chars_max": cap_lines * cpl,
        }
        reasons.append(empty["reason"])
        return {"status": "empty_frame", "pt": round(start_pt, 1),
                "lines": cap["lines"], "needed_px": round(cap["need_h"], 1),
                "box_h_emu": int(h_px * EMU_PER_PX),
                "reasons": reasons, "suggestions": suggestions}

    # ① 缩字号至下限。
    pt = start_pt
    chosen = None
    while pt >= min_pt - 1e-9:
        cap = capacity(stripped, w_px, h_px, pt)
        if cap["need_h"] <= cap["avail_h"]:
            chosen = (pt, cap)
            break
        pt -= 0.5
    if chosen is not None:
        pt, cap = chosen
        status = "ok" if abs(pt - start_pt) < 1e-9 else "shrunk"
        if status == "shrunk":
            reasons.append("缩字号 %.1f → %.1f pt 后容纳" % (start_pt, pt))
        return {"status": status, "pt": round(pt, 1), "lines": cap["lines"],
                "needed_px": round(cap["need_h"], 1),
                "box_h_emu": int(h_px * EMU_PER_PX),
                "reasons": reasons, "suggestions": {}}

    # ② 下限字号仍溢出：允许加高容器。
    cap = capacity(stripped, w_px, h_px, min_pt)
    new_h_px = cap["need_h"] + 2 * PAD_PX
    if grow:
        reasons.append("下限 %.1f pt 仍溢出，加高容器至 %.0f px"
                       % (min_pt, new_h_px))
        return {"status": "grown", "pt": round(min_pt, 1),
                "lines": cap["lines"], "needed_px": round(cap["need_h"], 1),
                "box_h_emu": int(new_h_px * EMU_PER_PX),
                "reasons": reasons, "suggestions": {}}

    # ③ 需减字。
    cpl = cap["chars_per_line"]
    target = cap["capacity_lines"] * cpl
    reasons.append("下限 %.1f pt 且不加高仍溢出，需减字" % min_pt)
    suggestions = {"target_chars": target, "current_chars": len(stripped),
                   "chars_per_line": cpl, "capacity_lines": cap["capacity_lines"]}
    return {"status": "reduce_chars", "pt": round(min_pt, 1),
            "lines": cap["lines"], "needed_px": round(cap["need_h"], 1),
            "box_h_emu": int(h_px * EMU_PER_PX),
            "reasons": reasons, "suggestions": suggestions}


# ----------------------------------------------------------------------
# pptx 遍历
# ----------------------------------------------------------------------
def _iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(sh.shapes)
        else:
            yield sh


def _has_blip(shape) -> bool:
    return shape._element.find(".//{%s}blip" % A_NS) is not None


def _is_decorative(shape) -> bool:
    """非文本槽位：箭头、自定义多边形（自由曲线）、图片占位框，不归容量闸管。"""
    name = shape.name or ""
    if "图片占位" in name:
        return True
    geom = shape._element.find(".//{%s}prstGeom" % A_NS)
    prst = geom.get("prst") if geom is not None else ""
    if not prst:  # 无 prstGeom → custGeom 自由多边形
        return True
    if "Arrow" in prst:
        return True
    return False


def _shape_pt(shape) -> float:
    if not shape.has_text_frame:
        return DEFAULT_PT
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if run.font.size is not None:
                return run.font.size.pt
        pPr = p._pPr
        if pPr is not None:
            defRPr = pPr.find("{%s}defRPr" % A_NS)
            if defRPr is not None and defRPr.get("sz"):
                return int(defRPr.get("sz")) / 100.0
    return DEFAULT_PT


def analyze_slide(slide, page_index: int, grow: bool = True) -> list[dict]:
    out: list[dict] = []
    for shape in _iter_shapes(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        if _has_blip(shape) or _is_decorative(shape):
            continue  # 图片槽位 / 箭头 / 自由多边形不归文本容量闸管。
        text = shape.text_frame.text or ""
        if shape.left is None or shape.width is None:
            continue
        w_px = shape.width / EMU_PER_PX
        h_px = shape.height / EMU_PER_PX
        pt = _shape_pt(shape)
        result = fit_text(text, w_px, h_px, start_pt=pt, grow=grow)
        result.update({
            "page": page_index,
            "shape_name": shape.name,
            "left_emu": int(shape.left or 0),
            "top_emu": int(shape.top or 0),
            "width_emu": int(shape.width or 0),
            "height_emu": int(shape.height or 0),
            "chars": len(text.strip()),
        })
        out.append(result)
    return out


def _apply_result(shape, result: dict) -> None:
    """只在 shrunk / grown 上真改：缩字号或加高；其余状态不动几何与文字。"""
    if result["status"] == "shrunk":
        sz = str(int(round(result["pt"] * 100)))
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                run._r.get_or_add_rPr().set("sz", sz)
    elif result["status"] == "grown":
        ext = shape._element.find(".//{%s}xfrm/{%s}ext" % (A_NS, A_NS))
        if ext is not None:
            ext.set("cy", str(result["box_h_emu"]))


def process_deck(path: str, only_page: int = 0, grow: bool = True) -> list[dict]:
    prs = Presentation(path)
    report: list[dict] = []
    for idx, slide in enumerate(prs.slides, start=1):
        if only_page and idx != only_page:
            continue
        report.extend(analyze_slide(slide, idx, grow=grow))
    return report


# ----------------------------------------------------------------------
# 自检
# ----------------------------------------------------------------------
def _selftest() -> int:
    failures: list[str] = []

    # ① CJK 长句在窄槽内需缩字号。
    narrow_w, narrow_h = 150.0, 200.0
    long_cjk = "这是一段用来验证窄槽长句必须缩字号的中文样例，内容足够长，" \
               "在窄宽度下一行放不下，需要把字号往下压才能容纳进给定高度。"
    r1 = fit_text(long_cjk, narrow_w, narrow_h, start_pt=18, min_pt=10)
    if r1["status"] not in ("shrunk", "grown"):
        failures.append("用例①：窄槽长句应 shrunk/grown，实际 %s" % r1["status"])
    if r1["status"] == "shrunk" and not r1["pt"] < 18:
        failures.append("用例①：shrunk 后字号应小于 18pt，实际 %.1f" % r1["pt"])

    # ② 数字 + 单位不被拆开。
    mixed = "容量 239 MB 与 1.4 GB，共 17,374,215 项，占 24.4%，误差 0.9%。"
    font = load_font(18)
    lines = break_lines(mixed, 120.0, font)
    joined = "".join(lines)
    for token in ("239 MB", "1.4 GB", "17,374,215", "24.4%", "0.9%"):
        if token not in joined:
            failures.append("用例②：token %s 被拆开；断行结果 %s" % (token, lines))

    # ③ 空文本框判 empty_frame 且给出字数区间。
    r3 = fit_text("", 500.0, 220.0, start_pt=18)
    if r3["status"] != "empty_frame":
        failures.append("用例③：空框应 empty_frame，实际 %s" % r3["status"])
    s3 = r3["suggestions"]
    if not (s3["suggested_chars_min"] >= 1
            and s3["suggested_chars_max"] >= s3["suggested_chars_min"]
            and s3["capacity_lines"] >= 1):
        failures.append("用例③：字数区间 / 容量行数字段非法：%s" % s3)

    # ④ 行首禁则不落在行首。
    punct = "逗号测试，句号测试。顿号测试、分号测试；问号测试？感叹号测试！"
    lines4 = break_lines(punct, 110.0, load_font(18))
    for ln in lines4:
        if ln and ln[0] in NO_LINE_START:
            failures.append("用例④：行首出现禁则标点 %r，断行 %s" % (ln[0], lines4))

    # 反证：文字充足的框应 ok，不得一律判空。
    ok_text = "第一段说明内容足够填满这一行，第二段补充细节，保证版面不空。"
    r_ok = fit_text(ok_text, 400.0, 90.0, start_pt=18)
    if r_ok["status"] == "empty_frame":
        failures.append("反证：文字充足的框不应判 empty_frame，实际 %s" % r_ok)

    if failures:
        for f in failures:
            print("FAIL:", f)
        return 1

    print("selftest OK：①窄槽长句缩字号 status=%s pt=%.1f；②数字单位整体不拆；"
          "③空框 empty_frame 字数区间 %d–%d、容量 %d 行；④行首无禁则标点；"
          "充足文字样例 status=%s 非空框。"
          % (r1["status"], r1["pt"], s3["suggested_chars_min"],
             s3["suggested_chars_max"], s3["capacity_lines"], r_ok["status"]))
    return 0


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="文本框容量闸：断行量测与空框判定")
    ap.add_argument("deck", nargs="?", help="待分析 .pptx")
    ap.add_argument("--page", type=int, default=0, help="只分析这一页（1 起）")
    ap.add_argument("--json", dest="json_out", help="报告写 JSON")
    ap.add_argument("--apply", dest="apply_out",
                    help="把 shrunk/grown 动作写回新 pptx")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()
    if not args.deck:
        ap.print_help()
        return 2
    if not os.path.exists(args.deck):
        print("输入不可读: %s" % args.deck)
        return 2

    report = process_deck(args.deck, only_page=args.page)

    if args.apply_out:
        prs = Presentation(args.deck)
        by_page = {i: slide for i, slide in enumerate(prs.slides, start=1)}
        for entry in report:
            if entry["status"] not in ("shrunk", "grown"):
                continue
            slide = by_page[entry["page"]]
            for shape in _iter_shapes(slide.shapes):
                if getattr(shape, "has_text_frame", False) \
                        and shape.name == entry["shape_name"]:
                    _apply_result(shape, entry)
                    break
        prs.save(args.apply_out)

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump({"frames": report}, f, ensure_ascii=False, indent=2)

    counts: dict[str, int] = {}
    for e in report:
        counts[e["status"]] = counts.get(e["status"], 0) + 1
    print("分析文本框 %d 个：%s" % (len(report), json.dumps(counts,
                                                           ensure_ascii=False)))
    for e in report:
        tag = {"ok": "OK", "shrunk": "缩字号", "grown": "加高",
               "reduce_chars": "需减字", "empty_frame": "空框"}[e["status"]]
        line = "  第%2d页 [%s] %s (%s)" % (e["page"], tag, e["shape_name"],
                                          "; ".join(e["reasons"]) or "—")
        if e["suggestions"]:
            line += " | 建议 %d–%d 字 / %.1fpt / %d 行" % (
                e["suggestions"].get("suggested_chars_min", 0),
                e["suggestions"].get("suggested_chars_max", 0),
                e["suggestions"].get("suggested_pt", 0),
                e["suggestions"].get("capacity_lines", 0))
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
