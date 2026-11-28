# -*- coding: utf-8 -*-
"""grid_flow.py — 兜底页均衡网格排布（python-pptx + EMU 路线）。

从上游 ``layout_flow.py`` 吸收三条规则，坐标一律输出 EMU：

1. **均衡列数**：列数先按 ``max_cols`` 上限取，再按行数回推，不允许
   「能塞 4 列就塞 4 列、最后一行只剩 1 个」。6 项 → 3×2；7 项 → 4+3；
   4 项 → 4×1。
2. **同行等高**：同一行的槽位同一高度；多行时行高统一取均分，不出现
   高低不齐。
3. **余量按比例分配**：剩余高度按行数均摊到各行（取整余数逐行轮补），
   不得把余量全给单独一行拉伸。

坐标口径：1 px @96 DPI = 9525 EMU；本脚本内部与 CLI 出入参均为 EMU，
不保留 px 硬编码画布尺寸。

CLI
    python3 scripts/grid_flow.py --n 7 --area x,y,w,h --max-cols 4 [--json]
    python3 scripts/grid_flow.py --selftest
"""
from __future__ import annotations

import argparse
import json
import sys

# 默认槽间距折算 EMU；默认内容区为整页 EMU（1 px @96 DPI = 9525 EMU）。
DEFAULT_GAP_X = 190500
DEFAULT_GAP_Y = 152400
DEFAULT_AREA_W = 12192000
DEFAULT_AREA_H = 6858000


def balanced_cols(n: int, max_cols: int) -> int:
    """均衡列数：不出现最后一行只剩 1 个。"""
    if n <= max_cols:
        return max(n, 1)
    rows = -(-n // max_cols)          # ceil(n / max_cols)
    return -(-n // rows)              # ceil(n / rows)


def layout(n: int, area: tuple[int, int, int, int],
           gap: tuple[int, int] = (DEFAULT_GAP_X, DEFAULT_GAP_Y),
           max_cols: int = 4) -> dict:
    """返回 ``{cols, rows, rects:[{x,y,w,h}], warnings[]}``，rects 为 EMU。"""
    x, y, w, h = area
    warnings: list[str] = []
    if n <= 0:
        return {"cols": 0, "rows": 0, "rects": [], "warnings": ["空列表"]}

    gx, gy = gap
    cols = balanced_cols(n, max_cols)
    rows = -(-n // cols)
    cw = (w - gx * (cols - 1)) // cols
    if cw <= 0:
        warnings.append("槽宽不足：区域宽 %d EMU 容不下 %d 列" % (w, cols))

    avail_h = h - gy * (rows - 1)
    # 余量按行数均摊；取整余数轮补，不集中到某一行。
    base = [avail_h // rows] * rows
    leftover = avail_h - sum(base)
    j = 0
    while leftover > 0 and base:
        base[j % len(base)] += 1
        leftover -= 1
        j += 1

    rects: list[dict] = []
    cy = y
    for r in range(rows):
        row_count = min(cols, n - r * cols)
        for c in range(row_count):
            rects.append({
                "x": int(x + c * (cw + gx)),
                "y": int(cy),
                "w": int(cw),
                "h": int(base[r]),
            })
        cy += base[r] + gy

    # 均衡自检：cols>1 时最后一行不得只剩 1 个。
    last_row = n - (rows - 1) * cols
    if cols > 1 and last_row == 1:
        warnings.append("末行仅 1 项，列数不均衡（%d 列 × %d 行）" % (cols, rows))

    # 越界自检。
    for rect in rects:
        if rect["x"] + rect["w"] > x + w or rect["y"] + rect["h"] > y + h:
            warnings.append("槽位越出内容区: %s" % rect)
            break

    return {"cols": cols, "rows": rows, "rects": rects, "warnings": warnings}


def _parse_area(text: str) -> tuple[int, int, int, int]:
    parts = [int(v.strip()) for v in text.split(",")]
    if len(parts) != 4:
        raise ValueError("--area 需为 x,y,w,h 四个 EMU 整数")
    return tuple(parts)  # type: ignore[return-value]


def _selftest() -> int:
    failures: list[str] = []
    # 整页 EMU 内容区（见 DEFAULT_AREA_*）。
    area = (0, 0, DEFAULT_AREA_W, DEFAULT_AREA_H)

    cases = {6: (3, 2), 7: (4, 2), 4: (4, 1)}
    for n, (ec, er) in cases.items():
        r = layout(n, area, max_cols=4)
        if (r["cols"], r["rows"]) != (ec, er):
            failures.append("用例 %d 项：应 %d×%d，实际 %d×%d"
                             % (n, ec, er, r["cols"], r["rows"]))
        # 同行等高；行高差不超过 1px 折算的取整误差。
        row_h: dict[int, set[int]] = {}
        for i, rect in enumerate(r["rects"]):
            row_h.setdefault(i // r["cols"], set()).add(rect["h"])
        for rh in row_h.values():
            if len(rh) != 1:
                failures.append("用例 %d 项：同行不等高 %s" % (n, rh))

    # 行分配：6 → 3+3，7 → 4+3。
    expect_rowsplit = {6: [3, 3], 7: [4, 3], 4: [4]}
    for n, split in expect_rowsplit.items():
        r = layout(n, area, max_cols=4)
        actual = [min(r["cols"], n - rr * r["cols"]) for rr in range(r["rows"])]
        if actual != split:
            failures.append("用例 %d 项：行分配应 %s，实际 %s" % (n, split, actual))

    # 反例：朴素「顶格 max_cols」会把 5 项排成 4+1；均衡算法应排成 3+2。
    bad_cols = 4
    naive_rows = -(-5 // bad_cols)
    naive_last = 5 - (naive_rows - 1) * bad_cols
    if naive_last != 1:
        failures.append("反例前提失效：朴素 4 列下 5 项末行应为 1")
    r5 = layout(5, area, max_cols=4)
    last5 = 5 - (r5["rows"] - 1) * r5["cols"]
    if (r5["cols"], r5["rows"], last5) != (3, 2, 2):
        failures.append("反例：5 项应纠正为 3+2，实际 %d×%d 末行 %d"
                         % (r5["cols"], r5["rows"], last5))

    # 余量不得全压给一行：各行高度差 ≤ 1 EMU（取整轮补误差内）。
    r7 = layout(7, area, max_cols=4)
    heights = {rect["h"] for rect in r7["rects"]}
    if max(heights) - min(heights) > 1:
        failures.append("反例：余量集中到单行，行高集合 %s" % heights)

    if failures:
        for f in failures:
            print("FAIL:", f)
        return 1

    print("selftest OK：6 项 3×2、7 项 4+3、4 项 4×1；5 项由朴素 4+1 "
          "纠正为 3+2；同行等高，余量轮补不拉伸单行。")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="兜底页均衡网格排布（EMU）")
    ap.add_argument("--n", type=int, default=0, help="槽位数")
    ap.add_argument("--area", default="0,0,%d,%d" % (DEFAULT_AREA_W,
                                                       DEFAULT_AREA_H),
                    help="内容区 x,y,w,h（EMU）")
    ap.add_argument("--max-cols", type=int, default=4)
    ap.add_argument("--gap", default="%d,%d" % (DEFAULT_GAP_X, DEFAULT_GAP_Y),
                    help="横向,纵向槽间距（EMU）")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()
    if args.n <= 0:
        ap.print_help()
        return 2

    area = _parse_area(args.area)
    gap = tuple(int(v.strip()) for v in args.gap.split(","))
    report = layout(args.n, area, gap=gap, max_cols=args.max_cols)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("%d 项 → %d 列 × %d 行；槽位 %d 个；warnings: %s"
              % (args.n, report["cols"], report["rows"], len(report["rects"]),
                 report["warnings"] or "无"))
        for i, rect in enumerate(report["rects"], start=1):
            print("  #%d x=%d y=%d w=%d h=%d" % (i, rect["x"], rect["y"],
                                               rect["w"], rect["h"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
