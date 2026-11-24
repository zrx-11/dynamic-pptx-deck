#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""20 类页型最小样例命中校验。

读取 `page-type-samples.json`，逐页走 `effect_router.plan_page`，断言：
页型命中期望、候选方案 ≥2 且含 1 个 template 与 1 个 custom、候选字段齐全。
退出码：0=全部通过；1=存在未命中或候选结构缺失。
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "scripts"))
import effect_router as er  # noqa: E402

REQUIRED_ALT_FIELDS = {"scheme_id", "scheme", "summary", "cost", "engine",
                       "tradeoff", "ref"}


def main() -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "page-type-samples.json"), encoding="utf-8") as f:
        data = json.load(f)

    rows = []
    ok = True
    for i, raw in enumerate(data["pages"], start=1):
        feat = er._normalize_page(raw, i)
        plan = er.plan_page(feat)
        alts = plan["alternatives"]
        schemes = {a["scheme"] for a in alts}
        fields_ok = all(REQUIRED_ALT_FIELDS <= set(a) for a in alts)
        valid = (plan["page_type"] == raw["expect"] and len(alts) >= 2
                 and "template" in schemes and "custom" in schemes and fields_ok)
        ok &= valid
        rows.append((i, raw["expect"], plan["page_type"], plan["confidence"],
                     len(alts), ", ".join(plan["basis"]),
                     "OK" if valid else "FAIL"))

    print("| 页 | 期望 | 命中 | 置信 | 候选数 | 依据 | 判定 |")
    print("|---|---|---|---|---|---|---|")
    for r in rows:
        print("| %d | %s | %s | %s | %d | %s | %s |" % r)
    print("ALL OK" if ok else "HAS FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
