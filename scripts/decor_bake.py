# -*- coding: utf-8 -*-
"""
decor_bake.py — 装饰层预渲染管线（仅 Pillow + 标准库）。

按参数化 spec 渲染渐变 / 光晕 / 噪点 / 噪点+模糊 / 暗角 / 纹理 / 模糊七类装饰图层，
每层出独立 PNG；id 为 bg 的图层参与整幅合成，并额外输出 baked-<sha8>.png。
同一份 spec、同一个 seed 两次运行产物字节级可复现（sha256 一致）。

CLI：
    python3 decor_bake.py --spec SPEC.json --out-dir DIR [--only ID] [--preset 名称]
    python3 decor_bake.py --preset 深底光斑 --out-dir DIR

spec.json：
    {"canvas": [1920, 1080],
     "layers": [
       {"id": "bg",  "type": "gradient",    "params": {"colors": ["#0E3B36", "#071E1C"], "angle": 120, "stops": [0, 1]}},
       {"id": "glow","type": "glow",        "params": {"center": [0.62, 0.38], "radius": 0.55, "color": "#7FE3C8", "strength": 0.55}},
       {"id": "noise","type": "noise",      "params": {"amount": 0.04, "scale": 2, "seed": 7}},
       {"id": "bn",  "type": "blur_noise",  "params": {"amount": 0.05, "scale": 8, "radius": 6, "seed": 11}},
       {"id": "vig", "type": "vignette",    "params": {"strength": 0.35, "radius": 0.75}},
       {"id": "tex", "type": "texture",     "params": {"kind": "grain", "amount": 0.03, "seed": 3}},
       {"id": "blur","type": "blur",        "params": {"radius": 4}}
     ]}

退出码：0 成功；2 用法错误或 spec / 输出目录不可用。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import sys

try:
    from PIL import Image, ImageDraw, ImageFilter
except ImportError:  # pragma: no cover
    Image = ImageDraw = ImageFilter = None

GENERATOR_VERSION = "decor_bake/1"
SUPPORTED_TYPES = ("gradient", "glow", "noise", "blur_noise",
                   "vignette", "texture", "blur")


# ----------------------------------------------------------------------
# 颜色与数值工具
# ----------------------------------------------------------------------
def _hex(color: str) -> tuple[int, int, int]:
    s = str(color).lstrip("#")
    if len(s) != 6:
        raise ValueError("颜色需为 #RRGGBB：%s" % color)
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _clamp01(v: float, default: float) -> float:
    try:
        return min(1.0, max(0.0, float(v)))
    except (TypeError, ValueError):
        return default


def _lerp(a: int, b: int, t: float) -> int:
    return round(a + (b - a) * t)


def _stop_lut(colors: list[str], stops: list[float] | None) -> list[list[int]]:
    """生成 256 级调色板：每通道一张 256 项查找表。"""
    if len(colors) < 2:
        raise ValueError("gradient 至少需要两种颜色")
    rgb = [_hex(c) for c in colors]
    if stops is None or len(stops) != len(rgb):
        stops = [i / (len(rgb) - 1) for i in range(len(rgb))]
    pos = [min(1.0, max(0.0, float(s))) for s in stops]
    lut: list[list[int]] = [[], [], []]
    for v in range(256):
        t = v / 255.0
        seg = 0
        for i in range(len(pos) - 1):
            if pos[i] <= t <= pos[i + 1]:
                seg = i
                break
        else:
            seg = len(pos) - 2
        span = pos[seg + 1] - pos[seg] or 1.0
        local = (t - pos[seg]) / span
        for ch in range(3):
            lut[ch].append(_lerp(rgb[seg][ch], rgb[seg + 1][ch], local))
    return lut


# ----------------------------------------------------------------------
# 图层渲染（每个函数返回与画布同尺寸的 RGBA Image；blur 类型在合成阶段处理）
# ----------------------------------------------------------------------
def _horizontal_ramp() -> Image.Image:
    """256×1 横向灰阶：左 0、右 255。"""
    ramp = Image.new("L", (256, 1))
    ramp.putdata(list(range(256)))
    return ramp


def _radial(center_bright: bool) -> Image.Image:
    """256×256 手工径向灰阶；center_bright=True 中心 255、边缘 0，反之亦然。"""
    import math
    size = 256
    rad = Image.new("L", (size, size))
    px = rad.load()
    c = (size - 1) / 2
    for yy in range(size):
        for xx in range(size):
            d = math.hypot(xx - c, yy - c) / c
            v = max(0.0, 1.0 - d)
            if not center_bright:
                v = 1.0 - v
            px[xx, yy] = int(round(v * 255))
    return rad


def _render_gradient(w: int, h: int, p: dict) -> Image.Image:
    import math
    angle = float(p.get("angle", 90))
    lut = _stop_lut(p["colors"], p.get("stops"))
    # 背板取对角线尺寸，旋转后居中裁剪，避免矩形外区域被填黑。
    diag = int(math.ceil(math.hypot(w, h)))
    ramp = _horizontal_ramp().resize((diag, diag))
    if angle:
        ramp = ramp.rotate(angle, resample=Image.BICUBIC, expand=False, fillcolor=0)
    left, top = (diag - w) // 2, (diag - h) // 2
    ramp = ramp.crop((left, top, left + w, top + h))
    channels = [ramp.point(lut[ch]) for ch in range(3)]
    return Image.merge("RGB", channels).convert("RGBA")


def _render_glow(w: int, h: int, p: dict) -> Image.Image:
    cx, cy = p.get("center", [0.5, 0.5])
    radius = float(p.get("radius", 0.5)) * min(w, h)
    color = _hex(p.get("color", "#FFFFFF"))
    strength = _clamp01(p.get("strength", 0.5), 0.5)
    box = max(1, int(2 * radius))
    radial = _radial(center_bright=True).resize((box, box))
    alpha = radial.point(lambda v: int(v * strength))
    layer = Image.new("RGBA", (w, h), color + (0,))
    glow = Image.new("RGBA", (box, box), color + (0,))
    glow.putalpha(alpha)
    x = int(cx * w - radius)
    y = int(cy * h - radius)
    layer.alpha_composite(glow, (x, y))
    return layer


def _render_noise(w: int, h: int, p: dict, rng: random.Random) -> Image.Image:
    amount = _clamp01(p.get("amount", 0.05), 0.05)
    scale = max(1, int(p.get("scale", 2)))
    gw, gh = max(1, w // scale), max(1, h // scale)
    small = Image.new("L", (gw, gh))
    px = small.load()
    for yy in range(gh):
        for xx in range(gw):
            px[xx, yy] = rng.randint(0, 255)
    noise = small.resize((w, h), Image.NEAREST)
    alpha = noise.point(lambda v: int(v * amount))
    gray = Image.merge("RGB", (noise, noise, noise))
    layer = gray.convert("RGBA")
    layer.putalpha(alpha)
    return layer


def _render_blur_noise(w: int, h: int, p: dict, rng: random.Random) -> Image.Image:
    layer = _render_noise(w, h, p, rng)
    radius = float(p.get("radius", 6))
    return layer.filter(ImageFilter.GaussianBlur(radius))


def _render_vignette(w: int, h: int, p: dict) -> Image.Image:
    strength = _clamp01(p.get("strength", 0.35), 0.35)
    radius = _clamp01(p.get("radius", 0.75), 0.75)
    dark = _radial(center_bright=False).resize((w, h))

    def shade(v: int) -> int:
        t = v / 255.0
        if t <= 1 - radius:
            return 0
        local = (t - (1 - radius)) / radius
        return int(min(1.0, local) * strength * 255)

    alpha = dark.point(shade)
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    layer.putalpha(alpha)
    return layer


def _render_texture(w: int, h: int, p: dict, rng: random.Random) -> Image.Image:
    kind = p.get("kind", "grain")
    amount = _clamp01(p.get("amount", 0.03), 0.03)
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    if kind in ("grain", "paper"):
        base = _render_noise(w, h, {"amount": amount, "scale": 1 if kind == "grain" else 3}, rng)
        if kind == "paper":
            base = base.filter(ImageFilter.GaussianBlur(1.2))
        layer = base
    elif kind == "dots":
        step = 18
        dot_r = 2
        draw = ImageDraw.Draw(layer)
        for yy in range(step, h, step):
            for xx in range(step, w, step):
                jx = rng.randint(-3, 3)
                jy = rng.randint(-3, 3)
                a = int(amount * 255)
                draw.ellipse([xx + jx - dot_r, yy + jy - dot_r,
                              xx + jx + dot_r, yy + jy + dot_r],
                             fill=(255, 255, 255, a))
    else:
        raise ValueError("texture.kind 仅支持 grain / paper / dots：%s" % kind)
    return layer


def render_layer(layer: dict, w: int, h: int) -> Image.Image:
    """按类型分派；blur 不产出独立位图，合成阶段直接作用于整幅画布。"""
    ltype = layer["type"]
    if ltype not in SUPPORTED_TYPES:
        raise ValueError("不支持的图层类型：%s" % ltype)
    params = layer.get("params", {})
    seed = int(params.get("seed", 0))
    rng = random.Random(seed)
    if ltype == "gradient":
        return _render_gradient(w, h, params)
    if ltype == "glow":
        return _render_glow(w, h, params)
    if ltype == "noise":
        return _render_noise(w, h, params, rng)
    if ltype == "blur_noise":
        return _render_blur_noise(w, h, params, rng)
    if ltype == "vignette":
        return _render_vignette(w, h, params)
    if ltype == "texture":
        return _render_texture(w, h, params, rng)
    # blur：独立占位透明层，效果在合成时施加。
    return Image.new("RGBA", (w, h), (0, 0, 0, 0))


# ----------------------------------------------------------------------
# 预设
# ----------------------------------------------------------------------
PRESETS: dict[str, dict] = {
    "深底光斑": {
        "canvas": [1920, 1080],
        "layers": [
            {"id": "bg", "type": "gradient",
             "params": {"colors": ["#0E3B36", "#071E1C"], "angle": 120, "stops": [0, 1]}},
            {"id": "glow", "type": "glow",
             "params": {"center": [0.62, 0.38], "radius": 0.55, "color": "#7FE3C8", "strength": 0.5}},
            {"id": "noise", "type": "noise",
             "params": {"amount": 0.03, "scale": 2, "seed": 7}},
            {"id": "vig", "type": "vignette",
             "params": {"strength": 0.2, "radius": 0.8}},
        ],
    },
    "暖橙渐变": {
        "canvas": [1920, 1080],
        "layers": [
            {"id": "bg", "type": "gradient",
             "params": {"colors": ["#E8833A", "#7A2E12"], "angle": 135, "stops": [0, 1]}},
            {"id": "glow", "type": "glow",
             "params": {"center": [0.3, 0.7], "radius": 0.5, "color": "#FFB27A", "strength": 0.35}},
            {"id": "tex", "type": "texture",
             "params": {"kind": "grain", "amount": 0.03, "seed": 3}},
        ],
    },
    "靛蓝噪点": {
        "canvas": [1920, 1080],
        "layers": [
            {"id": "bg", "type": "gradient",
             "params": {"colors": ["#2B3A67", "#141B33"], "angle": 90, "stops": [0, 1]}},
            {"id": "noise", "type": "noise",
             "params": {"amount": 0.05, "scale": 2, "seed": 11}},
            {"id": "bn", "type": "blur_noise",
             "params": {"amount": 0.05, "scale": 8, "radius": 6, "seed": 12}},
        ],
    },
    "暗角纹理": {
        "canvas": [1920, 1080],
        "layers": [
            {"id": "bg", "type": "gradient",
             "params": {"colors": ["#1A1D24", "#101216"], "angle": 100, "stops": [0, 1]}},
            {"id": "vig", "type": "vignette",
             "params": {"strength": 0.4, "radius": 0.7}},
            {"id": "tex", "type": "texture",
             "params": {"kind": "paper", "amount": 0.04, "seed": 5}},
        ],
    },
}


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _png_bytes(img: Image.Image) -> bytes:
    buf = __import__("io").BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def bake(spec: dict, out_dir: str, only: str | None = None) -> dict:
    canvas = spec.get("canvas", [1920, 1080])
    w, h = int(canvas[0]), int(canvas[1])
    layers = spec.get("layers", [])
    if only:
        layers = [lay for lay in layers if lay.get("id") == only]
    if not layers:
        raise ValueError("spec 中没有可输出的图层")

    os.makedirs(out_dir, exist_ok=True)
    composite = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    manifest_layers: list[dict] = []
    baked_name: str | None = None
    baked_sha: str | None = None

    for layer in layers:
        lid = layer["id"]
        ltype = layer["type"]
        params = layer.get("params", {})
        if ltype == "blur":
            radius = float(params.get("radius", 4))
            composite = composite.filter(ImageFilter.GaussianBlur(radius))
            img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        else:
            img = render_layer(layer, w, h)
            composite.alpha_composite(img)

        out_name = "%s.png" % lid
        data = _png_bytes(img)
        with open(os.path.join(out_dir, out_name), "wb") as f:
            f.write(data)
        manifest_layers.append({
            "id": lid, "type": ltype, "params": params,
            "out": out_name, "sha256": _sha256_bytes(data),
        })

    has_bg = any(lay.get("id") == "bg" for lay in layers)
    if has_bg:
        baked_data = _png_bytes(composite)
        baked_sha = _sha256_bytes(baked_data)
        baked_name = "baked-%s.png" % baked_sha[:8]
        with open(os.path.join(out_dir, baked_name), "wb") as f:
            f.write(baked_data)

    manifest = {
        "canvas": [w, h],
        "layers": manifest_layers,
        "baked": baked_name,
        "baked_sha256": baked_sha,
        "generator_version": GENERATOR_VERSION,
    }
    with open(os.path.join(out_dir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return manifest


def main(argv: list[str] | None = None) -> int:
    if Image is None:
        print("错误：未安装 Pillow，无法运行装饰层渲染")
        return 2

    ap = argparse.ArgumentParser(description="装饰层预渲染管线")
    ap.add_argument("--spec", help="spec.json 路径")
    ap.add_argument("--out-dir", required=True, help="产物输出目录")
    ap.add_argument("--only", help="仅输出指定 id 的图层")
    ap.add_argument("--preset", choices=sorted(PRESETS.keys()), help="内置预设名，直接出图")
    args = ap.parse_args(argv)

    if args.preset:
        spec = PRESETS[args.preset]
    elif args.spec:
        try:
            with open(args.spec, "r", encoding="utf-8") as f:
                spec = json.load(f)
        except (OSError, ValueError) as e:
            print("错误：spec 不可读：%s" % e)
            return 2
    else:
        ap.print_help()
        return 2

    try:
        manifest = bake(spec, args.out_dir, args.only)
    except (ValueError, OSError) as e:
        print("错误：%s" % e)
        return 2

    print("已输出 %d 个图层 → %s" % (len(manifest["layers"]), args.out_dir))
    if manifest["baked"]:
        print("合成图：%s（sha256 %s）" % (manifest["baked"], manifest["baked_sha256"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
