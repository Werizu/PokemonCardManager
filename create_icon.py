#!/usr/bin/env python3
"""Generates the Poke Inv app icon (.icns) in the v2 palette."""
import os
import subprocess
import sys
import tempfile
from PIL import Image, ImageDraw, ImageFilter


def draw_app_icon(size=1024):
    """App icon in the v2 palette: card stack with a Poké Ball, macOS icon grid."""
    ss = 2
    N = 1024 * ss
    u = lambda v: int(v * ss)
    img = Image.new("RGBA", (N, N), (0, 0, 0, 0))

    # Squircle body (macOS grid: 824px body, 100px inset) with a soft drop shadow
    shadow = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle([u(100), u(112), u(924), u(936)], u(185), fill=(0, 0, 0, 110))
    img.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(u(14))))
    grad = Image.new("RGBA", (N, N))
    gd = ImageDraw.Draw(grad)
    top, bot = (38, 43, 54), (13, 15, 19)
    for y in range(N):
        t = min(max((y - u(100)) / u(824), 0), 1)
        gd.line([(0, y), (N, y)], fill=tuple(int(a + (b - a) * t) for a, b in zip(top, bot)) + (255,))
    mask = Image.new("L", (N, N), 0)
    ImageDraw.Draw(mask).rounded_rectangle([u(100), u(100), u(924), u(924)], u(185), fill=255)
    img.paste(grad, (0, 0), mask)
    # thin top highlight
    hl = Image.new("L", (N, N), 0)
    ImageDraw.Draw(hl).rounded_rectangle([u(100), u(100), u(924), u(924)], u(185), outline=40, width=u(3))
    img.paste((255, 255, 255, 255), (0, 0), Image.composite(hl, Image.new("L", (N, N), 0), mask))

    def card(angle, dx, dy, fill, outline, out_w, alpha=255):
        w, h, r = u(430), u(600), u(44)
        layer = Image.new("RGBA", (N, N), (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        x0, y0 = (N - w) // 2 + u(dx), (N - h) // 2 + u(dy)
        d.rounded_rectangle([x0, y0, x0 + w, y0 + h], r, fill=fill, outline=outline, width=u(out_w))
        return layer.rotate(angle, resample=Image.BICUBIC, center=(N // 2, N // 2)), (x0, y0, w, h)

    sh, _ = card(-13, 6, 22, (0, 0, 0, 120), None, 0)
    img.alpha_composite(sh.filter(ImageFilter.GaussianBlur(u(16))))
    back, _ = card(-13, -10, 8, (42, 47, 57, 255), (70, 76, 90, 255), 6)
    img.alpha_composite(back)
    fsh, _ = card(6, 14, 28, (0, 0, 0, 140), None, 0)
    img.alpha_composite(fsh.filter(ImageFilter.GaussianBlur(u(18))))

    # Front card: gold frame, dark inner panel, Poké Ball
    w, h = u(430), u(600)
    front = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    d = ImageDraw.Draw(front)
    x0, y0 = (N - w) // 2 + u(14), (N - h) // 2 + u(6)
    d.rounded_rectangle([x0, y0, x0 + w, y0 + h], u(44), fill=(255, 203, 5, 255))
    i = u(26)
    d.rounded_rectangle([x0 + i, y0 + i, x0 + w - i, y0 + h - i], u(24), fill=(23, 26, 32, 255))
    cx, cy, R = x0 + w // 2, y0 + h // 2, u(132)
    gold, dark, lw = (255, 203, 5, 255), (23, 26, 32, 255), u(24)
    d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=gold)
    d.pieslice([cx - R + lw, cy - R + lw, cx + R - lw, cy + R - lw], 0, 180, fill=dark)
    d.rectangle([cx - R, cy - u(13), cx + R, cy + u(13)], fill=dark)
    for r, c in ((u(58), dark), (u(44), gold), (u(22), dark)):
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=c)
    img.alpha_composite(front.rotate(6, resample=Image.BICUBIC, center=(N // 2, N // 2)))

    return img.resize((size, size), Image.LANCZOS)


def create_icon(icns_path):
    with tempfile.TemporaryDirectory() as tmpdir:
        iconset = os.path.join(tmpdir, "icon.iconset")
        os.makedirs(iconset)
        for s in [16, 32, 128, 256, 512]:
            draw_app_icon(s).save(os.path.join(iconset, f"icon_{s}x{s}.png"))
            draw_app_icon(s * 2).save(os.path.join(iconset, f"icon_{s}x{s}@2x.png"))
        os.makedirs(os.path.dirname(icns_path), exist_ok=True)
        subprocess.run(["iconutil", "-c", "icns", iconset, "-o", icns_path], check=True)


if __name__ == "__main__":
    create_icon(sys.argv[1])
