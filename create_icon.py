#!/usr/bin/env python3
"""Generates the Poke Inv app icon (.icns) from scratch."""
import os
import sys
import tempfile
from PIL import Image, ImageDraw, ImageFont

def create_icon(icns_path):
    size = 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    margin = 20
    draw.rounded_rectangle(
        [margin, margin, size - margin, size - margin],
        radius=90,
        fill=(26, 26, 46, 255),
        outline=(255, 203, 5, 80),
        width=3,
    )

    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 320)
    except OSError:
        try:
            font = ImageFont.truetype("/System/Library/Fonts/SFNSMono.ttf", 300)
        except OSError:
            font = ImageFont.load_default()

    draw.text((size / 2, size / 2 - 15), "P", font=font, fill=(255, 203, 5, 255), anchor="mm")

    with tempfile.TemporaryDirectory() as tmpdir:
        iconset = os.path.join(tmpdir, "icon.iconset")
        os.makedirs(iconset)
        for s in [16, 32, 64, 128, 256, 512]:
            img.resize((s, s), Image.LANCZOS).save(os.path.join(iconset, f"icon_{s}x{s}.png"))
            if s <= 256:
                img.resize((s * 2, s * 2), Image.LANCZOS).save(
                    os.path.join(iconset, f"icon_{s}x{s}@2x.png"))
        os.makedirs(os.path.dirname(icns_path), exist_ok=True)
        os.system(f'iconutil -c icns "{iconset}" -o "{icns_path}"')

if __name__ == "__main__":
    create_icon(sys.argv[1])
