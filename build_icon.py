# -*- coding: utf-8 -*-
"""Tạo assets/app.ico (chạy một lần trước khi build exe). Cần: pip install pillow"""
from __future__ import annotations

import os

os.makedirs("assets", exist_ok=True)
out = os.path.join("assets", "app.ico")

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    raise SystemExit("Cài Pillow: pip install pillow")

W = 256
img = Image.new("RGB", (W, W), "#1e3a8a")
draw = ImageDraw.Draw(img)
font_path = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "segoeui.ttf")
try:
    font = ImageFont.truetype(font_path, 88)
except OSError:
    font = ImageFont.load_default()

text = "PO"
bbox = draw.textbbox((0, 0), text, font=font)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
draw.text(((W - tw) // 2, (W - th) // 2 - 8), text, fill="white", font=font)

# ICO nhiều kích thước cho Windows
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
imgs = [img.resize(s, Image.Resampling.LANCZOS) for s in sizes]
imgs[0].save(
    out,
    format="ICO",
    sizes=[(i.width, i.height) for i in imgs],
    append_images=imgs[1:],
)
print("OK:", os.path.abspath(out))
