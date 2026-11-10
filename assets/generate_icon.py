"""Render the black-and-white metadata monogram for Windows and Tk."""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
SCALE = 4
BLACK = "#000000"
WHITE = "#FFFFFF"

SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <title>文件源数据编辑</title>
  <rect x="8" y="8" width="240" height="240" rx="59" fill="#000"/>
  <path d="M44 183 V72 H64 L128 145 L192 72 H212 V183 H190 V102 L128 174 L66 102 V183 Z"
        fill="#fff"/>
</svg>
"""


image = Image.new("RGBA", (256 * SCALE, 256 * SCALE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((8 * SCALE, 8 * SCALE, 248 * SCALE, 248 * SCALE), radius=59 * SCALE, fill=BLACK)
outline = [(44, 183), (44, 72), (64, 72), (128, 145), (192, 72),
           (212, 72), (212, 183), (190, 183), (190, 102),
           (128, 174), (66, 102), (66, 183)]
draw.polygon([(x * SCALE, y * SCALE) for x, y in outline], fill=WHITE)

icon = image.resize((256, 256), Image.Resampling.LANCZOS)
icon.save(ROOT / "app-icon.png")
icon.save(ROOT / "app-icon.ico", format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
(ROOT / "app-icon.svg").write_text(SVG, encoding="utf-8")
