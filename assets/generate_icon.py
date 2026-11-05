"""Generate a minimal two-tone app icon at desktop and taskbar sizes."""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
SCALE = 4


image = Image.new("RGBA", (256 * SCALE, 256 * SCALE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

# One solid tile and one document silhouette keep the mark legible at 16 px.
draw.rounded_rectangle((8 * SCALE, 8 * SCALE, 248 * SCALE, 248 * SCALE), radius=56 * SCALE, fill="#11294A")
draw.rounded_rectangle((62 * SCALE, 42 * SCALE, 194 * SCALE, 214 * SCALE), radius=18 * SCALE, fill="#F6FAFF")
draw.polygon([(148 * SCALE, 42 * SCALE), (194 * SCALE, 88 * SCALE), (148 * SCALE, 88 * SCALE)], fill="#BFD9F4")

# Three metadata rows; the blue control point is the identifying detail.
draw.rounded_rectangle((84 * SCALE, 111 * SCALE, 170 * SCALE, 121 * SCALE), radius=5 * SCALE, fill="#7895BA")
draw.rounded_rectangle((84 * SCALE, 143 * SCALE, 170 * SCALE, 153 * SCALE), radius=5 * SCALE, fill="#7895BA")
draw.rounded_rectangle((84 * SCALE, 175 * SCALE, 149 * SCALE, 185 * SCALE), radius=5 * SCALE, fill="#7895BA")
draw.ellipse((122 * SCALE, 132 * SCALE, 164 * SCALE, 164 * SCALE), fill="#148BF2")

icon = image.resize((256, 256), Image.Resampling.LANCZOS)
icon.save(ROOT / "app-icon.png")
icon.save(ROOT / "app-icon.ico", format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
