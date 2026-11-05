"""Render the two-colour metadata monogram for Windows and Tk."""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
SCALE = 4
NAVY = "#102A43"
MINT = "#53DCC8"


def cubic(start, control_a, control_b, end, steps=32):
    points = []
    for step in range(steps + 1):
        t = step / steps
        u = 1 - t
        points.append((
            u**3 * start[0] + 3 * u**2 * t * control_a[0] + 3 * u * t**2 * control_b[0] + t**3 * end[0],
            u**3 * start[1] + 3 * u**2 * t * control_a[1] + 3 * u * t**2 * control_b[1] + t**3 * end[1],
        ))
    return points


# A single continuous stroke reads as an M for metadata at every icon size.
path = [(55, 183), (55, 91)]
path += cubic((55, 91), (55, 68), (75, 61), (90, 83))[1:]
path += cubic((90, 83), (105, 106), (116, 147), (128, 159))[1:]
path += cubic((128, 159), (140, 147), (151, 106), (166, 83))[1:]
path += cubic((166, 83), (181, 61), (201, 68), (201, 91))[1:]
path.append((201, 183))

image = Image.new("RGBA", (256 * SCALE, 256 * SCALE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((8 * SCALE, 8 * SCALE, 248 * SCALE, 248 * SCALE), radius=59 * SCALE, fill=NAVY)
draw.line([(round(x * SCALE), round(y * SCALE)) for x, y in path], fill=MINT, width=24 * SCALE, joint="curve")
radius = 12 * SCALE
for x, y in (path[0], path[-1]):
    draw.ellipse((x * SCALE - radius, y * SCALE - radius, x * SCALE + radius, y * SCALE + radius), fill=MINT)

icon = image.resize((256, 256), Image.Resampling.LANCZOS)
icon.save(ROOT / "app-icon.png")
icon.save(ROOT / "app-icon.ico", format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
