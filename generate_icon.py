"""Generate a sleek Nova app icon programmatically."""
from PIL import Image, ImageDraw, ImageFont
import math

SIZE = 256
CENTER = SIZE // 2
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Background circle - deep space
draw.ellipse([8, 8, SIZE-8, SIZE-8], fill=(10, 10, 15, 255))

# Outer ring glow (purple-cyan gradient effect)
for i in range(6, 0, -1):
    alpha = int(40 * (7 - i) / 7)
    r = int(100 + 24 * i / 6)
    g = int(40 + 138 * i / 6)
    b = int(200 + 12 * i / 6)
    offset = 18 + i * 4
    draw.ellipse(
        [offset, offset, SIZE - offset, SIZE - offset],
        outline=(r, g, b, alpha), width=3
    )

# Main orbital ring
draw.ellipse(
    [30, 30, SIZE-30, SIZE-30],
    outline=(124, 58, 237, 220), width=3
)

# Inner glow orb
for r in range(45, 0, -1):
    # Gradient from cyan center to purple edge
    t = r / 45.0
    red = int(6 + (124 - 6) * t)
    green = int(182 + (58 - 182) * t)
    blue = int(212 + (237 - 212) * t)
    alpha = int(200 * (1 - t * 0.7))
    draw.ellipse(
        [CENTER - r, CENTER - r, CENTER + r, CENTER + r],
        fill=(red, green, blue, alpha)
    )

# Central bright point
draw.ellipse(
    [CENTER - 8, CENTER - 8, CENTER + 8, CENTER + 8],
    fill=(255, 255, 255, 240)
)

# Small accent dots on the ring (like orbital particles)
for angle_deg in [45, 135, 225, 315]:
    angle = math.radians(angle_deg)
    radius = 83
    x = CENTER + radius * math.cos(angle)
    y = CENTER + radius * math.sin(angle)
    draw.ellipse(
        [x - 4, y - 4, x + 4, y + 4],
        fill=(6, 182, 212, 200)
    )

# Save as ICO with multiple sizes
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
imgs = [img.resize(s, Image.LANCZOS) for s in sizes]
imgs[-1].save("nova.ico", format="ICO", sizes=sizes, append_images=imgs[:-1])
print("✅ nova.ico generated successfully!")
