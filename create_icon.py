"""Generate a simple ICO icon for the application"""

import os

from PIL import Image, ImageDraw, ImageFont

# Create icon directory
os.makedirs("icon_temp", exist_ok=True)

# Generate icon at multiple sizes
sizes = [16, 32, 48, 64, 128, 256]

images = []
for size in sizes:
    # Create new image with transparency
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Draw gradient background (blue to green)
    for y in range(size):
        r = int(45 + (140 - 45) * y / size)
        g = int(140 + (200 - 140) * y / size)
        b = int(200 + (100 - 200) * y / size)
        draw.rectangle([(0, y), (size, y + 1)], fill=(r, g, b, 255))

    # Draw lock icon (simplified)
    margin = size // 4
    lock_width = size - 2 * margin
    lock_height = size // 2

    # Lock body
    body_top = margin + size // 6
    draw.rounded_rectangle(
        [(margin, body_top), (size - margin, size - margin)],
        radius=size // 10,
        fill=(255, 255, 255, 200),
    )

    # Lock shackle
    shackle_width = lock_width // 3
    shackle_top = margin
    draw.arc(
        [
            (size // 2 - shackle_width // 2, shackle_top),
            (size // 2 + shackle_width // 2, shackle_top + shackle_width),
        ],
        start=0,
        end=180,
        fill=(255, 255, 255, 200),
        width=size // 12,
    )

    # Keyhole
    keyhole_y = body_top + lock_height // 2
    draw.ellipse(
        [
            (size // 2 - size // 12, keyhole_y - size // 12),
            (size // 2 + size // 12, keyhole_y + size // 12),
        ],
        fill=(45, 140, 200, 255),
    )
    draw.rectangle(
        [
            (size // 2 - size // 16, keyhole_y),
            (size // 2 + size // 16, keyhole_y + lock_height // 3),
        ],
        fill=(45, 140, 200, 255),
    )

    images.append(img)

# Save as ICO
images[0].save(
    "icon.ico", format="ICO", sizes=[(s, s) for s in sizes], append_images=images[1:]
)

# Cleanup
import shutil

shutil.rmtree("icon_temp", ignore_errors=True)

print("✓ Icon created: icon.ico")
