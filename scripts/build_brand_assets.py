"""Package the established GAWA star on a transparent window background."""
from collections import deque
from pathlib import Path
from PIL import Image

assets = Path(__file__).resolve().parents[1] / "assets"
source = Image.open(assets / "gawa_source.png").convert("RGBA")
pixels = source.load()
width, height = source.size
queue = deque([(x, y) for x in range(width) for y in (0, height - 1)]
              + [(x, y) for y in range(height) for x in (0, width - 1)])
visited = set()
while queue:
    x, y = queue.popleft()
    if (x, y) in visited or not (0 <= x < width and 0 <= y < height):
        continue
    visited.add((x, y))
    r, g, b, alpha = pixels[x, y]
    if alpha and not (b > r and b > g):
        continue
    pixels[x, y] = (r, g, b, 0)
    queue.extend(((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)))
source.resize((48, 48), Image.Resampling.LANCZOS).save(assets / "gawa.png")
source.save(assets / "gawa.ico", sizes=[(n, n) for n in (16, 24, 32, 48, 64, 128, 256)])
