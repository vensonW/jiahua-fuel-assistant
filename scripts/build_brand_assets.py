"""Create size variants from the established GAWA blue brand asset."""
from pathlib import Path
from PIL import Image

assets = Path(__file__).resolve().parents[1] / "assets"
source = Image.open(assets / "gawa_source.png").convert("RGBA")
source.resize((48, 48), Image.Resampling.LANCZOS).save(assets / "gawa.png")
source.save(assets / "gawa.ico", sizes=[(n, n) for n in (16, 24, 32, 48, 64, 128, 256)])
