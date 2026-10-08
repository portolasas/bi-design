"""BI-PPT – PowerPoint im Boehringer-Ingelheim-Design erzeugen, prüfen, rendern und umwandeln."""
from pathlib import Path

_v = Path(__file__).resolve().parents[2] / "VERSION"
__version__ = _v.read_text(encoding="utf-8").strip() if _v.is_file() else "0"
