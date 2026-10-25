"""Lokale Einstellungen von bippt (je Rechner): ~/.bi-design/config.json.

Schlüssel
- asset_root   BrandDesign-Ordner (Logos, Icons, Vorlagen); Struktur siehe bi-designer tokens.json › assets
- vorlage      Standard-Vorlage (.pptx/.potx im BI-Master); leer = eigene Vorlage aus `bippt vorlage`
               (~/.bi-design/vorlagen/), sonst die des Plugins (vorlagen/), sonst <asset_root>/Vorlagen/
- vertraulichkeit  Vermerk in der Fußzeile (BI-Master v2.1), Standard „ONLY FOR INTERNAL USE“
- marke        Markenprofil aus bi-designer (Standard boehringer-ingelheim)
- render_dir   Arbeitsordner für PowerPoint-Export (Mac: einmalig in PowerPoint freigeben)

Fehlt die Datei, gilt die von BI-PPT 1.x (~/.bi-ppt/config.json), bis `bippt aufraeumen --ja` sie übernimmt.
"""
import json
import os
import platform
import sys
from pathlib import Path

from . import PLUGIN

HOME = Path(os.environ.get("BIPPT_HOME", Path.home() / ".bi-design"))
DATEI = HOME / "config.json"
ALT = None if "BIPPT_HOME" in os.environ else Path.home() / ".bi-ppt" / "config.json"
IST_MAC = sys.platform == "darwin"
IST_WIN = os.name == "nt"

_MAC_ASSETS = Path.home() / "Library/Mobile Documents/com~apple~CloudDocs/Arbeit/BrandDesign"
VORLAGE_NAME = "BI_Vorlage_16x9.pptx"


def _standard():
    asset = _MAC_ASSETS if (IST_MAC and _MAC_ASSETS.is_dir()) else None
    render = (Path.home() / "Documents" / "BI-PPT-Render") if IST_MAC else (HOME / "render")
    return {
        "asset_root": str(asset) if asset else "",
        "vorlage": "",
        "marke": "boehringer-ingelheim",
        "render_dir": str(render),
        "vertraulichkeit": "ONLY FOR INTERNAL USE",
    }


def _lesen(datei):
    try:
        werte = json.loads(datei.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        print(f"WARNUNG: {datei} nicht lesbar ({e}) – Standardwerte", file=sys.stderr)
        return {}
    werte.pop("quelle", None)  # Repo-Pfad des alten Installers
    if werte.get("vorlage") and Path(werte["vorlage"]).parent == Path.home() / ".bi-ppt" / "vorlagen":
        werte["vorlage"] = ""  # Kopie des alten Installers; jetzt aus dem Plugin
    return {k: v for k, v in werte.items() if v is not None}


def laden():
    cfg = _standard()
    if DATEI.is_file():
        cfg.update(_lesen(DATEI))
    elif ALT and ALT.is_file():
        cfg.update(_lesen(ALT))
    return cfg


def speichern(cfg):
    HOME.mkdir(parents=True, exist_ok=True)
    DATEI.write_text(json.dumps(cfg, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def asset_root(cfg=None):
    cfg = cfg or laden()
    p = cfg.get("asset_root") or ""
    return Path(p).expanduser() if p else None


def vorlage_pfad(cfg=None):
    """Standard-Vorlage oder None."""
    cfg = cfg or laden()
    if cfg.get("vorlage"):
        p = Path(cfg["vorlage"]).expanduser()
        return p if p.is_file() else None
    for p in (HOME / "vorlagen" / VORLAGE_NAME, PLUGIN / "vorlagen" / VORLAGE_NAME):
        if p.is_file():
            return p
    root = asset_root(cfg)
    if root:
        p = root / "Vorlagen" / VORLAGE_NAME
        if p.is_file():
            return p
    return None


def rechner():
    return f"{platform.system()} {platform.release()} · Python {platform.python_version()}"
