#!/usr/bin/env python3
"""bippt – Werkzeug des Agenten BI-PPT. Aufruf über den Starter:
  macOS:   ~/.bi-ppt/bippt <befehl> …
  Windows: & "$HOME\\.bi-ppt\\bippt.cmd" <befehl> …

Befehle
  umgebung                     Python, Pakete, Vorlage, Marke, Icons, Renderweg prüfen
  konfig [--setze k=v …]       Einstellungen zeigen/setzen (~/.bi-ppt/config.json)
  lesen <datei…> [-o md]       Quellen als Text (vtt, docx, pptx, pdf, xlsx, md)
  analyse <pptx> [--json]      Bestehendes Deck knapp beschreiben
  typen                        Folientypen auflisten
  icons <begriff…>             Brand-Icons suchen
  bauen <spec.json> [-o pptx]  Präsentation bauen
  pruefen <pptx>               Brand-Prüfung (Farben, Schriften, Kontrast, Überlauf, Master)
  rendern <pptx> [--alle] [--dpi N] [--folien 2,5] [--erzwingen]   PDF/PNG/Kontaktbogen zur Sichtprüfung
  umwandeln entwurf|stil <pptx> [-o ziel]            fremdes Deck ins BI-Design
  vorlage <bi-deck.pptx> [-o ziel]                   BI-Vorlage aus einem Deck im BI-Master erzeugen
  neu <ordner> [--name N]      Startgerüst: Spezifikation mit Beispielfolien anlegen
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except AttributeError:
    pass


def _umgebung(_):
    import importlib
    import platform
    from bippt import __version__, konfig
    print(f"BI-PPT {__version__} · {konfig.rechner()}")
    ok = True
    for mod, pflicht in (("pptx", True), ("lxml", True), ("docx", False), ("pymupdf", False), ("PIL", False), ("openpyxl", False)):
        try:
            importlib.import_module(mod)
            print(f"  ok      Paket {mod}")
        except ImportError:
            print(f"  {'FEHLT' if pflicht else 'optional'}  Paket {mod}")
            ok = ok and not pflicht
    cfg = konfig.laden()
    vl = konfig.vorlage_pfad(cfg)
    print(f"  {'ok' if vl else 'FEHLT':7} Vorlage {vl or '(bippt vorlage <BI-Deck> oder Installer erneut)'}")
    ok = ok and bool(vl)
    from bippt.marke import Marke
    m = Marke()
    print(f"  {'ok' if m.ordner else 'WARN':7} Marke {m.profil} aus {m.ordner or 'Notstand (bi-designer fehlt)'}")
    ar = konfig.asset_root(cfg)
    n_ic = len(list((ar / 'Icons_darkgreen').glob('*.png'))) if ar and (ar / 'Icons_darkgreen').is_dir() else 0
    print(f"  {'ok' if n_ic else 'WARN':7} Icons {n_ic} in {ar or '(asset_root nicht gesetzt: bippt konfig --setze asset_root=<BrandDesign>)'}")
    from bippt.rendern import wege
    w = wege()
    print(f"  {'ok' if w else 'WARN':7} Rendern über {', '.join(w) or 'nichts (PowerPoint/LibreOffice fehlt) – nur pruefen möglich'}")
    if "powerpoint" in w and konfig.IST_MAC:
        print(f"          render_dir {cfg['render_dir']} (einmalig in PowerPoint freigeben)")
    print("ERGEBNIS: " + ("OK" if ok else "UNVOLLSTÄNDIG"))
    return 0 if ok else 1


def _konfig(a):
    from bippt import konfig
    cfg = konfig.laden()
    for kv in a.setze or []:
        k, _, v = kv.partition("=")
        cfg[k.strip()] = v.strip()
    if a.setze:
        konfig.speichern(cfg)
    print(json.dumps(cfg, indent=2, ensure_ascii=False))
    return 0


def _lesen(a):
    from bippt.lesen import lesen
    teile = []
    for d in a.dateien:
        teile.append(f"# Quelle: {Path(d).name}\n\n{lesen(d)}")
    txt = "\n\n".join(teile)
    if a.o:
        Path(a.o).write_text(txt, encoding="utf-8")
        print(f"{a.o}: {len(txt)} Zeichen")
    else:
        print(txt)
    return 0


def _analyse(a):
    from bippt import analyse
    r = analyse.analyse(a.pptx)
    print(json.dumps(r, ensure_ascii=False, indent=1, default=str) if a.json else analyse.als_markdown(r, notizen=not a.ohne_notizen))
    return 0


def _typen(_):
    from bippt.folien import TYPEN
    for n, (fn, lay) in sorted(TYPEN.items()):
        print(f"{n:12} [{lay}]  {(fn.__doc__ or '').strip().splitlines()[0] if fn.__doc__ else ''}")
    return 0


def _icons(a):
    from bippt.marke import Marke
    m = Marke()
    treffer = m.icon_suchen(*a.begriffe)
    print(", ".join(treffer[:60]) or "keine Treffer")
    if len(treffer) > 60:
        print(f"… {len(treffer)} Treffer")
    return 0


def _bauen(a):
    from bippt.bauen import bauen
    _, w = bauen(a.spec, a.o)
    return 0


def _pruefen(a):
    from bippt.pruefen import bericht
    return 1 if bericht(a.pptx) else 0


def _rendern(a):
    from bippt.rendern import rendern
    folien = [int(x) for x in a.folien.split(",")] if a.folien else None
    rendern(a.pptx, a.o, dpi=a.dpi, alle=a.alle, weg=a.weg, folien=folien, erzwingen=a.erzwingen)
    return 0


def _umwandeln(a):
    from bippt import umwandeln
    (umwandeln.entwurf if a.weg == "entwurf" else umwandeln.stil)(a.pptx, a.o)
    return 0


def _vorlage(a):
    from bippt import konfig, vorlage
    ziel = Path(a.o) if a.o else konfig.HOME / "vorlagen" / konfig.VORLAGE_NAME
    namen = vorlage.extrahieren(a.pptx, ziel)
    print(f"VORLAGE: {ziel} – {len(namen)} Layouts")
    return 0


def _neu(a):
    ziel = Path(a.ordner)
    ziel.mkdir(parents=True, exist_ok=True)
    beispiel = Path(__file__).resolve().parent.parent / "skills" / "bi-ppt" / "beispiele" / "beispiel.spec.json"
    if not beispiel.is_file():
        beispiel = Path.home() / ".bi-ppt" / "beispiele" / "beispiel.spec.json"
    spec = json.loads(beispiel.read_text(encoding="utf-8"))
    spec["ausgabe"] = f"{a.name}.pptx"
    p = ziel / f"{a.name}.spec.json"
    if p.exists():
        raise SystemExit(f"{p} existiert schon")
    p.write_text(json.dumps(spec, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"NEU: {p}")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog="bippt", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("umgebung").set_defaults(fn=_umgebung)
    p = sub.add_parser("konfig"); p.add_argument("--setze", nargs="*"); p.set_defaults(fn=_konfig)
    p = sub.add_parser("lesen"); p.add_argument("dateien", nargs="+"); p.add_argument("-o"); p.set_defaults(fn=_lesen)
    p = sub.add_parser("analyse"); p.add_argument("pptx"); p.add_argument("--json", action="store_true")
    p.add_argument("--ohne-notizen", action="store_true"); p.set_defaults(fn=_analyse)
    sub.add_parser("typen").set_defaults(fn=_typen)
    p = sub.add_parser("icons"); p.add_argument("begriffe", nargs="+"); p.set_defaults(fn=_icons)
    p = sub.add_parser("bauen"); p.add_argument("spec"); p.add_argument("-o"); p.set_defaults(fn=_bauen)
    p = sub.add_parser("pruefen"); p.add_argument("pptx"); p.set_defaults(fn=_pruefen)
    p = sub.add_parser("rendern"); p.add_argument("pptx"); p.add_argument("-o"); p.add_argument("--dpi", type=int, default=80)
    p.add_argument("--alle", action="store_true"); p.add_argument("--weg", choices=["powerpoint", "libreoffice"])
    p.add_argument("--folien"); p.add_argument("--erzwingen", action="store_true"); p.set_defaults(fn=_rendern)
    p = sub.add_parser("umwandeln"); p.add_argument("weg", choices=["entwurf", "stil"]); p.add_argument("pptx")
    p.add_argument("-o"); p.set_defaults(fn=_umwandeln)
    p = sub.add_parser("vorlage"); p.add_argument("pptx"); p.add_argument("-o"); p.set_defaults(fn=_vorlage)
    p = sub.add_parser("neu"); p.add_argument("ordner"); p.add_argument("--name", default="Praesentation"); p.set_defaults(fn=_neu)
    a = ap.parse_args(argv)
    return a.fn(a)


if __name__ == "__main__":
    sys.exit(main())
