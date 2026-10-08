"""Selbsttest BI-PPT: python3 tests/test_bippt.py  (mit dem venv-Python: ~/.bi-ppt/venv/bin/python …)
Baut das Beispiel, prüft Struktur und Brand-Regeln, testet Basis-Modus, Analyse, Umwandeln (entwurf/stil).
Erwartet: ALLE FÄLLE BESTANDEN
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "tools"))
import os  # noqa: E402

os.environ.setdefault("BIPPT_HOME", tempfile.mkdtemp())  # keine echte config verändern
from pptx import Presentation  # noqa: E402

from bippt import analyse, bauen, konfig, pruefen, umwandeln  # noqa: E402

konfig.speichern({**konfig.laden(), "vorlage": str(REPO / "vorlagen" / "BI_Vorlage_16x9.pptx")})
fehler = []


def fall(name, bedingung, info=""):
    print(f"  {'ok  ' if bedingung else 'FEHL'} {name}{' – ' + info if info and not bedingung else ''}")
    if not bedingung:
        fehler.append(name)


with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    spec = td / "beispiel.spec.json"
    shutil.copy(REPO / "skills" / "bi-ppt" / "beispiele" / "beispiel.spec.json", spec)
    s = json.loads(spec.read_text(encoding="utf-8"))
    print("Beispiel bauen")
    ziel, warn = bauen.bauen(spec, still=True)
    fall("Datei gebaut", ziel.is_file())
    fall("keine Bau-Warnungen", not warn, "; ".join(warn))
    prs = Presentation(str(ziel))
    fall("Folienzahl", len(prs.slides) == len(s["folien"]), f"{len(prs.slides)} statt {len(s['folien'])}")
    ausgeblendet = [i for i, x in enumerate(prs.slides, 1) if x._element.get("show") == "0"]
    fall("Backup ausgeblendet", ausgeblendet == list(range(13, len(s["folien"]) + 1)), str(ausgeblendet))
    befunde = pruefen.pruefen(ziel)
    fall("Brand-Prüfung ohne FEHLER", not [b for b in befunde if b[0] == "FEHLER"],
         "; ".join(f"{b[1]}: {b[2]}" for b in befunde if b[0] == "FEHLER"))
    fuss = [ph.text_frame.text for ph in prs.slides[1].placeholders if "FOOTER" in str(ph.placeholder_format.type)]
    fall("Fußzeile mit Vermerk", fuss and fuss[0].endswith("ONLY FOR INTERNAL USE"), str(fuss))

    print("Basis-Modus")
    basis = {"basis": ziel.name, "ausgabe": "v2.pptx", "fusszeile": "V2",
             "folien": [{"basis": 1, "platzhalter": {"1": ["Projektteam-Update", "X | 01.01.2027"]}},
                        {"typ": "kernaussage", "titel": "Neu eingefügt", "text": "Test"},
                        {"basis": 3, "ersetzen": {"Stand 07.10.2026": "Stand 01.01.2027"}},
                        {"basis": 6, "ausblenden": True}]}
    bspec = td / "v2.spec.json"
    bspec.write_text(json.dumps(basis), encoding="utf-8")
    z2, w2 = bauen.bauen(bspec, still=True)
    p2 = Presentation(str(z2))
    fall("Basis: 4 Folien", len(p2.slides) == 4, str(len(p2.slides)))
    texte = " ".join(sh.text_frame.text for sh in p2.slides[2].shapes if sh.has_text_frame)
    fall("Basis: Ersetzung", "Stand 01.01.2027" in texte)
    fall("Basis: Ausblenden", p2.slides[3]._element.get("show") == "0")
    fall("Basis: Untertitel", "Projektteam-Update" in p2.slides[0].placeholders[1].text_frame.text)
    fall("Basis: keine Warnungen", not w2, "; ".join(w2))

    print("Analyse / Umwandeln")
    a = analyse.analyse(ziel)
    fall("Analyse: BI-Master erkannt", a["bi_master"])
    e = umwandeln.entwurf(ziel, td / "e.spec.json")
    es = json.loads(e.read_text(encoding="utf-8"))
    es["ausgabe"] = "e.pptx"
    e.write_text(json.dumps(es), encoding="utf-8")
    z3, _ = bauen.bauen(e, still=True)
    fall("Entwurf baubar", z3.is_file())
    leer = Presentation()
    sl = leer.slides.add_slide(leer.slide_layouts[1])
    sl.shapes.title.text = "Fremd"
    from pptx.dml.color import RGBColor
    r = sl.placeholders[1].text_frame.paragraphs[0].add_run()
    r.text = "rot"
    r.font.color.rgb = RGBColor(0xC0, 0x10, 0x10)
    r.font.name = "Comic Sans MS"
    leer.save(str(td / "fremd.pptx"))
    z4 = umwandeln.stil(td / "fremd.pptx", td / "fremd_stil.pptx")
    b4 = pruefen.pruefen(z4)
    fall("Stil: Farben/Schriften bereinigt",
         not [b for b in b4 if b[0] == "FEHLER" and ("Farben" in b[2] or "Schriften" in b[2])],
         "; ".join(b[2] for b in b4))

print("ALLE FÄLLE BESTANDEN" if not fehler else f"FEHLGESCHLAGEN: {', '.join(fehler)}")
sys.exit(1 if fehler else 0)
