"""Selbsttest bippt: python3 tests/test_bippt.py  (Python mit den Paketen aus plugin/tools/requirements.txt)
Baut das Beispiel, prüft Struktur und Brand-Regeln, testet Basis-Modus, Analyse, Umwandeln (entwurf/stil).
Erwartet: ALLE FÄLLE BESTANDEN
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

PLUGIN = Path(__file__).resolve().parents[1] / "plugin"
sys.path.insert(0, str(PLUGIN / "tools"))
import os  # noqa: E402

os.environ.setdefault("BIPPT_HOME", tempfile.mkdtemp())  # keine echte config verändern
from pptx import Presentation  # noqa: E402
from lxml import etree  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

from bippt import analyse, bauen, konfig, pruefen, umwandeln  # noqa: E402
from bippt.marke import Marke  # noqa: E402

konfig.speichern({**konfig.laden(), "vorlage": str(PLUGIN / "vorlagen" / "BI_Vorlage_16x9.pptx")})
fehler = []


def fall(name, bedingung, info=""):
    print(f"  {'ok  ' if bedingung else 'FEHL'} {name}{' – ' + info if info and not bedingung else ''}")
    if not bedingung:
        fehler.append(name)


print("Marke")
fall("Marke aus bi-designer im Plugin", Marke().ordner == PLUGIN / "skills" / "bi-designer" / "marken" / "boehringer-ingelheim")
fall("Standard-Vorlage aus dem Plugin", konfig.vorlage_pfad({"vorlage": ""}) == PLUGIN / "vorlagen" / "BI_Vorlage_16x9.pptx")

with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    spec = td / "beispiel.spec.json"
    shutil.copy(PLUGIN / "skills" / "bi-ppt" / "beispiele" / "beispiel.spec.json", spec)
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

    print("Projektteam-Update (Referenzmuster, Master-Regeln)")
    pspec = td / "projektteam.spec.json"
    shutil.copy(PLUGIN / "skills" / "bi-ppt" / "beispiele" / "projektteam.spec.json", pspec)
    zp, wp = bauen.bauen(pspec, still=True)
    fall("Projektteam: keine Bau-Warnungen", not wp, "; ".join(wp))
    bp = pruefen.pruefen(zp, woerter=200)
    fall("Projektteam: Prüfung ohne Fehler und Warnungen", not [b for b in bp if b[0] in ("FEHLER", "WARNUNG")],
         "; ".join(f"{b[1]}: {b[2]}" for b in bp if b[0] in ("FEHLER", "WARNUNG")))
    pp = Presentation(str(zp))
    fall("Agenda mit Titel", pp.slides[1].shapes.title.text_frame.text == "Agenda")
    geo = [sh._element.spPr.find(qn("a:prstGeom")) for sh in pp.slides[2].shapes if hasattr(sh._element, "spPr")]
    fall("Ampel Gelb–Rot als Halbkreis", any(g is not None and g.get("prst") == "chord" for g in geo))
    texte = {sh.text_frame.text for sh in pp.slides[8].shapes if sh.has_text_frame}
    fall("Quartalsskala ab 18 Monaten", {"Q1", "Q4", "2028"} <= texte, str(sorted(texte)[:8]))
    pfeile = [sh for sh in pp.slides[10].shapes if sh.shape_type == 1 and sh.auto_shape_type == 33]
    fall("Datenfluss mit Pfeilen je Zeile", len(pfeile) >= 7, str(len(pfeile)))
    farben = {str(sh.fill.fore_color.rgb) for sl in pp.slides for sh in sl.shapes
              if sh.shape_type == 1 and sh.fill.type == 1}
    fall("keine Abstufungen, nur Marke", farben <= Marke().erlaubte_farben() | {"FFFFFF"},
         str(farben - Marke().erlaubte_farben()))

    print("Prüfung: Abstufungen von Designfarben")
    ab = Presentation(str(zp))
    form = ab.slides[2].shapes.add_shape(1, 0, 0, 914400, 914400)
    form.fill.solid()
    fill = form._element.spPr.find(qn("a:solidFill"))
    for el in list(fill):
        fill.remove(el)
    sc = etree.SubElement(fill, qn("a:schemeClr"))
    sc.set("val", "accent2")
    etree.SubElement(sc, qn("a:lumMod")).set("val", "20000")
    etree.SubElement(sc, qn("a:lumOff")).set("val", "80000")
    ab.save(str(td / "abstufung.pptx"))
    ba = pruefen.pruefen(td / "abstufung.pptx", woerter=200)
    fall("Abstufung accent2 als FEHLER erkannt", any(b[0] == "FEHLER" and "Abstufungen" in b[2] for b in ba),
         "; ".join(b[2] for b in ba))

    print("Fremde Formen (Agenda-Fall: früherer Platzhalter ohne Geometrie, Freihand, AlternateContent)")
    from PIL import Image
    from bippt import aendern
    from bippt.bausteine import E
    vl = Presentation(str(PLUGIN / "vorlagen" / "BI_Vorlage_16x9.pptx"))
    fs = vl.slides.add_slide(next(lay for lay in vl.slide_layouts if lay.name == "Agenda"))
    fs.shapes.title.text = "Agenda"
    fs.background.fill.solid()
    fs.background.fill.fore_color.rgb = RGBColor.from_string("08312A")
    tb = fs.shapes.add_textbox(E(7.8), E(1.7), E(4.7), E(0.7))
    tb.name = "Textplatzhalter 2"
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = "Zielbild für den Prozess"
    r.font.color.rgb = RGBColor.from_string("00E47C")  # Akzentgrün auf dunklem Folienhintergrund: erlaubt
    tb._element.find(qn("p:nvSpPr")).find(qn("p:cNvSpPr")).attrib.pop("txBox")
    tb._element.spPr.remove(tb._element.spPr.find(qn("a:prstGeom")))
    Image.new("RGB", (64, 64), (0, 228, 124)).save(td / "clipart.png")
    fs.shapes.add_picture(str(td / "clipart.png"), E(7.1), E(1.76), E(0.43), E(0.43)).name = "BI Clipart Zielbild"
    ns = ('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
    fs.shapes._spTree.append(etree.fromstring(
        f'<p:contentPart {ns} r:id="rId99"><p:nvContentPartPr><p:cNvPr id="40" name="Freihand 39"/>'
        '<p:cNvContentPartPr/><p:nvPr/></p:nvContentPartPr></p:contentPart>'))
    fs.shapes._spTree.append(etree.fromstring(f'<mc:AlternateContent {ns}><mc:Choice Requires="p14"/>'
                                              '<mc:Fallback/></mc:AlternateContent>'))
    vl.save(str(td / "agenda_fremd.pptx"))
    af = analyse.analyse(td / "agenda_fremd.pptx", formen=True)["folien"][0]
    fall("Agenda-Fall: Analyse liest den Text der Form ohne Geometrie",
         any("Zielbild" in p["t"] for t in af["texte"] for p in t["absaetze"]), str(af["texte"]))
    gruende = " | ".join(f"{u['name']}: {u['grund']}" for u in af["unbekannt"])
    fall("Agenda-Fall: drei unbekannte Formen gemeldet", len(af["unbekannt"]) == 3 and "Textplatzhalter 2" in gruende
         and "Freihand 39" in gruende and "AlternateContent" in gruende, gruende)
    fall("Agenda-Fall: Titel und Bild gelesen", af["titel"] == "Agenda" and
         [b["form"] for b in af["bilder"]] == ["BI Clipart Zielbild"], f"{af['titel']} {af['bilder']}")
    bf = pruefen.pruefen(td / "agenda_fremd.pptx")
    fall("Agenda-Fall: Prüfung meldet die Form als Hinweis",
         any(b[0] == "HINWEIS" and "Textplatzhalter 2" in b[2] for b in bf), "; ".join(b[2] for b in bf))
    fall("Agenda-Fall: Akzentgrün auf dunklem Folienhintergrund ist kein Fehler",
         not any("Akzentgrün" in b[2] for b in bf), "; ".join(b[2] for b in bf))
    fall("Agenda-Fall: Entwurf zum Umwandeln", umwandeln.entwurf(td / "agenda_fremd.pptx", td / "af.spec.json").is_file())
    (td / "af_v2.spec.json").write_text(json.dumps({"basis": "agenda_fremd.pptx", "ausgabe": "af_v2.pptx", "folien": [
        {"basis": 1, "ersetzen": {"Zielbild": "Ziel"}}]}), encoding="utf-8")
    zf, wf = bauen.bauen(td / "af_v2.spec.json", still=True)
    texte = " ".join(sh.text_frame.text for sh in Presentation(str(zf)).slides[0].shapes if sh.has_text_frame)
    fall("Agenda-Fall: Basis-Modus ersetzt in der Form ohne Geometrie", "Ziel für den Prozess" in texte and not wf,
         f"{texte} {wf}")

    print("Änderungsmodus für Basisfolien (aendern, elemente)")
    Image.new("RGB", (1440, 1000), (0, 228, 124)).save(td / "alt.png")
    Image.new("RGB", (1440, 1000), (8, 49, 42)).save(td / "neu_gleich.png")
    Image.new("RGB", (1440, 2000), (229, 227, 222)).save(td / "neu_hoch.png")
    (td / "aendern.spec.json").write_text(json.dumps({"ausgabe": "aendern_basis.pptx", "fusszeile": "Test", "folien": [
        {"typ": "bild_text", "titel": "Bild", "bild": "alt.png", "punkte": ["A"]},
        {"typ": "frei", "titel": "Frei", "band": {"label": "Zu klären:", "text": "Alte Frage?"}, "elemente": [
            {"form": "text", "x": 1, "y": 1.5, "w": 5, "h": 1, "t": ["Erster Punkt", {"t": "Zweiter Punkt", "groesse": 14}]},
            {"form": "text", "x": 1, "y": 3, "w": 5, "h": 0.5, "t": "Weg damit"},
            {"form": "flaeche", "x": 7, "y": 1.5, "w": 2, "h": 1, "fill": "LG", "t": "Karte"}]},
        {"typ": "bild_text", "titel": "Bild 2", "bild": "alt.png", "punkte": ["B"]},
        {"typ": "kennzahlen", "titel": "Kennzahlen", "kacheln": [{"wert": "3", "label": "mit Bild", "bild": "alt.png"},
                                                                 {"wert": "4", "label": "ohne"}]},
        {"typ": "karten", "titel": "Karten", "karten": [{"titel": "Screenshot", "bild": "alt.png", "punkte": ["x"]},
                                                        {"titel": "Text", "punkte": ["y"]}]},
        {"typ": "prozess", "titel": "Prozess", "groesse": 12.5, "kopf_groesse": 13,
         "schritte": [{"titel": "Schritt A", "punkte": ["Eins"]}, {"titel": "Schritt B", "punkte": ["Zwei"]}]}]}),
        encoding="utf-8")
    za, wa = bauen.bauen(td / "aendern.spec.json", still=True)
    fall("Bildfeld, Schriftgröße: Bau ohne Warnungen", not wa, "; ".join(wa))
    pa = Presentation(str(za))
    bilder = [sum(1 for sh in pa.slides[i].shapes if sh.shape_type == 13) for i in (3, 4)]
    fall("Bildfeld in Kennzahlen- und Kartenfolie", bilder == [1, 1], str(bilder))
    gr = {r.text: r.font.size.pt for sh in pa.slides[5].shapes if sh.has_text_frame
          for p in sh.text_frame.paragraphs for r in p.runs if r.font.size}
    fall("Schrittkarten: groesse und kopf_groesse", gr.get("Eins") == 12.5 and gr.get("Schritt A") == 13, str(gr))
    fa = analyse.analyse(za, formen=True)["folien"]
    ph1, ph3 = fa[0]["bilder"][0]["form"], fa[2]["bilder"][0]["form"]
    nach_text = {fo["absaetze"][0][0]: fo["name"] for fo in fa[1]["formen"] if fo.get("absaetze")}
    tb1, tb2, karte_ = nach_text["Erster Punkt"], nach_text["Weg damit"], nach_text["Karte"]
    alt_crop = Presentation(str(za)).slides[2].shapes
    alt_crop = [(s.crop_top, s.crop_bottom) for s in alt_crop if s.name == ph3][0]
    (td / "aendern_v2.spec.json").write_text(json.dumps({"basis": "aendern_basis.pptx", "ausgabe": "aendern_v2.pptx",
                                                         "folien": [
        {"basis": 1, "aendern": [{"form": ph1, "bild": "neu_hoch.png", "ausschnitt": {"oben_px": 586}}]},
        {"basis": 2, "ersetzen": {"Alte Frage?": "Neue Frage?"}, "aendern": [
            {"form": tb1, "absaetze": [None, "**Zweiter** Punkt neu", "Dritter Punkt"], "h": 1.4},
            {"form": "Kernaussage", "laeufe": ["Offen:  ", None]},
            {"formen": [tb1, karte_], "dy": 0.5},
            {"form": tb2, "entfernen": True},
            {"form": "Gibt es nicht", "y": 1}],
         "elemente": [{"form": "band", "label": "Heute:", "text": "von Hand.", "stil": "hell", "y": 1.43, "h": 0.44}]},
        {"basis": 3, "aendern": [{"form": ph3, "bild": "neu_gleich.png"}]}]}), encoding="utf-8")
    zv, wv = bauen.bauen(td / "aendern_v2.spec.json", still=True)
    fall("aendern: nur die Warnung zur fehlenden Form", len(wv) == 1 and "Gibt es nicht" in wv[0], "; ".join(wv))
    pv = Presentation(str(zv))
    s1, s2, s3 = pv.slides[0], pv.slides[1], pv.slides[2]
    bild1 = [s for s in s1.shapes if s.name == ph1][0]
    sichtbar = 1440 * bild1.height / bild1.width / 2000  # Anteil der Bildhöhe, der bei voller Breite in den Rahmen passt
    fall("Bild getauscht, volle Breite, Zuschnitt ab Pixel 586", bild1.image.size == (1440, 2000)
         and abs(bild1.crop_top - 586 / 2000) < 0.001 and abs(bild1.crop_bottom - (1 - 586 / 2000 - sichtbar)) < 0.001
         and abs(bild1.crop_left) < 0.001 and abs(bild1.crop_right) < 0.001,
         f"{bild1.image.size} {bild1.crop_top:.3f} {bild1.crop_bottom:.3f} (erwartet {1 - 586 / 2000 - sichtbar:.3f})")
    bildrels = [r for r in s1.part.rels.values() if r.reltype.endswith("/image")]
    fall("altes Bild nicht mehr verknüpft", len(bildrels) == 1, str(len(bildrels)))
    bild3 = [s for s in s3.shapes if s.name == ph3][0]
    fall("gleiches Seitenverhältnis: Zuschnitt bleibt", bild3.image.size == (1440, 1000)
         and abs(bild3.crop_top - alt_crop[0]) < 0.001 and abs(bild3.crop_bottom - alt_crop[1]) < 0.001,
         f"{(bild3.crop_top, bild3.crop_bottom)} statt {alt_crop}")
    form1 = [s for s in s2.shapes if s.name == tb1][0]
    ps = form1.text_frame.paragraphs
    fall("Absätze gesetzt und ergänzt", [p.text for p in ps] == ["Erster Punkt", "Zweiter Punkt neu", "Dritter Punkt"],
         str([p.text for p in ps]))
    fall("neuer Absatz im Format des letzten, **fett** als eigener Lauf",
         ps[2].runs[0].font.size.pt == 14 and [(r.text, bool(r.font.bold)) for r in ps[1].runs]
         == [("Zweiter", True), (" Punkt neu", False)], str([(r.text, r.font.bold) for r in ps[1].runs]))
    fall("Form verschoben und vergrößert", abs(form1.top - E(2.0)) < 1000 and abs(form1.height - E(1.4)) < 1000,
         f"{form1.top / 914400:.2f} {form1.height / 914400:.2f}")
    karte2 = [s for s in s2.shapes if s.name == karte_][0]
    fall("Reihe mit formen + dy verschoben", abs(karte2.top - E(2.0)) < 1000, f"{karte2.top / 914400:.2f}")
    fall("Form entfernt", not [s for s in s2.shapes if s.name == tb2])
    baender = [s for s in s2.shapes if s.name == "Kernaussage"]
    unten = max(baender, key=lambda s: s.top)
    fall("Band: Läufe einzeln, Formatierung bleibt", [(r.text, bool(r.font.bold)) for r in
                                                       unten.text_frame.paragraphs[0].runs]
         == [("Offen:  ", True), ("Neue Frage?", False)], str([(r.text, r.font.bold) for r in
                                                              unten.text_frame.paragraphs[0].runs]))
    fall("elemente: helles Band als Untertitel", len(baender) == 2 and abs(min(s.top for s in baender) - E(1.43)) < 1000,
         str([round(s.top / 914400, 2) for s in baender]))
    fall("Prüfung der geänderten Fassung ohne Fehler", not [b for b in pruefen.pruefen(zv) if b[0] == "FEHLER"],
         "; ".join(b[2] for b in pruefen.pruefen(zv) if b[0] == "FEHLER"))
    gp = Presentation()
    gs = gp.slides.add_slide(gp.slide_layouts[6])
    gruppe = gs.shapes.add_group_shape()
    gruppe.name = "Reihe 1"
    gruppe.shapes.add_textbox(E(1), E(1), E(2), E(0.5)).text_frame.text = "in der Gruppe"
    from bippt.bausteine import Kontext
    gw = []
    aendern.aendern(Kontext(gp, gs, Marke(), ("", ""), gw, 1), [{"form": "Reihe 1", "dy": 0.5}], td)
    fall("aendern: ganze Gruppe ansprechen und verschieben", not gw and abs(gruppe.top - E(1.5)) < 1000,
         f"{gw} {gruppe.top / 914400:.2f}")
    p = Presentation().slides.add_slide(Presentation().slide_layouts[6]).shapes.add_textbox(0, 0, 9, 9).text_frame.paragraphs[0]
    for t, fett in (("Stand ", False), ("07.10.", True), ("2026 fertig", False)):
        r = p.add_run()
        r.text, r.font.bold = t, fett
    aendern._ersetzen_absatz(p, "07.10.2026", "01.01.2027")
    fall("ersetzen über Läufe: Format außerhalb des Treffers bleibt", [(r.text, bool(r.font.bold)) for r in p.runs]
         == [("Stand ", False), ("01.01.2027", True), (" fertig", False)], str([(r.text, r.font.bold) for r in p.runs]))

    print("Folienbeschreibung und Baustand")
    from bippt import beschreiben

    def abbruch(spec_, **kw):
        try:
            bauen.bauen(spec_, still=True, **kw)
        except SystemExit as e:
            return str(e)
        return ""

    md = zp.with_suffix(".md")
    text = md.read_text(encoding="utf-8")
    bau1 = beschreiben.baustand(md)
    fall("Beschreibung neben dem Deck, alle Folientitel", md.is_file() and text.startswith(beschreiben.START)
         and text.rstrip().endswith(beschreiben.ENDE)
         and all(f"Folie {i}: {sl.shapes.title.text_frame.text}" in text for i, sl in enumerate(pp.slides, 1)))
    fall("Beschreibung: Inhalt aus der Spec (Ampel, Termine, Gestaltung)", "Ampel Gelb–Rot" in text
         and "| KW 40 | IT-Termin |" in text and "Gestaltung: Layout „Only title“ · 3 Strangkarten" in text
         and "Folie 12: Rollen im Projektteam\nTyp `tabelle` · Backup (ausgeblendet)" in text, text[:400])
    fall("Baustand je Folie, ausgeblendet mit ~", bau1 and len(bau1) == 12 and bau1[-1].endswith("~")
         and not bau1[0].endswith("~"), str(bau1))
    bauen.bauen(pspec, still=True)
    fall("Neubau ohne Änderung: gleicher Baustand", beschreiben.baustand(md) == bau1)
    sp = json.loads(pspec.read_text(encoding="utf-8"))
    sp["folien"][7]["zeilen"][1]["wann"] = "KW 41"
    pspec.write_text(json.dumps(sp, ensure_ascii=False), encoding="utf-8")
    bauen.bauen(pspec, still=True)
    bau2 = beschreiben.baustand(md)
    fall("Spec-Änderung an Folie 8: nur Folie 8 anders", beschreiben.aenderungen(bau1, bau2) == ([8], 11, 0),
         str(beschreiben.aenderungen(bau1, bau2)))
    hand = Presentation(str(zp))
    hand.slides[2].shapes.title.text_frame.text = "Von Hand geändert"
    hand.save(str(zp))
    fehlt = abbruch(pspec)
    fall("Handänderung: Neubau bricht ab und nennt die Folie", "nach dem letzten Bau geändert (Folie 3)" in fehlt, fehlt)
    beschreiben.beschreiben(zp)
    fall("beschreiben <pptx>: Baustand bleibt, Folie 3 aus dem Deck, Rest übernommen",
         beschreiben.baustand(md) == bau2 and "Folie 3: Von Hand geändert" in md.read_text(encoding="utf-8")
         and "| KW 41 | IT-Termin |" in md.read_text(encoding="utf-8"))
    fall("Handänderung: weiter Abbruch nach beschreiben", "Folie 3" in abbruch(pspec))

    print("Basis-Modus auf das von Hand geänderte Deck")
    v3 = {"basis": zp.name, "ausgabe": "v3.pptx", "fusszeile": "V3", "folien": [{"basis": n} for n in range(1, 13)]}
    v3["folien"][7] = {"basis": 8, "ersetzen": {"Mitte 2027": "Ende 2027"}, "beschreibung": "Letzter Termin verschoben"}
    (td / "v3.spec.json").write_text(json.dumps(v3, ensure_ascii=False), encoding="utf-8")
    z3, w3 = bauen.bauen(td / "v3.spec.json", still=True)
    p3 = Presentation(str(z3))
    t3 = z3.with_suffix(".md").read_text(encoding="utf-8")
    fall("Basis: Folien behalten ihren Zustand (Backup ausgeblendet)",
         [i for i, s_ in enumerate(p3.slides, 1) if s_._element.get("show") == "0"] == [10, 11, 12] and not w3, str(w3))
    fall("Basis: Handänderung bleibt", p3.slides[2].shapes.title.text_frame.text == "Von Hand geändert")
    fall("Basis: unveränderte Folie übernimmt ihre Beschreibung",
         "Basisfolie 12 · Typ `tabelle` · Backup (ausgeblendet)" in t3 and "unverändert aus der Basis" in t3, t3[-600:])
    fall("Basis: geänderte Folie aus dem Deck mit Beschreibung",
         "beim Bau geändert: 1 Ersetzung\nBeschreibung: Letzter Termin verschoben" in t3
         and "- Ende 2027 | Schnittstelle validiert" in t3, t3[t3.find("Folie 8:"):][:500])
    bauen.bauen(td / "v3.spec.json", still=True)
    fall("Basis-Spec erneut gebaut: gleicher Baustand", beschreiben.baustand(z3.with_suffix(".md"))
         == beschreiben.fingerabdruecke(p3))
    v3["ausgabe"] = zp.name
    (td / "selbst.spec.json").write_text(json.dumps(v3, ensure_ascii=False), encoding="utf-8")
    fall("Basis = Ausgabe bricht ab", "dieselbe Datei" in abbruch(td / "selbst.spec.json"))
    zp.with_suffix(".md").unlink()
    fall("Deck ohne Baustand: Abbruch, --erzwingen baut", "ohne Baustand" in abbruch(pspec)
         and not abbruch(pspec, erzwingen=True) and beschreiben.baustand(md))

    print("Vorlage als dokument (z. B. von ppp)")
    (td / "Entwurf.md").write_text("# Entwurf aus ppp\n\nFolie 1: Titel\n", encoding="utf-8")
    sp.update({"ausgabe": "mit_vorlage.pptx", "dokument": "Entwurf.md", "zweck": "Test", "quellen": ["Entwurf.md"]})
    (td / "vorlage.spec.json").write_text(json.dumps(sp, ensure_ascii=False), encoding="utf-8")
    bauen.bauen(td / "vorlage.spec.json", still=True)
    bauen.bauen(td / "vorlage.spec.json", still=True)
    tv = (td / "Entwurf.md").read_text(encoding="utf-8")
    fall("Vorlage bleibt, Abschnitt einmal angehängt", tv.startswith("# Entwurf aus ppp\n\nFolie 1: Titel\n\n<!-- bippt")
         and tv.count(beschreiben.START) == 1 and "\n### Folie 3: " in tv and "Zweck: Test" in tv
         and not (td / "mit_vorlage.md").exists(), tv[:300])
    fall("analyse --folien liest nur diese Folien", [f["nr"] for f in analyse.analyse(zp, folien=[2, 5])["folien"]]
         == [2, 5])

    print("Messen beim Bau (Befunde brrp 28.10.2026): ersetzen, Titel, Wortbreite, Tabelle, Quelle")
    (td / "ers.spec.json").write_text(json.dumps({"ausgabe": "ers.pptx", "folien": [
        {"typ": "karten", "titel": "Rollen", "karten": [
            {"titel": "Team", "punkte": ["Entfällt mit GBS", "Prüft Chargen", "Meldet Abweichungen"]},
            {"titel": "SCM", "punkte": ["Plant Freigaben", "Fragt nach", "Eskaliert", "X1 und Y1"]}]}]}), encoding="utf-8")
    bauen.bauen(td / "ers.spec.json", still=True)
    zuordnung = {"Entfällt mit GBS": "Entfällt mit GBS United", "Prüft Chargen": "Prüft alle Chargen",
                 "Meldet": "Meldet sofort", "Plant Freigaben": "Plant die Freigaben", "Fragt nach": "Fragt gezielt nach",
                 "Eskaliert": "Eskaliert früh", "X1": "X1Y1", "Y1": "Z1", "Gibt es nicht": "egal"}
    (td / "ers_v2.spec.json").write_text(json.dumps({"basis": "ers.pptx", "ausgabe": "ers_v2.pptx", "folien": [
        {"basis": 1, "ersetzen": zuordnung}]}), encoding="utf-8")
    ze, we = bauen.bauen(td / "ers_v2.spec.json", still=True)
    te = "\n".join(sh.text_frame.text for sh in Presentation(str(ze)).slides[0].shapes if sh.has_text_frame)
    fall("ersetzen: nur die wirklich fehlende Ersetzung gemeldet", len(we) == 1 and "1 von 9 Ersetzungen" in we[0]
         and "Gibt es nicht" in we[0], "; ".join(we))
    fall("ersetzen: neuer Text, der den alten enthält, genau einmal", te.count("GBS United") == 1 and "United United"
         not in te and "Meldet sofort Abweichungen" in te and "Eskaliert früh" in te and "früh früh" not in te, te)
    fall("ersetzen: eingesetzter Text wird nicht von einem anderen Suchtext ersetzt", "X1Y1 und Z1" in te, te)

    from bippt.bausteine import umbrechen
    fall("Umbruch: nach Bindestrich und Gedankenstrich, nicht nach Schrägstrich",
         [w for w, _ in umbrechen("Power/Apps/Flows Ein–Aus–Schalter Projekt-Workshop", 0.74, 12)[1]]
         == ["Power/Apps/Flows", "Workshop"], str(umbrechen("Power/Apps/Flows Ein–Aus–Schalter Projekt-Workshop", 0.74, 12)))

    schritte = [{"ueber": f"KW {40 + i}", "titel": t, "punkte": [p]} for i, (t, p) in enumerate([
        ("Analyse", "SOP und WI lesen"), ("Projektworkshop", "Gemeinsames Verständnis"), ("Rundgang", "Aufzeichnung"),
        ("Messung", "Ausgangswerte erheben"), ("Klärung", "Anforderungen festhalten"), ("Prototyp", "Klickdummy"),
        ("Umsetzung", "Power Apps und Flows"), ("Einführung", "Schulung")])]
    karten2 = [{"titel": "Heute", "punkte": ["Mails"]}, {"titel": "Ziel", "punkte": ["Cockpit"]}]
    band_ = {"label": "Fazit:", "text": "Die letzte Zeile muss sichtbar sein."}

    def tab(n, **kw):
        return {"typ": "tabelle", "titel": f"Tabelle {n}", "kopf": ["Schritt", "Was", "Wer"], "breiten": [1, 3, 1],
                "zeilen": [[f"Schritt {i}", f"Beschreibung des Schritts {i}", "Team"] for i in range(1, n + 1)], **kw}

    (td / "messen.spec.json").write_text(json.dumps({"ausgabe": "messen.pptx", "folien": [
        {"typ": "karten", "titel": "Vom Rundgang zum Zielbild: Was der Workshop klären soll und was nicht",
         "karten": karten2},
        {"typ": "karten", "titel": "Vom Rundgang zum Zielbild: Was der Projektworkshop klären soll und was nicht",
         "karten": karten2},
        {"typ": "text", "titel": "Vom Rundgang zum Zielbild: Was der Projektworkshop klären soll und was nicht",
         "punkte": ["Eins"]},
        {"typ": "prozess", "titel": "Acht Schritte", "schritte": schritte},
        {"typ": "prozess", "titel": "Sechs Schritte", "schritte": schritte[:6]},
        tab(13, groesse=10, band=band_),
        tab(30, band=band_),
        {"typ": "prozess", "titel": "Quelle, Band, Karten", "quelle": "Quelle: SOP und WI", "band": band_,
         "schritte": [{"titel": "Eingang", "text": "Charge"}, {"titel": "Prüfung", "text": "Checkliste"}],
         "karten": karten2},
        tab(17, groesse=12, band=band_),
        {"typ": "text", "titel": "Text mit Band und Quelle", "punkte": ["Eins", "Zwei"], "band": band_,
         "quelle": "Quelle: Projektauftrag"}]}, ensure_ascii=False), encoding="utf-8")
    zm, wm = bauen.bauen(td / "messen.spec.json", still=True)
    pm = Presentation(str(zm))

    def warn_folie(n):
        return [w for w in wm if w.startswith(f"Folie {n}:")]

    def unterkante_(sh):
        return (sh.top + sh.height) / 914400

    fall("Titel mit 69 Zeichen passt in eine Zeile", not warn_folie(1), "; ".join(warn_folie(1)))
    fall("Titel mit 76 Zeichen auf karten: Warnung zwei Zeilen", any("Titel braucht 2 Zeilen" in w and "ragt in den Inhalt"
                                                                     in w for w in warn_folie(2)), "; ".join(wm))
    fall("zweizeiliger Titel über dem Inhaltsplatzhalter: keine Warnung", not warn_folie(3), "; ".join(warn_folie(3)))
    fall("pruefen meldet den zweizeiligen Titel", any(b_[0] == "WARNUNG" and b_[1] == 2 and "Titel braucht 2 Zeilen" in b_[2]
                                                     for b_ in pruefen.pruefen(zm, woerter=200)))
    w4 = " ".join(warn_folie(4))
    fall("8 Schrittkarten: zu breite Wörter gemeldet, Rat höchstens 6 Schritte", "Projektworkshop" in w4
         and "Verständnis" in w4 and "mitten im Wort" in w4 and "höchstens 6 Schritte" in w4 and len(warn_folie(4)) == 1, w4)
    fall("6 Schrittkarten: keine Warnung", not warn_folie(5), "; ".join(warn_folie(5)))
    for nr, n in ((6, 13), (8, None)):
        sl_ = pm.slides[nr - 1]
        band_top = min(sh.top for sh in sl_.shapes if sh.name == "Kernaussage") / 914400
        inhalt = [sh for sh in sl_.shapes if sh.name != "Kernaussage" and not sh.is_placeholder]
        unten_ = max(sh.top / 914400 + (sum(r.height for r in sh.table.rows) / 914400 if sh.has_table else sh.height / 914400)
                     for sh in inhalt)
        fall(f"Folie {nr}: Inhalt endet über dem Band", unten_ <= band_top + 0.005 and not warn_folie(nr),
             f"{unten_:.2f} > {band_top:.2f} {warn_folie(nr)}")
    fall("Tabelle mit 31 Zeilen: konkrete Warnung mit Bedarf und Platz", any("Tabelle: 31 Zeilen brauchen" in w and
                                                                             "unter dem Band" in w for w in warn_folie(7)),
         "; ".join(warn_folie(7)))
    fall("Tabelle mit fester groesse: Rat zur passenden Größe", any("groesse 9.5 (statt 12) passt" in w
                                                                    for w in warn_folie(9)), "; ".join(warn_folie(9)))
    quelle_ = [sh for sh in pm.slides[7].placeholders if sh.placeholder_format.idx == 13]
    kb = [sh for sh in pm.slides[7].shapes if sh.name == "Kernaussage"]
    fall("Quelle und Band: Band endet über der Quellenzeile", quelle_ and kb and unterkante_(kb[0]) <= quelle_[0].top / 914400,
         f"{unterkante_(kb[0]) if kb else '-'} / {quelle_[0].top / 914400 if quelle_ else '-'}")
    ph_ = [sh for sh in pm.slides[9].placeholders if sh.placeholder_format.idx == 1]
    lay_ = [sh for sh in pm.slides[9].slide_layout.placeholders if sh.placeholder_format.idx == 1]
    kb = [sh for sh in pm.slides[9].shapes if sh.name == "Kernaussage"]
    fall("Textfolie mit Band: Inhaltsplatzhalter behält Lage und Breite, endet über dem Band",
         ph_ and lay_ and kb and ph_[0]._element.spPr.find(qn("a:xfrm")).find(qn("a:off")) is not None
         and (ph_[0].left, ph_[0].top, ph_[0].width) == (lay_[0].left, lay_[0].top, lay_[0].width)
         and unterkante_(ph_[0]) <= kb[0].top / 914400 + 0.005,
         f"{[(sh.left, sh.top, sh.width, sh.height) for sh in ph_]} Layout {[(sh.left, sh.width) for sh in lay_]}")
    from pptx.util import Inches  # noqa: E402
    sl_ = pm.slides[9]
    sl_.shapes.add_textbox(Inches(1), Inches(2), 0, Inches(2)).text_frame.text = "Breite null"
    pm.save(str(td / "schmal.pptx"))
    fall("pruefen meldet ein Textfeld mit Breite 0 als FEHLER", any(b_[0] == "FEHLER" and b_[1] == 10 and "0.00 in breit"
                                                                    in b_[2] for b_ in pruefen.pruefen(td / "schmal.pptx")))

print("Bild und Text, Bilder prüfen")
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    from PIL import Image  # noqa: E402
    Image.new("RGB", (400, 120), (8, 49, 42)).save(td / "Boehringer_Logo_RGB_Dark_Green.png")
    lang = "Ein Punkt mit so viel Text, dass er in der schmalen Textspalte neben dem Bild mehrere Zeilen braucht"
    (td / "bild.spec.json").write_text(json.dumps({"ausgabe": "bild.pptx", "folien": [
        {"typ": "bild_text", "titel": "Bild und Text mit Band", "bild": "Boehringer_Logo_RGB_Dark_Green.png",
         "punkte": [lang] * 4, "band": {"label": "Fazit:", "text": "Bild und Text enden über dem Band."}}]},
        ensure_ascii=False), encoding="utf-8")
    zb, wb = bauen.bauen(td / "bild.spec.json", still=True)
    sl_ = Presentation(str(zb)).slides[0]
    kb = [sh for sh in sl_.shapes if sh.name == "Kernaussage"][0]
    pic = [sh for sh in sl_.shapes if sh.shape_type == 13 or (sh.is_placeholder and sh.placeholder_format.idx == 10)][0]
    tx = [sh for sh in sl_.placeholders if sh.placeholder_format.idx == 1][0]
    fall("bild_text mit Band: Bild und Text enden über dem Band",
         (pic.top + pic.height) <= kb.top + 4572 and (tx.top + tx.height) <= kb.top + 4572 and tx.width > 914400,
         f"Bild {(pic.top + pic.height) / 914400:.2f}, Text {(tx.top + tx.height) / 914400:.2f}, Band {kb.top / 914400:.2f}")
    fall("Warnung: BI-Logo als Bild", any("BI-Logo" in w for w in wb), "; ".join(wb))
    fall("Warnung: Bild unscharf (dpi)", any("dpi im sichtbaren Ausschnitt" in w for w in wb), "; ".join(wb))
    fall("Warnung: Text läuft über den Platzhalter", any("Text braucht etwa" in w and "läuft unten über" in w
                                                         for w in wb), "; ".join(wb))

print("Folienbibliothek")
from bippt import bibliothek  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
BIB = PLUGIN / "vorlagen" / "BI_Slide_Library.pptx"
fall("Bibliothek im Plugin", konfig.bibliothek_pfad({"bibliothek": ""}) == BIB)
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    Image.new("RGB", (2400, 1600), (0, 228, 124)).save(td / "foto.jpg")
    spalte = {"10": "Gebäude", "9": ["Neubau mit vier Etagen", "Bezug ab 03/2027"], "17": "Technik", "16": ["Probenlogistik"],
              "24": "Menschen", "23": ["120 Arbeitsplätze"]}
    (td / "bib.spec.json").write_text(json.dumps({"ausgabe": "bib.pptx", "fusszeile": "Probe | Bibliothek", "folien": [
        {"bib": 25, "titel": "Drei Bausteine", "texte": spalte},
        {"bib": 25, "titel": "Dieselbe Folie noch einmal", "texte": dict(spalte, **{"23": ["Viel zu lang " * 60]})},
        {"bib": 62, "titel": "Bild und Text", "texte": {"2": ["Offene Labore"]}, "bilder": {"11": "foto.jpg"}},
        {"bib": 119, "titel": "Kosten", "daten": {"16": {"titel": "Kosten in Mio. €", "kategorien": ["2025", "2026"],
                                                         "reihen": [{"name": "Bau", "werte": [10, 20]}]}}},
        {"bib": 2, "titel": "Inhalt mit Sprüngen"},
        {"bib": 114, "titel": "Zahlen je Werk", "tabellen": {"17": [["Werk", "2025"], ["Ingelheim", "1"], ["Biberach", "2"],
                                                                     ["Wien", "3"], ["Summe", "6"]]}},
        {"bib": 93, "titel": "Zeitplan mit Blindtext"}]}, ensure_ascii=False), encoding="utf-8")
    zl, wl = bauen.bauen(td / "bib.spec.json", still=True)
    pl = Presentation(str(zl))

    def form_(sl, sid):
        return next((sh for sh in sl.shapes if sh.shape_id == sid), None)

    tab_ = next(sh for sh in pl.slides[5].shapes if getattr(sh, "has_table", False))
    fall("tabellen: Zeilen und Spalten nach Angabe, Breite bleibt",
         [[c.text for c in r.cells] for r in tab_.table.rows][-1] == ["Summe", "6"] and len(tab_.table.rows) == 5
         and abs(sum(c.width for c in tab_.table.columns) - tab_.width) < 9144)
    fall("Bibliotheksfolien kopiert, auch dieselbe zweimal", len(pl.slides) == 7
         and [s_.slide_layout.name for s_ in pl.slides][:3] == ["Only title", "Only title", "Content with image portrait (1)"])
    fall("texte füllen die Formen nach id", form_(pl.slides[0], 10).text_frame.text == "Gebäude"
         and form_(pl.slides[1], 10).text_frame.text == "Gebäude" and pl.slides[0].shapes.title.text_frame.text == "Drei Bausteine")
    bild_ = form_(pl.slides[2], 11)
    fall("bilder: Foto im Bildplatzhalter", bild_ is not None and bild_.image.size == (2400, 1600))
    ch = form_(pl.slides[3], 16).chart
    fall("daten: Diagramm mit neuen Werten und Titel", list(ch.plots[0].categories) == ["2025", "2026"]
         and list(ch.series[0].values) == [10, 20] and ch.chart_title.text_frame.text == "Kosten in Mio. €")
    fall("Sprünge zu anderen Bibliotheksfolien entfernt", not any(r.reltype == RT.SLIDE for s_ in pl.slides
                                                                 for r in s_.part.rels.values()))
    fall("Notizen der Bibliothek kommen nicht mit", not any(s_.has_notes_slide for s_ in pl.slides))
    fuss_ = [sh for sh in pl.slides[0].placeholders if sh.placeholder_format.type is not None
             and "FOOTER" in str(sh.placeholder_format.type)]
    fall("Fußzeile mit Präsentationstitel", fuss_ and "Probe | Bibliothek" in fuss_[0].text_frame.text)
    fall("Warnung: Text zu lang für seine Form", any(w.startswith("Folie 2:") and "Text braucht etwa" in w for w in wl),
         "; ".join(wl))
    fall("Warnung: Blindtext übrig", any(w.startswith("Folie 7:") and "Blindtext" in w for w in wl)
         and not any(w.startswith("Folie 1:") and "Blindtext" in w for w in wl), "; ".join(wl))
    fall("Warnung: große Fläche halb leer", any(w.startswith("Folie 1:") and "nur zu" in w for w in wl), "; ".join(wl))
    import zipfile  # noqa: E402
    namen = zipfile.ZipFile(zl).namelist()
    fall("Paket ohne doppelte Teile", len(namen) == len(set(namen)))
    fall("pruefen ohne FEHLER", not [b_ for b_ in pruefen.pruefen(zl, woerter=300) if b_[0] == "FEHLER"])
    md_ = (td / "bib.md").read_text(encoding="utf-8")
    fall("Folienbeschreibung nennt Bibliotheksfolie und Inhalt spaltenweise",
         "Bibliotheksfolie 25" in md_ and "- Gebäude\n- Neubau mit vier Etagen" in md_)
    kat = bibliothek.katalog(BIB, [25])
    fall("katalog nennt ids, Platz und Beispieltext", "## 25: Text box: Three columns" in kat
         and "id 10 · Text" in kat and "Platz ~" in kat and "„Header box“" in kat and "Absätze: 4× Punkt" in kat)

print("Abwechslung und Layouts des Masters")
einerlei = [{"typ": "titel", "titel": "T"}] + [{"bib": 25} for _ in range(4)] + [{"typ": "tabelle"}]
wa = bibliothek.abwechslung(einerlei)
fall("Eintönig: drei gleiche Bauarten in Folge, mehr als die Hälfte, zu viel Text", any("Folien 2–4" in w for w in wa)
     and any("4 von 5" in w and "mehr als die Hälfte" in w for w in wa) and any("Textlastig" in w for w in wa), "; ".join(wa))
wieder = [{"typ": "titel"}, {"bib": 36}, {"bib": 36}, {"bib": 39}, {"bib": 81}, {"bib": 36}, {"typ": "status"}, {"bib": 81}]
fall("Wiederholen ist erlaubt: gleiche Folie zweimal in Folge und mehrfach ohne Warnung",
     not bibliothek.abwechslung(wieder), "; ".join(bibliothek.abwechslung(wieder)))
fall("Umlaute: Umschrift erkannt, normale Wörter nicht", bibliothek.schreibweise([{"titel": "Einfuehrung und Pruefung", "texte": {
    "1": ["aktuell manuell neue Steuerung Quelle Frequenz Prozesse"]}}]) == [
        "Umlaute: „Einfuehrung“, „Pruefung“ – echte Umlaute schreiben (ä, ö, ü); Eigennamen bleiben wie in der Quelle"], str(bibliothek.schreibweise([{"titel": "Einfuehrung"}])))
bunt = [{"typ": "titel"}, {"bib": 10}, {"bib": 36}, {"bib": 39}, {"bib": 81}, {"typ": "status"}, {"bib": 50}, {"bib": 182},
        {"typ": "kapitel", "ab_hier_ausblenden": True}, {"typ": "tabelle"}, {"typ": "tabelle"}, {"typ": "tabelle"}]
fall("Bunte Mischung ohne Warnung, Backup zählt nicht", not bibliothek.abwechslung(bunt), "; ".join(bibliothek.abwechslung(bunt)))
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    Image.new("RGB", (2400, 2700), (8, 49, 42)).save(td / "foto.jpg")
    (td / "lay.spec.json").write_text(json.dumps({"ausgabe": "lay.pptx", "folien": [
        {"typ": "layout", "layout": "Three contents", "titel": "Drei Spalten", "band": "Fazit",
         "felder": {"1": ["Eins", "Zwei"], "16": ["Drei"], "17": ["Vier"]}},
        {"typ": "layout", "layout": "Two contents", "titel": "Zwei Spalten", "felder": {"1": "Links", "16": "Rechts"}},
        {"typ": "layout", "layout": "Content with image portrait (1)", "titel": "Bild", "felder": {"1": ["Text"]},
         "bilder": {"10": "foto.jpg"}},
        {"typ": "layout", "layout": "Content with image portrait (1)", "titel": "Ohne Bild", "felder": {"9": "x"}}]},
        ensure_ascii=False), encoding="utf-8")
    zy, wy = bauen.bauen(td / "lay.spec.json", still=True)
    py = Presentation(str(zy))
    f1 = {sh.placeholder_format.idx: sh for sh in py.slides[0].placeholders}
    fall("Typ layout: echte Platzhalter des Master-Layouts", py.slides[0].slide_layout.name == "Three contents"
         and f1[16].text_frame.text == "Drei" and f1[17].text_frame.text == "Vier" and f1[17].width > 914400)
    kb = [sh for sh in py.slides[0].shapes if sh.name == "Kernaussage"]
    fall("Typ layout: Spalten enden über dem Band", kb and all(f1[i].top + f1[i].height <= kb[0].top + 4572 for i in (1, 16, 17)))
    fall("Typ layout: Bild im Platzhalter", any(sh.shape_type == 14 and sh.placeholder_format.idx == 10 and sh.image.size == (2400, 2700)
                                                for sh in py.slides[2].placeholders))
    fall("Typ layout: Warnung bei falschem Feld und leerem Bild", any("kein Textfeld 9" in w for w in wy)
         and any(w.startswith("Folie 4:") and "Bildplatzhalter 10 leer" in w for w in wy), "; ".join(wy))
    fall("layouts nennt die Felder", "**Three contents**: 1 Inhalt 3.9×4.8" in bibliothek.layouts(PLUGIN / "vorlagen" / "BI_Vorlage_16x9.pptx"))

print("Icons")
from bippt import marke as marke_mod  # noqa: E402
with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    for v, f in (("darkgreen", (8, 49, 42, 255)), ("white", (255, 255, 255, 255))):
        (td / f"Icons_{v}").mkdir()
        bild = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        bild.paste(f, (10, 10, 50, 50))
        bild.save(td / f"Icons_{v}" / "probe-icon.png")
    paket = td / "bi_icons.json"
    fall("iconpaket zählt die Icons", marke_mod.icon_paket_bauen(td, paket, groesse=32) == 1)
    alt_paket, marke_mod.ICON_PAKET = marke_mod.ICON_PAKET, paket
    marke_mod._PAKET.clear()
    weiss = marke_mod._aus_paket("probe-icon", "white")
    fall("Icon aus dem Paket in der Variantenfarbe", weiss is not None and Image.open(weiss).convert("RGBA").getpixel((16, 16)) == (255, 255, 255, 255))
    fall("Unbekanntes Icon: None", marke_mod._aus_paket("gibt-es-nicht", "white") is None)
    (td / "ic.spec.json").write_text(json.dumps({"ausgabe": "ic.pptx", "folien": [
        {"bib": 36, "titel": "Icon auf dunklem Band", "bilder": {"19": "icon:probe-icon", "23": "icon:gibt-es-nicht"},
         "texte": {"17": ["Eins", "Ein Satz mit genügend Wörtern für die Prüfung"], "16": ["Zwei", "Noch ein ganzer Satz mit Inhalt für die Folie"]}},
        {"bib": 81, "titel": "Zerhackt", "texte": {"18": ["Kopf", "Tool", "Programmieren", "Test"]}}]}, ensure_ascii=False), encoding="utf-8")
    zi, wi = bauen.bauen(td / "ic.spec.json", still=True)
    ic = next(sh for sh in Presentation(str(zi)).slides[0].shapes if sh.shape_id == 19)
    px = Image.open(__import__("io").BytesIO(ic.image.blob)).convert("RGBA").getpixel((16, 16))
    fall("bilder icon: auf dunklem Band die weiße Variante", px[:3] == (255, 255, 255), str(px))
    fall("Warnung: unbekanntes Icon", any("„gibt-es-nicht“ nicht gefunden" in w for w in wi), "; ".join(wi))
    fall("Warnung: Einzelwörter und wenig Text", any("Einzelwörtern" in w and "„Tool“" in w for w in wi) and any("Nur 4 Wörter" in w for w in wi), "; ".join(wi))
    marke_mod.ICON_PAKET = alt_paket
    marke_mod._PAKET.clear()

print("ALLE FÄLLE BESTANDEN" if not fehler else f"FEHLGESCHLAGEN: {', '.join(fehler)}")
sys.exit(1 if fehler else 0)
