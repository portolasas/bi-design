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

print("ALLE FÄLLE BESTANDEN" if not fehler else f"FEHLGESCHLAGEN: {', '.join(fehler)}")
sys.exit(1 if fehler else 0)
