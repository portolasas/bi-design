"""Präsentation aus einer Spezifikation (JSON) bauen.

Spezifikation (Auszug, vollständig in reference/folientypen.md):
{
  "ausgabe": "Projekt_Update_07.10.2026.pptx",   relativ zur Spec
  "vorlage": "…/BI_Vorlage_16x9.pptx",            optional; Standard aus ~/.bi-ppt/config.json
  "basis":   "Altes_Deck.pptx",                   optional; dann ist das alte Deck Vorlage UND Folienquelle
  "fusszeile": "Projekt | Anlass | 07.10.2026",
  "heute": "2026-10-07",
  "folien": [ {"typ": "titel", …}, {"basis": 7, "ausblenden": true, "ersetzen": {"alt": "neu"}}, … ]
}
Folien aus der Basis, die nicht in "folien" stehen, werden entfernt.
"""
import copy
import datetime as dt
import json
import re
import sys
from pathlib import Path

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

from . import konfig, vorlage
from .bausteine import Kontext, band
from .folien import TYPEN
from .marke import Marke


def _theme_schriften(prs):
    try:
        th = prs.slide_masters[0].part.part_related_by(RT.THEME)
        x = etree.fromstring(th.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        kopf = x.find(".//a:majorFont/a:latin", ns).get("typeface")
        txt = x.find(".//a:minorFont/a:latin", ns).get("typeface")
        return kopf, txt
    except (KeyError, AttributeError, etree.XMLSyntaxError):
        return "", ""


def _absatz_setzen(p, text):
    """Text eines Absatzes ersetzen, Format des ersten Laufs bleibt."""
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.add_run().text = text


def platzhalter_zeilen(ph, zeilen):
    """Zeilen in einen vorhandenen Platzhalter schreiben, Absatzformate bleiben erhalten."""
    zeilen = zeilen if isinstance(zeilen, list) else [zeilen]
    tf = ph.text_frame
    ps = list(tf.paragraphs)
    for i, z in enumerate(zeilen):
        if i < len(ps):
            _absatz_setzen(ps[i], z)
        else:
            neu = copy.deepcopy(ps[-1]._p)
            tf._txBody.append(neu)
            _absatz_setzen(tf.paragraphs[-1], z)
    for p in ps[len(zeilen):]:
        p._p.getparent().remove(p._p)


def _textrahmen(folie):
    for sh in folie.shapes:
        for x in (sh.shapes if sh.shape_type == 6 else [sh]):
            if x.has_text_frame:
                yield x.text_frame
            if getattr(x, "has_table", False) and x.has_table:
                for row in x.table.rows:
                    for c in row.cells:
                        yield c.text_frame


def ersetzen(folie, zuordnung):
    """Textstellen ersetzen (auch über Läufe hinweg, in Gruppen und Tabellen). Rückgabe: Anzahl Treffer."""
    n = 0
    for tf in _textrahmen(folie):
        for p in tf.paragraphs:
            voll = "".join(r.text for r in p.runs)
            neu = voll
            for a, b in zuordnung.items():
                if a in neu:
                    neu = neu.replace(a, b)
            if neu != voll:
                n += 1
                _absatz_setzen(p, neu)
    return n


def _fusszeile_basis(folie, text):
    for ph in folie.placeholders:
        if ph._element.find(".//" + qn("p:ph")).get("type") == "ftr" and ph.text_frame.paragraphs:
            _absatz_setzen(ph.text_frame.paragraphs[0], text)


def farben_angleichen(folie, marke):
    """Farben außerhalb von Marke + Vorlage auf die nächste Markenfarbe setzen (nur diese Folie)."""
    from .umwandeln import _naechste
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    palette = sorted(marke.erlaubte_farben() | {"000000", "FFFFFF"})
    n = 0
    for el in folie._element.iter(a + "srgbClr"):
        v = el.get("val", "").upper()
        if v not in palette:
            el.set("val", _naechste(v, palette))
            n += 1
    return n


def _pfad(basisordner, p):
    if not p:
        return None
    q = Path(p).expanduser()
    return q if q.is_absolute() else (basisordner / q)


def bauen(spec_pfad, ausgabe=None, still=False):
    spec_pfad = Path(spec_pfad)
    spec = json.loads(spec_pfad.read_text(encoding="utf-8"))
    ordner = spec_pfad.parent
    warnungen = []
    m = Marke(spec.get("marke"))
    basis = _pfad(ordner, spec.get("basis"))
    quelle = basis or _pfad(ordner, spec.get("vorlage")) or konfig.vorlage_pfad()
    if quelle is None:
        warnungen.append("Keine BI-Vorlage gefunden – Notbetrieb ohne Master/Logo. "
                         "`bippt vorlage <BI-Deck.pptx>` ausführen oder vorlage in der Spec angeben.")
    elif not Path(quelle).is_file():
        raise SystemExit(f"Vorlage/Basis nicht gefunden: {quelle}")
    prs = vorlage.oeffnen(quelle)
    if quelle is not None and not vorlage.ist_bi_master(prs):
        warnungen.append(f"{Path(quelle).name}: kein BI-Master – erst umwandeln (bippt umwandeln) oder BI-Vorlage nutzen")
    m.aus_vorlage(prs)
    alte = list(prs.slides)
    if not basis and alte:
        vorlage.alle_folien_entfernen(prs)
        alte = []
    kopf, txt = _theme_schriften(prs)
    schriften = ("+mj-lt", "+mn-lt") if "Boehringer" in kopf + txt else (m.schrift_kopf, m.schrift_text)
    heute = spec.get("heute") or dt.date.today().isoformat()
    cfg = konfig.laden()
    vermerk = spec.get("vertraulichkeit", cfg.get("vertraulichkeit", "ONLY FOR INTERNAL USE"))
    trenner = "     |     "  # wie im BI-Master v2.1
    fuss = trenner.join(x for x in (spec.get("fusszeile"), vermerk) if x) or None
    reihenfolge, ausblenden_ab = [], False
    for nr, fs in enumerate(spec.get("folien", []), 1):
        if "basis" in fs:
            i = int(fs["basis"]) - 1
            if not 0 <= i < len(alte):
                raise SystemExit(f"Folie {nr}: basis {fs['basis']} existiert nicht (Basis hat {len(alte)} Folien)")
            sl = alte[i]
            if fs.get("ersetzen"):
                treffer = ersetzen(sl, fs["ersetzen"])
                if treffer < len(fs["ersetzen"]):
                    warnungen.append(f"Folie {nr} (Basis {i + 1}): nur {treffer} von {len(fs['ersetzen'])} Ersetzungen gefunden")
            for idx, zeilen in (fs.get("platzhalter") or {}).items():
                ph = vorlage.platzhalter(sl, idx=int(idx))
                if ph is None:
                    warnungen.append(f"Folie {nr}: Platzhalter {idx} fehlt")
                else:
                    platzhalter_zeilen(ph, zeilen)
            if fs.get("titel") and sl.shapes.title is not None:
                platzhalter_zeilen(sl.shapes.title, fs["titel"])
            if fs.get("farben_angleichen", spec.get("basis_farben_angleichen")):
                n = farben_angleichen(sl, m)
                if n:
                    warnungen.append(f"Folie {nr} (Basis {i + 1}): {n} Nicht-Markenfarben angeglichen – im Render prüfen")
            if fuss and fs.get("fusszeile", True):
                _fusszeile_basis(sl, vermerk if sl.slide_layout.name.startswith("Title slide") else fuss)
        else:
            typ = fs.get("typ", "text")
            if typ not in TYPEN:
                raise SystemExit(f"Folie {nr}: unbekannter typ '{typ}' – vorhanden: {', '.join(sorted(TYPEN))}")
            fn, lay = TYPEN[typ]
            if fs.get("dunkel") and lay + "_dunkel" in vorlage.LAYOUTS:
                lay += "_dunkel"
            lay = fs.get("layout", lay)
            sl = prs.slides.add_slide(vorlage.layout(prs, lay))
            k = Kontext(prs, sl, m, schriften, warnungen, nr)
            k.heute = fs.get("heute", heute)
            k.dunkel = lay in vorlage.DUNKLE_LAYOUTS
            if fs.get("titel") is not None and sl.shapes.title is not None:
                sl.shapes.title.text = re.sub(r"\*\*|==", "", fs["titel"])
            fn(k, fs)
            if fs.get("band") and typ not in ("titel", "kapitel", "kernaussage"):
                band(k, fs["band"])
            if fs.get("quelle"):
                if not vorlage.quelle(sl, fs["quelle"]):
                    warnungen.append(f"Folie {nr}: Layout ohne Quellenzeile")
            vorlage.leere_platzhalter_entfernen(sl)
            # Titelfolien: nur der Vertraulichkeitsvermerk (How-to slides, BI-Master v2.1)
            text_fuss = (vermerk or None) if typ == "titel" else fuss
            vorlage.fusszeile(sl, text_fuss if fs.get("fusszeile", True) else None, nummer=typ != "titel")
        if fs.get("notizen"):
            sl.notes_slide.notes_text_frame.text = fs["notizen"]
        if fs.get("ab_hier_ausblenden"):
            ausblenden_ab = True
        if fs.get("ausblenden") or ausblenden_ab:
            sl._element.set("show", "0")
        elif sl._element.get("show") is not None:
            del sl._element.attrib["show"]
        reihenfolge.append(sl)
    # Reihenfolge herstellen, nicht verwendete Basisfolien entfernen
    lst = prs.slides._sldIdLst
    nach_part = {id(prs.part.related_part(e.rId)): e for e in lst}
    neu = [nach_part[id(s.part)] for s in reihenfolge]
    for e in list(lst):
        lst.remove(e)
        if e not in neu:
            prs.part.drop_rel(e.rId)
    for e in neu:
        lst.append(e)
    ziel = Path(ausgabe) if ausgabe else _pfad(ordner, spec.get("ausgabe") or spec_pfad.with_suffix(".pptx").name)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(ziel))
    if not still:
        sichtbar = sum(1 for s in reihenfolge if s._element.get("show") != "0")
        print(f"GEBAUT: {ziel} – {len(reihenfolge)} Folien ({sichtbar} sichtbar)")
        for w in warnungen:
            print(f"  WARNUNG {w}")
        print(f"ERGEBNIS: {'OK' if not warnungen else str(len(warnungen)) + ' Warnung(en)'}")
    return ziel, warnungen


if __name__ == "__main__":
    bauen(sys.argv[1])
