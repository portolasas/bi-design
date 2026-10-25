"""Grundbausteine für Folien im BI-Design: Flächen, Text, Linien, Punkte, Rauten, Pfeile, Icons, Band.

Maße in Zoll (inch). Farben nur über die Marke (Kürzel DG, AG, … oder Token-Namen).
Textangaben (überall gleich):
  "Text mit **fett** und ==Hervorhebung=="          ein Absatz
  ["Punkt 1", "Punkt 2"]                          mehrere Absätze
  {"t": "...", "fett": true, "groesse": 12, "farbe": "AG", "punkt": true, "ebene": 1, "ausr": "m", "nach": 4}
"""
import math
import re

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

AUSR = {"l": PP_ALIGN.LEFT, "m": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANKER = {"o": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "u": MSO_ANCHOR.BOTTOM}

# Fläche der Inhaltsseite (BI-Master, 13,33 × 7,5 in)
X0, XW = 0.52, 12.29
Y0, Y_ENDE = 1.45, 6.45
BAND_H = 0.44


def E(inch):
    return Emu(int(round(inch * 914400)))


def zoll(emu):
    return emu / 914400


class Kontext:
    """Eine Folie im Bau: Folie, Marke, Schriften, Warnungen."""

    def __init__(self, prs, folie, marke, schriften, warnungen, nr):
        self.prs, self.folie, self.m, self.nr = prs, folie, marke, nr
        self.kopf, self.text = schriften
        self.warnungen = warnungen
        self.dunkel = False

    def c(self, f):
        return self.m.c(f)

    def warn(self, txt):
        self.warnungen.append(f"Folie {self.nr}: {txt}")


# ---------------------------------------------------------------- Text
_MARK = re.compile(r"(\*\*.+?\*\*|==.+?==)")


def _laeufe(t):
    """'a **b** ==c==' → [(a,{}),(b,{fett}),(c,{hervor})]"""
    if isinstance(t, list):  # bereits Läufe [(text, fmt)]
        return t
    teile = []
    for teil in _MARK.split(str(t)):
        if not teil:
            continue
        if teil.startswith("**") and teil.endswith("**"):
            teile.append((teil[2:-2], {"fett": True}))
        elif teil.startswith("==") and teil.endswith("=="):
            teile.append((teil[2:-2], {"hervor": True}))
        else:
            teile.append((teil, {}))
    return teile


def absaetze(spec, punkt=False):
    if spec is None:
        return []
    if isinstance(spec, (str, dict)):
        spec = [spec]
    out = []
    for a in spec:
        if isinstance(a, dict):
            d = dict(a)
            d.setdefault("punkt", punkt)
        else:
            d = {"t": a, "punkt": punkt}
        out.append(d)
    return out


def klartext(spec):
    return " ".join(re.sub(r"\*\*|==", "", str(a.get("t", ""))) for a in absaetze(spec))


def text_hoehe(spec, breite, groesse, zeilenabstand=1.12, nach=3, punkt=False):
    """Grobe Höhe (in) eines Textes: Zeichenbreite ≈ 0,46 × Schriftgröße (am Render von Boehringer Forward Text
    kalibriert: 85 Zeichen à 11 pt auf 5,5 in ≈ 0,42; mit Reserve für Wortumbruch)."""
    h = 0.0
    for a in absaetze(spec, punkt):
        g = a.get("groesse", groesse)
        txt = re.sub(r"\*\*|==", "", str(a.get("t", "")))
        einzug = 0.17 * (1 + a.get("ebene", 0)) if a.get("punkt", punkt) else 0
        b = max(breite - einzug, 0.3)
        zz = max(1, math.ceil(len(txt) * g * 0.46 / 72 / b)) if txt else 1
        h += zz * g * zeilenabstand * 1.2 / 72 + a.get("nach", nach) / 72
    return h


def _aufzaehlung(p, farbe, ebene=0):
    pPr = p._p.get_or_add_pPr()
    marl = 0.17 * (ebene + 1)
    pPr.set("marL", str(E(marl)))
    pPr.set("indent", str(-E(0.17)))
    for tag in ("a:buNone", "a:buChar", "a:buClr", "a:buFont", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bc = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(bc, qn("a:srgbClr")).set("val", farbe)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•" if ebene == 0 else "–")


def _kein_punkt(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


def text(k, tf, spec, groesse=12, farbe="DG", fett=False, ausr="l", anker="o", rand=(0.08, 0.04, 0.08, 0.04),
         nach=3, zeile=1.1, punkt=False, kopf=False, hervor=None, umbruch=True):
    """Text in einen Textrahmen schreiben. hervor = Farbe für ==…== (Standard: AG auf Dunkel, sonst fett)."""
    tf.word_wrap = umbruch
    tf.auto_size = None
    tf.vertical_anchor = ANKER.get(anker, anker)
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [E(x) for x in rand]
    erste = True
    for a in absaetze(spec, punkt):
        p = tf.paragraphs[0] if erste else tf.add_paragraph()
        erste = False
        p.alignment = AUSR.get(a.get("ausr", ausr), PP_ALIGN.LEFT)
        p.line_spacing = a.get("zeile", zeile)
        p.space_after = Pt(a.get("nach", nach))
        p.space_before = Pt(a.get("vor", 0))
        fa = k.c(a.get("farbe", farbe))
        if a.get("punkt"):
            _aufzaehlung(p, a.get("punktfarbe") and k.c(a["punktfarbe"]) or fa, a.get("ebene", 0))
        elif a.get("ohne_punkt"):
            _kein_punkt(p)
        for txt, fmt in _laeufe(a.get("t", "")):
            r = p.add_run()
            r.text = txt
            f = r.font
            f.size = Pt(a.get("groesse", groesse))
            f.bold = fmt.get("fett", a.get("fett", fett))
            f.italic = a.get("kursiv", False)
            f.name = a.get("schrift") or (k.kopf if a.get("kopf", kopf) else k.text)
            col = fa
            if fmt.get("hervor"):
                if hervor:
                    col = k.c(hervor)
                else:
                    f.bold = True
            f.color.rgb = RGBColor.from_string(col)
    return tf


# ---------------------------------------------------------------- Formen
def _sauber(form):
    st = form._element.find(qn("p:style"))
    if st is not None:
        form._element.remove(st)
    spPr = form._element.spPr
    if spPr.find(qn("a:effectLst")) is None:
        spPr.append(etree.Element(qn("a:effectLst")))


def flaeche(k, x, y, w, h, fill=None, linie=None, lw=0.75, strich=False, form=MSO_SHAPE.RECTANGLE, adj=None,
            name=None, t=None, **kw):
    s = k.folie.shapes.add_shape(form, E(x), E(y), E(max(w, 0.01)), E(max(h, 0.01)))
    _sauber(s)
    if name:
        s.name = name
    if adj is not None:
        s.adjustments[0] = adj
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(k.c(fill))
    else:
        s.fill.background()
    if linie:
        s.line.color.rgb = RGBColor.from_string(k.c(linie))
        s.line.width = Pt(lw)
        if strich:
            s.line.dash_style = 4
    else:
        s.line.fill.background()
    if t is not None:
        if fill in ("DG", "darkGreen") and "farbe" not in kw:
            kw["farbe"] = "WH"
            kw.setdefault("hervor", "AG")
        text(k, s.text_frame, t, **kw)
    else:
        s.text_frame.margin_left = s.text_frame.margin_right = 0
    return s


def textfeld(k, x, y, w, h, t, warn_ueberlauf=True, **kw):
    s = k.folie.shapes.add_textbox(E(x), E(y), E(w), E(h))
    text(k, s.text_frame, t, **kw)
    if warn_ueberlauf and kw.get("umbruch", True):
        g = kw.get("groesse", 12)
        rand = kw.get("rand", (0.08, 0.04, 0.08, 0.04))
        bh = text_hoehe(t, w - rand[0] - rand[2], g, nach=kw.get("nach", 3), punkt=kw.get("punkt", False))
        if bh > h + 0.12:
            k.warn(f"Text evtl. zu lang für Feld ({bh:.2f} > {h:.2f} in): „{klartext(t)[:50]}…“")
    return s


def linie(k, x1, y1, x2, y2, farbe="DG", w=1.0, strich=False):
    c = k.folie.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
    st = c._element.find(qn("p:style"))
    if st is not None:
        c._element.remove(st)
    c.line.color.rgb = RGBColor.from_string(k.c(farbe))
    c.line.width = Pt(w)
    if strich:
        c.line.dash_style = 4
    return c


def punkt(k, x, y, d, fill, rand=None):
    return flaeche(k, x, y, d, d, fill=fill, linie=rand, lw=0.75, form=MSO_SHAPE.OVAL)


def raute(k, cx, cy, d, fill, rand="DG"):
    return flaeche(k, cx - d / 2, cy - d / 2, d, d, fill=fill, linie=rand, lw=1, form=MSO_SHAPE.DIAMOND)


def pfeil(k, x, y, w=0.28, h=0.22, fill="DG"):
    return flaeche(k, x, y, w, h, fill=fill, form=MSO_SHAPE.RIGHT_ARROW)


def icon(k, name, x, y, d, variante=None):
    """Brand-Icon (PNG) – Variante passend zum Hintergrund: hell → darkgreen, Dunkelgrün → white/accentgreen."""
    if not name:
        return None
    variante = variante or ("white" if k.dunkel else "darkgreen")
    p = k.m.icon_pfad(name, variante)
    if p is None:
        vorschlag = ", ".join(k.m.icon_suchen(name.split("-")[0])[:5])
        k.warn(f"Icon '{name}' ({variante}) nicht gefunden{' – ähnlich: ' + vorschlag if vorschlag else ''}")
        return None
    return k.folie.shapes.add_picture(str(p), E(x), E(y), E(d), E(d))


# Ampel = Signalfarben des BI-Masters v2.1 (Custom Colors „Signal Green/Yellow/Red“); YR = Gelb–Rot
# (Verzug, Eskalation erwartet) als Punkt mit gelber linker und roter rechter Hälfte.
AMPEL = {"G": "SGRN", "Y": "SYEL", "R": "SRED", "N": "WG", "grün": "SGRN", "gelb": "SYEL", "rot": "SRED", "grau": "WG"}
AMPEL_HALB = {"YR": ("SYEL", "SRED"), "Y/R": ("SYEL", "SRED"), "GY": ("SGRN", "SYEL"), "G/Y": ("SGRN", "SYEL")}


def _winkel(form, *werte):
    """Winkel-Anpassungen einer Form direkt setzen (in 1/60000 Grad); python-pptx skaliert sie sonst falsch."""
    av = form._element.spPr.find(qn("a:prstGeom")).find(qn("a:avLst"))
    for el in list(av):
        av.remove(el)
    for i, v in enumerate(werte, 1):
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", f"adj{i}")
        gd.set("fmla", f"val {int(v * 60000)}")


def ampel(k, x, y, d, wert, rand="WH"):
    halb = AMPEL_HALB.get(str(wert).upper())
    if not halb:
        return punkt(k, x, y, d, AMPEL.get(str(wert), AMPEL.get(str(wert).upper(), "WG")), rand=rand)
    s = punkt(k, x, y, d, halb[0], rand=rand)
    rechts = flaeche(k, x, y, d, d, fill=halb[1], form=MSO_SHAPE.CHORD)
    _winkel(rechts, 270, 90)  # Sehne von oben über rechts nach unten = rechte Hälfte
    return s


# Band unter dem Inhalt: dunkel (Standard, Kernaussage), akzent (Gesamtaussage, z. B. auf der Statusfolie),
# hell (Erläuterung, Fazit). Nur Theme- und Custom Colors des Masters, keine Abstufungen.
BAND_STIL = {"dunkel": ("DG", "WH", "AG"), "akzent": ("AG", "DG", "DG"), "hell": ("LG", "DG", "DG")}


def _band_text(spec):
    if isinstance(spec, dict):
        laeufe = []
        if spec.get("label"):
            laeufe.append((spec["label"] + "  ", {"fett": True, "hervor": True}))
        laeufe += [(t, f) for t, f in _laeufe(spec.get("text", ""))]
        return {"t": laeufe, "ausr": "m"}, False
    return {"t": spec, "ausr": "m"}, True


def band_hoehe(spec, groesse=12):
    """Höhe des Bands: eine Zeile 0,44 in, längere Texte zweizeilig (dann kürzen statt dreizeilig)."""
    if not spec:
        return 0
    t, fett = _band_text(spec)
    bedarf = text_hoehe(klartext(t["t"] if isinstance(t["t"], str) else "".join(x for x, _ in t["t"])),
                        XW - 0.4, groesse, zeilenabstand=1.05, nach=0) + 0.16
    return max(BAND_H, min(bedarf, 0.78))


def band(k, spec, y=None, h=None, groesse=12):
    """Band: "Text" (fett) oder {"label": "…:", "text": "…", "stil": "dunkel|akzent|hell"} (Label fett)."""
    if not spec:
        return None
    t, fett = _band_text(spec)
    stil = spec.get("stil", "dunkel") if isinstance(spec, dict) else "dunkel"
    if stil not in BAND_STIL:
        k.warn(f"Band: unbekannter stil '{stil}' – erlaubt: {', '.join(BAND_STIL)}")
        stil = "dunkel"
    fill, farbe, label = BAND_STIL[stil]
    h = h or band_hoehe(spec, groesse)
    if h >= 0.78 and text_hoehe(klartext(spec.get("text", "") if isinstance(spec, dict) else spec), XW - 0.4,
                                groesse, nach=0) > 0.62:
        k.warn("Band: mehr als zwei Zeilen – Kernaussage kürzen")
    y = Y_ENDE - h if y is None else y
    return flaeche(k, X0, y, XW, h, fill=fill, t=t, groesse=groesse, fett=fett, farbe=farbe, hervor=label,
                   anker="m", nach=0, zeile=1.05, rand=(0.2, 0.02, 0.2, 0.02), name="Kernaussage")


def kopfleiste(k, x, y, w, t, h=0.44, ampelwert=None, groesse=12, fill="DG"):
    """Kopfzeile einer Karte, optional mit Ampel rechts. fill AG (Hervorhebung) mit dunkelgrüner Schrift."""
    s = flaeche(k, x, y, w, h, fill=fill, t=t, groesse=groesse, fett=True, anker="m",
                rand=(0.18, 0, 0.5 if ampelwert else 0.18, 0))
    if ampelwert:
        ampel(k, x + w - 0.4, y + (h - 0.2) / 2, 0.2, ampelwert, rand="WH" if fill == "DG" else "DG")
    return s


def inhaltsunterkante(spec):
    """Unterkante des Inhaltsbereichs je nach Band."""
    return Y_ENDE - band_hoehe(spec.get("band")) - 0.18 if spec.get("band") else Y_ENDE
