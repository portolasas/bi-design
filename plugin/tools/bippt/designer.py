"""Deck aus der Erstellen-Funktion von Copilot (oder einem anderen Designer) in die Marke übertragen: `bippt umwandeln marke`.

Die Gestaltung bleibt, wie sie ist: Lage, Karten, Abläufe, Organigramm, Gantt. Getauscht wird, was die Marke festlegt: Farben nach
ihrer Rolle (nicht die nächste Farbe, sonst wird aus Blau Petrol), Schriften, Icons, Hintergrund, Logo und Fußzeile aus dem BI-Master.
Entscheidung des Nutzers (30.10.2026): Das Deck muss nicht in die Layouts des Masters passen, wenn die Elemente markenkonform sind.

Icons: Ohne Zuordnung werden sie nur umgefärbt und bleiben fremde Icons. Der Befehl listet sie je Folie mit dem Text daneben;
die Zuordnung {"<Folie>": {"<id>": "<Brand-Icon>"}} in einer JSON-Datei (--icons) tauscht sie gegen Brand-Icons.
"""
import colorsys
import io
import json
import re
from pathlib import Path

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu

from . import bibliothek, konfig, vorlage
from .marke import Marke

ZOLL = 914400
DG, AG, WH = "08312A", "00E47C", "FFFFFF"
SEITE, FEIN, WARM = "F6F5F3", "EFEEEB", "E5E3DE"  # Light Gray für die Seite, Gray 3 für Flächen darauf, Warm Gray
GRAU_TEXT, GRAU_LINIE, GRAU_FLAECHE = "5C5B59", "B7B6B2", "5C5B59"  # Gray 7, Gray 5
SIGNAL = {"gruen": ("1EAD1B", DG), "gelb": ("FF8413", "E18600"), "rot": ("F62B2B", "86251B")}  # Fläche, Text auf Hell
UNTEN = 6.62  # Unterkante des Inhalts (in): darunter liegen Logo und Fußzeile des Masters
FUSS_OBEN = 6.85  # ab hier beginnt die Fußzeile eines Designer-Decks
ICON_MAX = 1.2  # in; größere Bilder sind keine Icons
MIN_SZ = 900  # Hundertstel pt; Mindestgröße der Marke im Inhalt, Designer-Decks gehen bis unter 8 pt
ICON_KLEIN = 0.135  # in; kleiner sind es Häkchen und Punkte vor Listen, sie werden nur umgefärbt
EIGENE_LAGE = (qn("p:sp"), qn("p:pic"), qn("p:cxnSp"), qn("p:graphicFrame"))


# ----------------------------------------------------------------- Farben nach Rolle
def _rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _hsl(h):
    r, g, b = _rgb(h)
    t, l, s = colorsys.rgb_to_hls(r, g, b)
    return t * 360, s, l


def _lum(h):
    def k(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _rgb(h)
    return 0.2126 * k(r) + 0.7152 * k(g) + 0.0722 * k(b)


def _kontrast(a, b):
    x, y = sorted((_lum(a), _lum(b)), reverse=True)
    return (x + 0.05) / (y + 0.05)


def _familie(h):
    """Farbton in 30°-Stufen, nur für bunte Farben; None für Grau, sehr Helles und sehr Dunkles."""
    t, s, l = _hsl(h)
    return None if s < 0.3 or l > 0.9 or l < 0.08 else int(t // 30)


class Palette:
    """Regeln für ein Deck: Der häufigste Farbton ist die Hausfarbe des Designers und nie eine Signalfarbe."""

    def __init__(self, farben, seite):
        zahl = {}
        for f in farben:
            fam = _familie(f)
            if fam is not None:
                zahl[fam] = zahl.get(fam, 0) + 1
        self.haus = max(zahl, key=zahl.get) if zahl else None
        self.seite = seite

    def signal(self, h):
        fam = _familie(h)
        if fam is None or (self.haus is not None and min((fam - self.haus) % 12, (self.haus - fam) % 12) <= 1):
            return None
        t = _hsl(h)[0]
        return "gruen" if 75 <= t < 170 else "gelb" if 25 <= t < 70 else "rot" if t < 25 or t >= 335 else None

    def flaeche(self, h):
        s = self.signal(h)
        if s:
            return SIGNAL[s][0]
        _, sat, l = _hsl(h)
        if h == WH or l >= 0.985:
            return WH
        if l >= 0.9:
            return SEITE if h == self.seite else FEIN
        if l >= 0.7:
            return WARM
        if sat < 0.3:
            return GRAU_FLAECHE if _lum(h) > 0.02 else DG
        return DG if _lum(h) < 0.09 else AG

    def linie(self, h, grund):
        _, sat, l = _hsl(h)
        if _lum(grund) < 0.25:
            return AG if sat >= 0.3 and l < 0.7 else WARM
        if l >= 0.7:
            return WARM
        return GRAU_LINIE if sat < 0.3 else DG

    def text(self, h, grund):
        s = self.signal(h)
        if grund == AG:
            return DG
        t_sat, t_l = _hsl(h)[1:]
        if _lum(grund) < 0.25:
            neu = WH if t_l > 0.9 or t_sat < 0.3 else WARM if t_l >= 0.7 else AG
        elif s:
            neu = SIGNAL[s][1]
        else:
            neu = GRAU_TEXT if t_sat < 0.35 and 0.3 < t_l < 0.75 else DG
        return neu if _kontrast(neu, grund) >= 3 else max((DG, WH), key=lambda c: _kontrast(c, grund))


# ----------------------------------------------------------------- Lesen
def _box(el):
    """Lage (in) einer Form auf der Folie; None ohne eigene Lage (Platzhalter erben sie)."""
    xf = el.find(".//" + qn("a:xfrm")) if el.tag != qn("p:graphicFrame") else el.find(qn("p:xfrm"))
    if xf is None or xf.find(qn("a:off")) is None:
        return None
    o, e = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    x, y, w, h = (int(v) / ZOLL for v in (o.get("x"), o.get("y"), e.get("cx"), e.get("cy")))
    return x, y, w, h


def _text(el):
    return " ".join("".join(t.text or "" for t in p.iter(qn("a:t"))) for p in el.iter(qn("a:p"))).strip()


def _groesse(el):
    gr = [int(r.get("sz")) / 100 for r in el.iter(qn("a:rPr")) if r.get("sz") and r.getparent().find(qn("a:t")) is not None]
    return max(gr) if gr else 0


def _formen(folie):
    """Formen der Folie in Zeichenreihenfolge; Gruppen aufgelöst, sofern sie nicht skaliert sind."""
    def gehen(baum):
        for el in baum:
            if el.tag == qn("p:grpSp"):
                yield from gehen(el)
            elif el.tag in EIGENE_LAGE:
                yield el
    return list(gehen(folie.shapes._spTree))


def _hintergrund(folie):
    bg = folie._element.find(".//" + qn("p:bg"))
    if bg is None:
        return None, False
    f = bg.find(".//" + qn("a:srgbClr"))
    return (f.get("val").upper() if f is not None else None), bg.find(".//" + qn("a:blip")) is not None


def _deck_farben(prs):
    farben = []
    for s in prs.slides:
        farben += [el.get("val", "").upper() for el in s._element.iter(qn("a:srgbClr"))]
    return farben


def _kopfschrift(prs):
    """Schrift des größten Textes im Deck: Sie wird zur Kopfschrift der Marke, alle anderen zur Textschrift."""
    beste, gr = None, 0
    for s in prs.slides:
        for r in s._element.iter(qn("a:rPr")):
            lat = r.find(qn("a:latin"))
            if lat is not None and int(r.get("sz") or 0) > gr:
                beste, gr = lat.get("typeface"), int(r.get("sz"))
    return beste


# ----------------------------------------------------------------- Fußzeilen und Titel
def _gleicher_anfang(a, b):
    n = 0
    for x, y in zip(a.lower(), b.lower()):
        if x != y:
            break
        n += 1
    return n


def _fusstexte(prs, decktitel):
    """Ids der Fußtexte je Folie: Seitenzahl, Text am unteren Rand, der auf anderen Folien oder im Decktitel wiederkehrt.
    Inhalt am unteren Rand („Gesamtlaufzeit: …“) kehrt nicht wieder und bleibt."""
    unten = {}
    for nr, s in enumerate(prs.slides, 1):
        unten[nr] = [(el, _text(el)) for el in _formen(s) if _box(el) and _box(el)[1] >= FUSS_OBEN and _text(el)]
    weg = {}
    for nr, liste in unten.items():
        for el, t in liste:
            andere = [t2 for n2, l2 in unten.items() if n2 != nr for _, t2 in l2]
            if t == str(nr) or any(_gleicher_anfang(t, x) >= 20 for x in andere + [decktitel]):
                weg.setdefault(nr, []).append(el)
    return weg


def _aufraeumen(folie, nr, fuss):
    """Fremdes Beiwerk weg: Vollbilder, Fußbänder, Fußtexte, Hintergrund. Rückgabe: Zahl der entfernten Formen."""
    weg = list(fuss.get(nr, []))
    for el in _formen(folie):
        b = _box(el)
        if b is None or el in weg:
            continue
        if el.tag == qn("p:pic") and b[2] >= 13 and b[3] >= 7:
            weg.append(el)
        elif el.tag == qn("p:sp") and not _text(el) and b[2] >= 12 and b[1] >= FUSS_OBEN:
            weg.append(el)
    for el in weg:
        el.getparent().remove(el)
    return len(weg)


def _titel_setzen(folie):
    """Größter Text oben wird echter Titel: Er behält Lage und Format, Navigation und Prüfung kennen ihn dann."""
    oben = [el for el in _formen(folie) if el.tag == qn("p:sp") and _box(el) and _box(el)[1] < 1.1 and _text(el)]
    if not oben:
        return None
    el = max(oben, key=_groesse)
    nv = el.find(qn("p:nvSpPr"))
    c = nv.find(qn("p:cNvSpPr"))
    if c is not None and c.get("txBox"):
        del c.attrib["txBox"]
    pr = nv.find(qn("p:nvPr"))
    if pr.find(qn("p:ph")) is None:
        ph = pr.makeelement(qn("p:ph"), {"type": "title"})
        pr.insert(0, ph)
    return _text(el)


def _einpassen(folie):
    """Reicht der Inhalt in Logo und Fußzeile des Masters, staucht bippt ihn in der Höhe. Flächen werden flacher, Text und
    Icons behalten ihre Größe und rücken nur mit. Rückgabe: Faktor oder None."""
    formen = [(el, _box(el)) for el in _formen(folie) if _box(el)]
    if not formen:
        return None
    tief = max(b[1] + b[3] for _, b in formen)
    if tief <= UNTEN:
        return None
    start = min((b[1] for _, b in formen if b[1] >= 1.25), default=1.25)
    f = (UNTEN - start) / (tief - start)
    for el, (x, y, w, h) in formen:
        if y + h <= start:
            continue
        xf = el.find(".//" + qn("a:xfrm")) if el.tag != qn("p:graphicFrame") else el.find(qn("p:xfrm"))
        neu = start + (max(y, start) - start) * f
        xf.find(qn("a:off")).set("y", str(int(neu * ZOLL)))
        if el.tag == qn("p:cxnSp") or (el.tag == qn("p:sp") and not _text(el)):
            xf.find(qn("a:ext")).set("cy", str(int(h * f * ZOLL)))
    return f


# ----------------------------------------------------------------- Formen umstellen
def _farbe_setzen(kette, neu):
    f = kette.find(qn("a:solidFill"))
    if f is None:
        return
    for k in list(f):
        f.remove(k)
    x = f.makeelement(qn("a:srgbClr"), {"val": neu})
    f.append(x)


def _eigene_farbe(kette):
    s = kette.find(qn("a:solidFill")) if kette is not None else None
    x = s.find(qn("a:srgbClr")) if s is not None else None
    return x.get("val").upper() if x is not None else None


def _behaelter(gemalt, b):
    """Oberste gemalte Fläche unter dem Mittelpunkt von b; None auf der Seite."""
    cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
    for f in reversed(gemalt):
        if f[0] <= cx <= f[0] + f[2] and f[1] <= cy <= f[1] + f[3]:
            return f
    return None


def _icon_kante(b, gemalt):
    """Kantenlänge des Brand-Icons. Die Brand-Icons sind feine Linien und brauchen mehr Platz als die vollen Icons des
    Designers: In einem kleinen Kreis füllen sie ihn zu 70 %, sonst wachsen sie um ein Drittel."""
    kante = max(b[2], b[3])
    f = _behaelter(gemalt, b)
    if f and max(f[2], f[3]) < 0.7:
        return 0.7 * min(f[2], f[3])
    return kante * 1.33


def _mittig(el, b, gemalt):
    """Der Designer misst Textfelder auf seine Schrift. Sitzt ein Feld ohne Umbruch mittig in seiner Karte, soll es dort auch mit
    der schmaleren Markenschrift bleiben."""
    bp = el.find(".//" + qn("a:bodyPr"))
    f = _behaelter(gemalt, b)
    if bp is None or bp.get("wrap") != "none" or f is None or b[2] > f[2] * 0.95:
        return
    if abs((b[0] + b[2] / 2) - (f[0] + f[2] / 2)) < 0.04:
        for p in el.iter(qn("a:p")):
            pr = p.find(qn("a:pPr"))
            if pr is None:
                pr = p.makeelement(qn("a:pPr"), {})
                p.insert(0, pr)
            pr.set("algn", "ctr")


def _unter(gemalt, b, standard):
    """Farbe der obersten gemalten Fläche unter dem Mittelpunkt von b."""
    cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
    for x, y, w, h, f in reversed(gemalt):
        if x <= cx <= x + w and y <= cy <= y + h:
            return f
    return standard


def _schriften(el, kopf):
    for tag in ("a:rPr", "a:endParaRPr", "a:defRPr"):
        for r in el.iter(qn(tag)):
            for lat in r.findall(qn("a:latin")):
                lat.set("typeface", "+mj-lt" if lat.get("typeface") == kopf else "+mn-lt")
            for art in ("a:ea", "a:cs", "a:sym"):
                for x in r.findall(qn(art)):
                    r.remove(x)
            if r.get("sz") and int(r.get("sz")) < MIN_SZ:
                r.set("sz", str(MIN_SZ))


def _textfarben(el, pal, grund):
    for tag in ("a:rPr", "a:endParaRPr"):
        for r in el.iter(qn(tag)):
            alt = _eigene_farbe(r)
            if alt:
                _farbe_setzen(r, pal.text(alt, grund))
            elif r.getparent().find(qn("a:t")) is not None and r.find(qn("a:hlinkClick")) is None:
                f = r.makeelement(qn("a:solidFill"), {})
                f.append(f.makeelement(qn("a:srgbClr"), {"val": pal.text(DG, grund)}))
                r.insert(1 if r.find(qn("a:ln")) is not None else 0, f)
    for bu in el.iter(qn("a:buClr")):
        x = bu.find(qn("a:srgbClr"))
        if x is not None:
            x.set("val", pal.text(x.get("val").upper(), grund))


def _bildfarben(blob):
    """(Bild, Zahl der Farbstufen, häufigste deckende Farbe, Anteil deckender Pixel); Fotos erkennt man an vielen Stufen."""
    from PIL import Image
    im = Image.open(io.BytesIO(blob)).convert("RGBA")
    roh = im.tobytes()
    zahl = {}
    for i in range(0, len(roh), 4):
        if roh[i + 3] > 200:
            k = roh[i:i + 3]
            zahl[k] = zahl.get(k, 0) + 1
    stufen = {(k[0] // 16, k[1] // 16, k[2] // 16) for k in zahl}
    haeufig = max(zahl, key=zahl.get).hex().upper() if zahl else None
    return im, len(stufen), haeufig, sum(zahl.values()) / (len(roh) / 4)


def _farben_tauschen(im, regel):
    """Grafik (keine Fotos): jede Farbe nach ihrer Rolle, Transparenz bleibt."""
    from PIL import Image
    roh, aus, cache = im.tobytes(), bytearray(), {}
    for i in range(0, len(roh), 4):
        k = roh[i:i + 3]
        if k not in cache:
            cache[k] = bytes.fromhex(regel(k.hex().upper()))
        aus += cache[k] + roh[i + 3:i + 4]
    neu = io.BytesIO()
    Image.frombytes("RGBA", im.size, bytes(aus)).save(neu, "PNG")
    neu.seek(0)
    return neu


def _umgefaerbt(blob, farbe):
    from PIL import Image
    im = Image.open(io.BytesIO(blob)).convert("RGBA")
    flaeche = Image.new("RGBA", im.size, tuple(int(farbe[i:i + 2], 16) for i in (0, 2, 4)) + (255,))
    flaeche.putalpha(im.getchannel("A"))
    aus = io.BytesIO()
    flaeche.save(aus, "PNG")
    aus.seek(0)
    return aus


def _bild_einsetzen(folie, el, datei):
    """Bild der Form tauschen; die SVG-Fassung des Designers fällt weg, sonst zeigt PowerPoint weiter das fremde Icon."""
    _, rid = folie.part.get_or_add_image_part(datei)
    blip = el.find(".//" + qn("a:blip"))
    blip.set(qn("r:embed"), rid)
    for ext in blip.findall(qn("a:extLst")):
        blip.remove(ext)
    for src in el.findall(".//" + qn("a:srcRect")):
        src.getparent().remove(src)


def _bild(folie, el, b, pal, grund, gemalt, nr, m, zuordnung, liste, naechster):
    """Icons tauschen oder umfärben, Schmuck und flache Grafik in Markenfarben; nur Fotos bleiben und werden gemeldet."""
    rid = el.find(".//" + qn("a:blip")).get(qn("r:embed"))
    blob = folie.part.related_part(rid).blob
    sid = el.find(qn("p:nvPicPr")).find(qn("p:cNvPr")).get("id")
    w, h = b[2], b[3]
    try:
        im, stufen, farbe, deckend = _bildfarben(blob)
    except (OSError, ValueError):  # z. B. EMF: Pillow kann es nicht lesen
        liste["bilder"].append((nr, sid))
        return
    if stufen > 40 or max(im.size) > 1600 or farbe is None:
        liste["bilder"].append((nr, sid))
        return
    schmal = min(w, h) < 0.1 or max(w, h) / max(min(w, h), 0.01) > 2
    if schmal or w > ICON_MAX or h > ICON_MAX or stufen > 4:  # Linien, Zeilenraster, Flächen aus dem Browser
        regel = (lambda x: pal.linie(x, grund)) if schmal and min(w, h) < 0.1 else pal.flaeche
        _bild_einsetzen(folie, el, _farben_tauschen(im, regel))
        return
    name = (zuordnung.get(str(nr)) or {}).get(sid)
    if name and max(w, h) >= ICON_KLEIN:
        datei = m.icon_pfad(name, "white" if _lum(grund) < 0.25 else "darkgreen")
        if datei is None:
            liste["fehlt"].append((nr, sid, name))
        else:
            _bild_einsetzen(folie, el, str(datei))
            kante = _icon_kante(b, gemalt)
            xf = el.find(".//" + qn("a:xfrm"))
            xf.find(qn("a:off")).set("x", str(int((b[0] + w / 2 - kante / 2) * ZOLL)))
            xf.find(qn("a:off")).set("y", str(int((b[1] + h / 2 - kante / 2) * ZOLL)))
            xf.find(qn("a:ext")).set("cx", str(int(kante * ZOLL)))
            xf.find(qn("a:ext")).set("cy", str(int(kante * ZOLL)))
            liste["getauscht"] += 1
            return
    s = pal.signal(farbe)
    ziel = SIGNAL[s][0] if s and _kontrast(SIGNAL[s][0], grund) >= 2 else max((DG, WH), key=lambda c: _kontrast(c, grund))
    _bild_einsetzen(folie, el, _umgefaerbt(blob, ziel))
    if max(w, h) >= ICON_KLEIN:  # Häkchen und Punkte vor Listen sind Zeichen, kein Icon
        liste["offen"].append((nr, sid, naechster))


def _naechster_text(b, texte):
    if not texte:
        return ""
    cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
    def abstand(t):
        x, y, w, h = t[1]
        dx = max(x - cx, 0, cx - x - w)
        dy = max(y - cy, 0, cy - y - h)
        return dx + dy
    return min(texte, key=abstand)[2][:40]


def _umstellen(folie, nr, pal, seite, kopf, m, zuordnung, liste):
    """Farben, Schriften, Ecken, Schatten und Bilder einer Folie, in Zeichenreihenfolge: Was unten liegt, bestimmt die Farbe darüber."""
    gemalt = []
    formen = _formen(folie)
    texte = [(el, _box(el), _text(el)) for el in formen if _box(el) and _text(el)]
    for el in formen:
        b = _box(el)
        sp = el.find(qn("p:spPr"))
        if el.tag == qn("p:pic"):
            if b:
                _bild(folie, el, b, pal, _unter(gemalt, b, seite), gemalt, nr, m, zuordnung, liste, _naechster_text(b, texte))
            continue
        grund_form = None
        if sp is not None:
            for fx in sp.findall(qn("a:effectLst")):
                for k in list(fx):
                    fx.remove(k)
            geo = sp.find(qn("a:prstGeom"))
            if geo is not None and geo.get("prst") == "roundRect":  # die Marke kennt keine runden Ecken an Karten
                geo.set("prst", "rect")
                av = geo.find(qn("a:avLst"))
                if av is not None:
                    for k in list(av):
                        av.remove(k)
            verlauf = sp.find(qn("a:gradFill"))
            if verlauf is not None:
                erster = verlauf.find(".//" + qn("a:srgbClr"))
                voll = sp.makeelement(qn("a:solidFill"), {})
                voll.append(voll.makeelement(qn("a:srgbClr"), {"val": erster.get("val") if erster is not None else DG}))
                sp.replace(verlauf, voll)
            alt = _eigene_farbe(sp)
            if alt and b:
                unter = _unter(gemalt, b, seite)
                duenn = min(b[2], b[3]) < 0.08
                neu = pal.linie(alt, unter) if duenn else pal.flaeche(alt)
                if neu == WH and unter == WH and not duenn and b[2] * b[3] > 1:
                    neu = FEIN  # weiße Karte auf weißer Seite wäre unsichtbar
                _farbe_setzen(sp, neu)
                grund_form = neu
                if not duenn:
                    gemalt.append(b + (neu,))
            ln = sp.find(qn("a:ln"))
            alt_l = _eigene_farbe(ln)
            if alt_l and b:
                _farbe_setzen(ln, pal.linie(alt_l, grund_form or _unter(gemalt, b, seite)))
        if el.tag == qn("p:graphicFrame"):
            _tabelle(el, pal, seite)
        elif b or _text(el):
            grund = grund_form or (_unter(gemalt, b, seite) if b else seite)
            _textfarben(el, pal, grund)
            if b and not grund_form:
                _mittig(el, b, gemalt)
            bp = el.find(".//" + qn("a:bodyPr"))
            if b and b[2] < 0.3 and bp is not None and len(_text(el)) <= 4:  # Ziffer in Kreis oder Seitenangabe
                bp.set("wrap", "none")
        _schriften(el, kopf)
    _diagramme(folie, pal)


def _tabelle(el, pal, seite):
    for zelle in el.iter(qn("a:tc")):
        pr = zelle.find(qn("a:tcPr"))
        alt = _eigene_farbe(pr)
        grund = seite
        if alt:
            grund = pal.flaeche(alt)
            _farbe_setzen(pr, grund)
        _textfarben(zelle, pal, grund)


def _diagramme(folie, pal):
    for rel in list(folie.part.rels.values()):
        if rel.reltype == RT.CHART:
            teil = rel.target_part
            for x in teil._element.iter(qn("a:srgbClr")):
                x.set("val", pal.flaeche(x.get("val").upper()))
            for lat in teil._element.iter(qn("a:latin")):
                lat.set("typeface", "+mn-lt")


def _reste_entfernen(folie):
    """Bilder ohne Verweis (entfernte Vollbilder, SVG-Fassungen) nicht mitspeichern."""
    xml = folie.part.blob.decode("utf-8") if isinstance(folie.part.blob, bytes) else folie.part.blob
    for rid, rel in list(folie.part.rels.items()):
        if rel.reltype == RT.IMAGE and f'"{rid}"' not in xml:
            folie.part.rels._rels.pop(rid)


# ----------------------------------------------------------------- Titelfolie
def _ist_titelfolie(folie, nr):
    if nr != 1:
        return False
    bg, bild = _hintergrund(folie)
    voll = any(el.tag == qn("p:pic") and _box(el) and _box(el)[2] >= 13 and _box(el)[3] >= 7 for el in _formen(folie))
    return voll or bild or (bg is not None and _lum(bg) < 0.25) or "title" in folie.slide_layout.name.lower()


def _titelfolie(prs, quelle, vermerk):
    """Titelfolie des Masters mit den Texten des Designers; kein Text geht verloren: größter = Titel, der Rest als Untertitel."""
    texte = [(el, _box(el), _text(el)) for el in _formen(quelle) if _box(el) and _text(el)]
    texte.sort(key=lambda t: t[1][1])
    if not texte:
        return prs.slides.add_slide(vorlage.layout(prs, "titel")), ""
    gross = max(texte, key=lambda t: _groesse(t[0]))
    rest = [t for t in texte if t is not gross]
    unten = [t[2] for t in rest if t[1][1] >= FUSS_OBEN]
    mitte = [t for t in rest if t[1][1] < FUSS_OBEN]
    haupt = [t[2] for t in mitte if t[1][1] > gross[1][1]]
    darueber = [t[2] for t in mitte if t[1][1] <= gross[1][1]]
    sl = prs.slides.add_slide(vorlage.layout(prs, "titel"))
    sl.shapes.title.text = gross[2]
    ph = vorlage.platzhalter(sl, idx=1)
    zeilen = [(t, 0) for t in haupt] + [(t, 1) for t in darueber] + ([(" | ".join(unten), 1)] if unten else [])
    if ph is not None and zeilen:
        tf = ph.text_frame
        for i, (t, stufe) in enumerate(zeilen):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = t
            p.level = stufe
    vorlage.leere_platzhalter_entfernen(sl)
    vorlage.fusszeile(sl, vermerk, nummer=False)
    return sl, gross[2]


# ----------------------------------------------------------------- Ablauf
def marke(pfad, ziel=None, icons=None, fusszeile=None):
    """Überträgt das Deck pfad in die Marke. Rückgabe: (Zieldatei, Ergebnisse)."""
    quelle = vorlage.oeffnen(pfad)
    vl = konfig.vorlage_pfad()
    if vl is None:
        raise SystemExit("Keine BI-Vorlage gefunden")
    prs = vorlage.oeffnen(vl)
    vorlage.alle_folien_entfernen(prs)
    m = Marke()
    m.aus_vorlage(prs)
    zuordnung = json.loads(Path(icons).read_text(encoding="utf-8")) if icons else {}
    seiten = [_hintergrund(s)[0] for s in quelle.slides]
    haeufig = max(set(x for x in seiten if x) or {WH}, key=seiten.count)
    pal = Palette(_deck_farben(quelle), haeufig)
    kopf = _kopfschrift(quelle)
    vermerk = konfig.laden().get("vertraulichkeit", "ONLY FOR INTERNAL USE")
    trenner = "     |     "

    decktitel, titelnr = "", None
    if len(quelle.slides) and _ist_titelfolie(quelle.slides[0], 1):
        titelnr = 1
        kand = [el for el in _formen(quelle.slides[0]) if _text(el)]
        decktitel = _text(max(kand, key=_groesse)) if kand else ""
    fuss = _fusstexte(quelle, decktitel)
    liste = {"getauscht": 0, "offen": [], "fehlt": [], "bilder": []}
    gestaucht, entfernt = [], 0
    fuss_text = trenner.join(x for x in (fusszeile or decktitel, vermerk) if x)
    for nr, q in enumerate(quelle.slides, 1):
        if nr == titelnr:
            _titelfolie(prs, q, vermerk)
            continue
        bg, _ = _hintergrund(q)
        dunkel = bg is not None and _lum(bg) < 0.25
        sl, _ = bibliothek.kopieren(prs, quelle, nr, layout="One content dark" if dunkel else "Only title")
        fuss_hier = {nr: []}
        namen = {el.find(".//" + qn("p:cNvPr")).get("id") for el in fuss.get(nr, [])}
        fuss_hier[nr] = [el for el in _formen(sl) if el.find(".//" + qn("p:cNvPr")).get("id") in namen]
        entfernt += _aufraeumen(sl, nr, fuss_hier)
        csld = sl._element.find(qn("p:cSld"))
        for alt in csld.findall(qn("p:bg")):
            csld.remove(alt)
        seite = DG if dunkel else WH
        if not dunkel and bg is not None and bg != WH:
            seite = SEITE
            sl.background.fill.solid()
            from pptx.dml.color import RGBColor
            sl.background.fill.fore_color.rgb = RGBColor.from_string(SEITE)
        _titel_setzen(sl)
        f = _einpassen(sl)
        if f:
            gestaucht.append(f"{nr} ({f:.0%})")
        _umstellen(sl, nr, pal, seite, kopf, m, zuordnung, liste)
        vorlage.fusszeile(sl, fuss_text)
        _reste_entfernen(sl)
    ziel = Path(ziel) if ziel else Path(pfad).with_name(Path(pfad).stem + "_BI.pptx")
    prs.save(str(ziel))
    print(f"MARKE: {ziel} – {len(prs.slides)} Folien, Gestaltung übernommen")
    print(f"  Fremdes Beiwerk entfernt: {entfernt} Formen · in der Höhe eingepasst: {', '.join(gestaucht) or 'keine Folie'}")
    print(f"  Icons: {liste['getauscht']} durch Brand-Icons ersetzt, {len(liste['offen'])} nur umgefärbt")
    for nr, sid, name in liste["fehlt"]:
        print(f"  WARNUNG Folie {nr} #{sid}: Brand-Icon „{name}“ nicht gefunden (bippt icons <Begriff>)")
    if liste["offen"]:
        print("  OFFENE ICONS – Folie #id neben Text; Zuordnung in eine JSON-Datei, dann erneut mit --icons:")
        for nr, sid, text in liste["offen"]:
            print(f"    {nr} #{sid} neben „{text}“")
    for nr, sid in liste["bilder"]:
        print(f"  HINWEIS Folie {nr} #{sid}: Bild bleibt wie es ist (Foto oder Grafik), Farben prüfen")
    print("ERGEBNIS: " + ("OK" if not liste["offen"] and not liste["fehlt"] else
                          f"{len(liste['offen']) + len(liste['fehlt'])} Icon(s) offen"))
    return ziel, liste
