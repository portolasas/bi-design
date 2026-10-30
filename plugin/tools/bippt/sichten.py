"""Prüfung am gerenderten Deck, die ohne Augen auskommt: `bippt sichten` rendert und misst auf PDF und Bild, was `bauen` und
`pruefen` nur schätzen. Gedacht für Umgebungen, in denen der Agent seine Folien nicht ansehen kann (Copilot: LibreOffice ja,
Blick unklar); sie ersetzt dort den Kontaktbogen, so weit das geht.

Gemessen wird je Folie:
- Überlappung: Textzeilen, die einander überdecken.
- Überstand: Textzeile, die über den Rand der Folie oder einer gefüllten Form ragt.
- Leere: größte Fläche der Inhaltszone ohne jedes Detail (einfarbige Zellen, auch einfarbig gefüllte Karten).
Rendert LibreOffice ohne die Boehringer-Schriften, weicht der Umbruch ab: Überlappung und Überstand sind dann nur Hinweise; die
Leere hängt kaum an der Schrift und zählt weiter.
"""
from pathlib import Path

from pptx import Presentation

from .rendern import rendern

ZOLL = 72  # PDF-Punkte je Zoll; Folie 13,33 × 7,5 in
ZONE = (0.5, 1.3, 12.83, 6.35)  # Inhaltszone (in): unter dem Titel, über Logo und Fußzeile
ZELLE = 0.25  # Kantenlänge einer Prüfzelle (in)
KONTRAST = 8  # Graustufen: Zelle mit weniger Spanne gilt als leer
LEER_FLAECHE = 0.40  # am Render der 184 Folien der Slide Library kalibriert: nur Rahmenfolien (Disclaimer, Outro) liegen darüber
TOLERANZ = 0.04  # in; Überstand darunter zählt nicht
RAHMEN_LAYOUTS = ("title slide", "chapter", "divider", "agenda", "section", "thank", "disclaimer", "outro", "contact", "one content dark")  # hier ist Leere gewollt


def _zeilen(seite):
    """Textzeilen der Seite: (Text, Schriftgröße, Box in Zoll, Tintenkern in Zoll). Der Kern lässt Ober- und Unterlänge weg."""
    zeilen = []
    for block in seite.get_text("dict")["blocks"]:
        for z in block.get("lines", []):
            text = "".join(sp["text"] for sp in z["spans"]).strip()
            if not text:
                continue
            x0, y0, x1, y1 = (v / ZOLL for v in z["bbox"])
            rand = (y1 - y0) * 0.25
            zeilen.append({"t": text, "g": max(sp["size"] for sp in z["spans"]), "box": (x0, y0, x1, y1),
                           "kern": (x0, y0 + rand, x1, y1 - rand)})
    return zeilen


def _auszug(t, n=28):
    return t if len(t) <= n else t[:n - 1] + "…"


def _ueberlappung(zeilen):
    funde = []
    for i, a in enumerate(zeilen):
        for b in zeilen[i + 1:]:
            if a["t"] == b["t"] or min(len(a["t"]), len(b["t"])) < 5:  # Nummern und Marken sitzen oft bewusst am Text
                continue
            breite = min(a["kern"][2], b["kern"][2]) - max(a["kern"][0], b["kern"][0])
            hoehe = min(a["kern"][3], b["kern"][3]) - max(a["kern"][1], b["kern"][1])
            if breite > 0.1 and hoehe > 0.01:
                funde.append(f"Überlappung: „{_auszug(a['t'])}“ und „{_auszug(b['t'])}“")
    return funde


def _formen(seite):
    """Gefüllte Flächen der Seite als Rechtecke (in), ohne Vollflächen und Kleinteile."""
    flaechen = []
    for d in seite.get_drawings():
        if d.get("fill") is None:
            continue
        r = d["rect"]
        x0, y0, x1, y1 = r.x0 / ZOLL, r.y0 / ZOLL, r.x1 / ZOLL, r.y1 / ZOLL
        if x1 - x0 < 0.6 or y1 - y0 < 0.25 or (x1 - x0 > 13 and y1 - y0 > 7):
            continue
        flaechen.append((x0, y0, x1, y1))
    return flaechen


def _ueberstand(seite, zeilen):
    funde = []
    breite, hoehe = seite.rect.width / ZOLL, seite.rect.height / ZOLL
    flaechen = _formen(seite)
    for z in zeilen:
        x0, y0, x1, y1 = z["kern"]
        if len(z["t"].split()) < 3:  # Beschriftungen an Balken und Pfeilen ragen bewusst über ihre Form
            continue
        if x0 < -TOLERANZ or y0 < -TOLERANZ or x1 > breite + TOLERANZ or y1 > hoehe + TOLERANZ:
            funde.append(f"Überstand: „{_auszug(z['t'])}“ ragt über den Folienrand")
            continue
        for fx0, fy0, fx1, fy1 in flaechen:
            innen = x0 < fx1 and x1 > fx0 and y0 < fy1 and y1 > fy0
            ganz_innen = x0 >= fx0 - TOLERANZ and x1 <= fx1 + TOLERANZ and y0 >= fy0 - TOLERANZ and y1 <= fy1 + TOLERANZ
            if innen and not ganz_innen:
                funde.append(f"Überstand: „{_auszug(z['t'])}“ ragt über den Rand einer Form")
                break
    return funde


def _groesstes_rechteck(zeilen):
    """Größtes Rechteck aus lauter wahren Zellen (Zellen zählen), Histogramm je Reihe."""
    if not zeilen:
        return 0
    hoehen, best = [0] * len(zeilen[0]), 0
    for reihe in zeilen:
        hoehen = [h + 1 if leer else 0 for h, leer in zip(hoehen, reihe)]
        stapel = []
        for i, h in enumerate(hoehen + [0]):
            start = i
            while stapel and stapel[-1][1] >= h:
                start, oben = stapel.pop()
                best = max(best, oben * (i - start))
            stapel.append((start, h))
    return best


def _leere(seite, pymupdf):
    """(Anteil leerer Zellen, größte leere Fläche als Anteil) der Inhaltszone; eine Zelle ist leer, wenn sie einfarbig ist."""
    bild = seite.get_pixmap(dpi=36, colorspace=pymupdf.csGRAY, alpha=False)
    p = 36 * ZELLE
    x0, y0, x1, y1 = (v * 36 for v in ZONE)
    raster = []
    y = y0
    while y + p <= y1 + 0.01:
        reihe, x = [], x0
        while x + p <= x1 + 0.01:
            werte = [bild.samples[bild.stride * j + i] for j in range(int(y), int(y + p)) for i in range(int(x), int(x + p))]
            reihe.append(max(werte) - min(werte) < KONTRAST)
            x += p
        raster.append(reihe)
        y += p
    zellen = sum(map(len, raster)) or 1
    return sum(map(sum, raster)) / zellen, _groesstes_rechteck(raster) / zellen


def messen(pdf, layouts=None):
    """Befunde je Seite eines gerenderten Decks: {"nr", "leer", "flaeche", "text": [...], "leere": str|None}."""
    try:
        import pymupdf
    except ImportError:
        import fitz as pymupdf
    doc = pymupdf.open(str(pdf))
    ergebnis = []
    for nr, seite in enumerate(doc, 1):
        zeilen = _zeilen(seite)
        layout = (layouts[nr - 1] if layouts and nr <= len(layouts) else "").lower()
        rahmen = any(n in layout for n in RAHMEN_LAYOUTS)
        leer, flaeche = (0.0, 0.0) if rahmen else _leere(seite, pymupdf)
        befund = None
        if flaeche >= LEER_FLAECHE:
            befund = (f"Leere: Eine leere Fläche nimmt {flaeche:.0%} der Inhaltszone ein – mehr Inhalt oder eine "
                      "Bildform aus reference/bibliothek.md")
        ergebnis.append({"nr": nr, "leer": leer, "flaeche": flaeche, "leere": befund,
                         "text": _ueberlappung(zeilen) + _ueberstand(seite, zeilen)})
    schriften = {f[3].split("+")[-1] for seite in doc for f in seite.get_fonts()}
    doc.close()
    return ergebnis, schriften


def sichten(pptx, ausgabe=None, weg=None, folien=None):
    """Rendert das Deck (auch Ausgeblendetes, damit die Nummern stimmen) und meldet die Befunde. Rückgabe: Zahl der Befunde."""
    pptx = Path(pptx)
    layouts = [s.slide_layout.name for s in Presentation(str(pptx)).slides]
    pdf, _, _ = rendern(pptx, ausgabe, dpi=80, alle=True, weg=weg, blatt=True)
    ergebnis, schriften = messen(pdf, layouts)
    marke = any("Boehringer" in s for s in schriften)
    print(f"SICHTUNG {pptx.name}: {len(ergebnis)} Folien · Schrift: "
          + ("Boehringer" if marke else f"Ersatz ({', '.join(sorted(schriften)) or 'keine'}), Umbruch weicht ab"))
    zahl = 0
    for f in ergebnis:
        if folien and f["nr"] not in folien:
            continue
        if f["leere"]:
            print(f"  BEFUND Folie {f['nr']}: {f['leere']}")
            zahl += 1
        for t in f["text"]:
            print(f"  {'BEFUND' if marke else 'HINWEIS'} Folie {f['nr']}: {t}")
            zahl += 1 if marke else 0
    print(f"ERGEBNIS: {'OK' if not zahl else str(zahl) + ' Befund(e)'}")
    return zahl
