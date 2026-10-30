"""Folienbibliothek: fertig gestaltete Folien der BI Slide Library (gleicher Master wie die Vorlage) in ein Deck kopieren
– auch mehrfach – und füllen.

Spec je Folie: {"bib": 25, "titel": "…", "texte": {"5": "…", "6": ["Absatz", {"t": "…", "wie": 2}]},
"bilder": {"7": "foto.jpg" | {"datei": "foto.jpg", "ausschnitt": "oben"}}, "daten": {"4": {"kategorien": […],
"reihen": [{"name": "…", "werte": […]}], "titel": "…" | ""}}, "tabellen": {"6": [["Kopf", …], ["Zelle", null, …]]},
"entfernen": [8, 9]} und wie bei Basisfolien "ersetzen", "platzhalter",
"aendern", "elemente", "notizen", "ausblenden". Die ids nennt `bippt katalog --folien 25`. Texte bekommen das Format des
vorhandenen Absatzes (aendern › absaetze). Nach dem Füllen warnt bippt, wenn ein Text mehr Platz braucht, als seine Form
hat, und wenn Blindtext der Bibliothek („Lorem ipsum“, „20XX“) übrig ist.
"""
import copy
import re
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, _Relationship
from pptx.oxml.ns import qn

from . import konfig, vorlage
from .aendern import absaetze_setzen, aendern
from .bausteine import umbrechen
from .formen import alle, finden

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WEGLASSEN = {RT.NOTES_SLIDE, RT.COMMENTS}  # Hinweise der Bibliothek und Kommentare kommen nicht mit
NACH_OBEN = {RT.SLIDE, RT.SLIDE_LAYOUT, RT.SLIDE_MASTER, RT.NOTES_MASTER, RT.HANDOUT_MASTER}
BLINDTEXT = re.compile(r"(?i)\blorem\b|\bipsum\b|\bdolor sit\b|\bmaecenas\b|\bconsectetu?r\b|\bvestibulum\b|"
                       r"\bpraesent\b|\bfusce\b|\bpellentesque\b|20XX|XX,XXX|\bPage 00\b|Author \| Position|"
                       r"First name Last name|Name of the current chapter|Optional short description|Diagram title|"
                       r"Data series \d|Category \d|Header box|\bHeadline\b|Image and text|Data overview|Excel based chart")
FELDER = {"bib", "titel", "texte", "bilder", "daten", "tabellen", "entfernen", "ersetzen", "platzhalter", "aendern", "elemente",
          "notizen", "ausblenden", "ab_hier_ausblenden", "beschreibung", "fusszeile", "farben_angleichen"}


def pfad(spec, ordner):
    """Bibliothek aus der Spec (relativ zu ihr), sonst aus den Einstellungen bzw. dem Plugin."""
    if spec.get("bibliothek"):
        p = Path(spec["bibliothek"]).expanduser()
        return p if p.is_absolute() else Path(ordner) / p
    return konfig.bibliothek_pfad()


# ----------------------------------------------------------------- Kopieren
def _muster(partname):
    """„/ppt/media/image12.png“ → „/ppt/media/image%d.png“ (für next_partname)."""
    m = re.match(r"^(.*?)(\d+)(\.[^./]+)$", partname)
    if m:
        return m.group(1) + "%d" + m.group(3)
    stamm, _, endung = partname.rpartition(".")
    return f"{stamm}%d.{endung}"


def _rel(teil, rid, reltype, ziel, extern=False):
    """Beziehung mit fester rId (die XML des kopierten Teils verweist auf sie)."""
    teil.rels._rels[rid] = _Relationship(teil.rels._base_uri, rid, reltype,
                                         target_mode=RTM.EXTERNAL if extern else RTM.INTERNAL, target=ziel)


def _teil_kopieren(quelle, vater, rid, reltype, paket, kopien):
    """Teil (Bild, Diagramm mit Arbeitsmappe, OLE-Objekt, Tags …) samt Unterteilen ins Zielpaket kopieren und an vater
    hängen. Jeder Teil wird erst eingehängt, dann findet next_partname ihn beim nächsten Namen."""
    neu = kopien.get(id(quelle))
    if neu is not None:
        _rel(vater, rid, reltype, neu)
        return
    neu = PartFactory(paket.next_partname(_muster(str(quelle.partname))), quelle.content_type, paket, quelle.blob)
    kopien[id(quelle)] = neu
    _rel(vater, rid, reltype, neu)
    for rid2, r2 in quelle.rels.items():
        if r2.is_external:
            _rel(neu, rid2, r2.reltype, r2.target_ref, extern=True)
        elif r2.reltype not in NACH_OBEN:
            _teil_kopieren(r2.target_part, neu, rid2, r2.reltype, paket, kopien)


def _verweise_entfernen(el, rid):
    """Sprünge zu anderen Folien der Bibliothek (Agenda „Page 003“) entfernen; ihr Ziel kommt nicht mit."""
    for x in list(el.iter()):
        if any(k.startswith(R) and v == rid for k, v in x.attrib.items()):
            if x.tag in (qn("a:hlinkClick"), qn("a:hlinkHover")):
                x.getparent().remove(x)
            else:
                for k in [k for k, v in x.attrib.items() if k.startswith(R) and v == rid]:
                    del x.attrib[k]


def _inhalt_uebernehmen(ziel, quelle):
    """Kinder von p:sld kopieren. p:cSld und p:spTree der neuen Folie bleiben dieselben Objekte – python-pptx hat sie
    sich schon gemerkt (slide.shapes)."""
    z_csld, q_csld = ziel.find(qn("p:cSld")), quelle.find(qn("p:cSld"))
    z_baum = z_csld.find(qn("p:spTree"))
    for kind in list(z_baum):
        z_baum.remove(kind)
    for kind in q_csld.find(qn("p:spTree")):
        z_baum.append(copy.deepcopy(kind))
    for kind in list(z_csld):
        if kind is not z_baum:
            z_csld.remove(kind)
    for kind in q_csld:
        if kind.tag == qn("p:spTree"):
            continue
        if kind.tag == qn("p:bg"):
            z_baum.addprevious(copy.deepcopy(kind))
        else:
            z_csld.append(copy.deepcopy(kind))
    for k, v in q_csld.attrib.items():
        z_csld.set(k, v)
    for kind in list(ziel):
        if kind is not z_csld:
            ziel.remove(kind)
    for kind in quelle:
        if kind.tag != qn("p:cSld"):
            ziel.append(copy.deepcopy(kind))
    for k, v in quelle.attrib.items():
        ziel.set(k, v)


def _layout(prs, name):
    for m in prs.slide_masters:
        for lay in m.slide_layouts:
            if lay.name == name:
                return lay, None
    return vorlage.layout(prs, "nur_titel"), f"Layout „{name}“ fehlt in der Vorlage – „Only title“ genommen"


def kopieren(prs, bib, nr):
    """Folie nr (ab 1) der Bibliothek bib als neue Folie ans Ende von prs. Rückgabe (Folie, Warnung oder None)."""
    if not 1 <= nr <= len(bib.slides):
        raise SystemExit(f"bib {nr} existiert nicht (Bibliothek hat {len(bib.slides)} Folien)")
    quelle = bib.slides[nr - 1]
    layout, warnung = _layout(prs, quelle.slide_layout.name)
    folie = prs.slides.add_slide(layout)
    _inhalt_uebernehmen(folie._element, quelle._element)
    teil, kopien = folie.part, {}
    teil.rels._rels.clear()  # rId der Vorlage aus add_slide; die kopierte XML nutzt die rIds der Quelle
    for rid, rel in quelle.part.rels.items():
        if rel.reltype == RT.SLIDE_LAYOUT:
            _rel(teil, rid, rel.reltype, layout.part)
        elif rel.reltype in WEGLASSEN:
            continue
        elif rel.reltype in NACH_OBEN:
            _verweise_entfernen(folie._element, rid)
        elif rel.is_external:
            _rel(teil, rid, rel.reltype, rel.target_ref, extern=True)
        else:
            _teil_kopieren(rel.target_part, teil, rid, rel.reltype, prs.part.package, kopien)
    # Bedienhinweise der Bibliothek liegen neben der Folie („To insert a picture …“) – sie kommen nicht mit
    for sh in list(folie.shapes):
        if sh.left is None or sh.width is None:
            continue
        if sh.left + sh.width <= 0 or sh.left >= prs.slide_width or sh.top + sh.height <= 0 or sh.top >= prs.slide_height:
            sh._element.getparent().remove(sh._element)
    return folie, warnung


# ----------------------------------------------------------------- Füllen
def _daten(k, sid, d):
    t = [sh for sh in finden(k.folie, sid=sid) if getattr(sh, "has_chart", False)]
    if not t:
        k.warn(f"daten: kein Diagramm mit id {sid} (bippt katalog --folien n)")
        return
    cd = CategoryChartData()
    cd.categories = d.get("kategorien", [])
    for r in d.get("reihen", []):
        cd.add_series(r.get("name", ""), r.get("werte", []))
    chart = t[0].chart
    try:
        chart.replace_data(cd)
    except (ValueError, KeyError, TypeError, AttributeError) as e:
        k.warn(f"daten (id {sid}): Diagrammart lässt sich so nicht füllen ({e}) – Typ diagramm nehmen")
    if "titel" in d:  # "" oder null: ohne Diagrammtitel
        if d["titel"]:
            chart.has_title = True
            chart.chart_title.text_frame.text = d["titel"]
        else:
            chart.has_title = False


def _tabelle(k, sid, zeilen):
    """Tabelle zeilenweise füllen (mit Kopfzeile); Zellformat bleibt, null = Zelle bleibt. Fehlende Zeilen sind Kopien
    der letzten, überzählige entfallen – die letzte Zeile der Bibliothek ist oft eine Summenzeile. Überzählige Spalten
    entfallen, ihre Breite geht anteilig an die übrigen (die Tabelle behält ihre Breite)."""
    t = [sh for sh in finden(k.folie, sid=sid) if getattr(sh, "has_table", False)]
    if not t:
        k.warn(f"tabellen: keine Tabelle mit id {sid} (bippt katalog --folien n)")
        return
    tbl = t[0].table._tbl
    while len(tbl.tr_lst) < len(zeilen):
        tbl.tr_lst[-1].addnext(copy.deepcopy(tbl.tr_lst[-1]))
    while len(tbl.tr_lst) > max(len(zeilen), 1):
        tbl.tr_lst[-1].getparent().remove(tbl.tr_lst[-1])
    spalten = max((len(z) for z in zeilen), default=0)
    grid = tbl.tblGrid.gridCol_lst
    if 0 < spalten < len(grid):
        weg = sum(int(g.get("w")) for g in grid[spalten:])
        rest = sum(int(g.get("w")) for g in grid[:spalten])
        for g in grid[:spalten]:
            g.set("w", str(int(g.get("w")) + int(weg * int(g.get("w")) / rest)))
        for g in grid[spalten:]:
            g.getparent().remove(g)
        for tr in tbl.tr_lst:
            for tc in tr.tc_lst[spalten:]:
                tr.remove(tc)
    for row, werte in zip(t[0].table.rows, zeilen):
        if len(werte) > len(row.cells):
            k.warn(f"tabellen (id {sid}): {len(werte)} Werte für {len(row.cells)} Spalten – Rest entfällt")
        for c, w in zip(row.cells, werte):
            if w is not None:
                absaetze_setzen(c.text_frame, w if isinstance(w, list) else [str(w)])


def _ebene(sh, p):
    """Absatzformat wie PowerPoint es auflöst: a:pPr des Absatzes, sonst lvlNpPr der lstStyle der Form. Platzhalter
    erben den Rest vom Master: Fließtext 16 pt mit 140 % Zeilenabstand (BI-Master, am Render gemessen)."""
    ppr = p._p.find(qn("a:pPr"))
    lvl = int(ppr.get("lvl", 0)) if ppr is not None else 0
    lst = sh.text_frame._txBody.find(qn("a:lstStyle"))
    stil = lst.find(qn(f"a:lvl{lvl + 1}pPr")) if lst is not None else None
    quellen = [x for x in (ppr, stil) if x is not None]

    def wert(tag):
        return next((x.find(qn(tag)) for x in quellen if x.find(qn(tag)) is not None), None)

    def attr(name, std):
        return next((x.get(name) for x in quellen if x.get(name) is not None), std)

    ph = sh.is_placeholder
    groesse = next((r.font.size.pt for r in p.runs if r.font.size), None)
    if groesse is None:
        d = wert("a:defRPr")
        groesse = int(d.get("sz")) / 100 if d is not None and d.get("sz") else (16.0 if ph else 18.0)
    zs = wert("a:lnSpc")
    pct = zs.find(qn("a:spcPct")) if zs is not None else None
    faktor = int(pct.get("val")) / 100000 if pct is not None else (1.4 if ph else 1.0)

    def abstand(tag):
        el = wert(tag)
        if el is None:
            return 0.0
        pts, pc = el.find(qn("a:spcPts")), el.find(qn("a:spcPct"))
        if pts is not None:
            return int(pts.get("val")) / 100 / 72
        return int(pc.get("val")) / 100000 * groesse * 1.2 / 72 if pc is not None else 0.0

    return {"groesse": groesse, "faktor": faktor, "einzug": int(attr("marL", 0)) / 914400,
            "vor": abstand("a:spcBef"), "nach": abstand("a:spcAft")}


def _absatzarten(sh):
    """Aufbau eines Textfelds je Absatz: „fett“ (Überschrift), „Text“ oder „Punkt“ (Aufzählung), gleiche zusammengefasst –
    so weiß, wer texte schreibt, welcher Absatz welches Format bekommt."""
    arten = []
    lst = sh.text_frame._txBody.find(qn("a:lstStyle"))
    for p in sh.text_frame.paragraphs:
        ppr = p._p.find(qn("a:pPr"))
        lvl = int(ppr.get("lvl", 0)) if ppr is not None else 0
        stil = lst.find(qn(f"a:lvl{lvl + 1}pPr")) if lst is not None else None
        punkt = False
        for x in (ppr, stil):
            if x is not None and x.find(qn("a:buNone")) is not None:
                break
            if x is not None and x.find(qn("a:buChar")) is not None:
                punkt = True
                break
        art = "Punkt" if punkt else "fett" if p.runs and all(r.font.bold for r in p.runs if r.text.strip()) else "Text"
        if arten and arten[-1][0] == art:
            arten[-1][1] += 1
        else:
            arten.append([art, 1])
    return ", ".join(f"{n}× {a}" if n > 1 else a for a, n in arten)


def _schrift(sh):
    """Schriftgröße (pt) des ersten Absatzes."""
    ps = sh.text_frame.paragraphs
    return _ebene(sh, ps[0])["groesse"] if ps else 18.0


def bedarf(sh):
    """Höhe (in), die der Text der Form braucht: Umbruch wie PowerPoint (umbrechen), Zeilen- und Absatzabstand und
    Einzug wie _ebene. Ohne Textrahmen oder Breite: None."""
    if not getattr(sh, "has_text_frame", False) or not sh.width:
        return None
    tf = sh.text_frame
    rand_lr = ((tf.margin_left if tf.margin_left is not None else 91440)
               + (tf.margin_right if tf.margin_right is not None else 91440)) / 914400
    h = ((tf.margin_top if tf.margin_top is not None else 45720)
         + (tf.margin_bottom if tf.margin_bottom is not None else 45720)) / 914400
    ps = tf.paragraphs
    for i, p in enumerate(ps):
        e = _ebene(sh, p)
        fett = any(r.font.bold for r in p.runs)
        zeilen = umbrechen(p.text or " ", sh.width / 914400 - rand_lr - e["einzug"], e["groesse"], fett=fett)[0]
        h += zeilen * e["groesse"] * 1.2 * e["faktor"] / 72 + (e["vor"] if i else 0) + (e["nach"] if i < len(ps) - 1 else 0)
    return h


def _platz(k, sid):
    for sh in finden(k.folie, sid=sid):
        b = bedarf(sh)
        if b is None or not sh.height:
            continue
        hoehe = sh.height / 914400
        if b > hoehe * 1.15 + 0.08:
            k.warn(f"„{sh.name}“ (id {sid}): Text braucht etwa {b:.2f} in, die Form hat {hoehe:.2f} in – kürzen "
                   "(bippt katalog nennt den Platz)")


def _leer(k, sid):
    """Große Fläche nur halb gefüllt: So wirkt die Folie unfertig (Kopfboxen mit drei kurzen Punkten)."""
    for sh in finden(k.folie, sid=sid):
        b = bedarf(sh)
        if b is None or not sh.height or sh.height < 914400 * 2 or getattr(sh.fill, "type", None) is None:
            continue
        if b < sh.height / 914400 * 0.5:
            k.warn(f"„{sh.name}“ (id {sid}): Fläche nur zu {b / (sh.height / 914400):.0%} gefüllt – eine Folie mit weniger "
                   "Fläche nehmen (z. B. Fazit-Band 36–38, Fakten 39–44) oder die Form mit aendern h kürzen")
            return


def layouts(pptx):
    """Markdown: Layouts des Masters mit ihren Feldern für den Typ layout."""
    z = ["# Layouts des Masters", 'Füllen: `{"typ": "layout", "layout": "<Name>", "titel": "…", "felder": {"idx": …}, '
         '"bilder": {"idx": "foto.jpg"}}`. Leere Textfelder entfallen; Bildplatzhalter brauchen ein Bild.']
    for lay in vorlage.oeffnen(pptx).slide_layouts:
        felder = []
        for ph in lay.placeholders:
            art = str(ph.placeholder_format.type).split(".")[-1].split(" ")[0]
            if art in ("FOOTER", "SLIDE_NUMBER", "DATE", "TITLE", "CENTER_TITLE"):
                continue
            name = {"PICTURE": "Bild", "OBJECT": "Inhalt", "BODY": "Text", "SUBTITLE": "Untertitel"}.get(art, art.lower())
            if ph.placeholder_format.idx == 13:
                name = "Quelle"
            felder.append(f"{ph.placeholder_format.idx} {name} {ph.width / 914400:.1f}×{ph.height / 914400:.1f}")
        z.append(f"- **{lay.name}**: " + (" · ".join(felder) or "nur Titel"))
    return "\n".join(z) + "\n"


# Formfamilien für die Abwechslung im Deck: Folien gleicher Bauart zählen zusammen
FAMILIEN = [(range(22, 28), "Textspalten mit Kopfbox", True), (range(28, 36), "Texte mit Icons", False),
            (range(36, 39), "Bereiche mit Fazit-Band", False), (range(39, 45), "nummerierte Fakten", False),
            (range(45, 46), "Textboxen", True), (range(46, 55), "Zitate", False), (range(56, 61), "Abwägung", False),
            (range(62, 76), "Bild und Text", False), (range(77, 87), "Phasen", False), (range(87, 93), "Trend", False),
            (range(93, 98), "Zeitstrahl", False), (range(98, 102), "Projektplan", False),
            (range(103, 113), "Organisation", False), (range(114, 118), "Tabellen", True),
            (range(119, 135), "Diagramme", False), (range(136, 152), "Karten", False),
            (range(153, 164), "Grafiken", False), (range(165, 176), "Grafiken", False)]
RAHMEN = {"titel", "kapitel", "agenda", "abschluss"}  # zählen nicht als Inhaltsfolien
RAHMEN_BIB = set(range(1, 22)) | {55, 61, 76, 102, 113, 118, 135, 152, 164, 176, 180, 181, 182, 183, 184}
TEXT_TYPEN = {"text", "tabelle"}


def _form(fs):
    """(Formname, reine Textfolie?) einer neuen Folie; None für Rahmenfolien, Basisfolien."""
    if "basis" in fs:
        return None
    if "bib" in fs:
        n = int(fs["bib"])
        if n in RAHMEN_BIB:
            return None
        return next(((f, t) for r, f, t in FAMILIEN if n in r), (f"Bibliothek {n}", False))
    typ = fs.get("typ", "text")
    if typ in RAHMEN:
        return None
    if typ == "layout":
        name = str(fs.get("layout", ""))
        return (f"Layout {name}", not fs.get("bilder") and any(x in name for x in ("content", "Agenda")))
    return (typ, typ in TEXT_TYPEN)


def abwechslung(folien):
    """Warnungen zur Gestaltung des Decks über alle sichtbaren neuen Folien: dieselbe Bauart hintereinander oder zu oft,
    zu viele reine Textfolien. Erst ab vier Inhaltsfolien; Backup zählt nicht."""
    formen, ausblenden, w = [], False, []
    for nr, fs in enumerate(folien, 1):
        ausblenden = ausblenden or bool(fs.get("ab_hier_ausblenden"))
        if ausblenden or fs.get("ausblenden"):
            continue
        f = _form(fs)
        if f:
            formen.append((nr, f[0], f[1], fs.get("bib")))
    if len(formen) < 4:
        return w
    for a, b in zip(formen, formen[1:]):
        if a[1] == b[1] and a[1] != "timeline":  # Gesamtüberblick und kritischer Pfad gehören zusammen
            w.append(f"Abwechslung: Folien {a[0]} und {b[0]} haben dieselbe Bauart ({a[1]}) – eine andere Bildform wählen")
    zahl, gleiche = {}, {}
    for _, name, _, bib in formen:
        zahl[name] = zahl.get(name, 0) + 1
        if bib:
            gleiche[bib] = gleiche.get(bib, 0) + 1
    w += [f"Abwechslung: Bibliotheksfolie {n} steht {c}× im Deck – höchstens einmal, sonst wirkt es eintönig"
          for n, c in gleiche.items() if c > 1]
    w += [f"Abwechslung: {c} von {len(formen)} Inhaltsfolien sind „{n}“ – höchstens zwei je Bauart"
          for n, c in zahl.items() if c > 2]
    text = sum(1 for x in formen if x[2])
    if text > len(formen) / 3:
        w.append(f"Abwechslung: {text} von {len(formen)} Inhaltsfolien sind reine Text- oder Tabellenfolien, höchstens ein "
                 "Drittel – Aussagen in Bildformen umsetzen (Skill bi-ppt › Aussage → Bildform)")
    return w


def blindtext(folie):
    """Formen mit übrig gebliebenem Blindtext der Bibliothek: [(id, Name, Auszug)]."""
    treffer = []
    for sh in alle(folie.shapes):
        if sh.is_placeholder and sh.placeholder_format.type is not None and \
                str(sh.placeholder_format.type).split(".")[-1].split(" ")[0] in ("FOOTER", "SLIDE_NUMBER", "DATE"):
            continue
        if getattr(sh, "has_text_frame", False) and BLINDTEXT.search(sh.text_frame.text):
            treffer.append((sh.shape_id, sh.name, re.sub(r"\s+", " ", sh.text_frame.text)[:40]))
        if getattr(sh, "has_chart", False) and sh.chart.has_title and BLINDTEXT.search(
                sh.chart.chart_title.text_frame.text):
            treffer.append((sh.shape_id, sh.name, "Diagrammtitel " + sh.chart.chart_title.text_frame.text[:30]))
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                for c in row.cells:
                    if BLINDTEXT.search(c.text):
                        treffer.append((sh.shape_id, sh.name, re.sub(r"\s+", " ", c.text)[:40]))
                        break
    return treffer


def fuellen(k, fs, ordner):
    """texte, bilder, entfernen, daten einer Bibliotheksfolie; danach Platz und Blindtext prüfen."""
    fremd = set(fs) - FELDER
    if fremd:
        k.warn(f"unbekannte Felder für bib: {', '.join(sorted(fremd))}")
    ops = []
    for sid, wert in (fs.get("texte") or {}).items():
        ops.append({"id": int(sid), "absaetze": wert if isinstance(wert, list) else [wert]})
    for sid, wert in (fs.get("bilder") or {}).items():
        if isinstance(wert, dict):
            ops.append({"id": int(sid), "bild": wert.get("datei"), "ausschnitt": wert.get("ausschnitt")})
        else:
            ops.append({"id": int(sid), "bild": wert})
    for sid in fs.get("entfernen") or []:
        ops.append({"id": int(sid), "entfernen": True})
    if ops:
        aendern(k, ops, ordner)
    for sid, d in (fs.get("daten") or {}).items():
        _daten(k, int(sid), d)
    for sid, zeilen in (fs.get("tabellen") or {}).items():
        _tabelle(k, int(sid), zeilen)
    for sid in fs.get("texte") or {}:
        _platz(k, int(sid))
        _leer(k, int(sid))


# ----------------------------------------------------------------- Katalog
ARTEN = {"PICTURE": "Bild", "CHART": "Diagramm", "TABLE": "Tabelle", "GROUP": "Gruppe", "LINE": "Linie",
         "EMBEDDED_OLE_OBJECT": "OLE", "FREEFORM": "Freiform", "TEXT_BOX": "Text", "AUTO_SHAPE": "Form"}


def _art(sh):
    t = str(sh.shape_type).split(".")[-1].split(" ")[0] if sh.shape_type is not None else ""
    if sh.is_placeholder:
        if sh.placeholder_format.idx == 13:
            return "Quellenzeile"
        pt = str(sh.placeholder_format.type).split(".")[-1].split(" ")[0]
        return {"TITLE": "Titel", "CENTER_TITLE": "Titel", "PICTURE": "Bildplatzhalter", "BODY": "Textplatzhalter",
                "OBJECT": "Inhaltsplatzhalter", "SUBTITLE": "Untertitel"}.get(pt, pt.lower())
    if t == "PICTURE" and sh.top and sh.top > 914400 * 6.4 and sh.width < 914400 * 1.6:
        return "Logo (nicht ändern)"  # Logovariante für farbige Hintergründe, deckt das Logo des Masters
    if t == "PICTURE" and sh.width and sh.width < 914400 * 0.8:
        return "Icon"
    return ARTEN.get(t, t.lower() or "?")


def katalog(pptx, folien=None):
    """Markdown: je Folie Titel, Layout und die Formen zum Füllen – id, Art, Lage, Schrift, Platz, Beispieltext."""
    prs = vorlage.oeffnen(pptx)
    z = [f"# Katalog `{Path(pptx).name}` · {len(prs.slides)} Folien",
         "Je Form: id · Art · Lage x, y, Breite × Höhe (in) · Schrift (pt) · Platz (Zeilen × Zeichen) · Absätze "
         "(fett = Überschrift, Text, Punkt = Aufzählung; der n-te Absatz von texte bekommt das Format des n-ten) · "
         "Beispieltext. "
         "Füllen mit `{\"bib\": n, \"texte\": {\"id\": …}, \"bilder\": {\"id\": …}, \"daten\": {\"id\": …}, "
         "\"tabellen\": {\"id\": [[…]]}, \"entfernen\": [id]}`. OLE-Objekte (think-cell) und Fußzeile nicht anfassen."]
    for nr, sl in enumerate(prs.slides, 1):
        if folien and nr not in folien:
            continue
        t = re.sub(r"\s+", " ", sl.shapes.title.text_frame.text.strip()) if sl.shapes.title is not None else ""
        dunkel = " · dunkel" if "dark" in sl.slide_layout.name else ""
        z += ["", f"## {nr}: {t or '(ohne Titel)'} [{sl.slide_layout.name}]{dunkel}"]
        formen = sorted(alle(sl.shapes), key=lambda f: (round((f.top or 0) / 914400, 1), (f.left or 0)))
        for sh in formen:  # von oben nach unten, links nach rechts
            if sh.left is not None and sh.width is not None and (sh.left + sh.width <= 0 or sh.left >= prs.slide_width):
                continue  # Bedienhinweis neben der Folie, kommt nicht mit
            if sh.is_placeholder and str(sh.placeholder_format.type).split(".")[-1].split(" ")[0] in (
                    "FOOTER", "SLIDE_NUMBER", "DATE"):
                continue
            art = _art(sh)
            klein = sh.width is not None and sh.width < 228600 and sh.height < 228600
            if art in ("Linie", "OLE") or klein and not (getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip()):
                continue  # think-cell-Anker, Trennlinien
            lage = (f"{sh.left / 914400:.1f}, {sh.top / 914400:.1f}, {sh.width / 914400:.1f} × {sh.height / 914400:.1f}"
                    if sh.width is not None and sh.left is not None else "geerbt")
            teile = [f"id {sh.shape_id}", art, lage]
            if art == "Titel":
                teile.append("eine Zeile, höchstens ~68 Zeichen")
            elif getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip() and sh.width and sh.height:
                e = _ebene(sh, sh.text_frame.paragraphs[0])
                g = e["groesse"]
                zeilen = max(1, int((sh.height / 914400 - 0.1) / (g * 1.2 * e["faktor"] / 72)))
                zeichen = max(1, int((sh.width / 914400 - 0.3 - e["einzug"]) / (g * 0.55 / 72)))
                text = re.sub(r"\s+", " ", sh.text_frame.text.strip())
                teile += [f"{g:g} pt", f"Platz ~{zeilen} × {zeichen}", f"Absätze: {_absatzarten(sh)}",
                          f"„{text[:50]}{'…' if len(text) > 50 else ''}“"]
            elif getattr(sh, "has_chart", False):
                c = sh.chart
                teile.append(f"{str(c.chart_type).split('.')[-1].split(' ')[0].lower()}, "
                             f"{len(list(c.plots[0].categories)) if c.plots else 0} Kategorien, "
                             f"{len(list(c.series))} Reihen")
            elif getattr(sh, "has_table", False):
                teile.append(f"{len(sh.table.rows)} × {len(sh.table.columns)} Zellen")
            z.append("- " + " · ".join(x for x in teile if x))
    return "\n".join(z) + "\n"
