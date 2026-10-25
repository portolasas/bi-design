"""Vorlage (BI-Master) öffnen, Layouts finden, Fußzeile/Foliennummer/Quelle setzen, Vorlage extrahieren."""
import copy
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

# Kurzname → Layoutname im BI-Master (englische Fassung). Weitere Vorlagen: Namen hier ergänzen.
LAYOUTS = {
    "titel": ["Title slide (bright)"],
    "titel_dunkel": ["Title slide (dark)"],
    "titel_bild": ["Title slide image (small)"],
    "agenda": ["Agenda"],
    "kapitel": ["Chapter divider (1)"],
    "kapitel_dunkel": ["Chapter divider (2)"],
    "inhalt": ["One content"],
    "inhalt_dunkel": ["One content dark"],
    "zwei_inhalte": ["Two contents"],
    "drei_inhalte": ["Three contents"],
    "inhalt_bild": ["Content and image"],
    "nur_titel": ["Only title"],
    "leer": ["Empty"],
    "abschluss": ["One content dark Outro"],
    "disclaimer": ["Disclaimer"],
    "kontakt": ["Contact"],
}
DUNKLE_LAYOUTS = {"titel_dunkel", "kapitel_dunkel", "inhalt_dunkel", "abschluss"}


def oeffnen(pfad=None):
    """Vorlage/Basis öffnen. Ohne Pfad: leere 16:9-Präsentation (Notbetrieb ohne BI-Master)."""
    if pfad is None:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
        return prs
    p = Path(pfad)
    if p.suffix.lower() in (".potx", ".pptm", ".potm"):
        return _als_pptx(p)
    return Presentation(str(p))


def _als_pptx(p):
    """.potx/.pptm lassen sich mit python-pptx nur über den Content-Type öffnen."""
    import io
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(p) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = re.sub(rb"application/vnd\.ms-powerpoint\.(template|presentation)\.macroEnabled\.main\+xml|"
                              rb"application/vnd\.openxmlformats-officedocument\.presentationml\.template\.main\+xml",
                              b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", data)
            zout.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def ist_bi_master(prs):
    namen = {l.name for l in prs.slide_layouts}
    return "Only title" in namen and "Title slide (bright)" in namen


def layout(prs, kurz):
    namen = LAYOUTS.get(kurz, [kurz])
    for l in prs.slide_layouts:
        if l.name in namen:
            return l
    # Ersatz nach Struktur (fremde Vorlage / Notbetrieb)
    ersatz = {"titel": 0, "titel_dunkel": 0, "titel_bild": 0, "kapitel": 2, "kapitel_dunkel": 2, "agenda": 1,
              "inhalt": 1, "inhalt_dunkel": 1, "zwei_inhalte": 3, "drei_inhalte": 3, "inhalt_bild": 8,
              "nur_titel": 5, "leer": 6, "abschluss": 1, "disclaimer": 6, "kontakt": 1}
    i = ersatz.get(kurz, 5)
    return prs.slide_layouts[min(i, len(prs.slide_layouts) - 1)]


def platzhalter(folie, typ=None, idx=None):
    for ph in folie.placeholders:
        pf = ph.placeholder_format
        if idx is not None and pf.idx == idx:
            return ph
        if typ is not None and str(pf.type).split(".")[-1].split(" ")[0] in typ:
            return ph
    return None


def _layout_ph(folie, phtyp):
    for sp in folie.slide_layout.placeholders:
        el = sp._element.find(".//" + qn("p:ph"))
        if el is not None and el.get("type") == phtyp:
            return sp
    return None


def _hat_ph(folie, phtyp):
    for sp in folie.shapes._spTree.iter(qn("p:ph")):
        if sp.get("type") == phtyp:
            return True
    return False


def _ph_anlegen(folie, layout_sp, inhalt_xml):
    """Platzhalter vom Layout übernehmen (Position erbt), Text ersetzen."""
    sp = copy.deepcopy(layout_sp._element)
    spPr = sp.find(qn("p:spPr"))
    for ch in list(spPr):
        spPr.remove(ch)
    tx = sp.find(qn("p:txBody"))
    for p in tx.findall(qn("a:p")):
        tx.remove(p)
    lst = tx.find(qn("a:lstStyle"))
    if lst is not None:
        for ch in list(lst):
            lst.remove(ch)
    tx.append(etree.fromstring(inhalt_xml))
    ids = [int(x.get("id")) for x in folie.shapes._spTree.iter(qn("p:cNvPr")) if x.get("id", "").isdigit()]
    sp.find(".//" + qn("p:cNvPr")).set("id", str(max(ids + [1]) + 1))
    folie.shapes._spTree.append(sp)
    return sp


_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def fusszeile(folie, text, nummer=True):
    """Fußzeile und Foliennummer wie im Master setzen (python-pptx legt sie bei neuen Folien nicht an)."""
    if text is not None:
        vorhanden = [ph for ph in folie.placeholders if ph._element.find(".//" + qn("p:ph")).get("type") == "ftr"]
        if vorhanden:
            tf = vorhanden[0].text_frame
            tf.text = text
        else:
            lsp = _layout_ph(folie, "ftr")
            if lsp is not None:
                t = (text.replace("&", "&amp;").replace("<", "&lt;"))
                _ph_anlegen(folie, lsp, f'<a:p {_A}><a:r><a:rPr lang="de-DE"/><a:t>{t}</a:t></a:r></a:p>')
    if nummer and not _hat_ph(folie, "sldNum"):
        lsp = _layout_ph(folie, "sldNum")
        if lsp is not None:
            _ph_anlegen(folie, lsp, f'<a:p {_A}><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" '
                                    f'type="slidenum"><a:rPr lang="de-DE"/><a:t>‹#›</a:t></a:fld></a:p>')


def quelle(folie, text):
    """Quellen-/Hinweiszeile (BI-Master: Textplatzhalter idx 13 unten)."""
    ph = platzhalter(folie, idx=13)
    if ph is not None and text:
        ph.text_frame.text = text
    return ph is not None


def leere_platzhalter_entfernen(folie, behalten=("ftr", "sldNum")):
    for ph in list(folie.placeholders):
        t = ph._element.find(".//" + qn("p:ph")).get("type")
        if t in behalten:
            continue
        if ph.has_text_frame and not ph.text_frame.text.strip():
            ph._element.getparent().remove(ph._element)


def alle_folien_entfernen(prs):
    lst = prs.slides._sldIdLst
    for el in list(lst):
        prs.part.drop_rel(el.rId)
        lst.remove(el)


def extrahieren(quelle_pptx, ziel):
    """Vorlage aus einer Präsentation im BI-Master erzeugen: alle Folien raus, Master/Layouts/Theme bleiben.
    Doppelte Layouts aus eingefügten Folien („1_…“, „2_…“) werden entfernt."""
    prs = oeffnen(quelle_pptx)
    if not ist_bi_master(prs):
        print("WARNUNG: Quelle hat keine BI-Master-Layouts ('Only title', 'Title slide (bright)').")
    alle_folien_entfernen(prs)
    for master in prs.slide_masters:
        namen = {l.name for l in master.slide_layouts}
        for l in list(master.slide_layouts):
            basis = re.sub(r"^\d+_", "", l.name)
            if basis != l.name and basis in namen:
                master.slide_layouts.remove(l)
    Path(ziel).parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(ziel))
    return [l.name for l in prs.slide_layouts]
