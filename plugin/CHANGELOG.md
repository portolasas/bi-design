# Changelog

## 2.6.0 – 2026-10-30
Befund aus Erprobung 6 im Copilot Agent Builder (Steering-Deck): Der Agent nahm sechsmal Bibliotheksfolie 25, die Nummer aus
den Beispielen der Doku; der Nutzer fand das zu textlastig und will den Inhalt erst im Dialog erarbeiten.
- **Skill `bi-inhalt`:** Inhalt im Dialog in fünf Schritten (Auftrag, Kernbotschaft, Storyline, Folie für Folie, Briefing).
  bi-ppt baut aus dem abgenommenen Briefing ohne neue Rückfragerunde.
- **Bildform zuerst:** Tabelle Aussage → Bildform in `reference/bibliothek.md`, Storyboard als Schritt vor der Spec; die
  Doku nennt keine feste Foliennummer mehr als Beispiel.
- **`bauen` prüft die Abwechslung** über alle sichtbaren neuen Folien: dieselbe Bauart hintereinander, dieselbe Bibliotheksfolie
  mehrfach, mehr als zwei Folien je Bauart, mehr als ein Drittel reiner Text oder Tabelle. Dazu Warnung bei großen Flächen,
  die nur halb gefüllt sind.
- **Typ `layout`:** echte Platzhalter jedes Master-Layouts (Two/Three contents, Content and margin, Key message, Bild hochkant);
  `bippt layouts` nennt die Felder.

## 2.5.1 – 2026-10-29
- `bippt katalog` nennt je Textfeld den Aufbau der Absätze („Absätze: fett, Text, 2× Punkt“): Der n-te Absatz von
  `texte` bekommt das Format des n-ten. Am Musterdeck aus der Bibliothek fiel auf, dass sonst Überschrift, Text und
  Aufzählung vertauscht werden. `reference/bibliothek.md`: anderes Format mit `{"t": …, "wie": n}`.

## 2.5.0 – 2026-10-29
Folienbibliothek und Befunde aus der Erprobung im Copilot Agent Builder (Laborneubau-Deck, 29.10.2026). Der Nutzer fand
die Decks zu schematisch und unzuverlässig gerendert und hat die offizielle BI Slide Library bereitgestellt.
- **Folienbibliothek:** `vorlagen/BI_Slide_Library.pptx` (Slide Library 3.3, 184 Folien, gleicher Master). Spec-Folien
  `{"bib": n, "titel", "texte": {id: …}, "bilder": {id: …}, "daten": {id: {kategorien, reihen, titel}}, "tabellen":
  {id: [[…]]}, "entfernen": [id]}` kopieren Folie n samt Bildern, Diagrammen, OLE-Objekten und Tags ins Deck – auch
  mehrfach – und füllen sie; danach wirken `ersetzen`, `aendern`, `elemente` wie bei Basisfolien. Sprünge zu anderen
  Bibliotheksfolien, Notizen und Bedienhinweise neben der Folie kommen nicht mit; die Fußzeile wird gesetzt.
- `bippt katalog [--folien n]`: je Form id, Art, Lage, Schrift, Platz (Zeilen × Zeichen) und Beispieltext.
- `bauen` warnt bei Bibliotheksfolien, wenn ein Text mehr Platz braucht, als seine Form hat (Absatzformat wie
  PowerPoint aufgelöst), und wenn Blindtext übrig ist („Lorem ipsum“, „20XX“, „Header box“, „Diagram title“).
- Folienbeschreibung: „Bibliotheksfolie n“, Inhalt spaltenweise gelesen.
- Referenz `reference/bibliothek.md`: wann Bibliothek, wann Generator; Übersicht aller Folien; Regeln zu Blindtext,
  Icons, Bildern und Farbvarianten. Skill §3 und Arbeitsweise: gestaltete Folie vor generierter.
- Fehler: `bild_text` mit Band oder Quelle – Bild und Text liefen unter das Band. Beide enden jetzt darüber; das
  eingesetzte Bild behält den gekürzten Rahmen.
- Warnungen: Aufzählung in `text` und `bild_text` läuft über den Platzhalter (Maße des Masters, am Render gemessen);
  BI-Logo oder -Symbol als Bild (kommt aus dem Master); Bild unter 100 dpi im sichtbaren Ausschnitt.
- Tests: Bild und Text mit Band, die drei Warnungen, Bibliothek (Kopie doppelt, Texte, Bild, Diagramm, Tabelle,
  Sprünge, Notizen, Fußzeile, Platz, Blindtext, Katalog).

## 2.4.1 – 2026-10-29
Befund aus der ersten Erprobung im Copilot Agent Builder (Steering-Deck aus einem Projektauftrag, 29.10.2026).
- Fehler: Typ `text` mit `band` setzte nur die Höhe des Inhaltsplatzhalters. Weil der Platzhalter seine Lage vom Layout
  erbt, entstand eine Form mit Breite 0 und ohne Lage – der Text stand in PowerPoint einen Buchstaben je Zeile am
  linken Rand. Betraf alle Textfolien mit Band seit 1.0.0, in VS Code wie in Copilot. Jetzt schreibt bippt Lage und
  Breite mit.
- `pruefen` meldet Textfelder unter 0,3 in Breite als FEHLER. So fällt der Fehler auch ohne Rendern auf (Copilot) und
  in Decks, die mit älteren Versionen gebaut wurden: dort die Spec neu bauen.
- Tests: Textfolie mit Band und Quelle, Textfeld mit Breite 0.

## 2.4.0 – 2026-10-29
- bippt läuft auch als Skill-Paket für Agenten im Microsoft 365 Copilot Agent Builder (Repo `portolasas/Agents`,
  Ordner `bi-ppt`; Nutzer, 29.10.2026). Im Paket liegen Markenprofil und Beispiele flach neben `tools/`, weil Agent
  Builder höchstens drei Ordnerebenen annimmt: bippt sucht das Profil zusätzlich in `<plugin>/marken/`, `bippt neu` das
  Beispiel in `<plugin>/beispiele/`. Ohne Benutzerordner (Sandbox) nimmt bippt den Temp-Ordner statt abzubrechen.

## 2.3.3 – 2026-10-29
- Installation in VS Code ohne JSON: Marktplatz `portolasas/bi-design` in den Einstellungen eintragen
  (`chat.plugins.marketplaces` › „Element hinzufügen“) oder unter Customizations › Configure Marketplaces;
  README und `reference/installation.md` (Nutzer, 29.10.2026; wie ppp 0.4.2).

## 2.3.2 – 2026-10-28
Befunde aus brrp (28.10.2026, Deck mit Basisfolien und neuen Folien): falsche Warnungen, doppelte Ersetzungen und
Überdeckungen, die erst im Render auffielen.
- Fehler: `ersetzen` lief je Eintrag einmal über die ganze Folie. Folgen: falsche Warnungen („8 von 9 Ersetzungen nicht
  gefunden“, obwohl alle ersetzt waren) und mehrfach ersetzter Text, wenn der neue Text den alten enthält („GBS United
  United“). Jetzt ein Aufruf je Folie; alle Einträge wirken in einem Durchgang von links nach rechts, eingesetzter Text
  wird nie noch einmal ersetzt (auch nicht von einem anderen Eintrag).
- Titel: `bauen` und `pruefen` schätzen die Zeilen des Titels (Schrift, Größe, Ränder und Anker geerbt aus Layout und
  Master) und warnen, wenn eine zweite Zeile in den Inhalt darunter ragt (Typen auf „Only title“, ab etwa 68–72 Zeichen)
  oder ein unten verankerter Titel (Titel-, Kapitelfolie) über Logo oder Folienrand wächst. Der Inhalt rückt nicht nach
  unten – Titel kürzen.
- Wortbreite: Textbreiten nach den Zeichenbreiten von Boehringer Forward (±2 %, am Render geprüft), Umbruch wie in
  PowerPoint. `bauen` meldet je Folie Wörter, die breiter als ihr Textfeld sind und mitten im Wort brechen (Karten,
  Schrittkarten, Tabellenzellen, Boxen); bei `prozess` mit der Zahl Schritte, die nebeneinander passen.
- Tabelle: Zeilenhöhe aus Umbruch, Schriftgröße und Spaltenbreite (am Render kalibriert) statt fest mindestens 0,36 in –
  13 Zeilen à 10 pt passen jetzt über das Band. Ohne `groesse` wird die Schrift bis 9 pt kleiner; passt es nicht, warnt
  `bauen` mit Bedarf und Platz statt „viele Zeilen – Höhe prüfen“ (bei fester `groesse` mit der passenden Größe).
- Quelle: Mit `quelle` enden Band und Inhalt über der Quellenzeile; bisher lag die Quelle unter dem Band.
- Tests: Ersetzen (falsche Warnung, Wiederholung, Kette), Umbruchzeichen, Titel, Schrittkarten, Tabelle, Quelle und Band.

## 2.3.1 – 2026-10-28
- Sichern wie bei ppp (Nutzer): Liegt der Arbeitsordner in einem Git-Repo, holt bi-ppt zu Beginn mit `git pull` und
  committet nach jedem abgeschlossenen Schritt (Deck gebaut und geprüft, Änderung fertig, Umwandlung fertig) ohne
  Rückfrage und pusht nach `main` – PPTX, Spec, Folienbeschreibung und neue Bilder, ein Commit je Schritt. Bisher nur auf
  Anweisung. Der Abschluss nennt den Commit.

## 2.3.0 – 2026-10-28
Anforderung des Nutzers: Zu jeder Präsentation gehört ein Dokument, das den Inhalt jeder Folie beschreibt, damit andere
Agenten und Sitzungen das Deck lesen, ohne die PPTX auszulesen; beim Überarbeiten ändert bi-ppt nur, was sich wirklich
ändert, statt die ganze Präsentation neu zu bauen.
- Folienbeschreibung: `bippt bauen` schreibt neben die PPTX `<ausgabe>.md` mit Typ, Gestaltung (Layout, Aufbau, Band,
  Icons, Bilder), Kernaussage, Inhalt (Ampeln, Termine, Tabellen, Diagrammdaten) und Notizen je Folie. Neue Folien aus der
  Spec, unveränderte Basisfolien aus der Folienbeschreibung der Basis, sonst aus dem Deck. Neue Felder: `zweck`,
  `quellen`, `dokument` (Kopf), `beschreibung` (je Folie).
- Vorlage: `"dokument": "<pfad>.md"` macht eine vorhandene Beschreibung (z. B. von ppp) zur Folienbeschreibung; ihr Text
  bleibt, bippt pflegt nur seinen Abschnitt zwischen den Marken.
- Baustand: Fingerabdruck je Folie in der Folienbeschreibung. `bippt bauen` meldet `ÄNDERUNGEN` (welche Folien sich
  gegenüber dem letzten Bau geändert haben) und bricht ab, wenn das Deck nach dem Bau geändert wurde (z. B. in
  PowerPoint) oder ohne Baustand schon existiert; `--erzwingen` überschreibt. Öffnen und Speichern in PowerPoint lässt den
  Fingerabdruck gleich (geprüft mit PowerPoint für Mac).
- Basisfolien behalten ihre Sichtbarkeit (bisher wurden ausgeblendete Basisfolien ohne `ausblenden` wieder eingeblendet);
  `"ausblenden": false` blendet ein. `basis` gleich `ausgabe` bricht ab.
- Neu: `bippt beschreiben <pptx>` (Deck ohne Spec oder nach Handänderungen), `bippt analyse --folien 3,7`.
- Skill und Agenten: Modus Ä (einzelne Folien ändern, §3a); bestehende Decks zuerst über die Folienbeschreibung lesen, das
  Deck nur für einzelne Folien; Abschluss nennt Folienbeschreibung und geänderte Folien.
- Tests: Folienbeschreibung, Baustand, Handänderung, Basis-Modus auf ein geändertes Deck, Vorlage, `--folien`.

## 2.2.0 – 2026-10-27
Befunde aus brrp (26./27.10.2026): `bippt` brach an einer fremden Präsentation ab, und bi-ppt musste bestehende Folien mit
einem eigenen python-pptx-Skript nacharbeiten.
- Fehler: `analyse`, `pruefen`, `umwandeln entwurf` und der Basis-Modus brachen mit „Shape instance of unrecognized shape
  type“ ab (Form ohne Geometrie, z. B. ein früherer Platzhalter). Das neue Modul `formen.py` geht Formen sicher durch:
  Formen ohne Geometrie, Freihand (`p:contentPart`), SmartArt und `mc:AlternateContent` werden gemeldet (`analyse`:
  „Unbekannt“, `pruefen`: HINWEIS); ihr Text wird gelesen, sonst werden sie übersprungen.
- Fehler: `pruefen` meldete Akzentgrün-Text auf dunklem Folienhintergrund (`p:bg` von Folie, Layout oder Master) als FEHLER.
- Fehler: `umwandeln entwurf` brach an verknüpften Bildern ab und legte `<Quelle>_medien/` auch mit `-o` neben die Quelle;
  jetzt neben die Spec und nur, wenn es Bilder gibt.
- Basisfolien: `aendern` spricht Formen über Name (`nr` bei doppelten Namen), `id` oder `formen` (mehrere, auch Gruppen) an
  und tauscht Bilder in Bildform oder Bildplatzhalter (Rahmen bleibt; Zuschnitt wie bisher, `oben`/`mitte`/`unten`/`links`/
  `rechts` oder `{"oben_px": n}`; eine SVG-Fassung des alten Bildes entfällt), setzt, ergänzt oder entfernt Absätze, setzt
  Läufe einzeln, verschiebt (`x`, `y`, `w`, `h`, `dx`, `dy`), benennt um und entfernt. `elemente` legt neue Formen wie im
  Typ `frei` auf die Basisfolie.
- Gemischte Formatierung bleibt: `ersetzen` ändert nur die getroffenen Läufe; gesetzter Text mit `**fett**` übernimmt das
  Format eines fetten bzw. normalen Laufs im Absatz. `bauen` nennt die nicht gefundenen Ersetzungen.
- `bippt analyse --formen` listet je Form id, Name, Art, Lage und Absätze mit Läufen.
- `frei`: Formen `band` (mit `y`, `h`, z. B. helles Band als Untertitel) und `bild` (eingepasst).
- Bildfeld: Block `bild` in Karten, `bild` in Kennzahlen-Kacheln. `prozess`: `groesse` und `kopf_groesse` (Schrittkarten).
- Bildpfade (`bild_text`, Bildfelder, `aendern`) relativ zur Spec, sonst wie bisher zum Arbeitsordner.
- Skill, Agenten und Referenzen: bestehende Folien über `aendern` nacharbeiten, kein eigenes python-pptx-Skript.
- Tests: Agenda-Fall (nachgebaut, ohne die Originaldatei), Änderungsmodus, Bildfelder, Schriftgröße der Schrittkarten.

## 2.1.0 – 2026-10-25
Nachbau-Abgleich mit dem CPV-Projektteam-Update (24.09.2026) und den How-to slides des BI-Masters v2.1. Der Master ist
verbindlich (Nutzer): nur Theme- und Custom Colors, keine Abstufungen, keine Standardfarben.
- Fehler: In der Prozesskette war der Titel eines erledigten Schritts unsichtbar (dunkle Schrift auf Dunkelgrün).
- `prozess` mit Schrittkarten: Kopf in Statusfarbe mit Überzeile (`ueber`, z. B. Termin) und Titel, darunter `punkte`,
  Pfeile zwischen den Köpfen; Kopfhöhe nach dem längsten Titel.
- Bänder: `stil` dunkel/akzent/hell, Höhe nach Text (bis zwei Zeilen), Label fett, Text normal; Warnung ab drei Zeilen.
- Karten: `hervor` (Kopfleiste Accent Green, dunkelgrüne Schrift), `hoehe` inhalt (Standard bei `karten`) oder voll.
- Neuer Typ `fluss` (Datenfluss in Spalten mit Pfeilen, eine Box kann über alle Zeilen spannen).
- Timeline: Quartalsskala (automatisch ab 18 Monaten), Beschriftungen nie über den Folienrand, Zeilen ab 0,32 in.
- Ampel `YR` (Gelb–Rot) als halb gelber, halb roter Punkt; auch in Legende und Vergleich.
- Agenda ohne Titel bekommt „Agenda“; `massnahmen` bis ca. 8 Zeilen.
- Diagramme in der automatischen Farbfolge des Masters (Akzent 1–6).
- `pruefen`: Abstufungen von Designfarben (lumMod/lumOff/tint/shade außerhalb von Formstilen) sind FEHLER, abgestuftes
  Weiß/Schwarz ein Hinweis; `--woerter N` (Steering 120, Projektteam 200, Standard 170).
- `farben_angleichen` gleicht auch Abstufungen in Basisfolien an die nächste Markenfarbe an.
- `bippt umgebung` prüft Python ≥ 3.9; Skill, Agent und Installationshilfe nennen den Python-Aufruf wie ppp (Windows
  `python`, Mac `python3`), aber ohne `-I`, weil `-I` die mit `pip --user` installierten Pakete ausblendet.
- Arbeitsweise: Regeln des Masters, Dichte je Publikum, Storyline Projektteam-Update, Textregeln (Aktionstitel, ganze
  Aussagen, Fakt – Konsequenz, Varianten gleich gliedern, feste Band-Labels, Termine). Neues anonymisiertes Muster
  `beispiele/projektteam.spec.json`, zugleich Testfall.
- bi-designer: Master-Regeln in MARKE.md, Office-Rezept und Regeln (keine Abstufungen, Diagrammfolge Akzent 1–6).

## 2.0.0 – 2026-10-25
Aus dem Repo BI-PPT wird das Plugin **bi-design** im privaten Marktplatz `portolasas/bi-design` (Entscheidung des Nutzers).
- Ein Plugin für VS Code und Claude Code (Aufbau A wie ppp: `plugin.json` für VS Code, `.claude-plugin/plugin.json` für
  Claude Code): Agent `bi-ppt`, Skills `bi-ppt` und `bi-designer`, Generator `bippt`, Vorlage BI-Master v2.1.
- Skill **bi-designer** aus PPP 0.17 übernommen (Stand 08.10.); Verweise auf PPP 0.17 durch ppp und `~/.bi-design/config.json`
  ersetzt. Er ist Skill, kein eigener Agent mehr; ppp und bi-ppt nutzen ihn.
- Installer, `/bi-ppt-update`, Starter `~/.bi-ppt/bippt` und venv entfallen. bippt läuft aus der Kopie des Plugins mit dem
  System-Python; `bippt umgebung` nennt den Aufruf, `bippt einrichten` installiert die Pakete mit `pip --user`.
- `bippt aufraeumen [--ja]` zeigt bzw. löscht alte Kopien von BI-PPT 1.x und PPP 0.17 in `~/.copilot`, `~/.claude` und
  `~/.bi-ppt` (nur mit Merkmal von BI-PPT oder BI-Designer) und übernimmt `~/.bi-ppt/config.json`; `bippt umgebung` meldet sie.
- Einstellungen in `~/.bi-design/config.json`; Marke aus dem Skill im Plugin; Vorlage aus dem Plugin, eigene über `konfig`.
- Tests: `tests/test_plugin.py` (Aufbau, Version, Aufräumen) neben `tests/test_bippt.py`.

## 1.0.1 – 2026-10-08
- Rendern (Mac): PowerPoint wird nach dem Export zuverlässig beendet (quit direkt nach close ging verloren;
  jetzt bis zu 3 Versuche). Sonst lief der nächste Export in derselben Sitzung (verstümmelte Titel-Schrift).
- Basisfolien: `farben_angleichen` / `basis_farben_angleichen` setzt Nicht-Markenfarben alter Folien auf die nächste
  Markenfarbe. Timeline-Beschriftungen dürfen Heute-/Markerlinien notfalls kreuzen (kein Fehlalarm), Legende ohne
  Randüberstand, Höhenschätzung am Render kalibriert. `bippt typen` mit Kurzbeschreibung aller Typen.

## 1.0.0 – 2026-10-08
Erste Fassung, abgeleitet aus dem CPV-Steering-Committee-Update (Copilot-Chat 07.10.2026) und dem BI-Master v2.1.
- Agent `bi-ppt` für VS Code (Copilot) und Claude Code (Skill `/bi-ppt`, Subagent `bi-ppt`), Befehl `/bi-ppt-update`.
- Arbeitsweise: Modi Neu / Neue Version / Umwandeln / Prüfen, eine gebündelte Rückfragerunde, Storyline nach
  Publikum, Aktionstitel, Kernaussage-Band, Backup ausgeblendet.
- Generator `bippt` (Python): `umgebung`, `konfig`, `lesen` (vtt/docx/pptx/pdf/xlsx), `analyse`, `typen`, `icons`,
  `bauen`, `pruefen`, `rendern`, `umwandeln entwurf|stil`, `vorlage`, `neu`.
- Folientypen: titel, agenda, kapitel, text, status, vergleich, prozess, karten (Blöcke punkte/text/phasen/schritte/
  von_nach/kennzahl/hinweis), timeline (Monate/KW, Heute, Marker, Schatten, Legende, Kollisionsprüfung der
  Beschriftungen), massnahmen, tabelle (echt), kennzahlen, diagramm (echt), bild_text, kernaussage, abschluss, frei.
- Basis-Modus: Folien einer bestehenden PPTX behalten/ausblenden/ersetzen/umordnen, Rest entfällt.
- Marke modular aus PPP-Skill bi-designer; Custom Colors des Masters (Gray 1–8, Signalfarben) aus der Vorlage; Ampel
  = Signalfarben; Fußzeile „Titel | Vertraulichkeit“ wie im Master, Titelfolie nur Vermerk.
- Vorlage: `vorlagen/How-to slides.pptx` (Original) → `BI_Vorlage_16x9.pptx` (ohne Folien, ohne doppelte Layouts).
- Rendern: Mac über PowerPoint im fest freigegebenen Ordner `~/Documents/BI-PPT-Render` (keine Zugriffsdialoge mehr
  je Datei), frischer PowerPoint-Start je Export (sonst verstümmelte Titel-Schrift ab dem 2. PDF-Export);
  Windows über PowerPoint-COM; LibreOffice als Ausweichweg.
