# Changelog

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
