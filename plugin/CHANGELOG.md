# Changelog

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
