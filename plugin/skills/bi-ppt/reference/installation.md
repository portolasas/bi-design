# Installation, Update, Fehlerbilder

Quelle: Plugin **bi-design** aus dem privaten Marktplatz `portolasas/bi-design`. Es bringt Agent `bi-ppt`, die Skills
`bi-ppt` und `bi-designer`, den Generator `bippt` und die Vorlage mit. Einstellungen je Rechner: `~/.bi-design/config.json`.

## Installieren
- **VS Code:** in den Benutzereinstellungen (JSON) `"chat.plugins.enabled": true` und
  `"chat.plugins.marketplaces": [..., "portolasas/bi-design"]`; in der Erweiterungen-Ansicht unter Agent Plugins
  **bi-design** installieren; „Developer: Reload Window“; Agent **bi-ppt** im Agenten-Menü wählen.
- **Claude Code:** `/plugin marketplace add portolasas/bi-design`, `/plugin install bi-design@bi-design`, neue Sitzung.
  Skill `/bi-design:bi-ppt` im Hauptchat oder Subagent `bi-ppt` für abgegrenzte Bauaufträge.
- Dann `bippt umgebung`; Pakete mit `bippt einrichten`, alte Kopien mit `bippt aufraeumen --ja`.

## Aktualisieren
VS Code holt neue Versionen selbst (bei `extensions.autoUpdate` etwa alle 24 Stunden), danach „Developer: Reload Window“.
Claude Code: `/plugin marketplace update bi-design`, dann `/plugin uninstall bi-design@bi-design` und
`/plugin install bi-design@bi-design`, neue Sitzung.

## `bippt umgebung` meldet …
| Meldung | Abhilfe |
|---|---|
| Paket pptx/lxml FEHLT | `bippt einrichten` (nach Zustimmung). Proxy: `pip config set global.proxy http://<proxy>:<port>`. Kein venv anlegen: Gruppenrichtlinien sperren es oft |
| Python fehlt oder < 3.9 | Windows ohne Adminrechte: `winget install -e --id Python.Python.3.12 --scope user`, neues Terminal |
| ALT … | `bippt aufraeumen --ja` (übernimmt `~/.bi-ppt/config.json`), dann VS Code neu laden bzw. neue Sitzung |
| Vorlage FEHLT | Plugin neu installieren, oder `bippt vorlage <BI-Deck.pptx>` (jedes Deck im BI-Master taugt) |
| Marke … Notstand | Profil fehlt in `skills/bi-designer/marken/`; Farben laufen trotzdem (eingebaut) |
| Icons 0 | `bippt konfig --setze asset_root=<BrandDesign-Ordner>` (Ordner mit Icons_darkgreen/ …) |
| Rendern über nichts | PowerPoint fehlt → nur `pruefen`; optional LibreOffice |

## Rendern – bekannte Fallen
- **Mac, Dialog „Dateizugriff erteilen“:** PowerPoint ist sandboxed. Einmalig „Auswählen…“ → Ordner
  `Dokumente/BI-PPT-Render` (den Ordner, nicht eine Datei) → „Zugriff erteilen“. bippt gibt PowerPoint nie andere Ordner.
- **Mac, Buchstaben fehlen in Titeln im PDF:** PowerPoint-Fehler ab dem 2. PDF-Export einer Sitzung. bippt startet
  PowerPoint deshalb je Export neu; sind eigene Präsentationen offen, bricht es ab → schließen und erneut.
- **Mac, Fehler -9074 / Zeitüberschreitung:** Freigabe fehlt oder Dialog offen → in PowerPoint nachsehen.
- **Windows:** PowerPoint per COM, keine Dialoge; läuft PowerPoint mit offenen Dateien, bleiben sie unberührt.
