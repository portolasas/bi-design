# Installation, Update, Fehlerbilder

Quelle: Repo **BI-PPT** (`portolasas/BI-PPT`, lokal z. B. `…/Arbeit/BI-PPT`). Installiert wird nach
`~/.bi-ppt` (Werkzeug, venv, Vorlage), `~/.copilot` (VS Code: Agent + Skills) und `~/.claude` (Claude Code: Skills + Agent).

## Installieren / aktualisieren
- macOS: `python3 <repo>/install/install.py` (Update: `/bi-ppt-update` oder `python3 <repo>/install/install.py --update`)
- Windows: `py <repo>\install\install.py` (oder `python …`). Ohne Python: siehe INSTALL.md im Repo (PowerShell-Block
  installiert Agent und Skills; Generator braucht Python 3.9+ – `winget install -e --id Python.Python.3.12 --scope user`).
Danach VS Code neu laden (Developer: Reload Window) bzw. neue Claude-Code-Sitzung.

## `bippt umgebung` meldet …
| Meldung | Abhilfe |
|---|---|
| Paket pptx FEHLT | Installer erneut (`--nur-venv`); Proxy: `pip` mit `--proxy` bzw. Firmen-Index |
| Vorlage FEHLT | Installer erneut, oder `bippt vorlage <BI-Deck.pptx>` (jedes Deck im BI-Master taugt) |
| Marke … Notstand | PPP installieren (`/ppp-update`) – bi-designer fehlt; Farben laufen trotzdem (eingebaut) |
| Icons 0 | `bippt konfig --setze asset_root=<BrandDesign-Ordner>` (Ordner mit Icons_darkgreen/ …) |
| Rendern über nichts | PowerPoint fehlt → nur `pruefen`; optional LibreOffice |

## Rendern – bekannte Fallen
- **Mac, Dialog „Dateizugriff erteilen“:** PowerPoint ist sandboxed. Einmalig „Auswählen…“ → Ordner
  `Dokumente/BI-PPT-Render` (den Ordner, nicht eine Datei) → „Zugriff erteilen“. bippt gibt PowerPoint nie andere Ordner.
- **Mac, Buchstaben fehlen in Titeln im PDF:** PowerPoint-Fehler ab dem 2. PDF-Export einer Sitzung. bippt startet
  PowerPoint deshalb je Export neu; sind eigene Präsentationen offen, bricht es ab → schließen und erneut.
- **Mac, Fehler -9074 / Zeitüberschreitung:** Freigabe fehlt oder Dialog offen → in PowerPoint nachsehen.
- **Windows:** PowerPoint per COM, keine Dialoge; läuft PowerPoint mit offenen Dateien, bleiben sie unberührt.
