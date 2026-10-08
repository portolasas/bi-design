# BI-PPT installieren und aktualisieren

Voraussetzungen: VS Code mit GitHub Copilot und/oder Claude Code; **Python 3.9+** für den Generator;
PowerPoint (Mac/Windows) für die Sichtprüfung (optional). Für Farben/Icons den PPP-Skill **bi-designer**
(`/ppp-update`) – fehlt er, laufen eingebaute Markenfarben.

## macOS
```bash
git clone https://github.com/portolasas/BI-PPT.git   # oder vorhandenen Ordner nutzen
python3 BI-PPT/install/install.py
```
Einmalig beim ersten Rendern: PowerPoint fragt „Dateizugriff erteilen“ → „Auswählen…“ → Ordner
`Dokumente/BI-PPT-Render` markieren → „Zugriff erteilen“.

## Windows
```powershell
git clone https://github.com/portolasas/BI-PPT.git
py BI-PPT\install\install.py
```
Kein Python? Ohne Adminrechte: `winget install -e --id Python.Python.3.12 --scope user` (oder Microsoft Store
„Python 3.12“), neues Terminal, dann den Installer. Hinter einem Proxy vorher `py -m pip config set global.proxy http://<proxy>:<port>`.

Nur Agent/Skills ohne Python installieren (Generator fehlt dann; der Agent kann Spec und Inhalte vorbereiten,
gebaut wird auf einem Rechner mit Python) – in PowerShell, `$Repo` anpassen:
```powershell
& { $Repo = "$HOME\BI-PPT"
  New-Item -ItemType Directory -Force "$HOME\.copilot\agents","$HOME\.copilot\skills" | Out-Null
  Copy-Item "$Repo\agents\bi-ppt.agent.md" "$HOME\.copilot\agents\" -Force
  foreach ($s in 'bi-ppt','bi-ppt-update') { Remove-Item "$HOME\.copilot\skills\$s" -Recurse -Force -ErrorAction SilentlyContinue
    Copy-Item "$Repo\skills\$s" "$HOME\.copilot\skills\$s" -Recurse -Force }
  "BI-PPT Agent und Skills installiert (ohne Generator)" }
```

## Danach
- VS Code: Strg/Cmd+Umschalt+P › „Developer: Reload Window“, neuer Chat, Agent **bi-ppt** wählen.
- Claude Code: neue Sitzung, `/bi-ppt` (Skill) – oder Subagent `bi-ppt` für abgegrenzte Bauaufträge.
- Prüfen: `~/.bi-ppt/bippt umgebung` (Windows: `& "$HOME\.bi-ppt\bippt.cmd" umgebung`) → `ERGEBNIS: OK`.

## Aktualisieren
Im Chat `/bi-ppt-update` – oder `python3 <repo>/install/install.py --update` (holt per `git pull` und spielt ein).

## Anweisung für einen Agenten („Installiere BI-PPT nach INSTALL.md“)
1. Python prüfen: macOS `python3 --version`, Windows `py --version` (sonst `python --version`). < 3.9 oder fehlt →
   Windows-Hinweis oben nennen und, nur wenn der Nutzer zustimmt, den PowerShell-Block „ohne Python“ ausführen.
2. `install/install.py` dieses Repos mit diesem Python ausführen, Ausgabe unverändert zeigen.
3. Erfolg = `BI-PPT <Version> installiert` und `ERGEBNIS: OK`. Dem Nutzer „Danach“ nennen. Nichts weiter ändern.
