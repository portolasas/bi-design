---
name: bi-ppt-update
description: BI-PPT aktualisieren – holt den neuesten Stand aus dem Repo BI-PPT (git pull) und installiert Agent, Skills, Werkzeug und Vorlage neu für VS Code (Copilot) und Claude Code. Aufruf /bi-ppt-update.
---

# BI-PPT aktualisieren

1. Quelle aus `~/.bi-ppt/config.json` (Feld `quelle`) lesen – das ist der Repo-Ordner BI-PPT.
2. Ausführen (ein Befehl, Ausgabe unverändert zeigen):
   - macOS: `python3 "<quelle>/install/install.py" --update`
   - Windows (PowerShell): `py "<quelle>\install\install.py" --update`
3. Erfolg: Zeile `BI-PPT <Version> installiert` und `bippt umgebung … ERGEBNIS: OK`.
4. Dem Nutzer sagen: VS Code neu laden (Developer: Reload Window) bzw. neue Claude-Code-Sitzung.
Fehlt `quelle`: nach dem Repo-Ordner fragen. Keine weiteren Dateien ändern.
