# BI-PPT

Agent für **PowerPoint im Boehringer-Ingelheim-Design** – für VS Code (GitHub Copilot) und Claude Code, macOS und Windows.

- **Neu:** aus Informationen (Transkripte, Projektaufträge, Notizen, Excel) eine Präsentation im BI-Master v2.1 bauen,
  mit gezielten Rückfragen.
- **Neue Version:** eine bestehende Präsentation als Grundlage nehmen (z. B. Projektteam-Update → Steering-Update):
  Folien behalten, ausblenden, Texte ersetzen, neue Folien ergänzen.
- **Umwandeln:** Präsentationen in fremdem Design ins BI-Design bringen (neu bauen oder Stil angleichen).
- **Prüfen:** Brand-Regeln (Farben, Schriften, Kontrast, Überlauf, Master) und Sichtprüfung per Render.

Arbeitsweise: Der Agent schreibt eine **Spezifikation (JSON)**, der Generator **bippt** baut daraus die Folien aus
Bausteinen (Status, Vergleich, Prozess, Karten, Timeline/Gantt, Maßnahmen, Kennzahlen, Diagramm, Tabelle …).
Änderungen später = Spec ändern + neu bauen. Marke kommt modular aus dem PPP-Skill **bi-designer**
(`tokens.json`, Icons), die Folienvorlage aus dem offiziellen Master (`vorlagen/How-to slides.pptx`).

## Inhalt
| Ordner | Zweck | Installiert nach |
|---|---|---|
| `agents/bi-ppt.agent.md` | Copilot-Agent (Agent-Picker in VS Code) | `~/.copilot/agents/` |
| `skills/bi-ppt/` | Arbeitsweise, Rückfragen, Folientypen, Umwandeln, Beispiel-Spec | `~/.copilot/skills/`, `~/.claude/skills/` |
| `skills/bi-ppt-update/` | Befehl `/bi-ppt-update` | wie oben |
| `claude/agents/bi-ppt.md` | Claude-Code-Subagent | `~/.claude/agents/` |
| `tools/` | Generator/CLI `bippt` (Python, python-pptx) | `~/.bi-ppt/tools/` |
| `vorlagen/` | BI-Master v2.1 (Original) + bereinigte Arbeitsvorlage | `~/.bi-ppt/vorlagen/` |
| `install/install.py` | Installer/Updater (Mac + Windows) | – |
| `tests/` | Selbsttest (`python3 tests/test_bippt.py`) | – |

Installation: [INSTALL.md](INSTALL.md). Änderungen: [CHANGELOG.md](CHANGELOG.md).

## Regeln für Änderungen am Repo
- Immer direkt nach `main` committen und pushen; jede Änderung erhöht `VERSION` und bekommt einen CHANGELOG-Eintrag.
- Neuer Folientyp: `tools/bippt/folien.py` (`@folientyp`), Doku in `skills/bi-ppt/reference/folientypen.md`,
  Beispiel in `skills/bi-ppt/beispiele/beispiel.spec.json`, `tests/test_bippt.py` muss grün sein.
- Keine Projekt- oder Personendaten, keine Schriftdateien der Marke ins Repo.
