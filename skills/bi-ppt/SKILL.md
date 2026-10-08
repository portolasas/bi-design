---
name: bi-ppt
description: BI-PPT – erstellt PowerPoint-Präsentationen im Boehringer-Ingelheim-Design (offizieller BI-Master v2.1) aus gegebenen Informationen (Transkripte, Projektaufträge, Notizen, Excel), erzeugt neue angepasste Versionen bestehender Präsentationen (z. B. Projektteam-Update → Steering-Committee-Update) und wandelt Präsentationen in fremdem Design ins BI-Design um. Arbeitet mit Spezifikation (JSON) + Generator bippt (Folientypen Status, Vergleich, Prozess, Timeline/Gantt, Maßnahmen, Kennzahlen, Diagramm, Tabelle …), prüft Brand-Regeln und rendert zur Sichtprüfung. Laden bei jeder Anfrage zu Folien, Präsentation, Deck, PowerPoint, PPT, Steering-/Projekt-Update, „im BI-Design“ oder „Folien umwandeln“.
---

# BI-PPT – Präsentationen im BI-Design

Werkzeug: `bippt` (Python, python-pptx), installiert unter `~/.bi-ppt`.
- macOS: `~/.bi-ppt/bippt <befehl>` · Windows (PowerShell): `& "$HOME\.bi-ppt\bippt.cmd" <befehl>`
- Marke (Farben, Icons): Skill **bi-designer** aus PPP (`marken/<profil>/tokens.json`, `icons.txt`); Custom Colors des
  BI-Masters (Gray 1–8, Signal Red/Yellow/Green) liest bippt aus der Vorlage. Nie Farben frei erfinden.
- Vorlage: `~/.bi-ppt/vorlagen/BI_Vorlage_16x9.pptx` (BI-Master v2.1, aus `How-to slides.pptx` des Repos).

Antworte deutsch, sachlich, knapp. Grundsatz: **teuer denken, billig bauen** – Inhalt und Storyline gründlich,
Umsetzung immer über Spezifikation + `bippt bauen` (nie Formen einzeln per Hand-Skript), Änderungen = Spec ändern.

## 0. Start (jede Sitzung, 1 Befehl)
`bippt umgebung` → muss mit `ERGEBNIS: OK` enden. Sonst: [reference/installation.md](reference/installation.md).

## 1. Modus erkennen
| Modus | Auslöser | Ablauf |
|---|---|---|
| **N – Neu** | Infos/Quellen, keine Basis-PPT | §2 → §3 → §4 |
| **V – Neue Version** | bestehende PPT als Grundlage („auf Basis von …“, Update für anderes Gremium) | §2 (+ `bippt analyse`) → §3 mit `basis` → §4 |
| **U – Umwandeln** | PPT in fremdem/altem Design → BI-Design | [reference/umwandeln.md](reference/umwandeln.md) |
| **P – Prüfen** | „passt das zum Brand?“ | `bippt pruefen` + `bippt rendern` → Befunde mit Folie, Regel, Fix |

## 2. Verstehen (token-sparsam)
1. Quellen nur über `bippt lesen <dateien…> -o <arbeitsordner>/_quellen.md` (VTT, DOCX, PPTX, PDF, XLSX) und
   die Datei dann gezielt lesen. Basis-PPT: `bippt analyse <pptx>` (Text je Folie, Layout, ausgeblendet).
2. Ordner, die der Nutzer ausschließt (z. B. `Archiv/`), nie öffnen. Aktuellste Quelle hat Vorrang
   (z. B. Monolog/Transkript vor Projektplan vor alter Präsentation) – Widersprüche nicht still auflösen.
3. **Eine gebündelte Rückfragerunde** (Tool ask_user / AskUserQuestion, je Frage Empfehlung zuerst) nach
   [reference/fragen.md](reference/fragen.md): nur was Inhalt oder Gestaltung wirklich ändert – Datum, Publikum/Zweck,
   Status/Ampeln, Konflikte zwischen Quellen (mit Zitat beider Stellen), Termine, Entscheidungsbedarf, Umfang.
   Alles andere mit Standard entscheiden und im Abschluss nennen. Bei Konflikten oder Unsicherheiten immer fragen.

## 3. Storyline und Spezifikation
1. Storyline nach [reference/arbeitsweise.md](reference/arbeitsweise.md) (Publikum → Folienzahl, Aktionstitel,
   eine Kernaussage je Folie als Band, Details ins ausgeblendete Backup).
2. Spezifikation `<Name>.spec.json` **neben die Ausgabe** schreiben (Startgerüst: `bippt neu <ordner> --name <Name>`).
   Folientypen und Felder: [reference/folientypen.md](reference/folientypen.md) (`bippt typen`). Icons: `bippt icons <begriff>`.
3. Modus V: `"basis": "<alte.pptx>"`; jede Folie entweder neu (`"typ"`) oder `{"basis": n, "ausblenden"?, "ersetzen"?,
   "platzhalter"?}`. Nicht aufgeführte Basisfolien entfallen. Titel-Folie der Basis über `platzhalter` aktualisieren.

## 4. Bauen – Prüfen – Sehen (max. 3 Runden)
```
bippt bauen <Name>.spec.json          → ERGEBNIS: OK | Warnungen (Überlauf, Icons, Timeline-Beschriftung)
bippt pruefen <Name>.pptx             → ERGEBNIS: OK (keine FEHLER)
bippt rendern <Name>.pptx [--alle]    → kontaktbogen.png ansehen, nur auffällige Folien einzeln (folie_NN.png)
```
Sichtprüfung je Folie: nichts überlappt/abgeschnitten, Text lesbar (≥ 9 pt im Inhalt), Kernaussage stimmt,
Ampeln/Termine = Quellen, keine leeren Flächen > ⅓ Folie. Fixes immer in der Spec, dann neu bauen.
Rendern auf dem Mac: PowerPoint muss ohne offene eigene Präsentationen sein (bippt startet es frisch und beendet es);
der Ordner `~/Documents/BI-PPT-Render` ist einmalig in PowerPoint freigegeben – nie andere Ordner an PowerPoint geben.
Kein Renderweg? Dann nur `pruefen` und das im Abschluss sagen.

## 5. Abschluss
- Ausgabe neben die Quellen (oder wohin der Nutzer sagt); Dateiname `<Projekt>_<Anlass>_<TT.MM.JJJJ>.pptx`;
  Spec daneben lassen (spätere Änderungen = Spec + `bippt bauen`). `_render_*`-Ordner und `_quellen.md` löschen,
  wenn der Nutzer sie nicht braucht.
- Git-Repo und Nutzer will veröffentlichen: committen und direkt nach `main` pushen (nur auf Anweisung).
- Antwort: Pfad · Folienliste (sichtbar/Backup) mit je einer Zeile Inhalt · getroffene Annahmen · offene Punkte ·
  Prüfergebnis (`pruefen`, Render ja/nein).

## 6. Regeln (BI-Master v2.1 + bi-designer)
- Nur Layouts der Vorlage, nur Theme- und Custom Colors; Ampel = Signal Green/Yellow/Red; Akzentgrün nie als Text/
  dünne Linie auf Hell (nur auf Dunkelgrün oder als Fläche hinter dunkelgrünem Text).
- Schriften = Theme (Boehringer Forward Head/Text); Textebenen 1–4 16 pt, ab Ebene 5 12 pt im Fließtext-Layout.
- Fußzeile: `<Präsentationstitel>     |     <Vertraulichkeit>` (Standard „ONLY FOR INTERNAL USE“); Titelfolie nur Vermerk.
- Icons nur aus dem Brand-Set, Variante nach Hintergrund. Logo kommt aus dem Master – nie selbst einfügen.
- Nicht im Paket definiert (Bildsprache, Freiraum) → „(offen, Brand Hub prüfen)“ statt raten.

## 7. Erweitern
Neuer Folientyp: Funktion mit `@folientyp("name")` in `tools/bippt/folien.py` (Bausteine aus `bausteine.py`), Doku-Zeile in
`reference/folientypen.md`, Beispiel in `beispiele/beispiel.spec.json`, Version + CHANGELOG im Repo BI-PPT, Commit direkt
nach `main`, dann `/bi-ppt-update`. Neue Marke: Profil in bi-designer + eigene Vorlage (`bippt vorlage <deck> -o …`).
