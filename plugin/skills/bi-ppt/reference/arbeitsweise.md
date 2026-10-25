# Arbeitsweise – Storyline und Gestaltung

Abgeleitet aus dem CPV-Steering-Committee-Update (07.10.2026, vom Nutzer abgenommen) und den How-to slides des BI-Masters.

## Publikum → Umfang
| Publikum | Sichtbare Folien | Schwerpunkt | Backup |
|---|---|---|---|
| Steering Committee / Management | 6–9 | Status, Änderungen, strategische Themen, Timelines, Entscheidungsbedarf, nächste Schritte | Detailfolien ausgeblendet |
| Projektteam | 8–14 | Arbeitsstand je Strang, Aufgaben, Abhängigkeiten, Details | selten |
| Sounding Board / Fachgremium | 5–8 | Lösungsidee, Varianten, Bitte um Feedback | Varianten-Details |
| Onboarding / Info | frei | Erklären, Kontext, Kontakte | – |

## Standard-Storyline Status-Update (Steering)
1. `titel` – Projekt, Anlass, Autor | Datum
2. `status` – Status auf einen Blick (Stränge, Ampel, Phasen, nächster Schritt) + Band mit Gesamtaussage
3. `vergleich` – Was hat sich seit dem letzten Update geändert? (nur bei Folge-Updates)
4. 1–2 Strang-Folien (`prozess`, `karten`) – nur die strategisch relevanten Themen
5. `timeline` (Monate, Gesamtüberblick) + bei kritischem Pfad `timeline` (Kalenderwochen)
6. `massnahmen` – Nächste Schritte und Termine + Band „Kritischer Pfad“
7. Entscheidungsbedarf nur, wenn wirklich nötig (`kernaussage` oder Band) – sonst weglassen
8. `kapitel` „Backup“ mit `ab_hier_ausblenden` + aktualisierte Detailfolien der Basis

## Folien-Regeln
- **Aktionstitel:** Titel = Aussage, nicht Thema („Strang A – Skript-Lösung statt Zwischendatenbank“, nicht „Strang A“).
- **Eine Kernaussage je Folie** als `band` {label, text}; Label kurz (2–4 Wörter) mit Doppelpunkt oder Gedankenstrich.
- **Grafik vor Text:** Abläufe → `prozess`, Termine → `timeline`, alt/neu → `vergleich`, Zahlen → `kennzahlen`/`diagramm`.
- **Nicht überladen:** ≤ 5 Punkte je Karte, ≤ 2 Zeilen je Punkt, ≤ ~120 Wörter je Folie; Rest ins Backup.
- **Ampeln:** G = im Plan, Y = beobachten/Freigabe ausstehend, R = Eskalation. Immer mit Legende (status macht das).
- **Termine:** TT.MM.JJJJ, Monate „10/2026“, Kalenderwochen ISO („KW 43“); Heute-Linie in Timelines.
- **Namen:** Initial + Nachname („T. Schmitt“); keine Personendaten über das Nötige hinaus.
- **Änderungen ggü. Vorversion** immer explizit zeigen (Vergleich oder Band), nicht nur stillschweigend aktualisieren.

## Neue Version einer bestehenden Präsentation (Modus V)
1. `bippt analyse alt.pptx` → Folienliste mit Nummern.
2. Je Folie entscheiden: behalten (ggf. `ersetzen`), ausblenden (Backup), entfernen, oder durch neuen Typ ersetzen.
3. Veraltete Aussagen in behaltenen Folien über `ersetzen` korrigieren (exakter Textausschnitt aus `analyse`).
   `bippt bauen` warnt, wenn eine Ersetzung nicht gefunden wurde.
4. Titelfolie: `{"basis": 1, "platzhalter": {"1": ["Neuer Anlass", "Name | Datum"]}}`.

## Effizienz
- Quellen einmal lesen (`bippt lesen -o _quellen.md`), danach nur gezielt nachschlagen.
- Erst komplette Spec, dann bauen – nicht Folie für Folie rendern.
- Render: Kontaktbogen (80 dpi) reicht für die Übersicht; Einzelfolien nur bei Auffälligkeiten (`--folien 3,6 --dpi 110`).
- Korrekturschleifen ≤ 3; was danach offen ist, im Abschluss nennen.
