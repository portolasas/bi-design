# Spezifikation und Folientypen

Vollständiges Beispiel mit allen Typen: [../beispiele/beispiel.spec.json](../beispiele/beispiel.spec.json) (`bippt neu`).

## Kopf der Spezifikation
| Feld | Bedeutung |
|---|---|
| `ausgabe` | Dateiname/Pfad der PPTX (relativ zur Spec) |
| `fusszeile` | Präsentationstitel in der Fußzeile; Vermerk wird angehängt: `Titel     \|     ONLY FOR INTERNAL USE` |
| `vertraulichkeit` | Vermerk (Standard aus config, „ONLY FOR INTERNAL USE“); `""` = keiner |
| `heute` | Datum der Heute-Linie in Timelines (Standard: heute) |
| `vorlage` | andere Vorlage (Standard: BI-Master v2.1) |
| `basis` | bestehende PPTX (Modus V) – dann ist sie Vorlage und Folienquelle |
| `marke` | bi-designer-Profil (Standard boehringer-ingelheim) |
| `folien` | Liste der Folien in Reihenfolge |

## Felder jeder neuen Folie
`typ`, `titel`, `band` (`"Text"` oder `{"label": "…:", "text": "…"}`), `quelle` (Quellenzeile unten), `notizen`,
`ausblenden`, `ab_hier_ausblenden` (z. B. am Backup-Kapitel), `dunkel` (dunkle Layoutvariante, wo vorhanden),
`layout` (Layout-Kurzname erzwingen), `fusszeile: false`.

Text überall: `"… **fett** … ==hervorgehoben== …"`, Liste = mehrere Absätze, oder
`{"t": "…", "fett": true, "groesse": 12, "farbe": "DG", "punkt": true, "ebene": 1}`.
Farben als Kürzel: DG, AG, WH, LG, WG, YEL, MYEL, LRED, RED, MRED, DRED, LBLU, BLU, MBLU, DBLU, LVIO, VIO, MVIO, DVIO,
SGRN, SYEL, SRED, G1…G8 – oder Token-Namen aus bi-designer. Ampelwerte: G, Y, R, N.

## Basisfolien (Modus V)
`{"basis": 7, "ausblenden": true, "ersetzen": {"alter Text": "neuer Text"}, "platzhalter": {"1": ["Zeile 1", "Zeile 2"]}, "titel": "…"}`
`"farben_angleichen": true` (je Folie) bzw. `"basis_farben_angleichen": true` (Kopf) setzt Nicht-Markenfarben alter
Folien auf die nächste Markenfarbe – sinnvoll, wenn `bippt pruefen` „Farben außerhalb der Marke“ in Basisfolien meldet.

## Typen
| typ | Layout | Felder | Einsatz |
|---|---|---|---|
| `titel` | Title slide (bright) / `dunkel` | `titel`, `untertitel` [Zeilen] | Titelfolie |
| `agenda` | Agenda | `punkte` | Agenda |
| `kapitel` | Chapter divider (1) / `dunkel` | `titel`, `untertitel` | Abschnitt, „Backup“ |
| `text` | One content | `punkte` (mit `ebene`) | Fließtext/Aufzählung im Master-Stil |
| `status` | Only title | `phasen` [Namen], `straenge` [{`titel`, `ampel`, `stand` [2=fertig,1=aktiv,0=offen], `punkte`, `naechster`}], `legende` {G: …} | Status je Strang |
| `vergleich` | Only title | `spalten` [3], `zeilen` [[thema, alt, neu]]; Zelle `{"ampel": [..], "labels": [..]}` | Änderungen alt → neu |
| `prozess` | Only title | `kopf` {label,text}, `schritte` [{`icon`, `titel`, `text`, `status`}], `hinweis`, `hinweis_ab`, `karten` | Ablauf/Prozesskette |
| `karten` | Only title | `kopf`, `karten` [{`titel`, `ampel` \| `akzent`+`icon`, `stil` leiste/akzent/ohne, `hell`, Blöcke}] | 2–4 Bereiche nebeneinander |
| `timeline` | Only title | `skala` monat/woche (auto), `von`, `bis`, `heute`, `zeilen` [{`titel`, `elemente`}], `marker`, `schatten`, `legende`, `fakten` \| `boxen` | Gantt/Roadmap |
| `massnahmen` | Only title | `zeilen` [{`wann`, `was`, `details`, `tag`}] | Nächste Schritte |
| `tabelle` | Only title | `kopf`, `zeilen`, `breiten`, `groesse`, `erste_spalte_fett`; Zelle `{"ampel": "G", "t": "…"}` | echte, editierbare Tabelle |
| `kennzahlen` | Only title | `kacheln` [{`wert`, `label`, `text`, `ampel`, `icon`}], `spalten` | KPIs |
| `diagramm` | Only title | `art` saeulen/gestapelt/balken/linie/torte, `kategorien`, `reihen` [{name, werte}], `werte_anzeigen`, `fazit` {titel, punkte, icon} | echtes Diagramm (editierbar) |
| `bild_text` | Content and image | `bild` (Pfad), `punkte` | Bild + Text |
| `kernaussage` | One content dark | `titel` (Aussage), `text` | Botschaft / Entscheidung |
| `abschluss` | One content dark Outro | `titel` | „Vielen Dank“ |
| `frei` | Only title | `elemente` [{`form`: flaeche/text/linie/icon/punkt/pfeil/karte, x, y, w, h …}] in Zoll | Sonderfälle ohne Code |

## Blöcke in Karten (`inhalt` oder Kurzfelder)
`{"punkte": [...]}` · `{"text": "..."}` · `{"phasen": {"namen": [...], "stand": [2,1,0]}}` ·
`{"schritte": [{"status": "fertig|aktiv|offen", "titel", "text"}]}` ·
`{"von_nach": {"von": {"klein", "gross"}, "nach": {...}}}` · `{"kennzahl": {"wert", "label"}}` ·
`{"hinweis": "..." | {"titel", "text", "strich": true}, "unten": true}`.
Kurzform: `punkte`, `schritte`, `phasen`, `text`, `von_nach`, `naechster` (→ Hinweis „Nächster Schritt“ unten), `hinweis`.

## Timeline-Elemente
`{"art": "balken", "von", "bis", "text", "status", "pos"}` · `{"art": "meilenstein", "datum", "text", "status", "pos"}` ·
`{"art": "ueberbrueckung", "von", "bis"}` (rote Linie). `status`: fertig, plan, aktiv, klaerung, offen, extern, kritisch.
Daten: `2026-10-07`, `07.10.2026`, `2026-10` (Monat; bei `bis` = Monatsende), `2026-W43` oder `KW43/2026`.
`pos`: rechts/links/oben/unten – nur nötig, wenn `bauen` „überlappt“ meldet.
`marker` [{datum, text}] = durchgezogene Linie (z. B. Cut-Over); `schatten` [{von, bis, text}] = graue Zone (Feiertage).

## Layout-Kurznamen (`layout`)
titel, titel_dunkel, titel_bild, agenda, kapitel, kapitel_dunkel, inhalt, inhalt_dunkel, zwei_inhalte, drei_inhalte,
inhalt_bild, nur_titel, leer, abschluss, disclaimer, kontakt – oder der exakte Layoutname des Masters.
