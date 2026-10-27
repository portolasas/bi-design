# Umwandeln – fremdes Design → BI-Design

## Entscheiden (eine Frage, Empfehlung zuerst)
| Weg | Wann | Ergebnis |
|---|---|---|
| **Neu bauen (Empfehlung)** | Inhalt zählt, ≤ ~30 Folien, Folien sollen BI-Layouts nutzen | 100 % BI-Master, Inhalte verdichtet |
| **Stil angleichen** | sehr viele Folien, komplexe Grafiken/Bilder, schnell gebraucht | Farben/Schriften BI, alter Master bleibt (Befund) |
| Kombination | Kernfolien neu, Detail-/Bildfolien als Stil-Kopie | zwei Dateien, Nutzer fügt zusammen |

## Weg 1 – Neu bauen
1. `bippt analyse fremd.pptx` (Überblick) und `bippt pruefen fremd.pptx` (Ausgangsbefunde, für den Abschluss).
2. `bippt umwandeln entwurf fremd.pptx [-o ziel.spec.json]` → `fremd_BI.spec.json` (+ `fremd_medien/` mit Bildern
   neben der Spec; verknüpfte Bilder werden nicht übernommen).
   Der Entwurf ordnet zu: Titel → `titel`, Diagramm → `diagramm` (Daten übernommen), Tabelle → `tabelle`,
   Bild+Text → `bild_text`, 2–4 Textblöcke → `karten`, sonst `text`. Originalnotizen bleiben.
3. Spec verdichten (das ist die eigentliche Arbeit): Aktionstitel, passende Typen (Abläufe → `prozess`, Termine →
   `timeline`, alt/neu → `vergleich`, Zahlen → `kennzahlen`), Band je Folie, Felder `_quelle_folie`/`_hinweis` löschen.
   Inhaltlich nichts erfinden – Unklares fragen (fragen.md, Modus U).
4. `bippt bauen` → `pruefen` → `rendern` wie üblich. Abschluss: Folienzuordnung alt → neu als Tabelle.

## Weg 2 – Stil angleichen
`bippt umwandeln stil fremd.pptx [-o ziel.pptx]` – ersetzt jede Nicht-Markenfarbe durch die nächste Markenfarbe
(Helligkeit gewichtet), setzt fremde Schriften auf die Theme-Schriften und übernimmt Farb- und Schriftschema des
BI-Masters ins Theme. Danach `pruefen` und `rendern`: auf Kontrast (Akzentgrün auf Hell!) und kaputte Grafiken prüfen.
Grenzen: Layouts, Logo und Hintergrundbilder des alten Masters bleiben; Bilder werden nicht umgefärbt.
