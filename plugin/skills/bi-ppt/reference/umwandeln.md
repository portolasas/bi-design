# Umwandeln – fremdes Design → BI-Design

## Entscheiden (eine Frage, Empfehlung zuerst)
| Weg | Wann | Ergebnis |
|---|---|---|
| **In die Marke übertragen (Empfehlung)** | Deck ist gut gestaltet, z. B. aus der Erstellen-Funktion von Copilot | Gestaltung und Inhalt bleiben; Farben, Schriften, Icons, Logo und Fußzeile der Marke |
| **Neu bauen** | Inhalt muss verdichtet werden oder Folien sollen BI-Layouts nutzen | 100 % BI-Master, Inhalte verdichtet |
| **Stil angleichen** | sehr viele Folien, komplexe Grafiken/Bilder, schnell gebraucht | Farben/Schriften BI, alter Master bleibt (Befund) |
| Kombination | Kernfolien neu, Detail-/Bildfolien als Stil-Kopie | zwei Dateien, Nutzer fügt zusammen |

## In die Marke übertragen
Entscheidung des Nutzers (30.10.2026): Das Deck muss nicht in die Layouts des Masters passen, wenn die Elemente markenkonform sind.
Die Erstellen-Funktion von Copilot gestaltet in HTML und übersetzt das in PowerPoint-Formen; das Ergebnis ist oft besser als alles, was
wir bauen. Darum nehmen wir es als Quelle.
1. `bippt umwandeln marke fremd.pptx -o ziel.pptx`. Der Bericht nennt die OFFENEN ICONS: Folie, id und der Text daneben.
2. Je Icon nach der Bedeutung des Textes ein Brand-Icon wählen (`bippt icons <englischer Begriff>`, ein Begriff je Aufruf) und in
   eine JSON-Datei schreiben: `{"2": {"7": "clock", "10": "office-printer"}}`. Dann erneut mit `--icons icons.json`. Fertig bei
   „0 nur umgefärbt“. Häkchen und Punkte vor Listen zählen nicht.
3. `bippt pruefen` und `bippt sichten`. Überlappung und Überstand beheben. Leere nur nennen: Die Gestaltung stammt vom Designer.
4. Inhalt nicht ändern und nicht kürzen. Fotos färbt bippt nicht um; es meldet sie.

Der Befehl tauscht Farben nach ihrer Rolle: dunkle Flächen werden Dark Green, Hervorhebungen Accent Green mit dunklem Text, helle Flächen
Warm und Light Gray, gedämpfter Text Gray 7. Runde Ecken und Schatten fallen weg. Fußtexte des Designers weichen der Fußzeile des
Masters, und was in Logo oder Fußzeile ragt, wird in der Höhe eingepasst. Die Titelfolie kommt aus dem Master.

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
