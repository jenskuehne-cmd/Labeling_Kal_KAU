# Thread-Handoff: Baue Streamlit-Relabeling-App

## Zweck
Dieses Dokument holt den Codex-Thread `Baue Streamlit-Relabeling-App` in das Projekt, ohne die App-Logik direkt zu verändern.

## Ausgangslage aus dem Thread
Ziel ist eine lokale Streamlit-App für die Priorisierung von Relabeling-Messstellen auf Basis einer CSV-Datei mit ca. 7000 Zeilen.

Die erste Ausbaustufe soll folgende Funktionen abdecken:
- CSV-Upload
- Tabelle anzeigen
- Spaltennamen anzeigen
- einfache Filter für Textspalten
- Regex Include / Regex Exclude auf frei wählbarer Spalte
- Export der gefilterten Daten als CSV
- sauberer, modularer und erweiterbarer Code

## Vom Thread vorgeschlagene Baseline
Der Thread hat als Startpunkt diesen Funktionsumfang festgehalten:
- `app.py`
- `requirements.txt`
- `streamlit`
- `pandas`
- `openpyxl`

## Fachliche Leitplanken
- Exploration und Entscheidungslogik sollen getrennt bleiben.
- Filteransichten sind keine finale Regel-Engine.
- Kriterien-Sets sollen versionierbar sein.
- Referenzlisten und manuelle Overrides sollen später sauber integrierbar bleiben.

## Bisherige Projektstruktur, die dafür relevant ist
Im aktuellen Repo existieren bereits:
- `app.py`
- `requirements.txt`
- `criteria_sets/`
- `reference_lists/`
- `data/`
- `ANLEITUNG.md`
- `To Do und vorgehen Logik.md`

Das heißt:
- Es gibt schon eine App-Basis.
- Der Thread ist kein Neubau, sondern ein Kontext- und Ausbaupfad für die vorhandene Anwendung.

## Nicht direkt übernehmen
Der Thread sollte nicht 1:1 als Blindkopie in die App gepatcht werden, wenn dabei bestehende Funktionen, Presets oder State-Logik brechen würden.

Empfohlene Reihenfolge:
1. Erst den bisherigen Stand der App prüfen.
2. Dann die Thread-Ziele in bestehende Strukturen einordnen.
3. Erst danach gezielte Änderungen an UI, Filtern oder Datenfluss machen.

## Offene nächste Schritte
- Abgleichen, welche Teile des Threads bereits im aktuellen `app.py` angekommen sind.
- Prüfen, ob eine separate Minimal-App oder nur ein Ausbau der bestehenden App sinnvoll ist.
- Die erste Ausbaustufe klar schneiden:
  - reine CSV-Filterung
  - oder schon Kriterien-Set / Regeln / Priorisierung

## Entscheidungspunkt
Wenn du möchtest, kann dieser Thread-Handoff jetzt als Arbeitsgrundlage dienen, ohne dass an der App selbst sofort etwas geändert wird.
