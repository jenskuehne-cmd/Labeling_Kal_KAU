# Abgleich: Thread "Baue Streamlit-Relabeling-App" vs. aktuelles `app.py`

## Kurzfazit
Der Thread hatte einen sehr schlanken Startpunkt. Das aktuelle `app.py` geht deutlich darueber hinaus und deckt die Thread-Ziele bereits weitgehend ab. Es gibt also keinen 1:1-Importbedarf fuer die App-Logik, sondern vor allem einen inhaltlichen Abgleich und eine Priorisierung der naechsten Schritte.

## Was der Thread forderte
- CSV-Upload
- Tabelle anzeigen
- Spaltennamen anzeigen
- einfache Filter fuer Textspalten
- Regex Include / Regex Exclude auf frei waehlbarer Spalte
- Export der gefilterten Daten als CSV
- modularer, erweiterbarer Code

## Was im aktuellen `app.py` schon vorhanden ist

### Bereits umgesetzt
- CSV-Auswahl und Upload im Sidebar-Bereich
- Anzeige der Spalten der aktuellen CSV
- drei Filteransichten A/B/C
- Regex-Filter
- Kategorie-Filter fuer Textspalten
- numerische Filter
- Datumsfilter
- Export als CSV
- Export als CSV + Filterzustand im ZIP
- persistenter Arbeitsstand ueber `runs/last_session_state.json`
- Kriterien-Sets als JSON
- QC-Referenzliste
- manuelle Overrides
- Standortspezifische Auswertung
- Phasenplan und Kapazitaetslogik
- Kriterien-Vergleich
- Hilfeseiten und Anleitung

### In Relation zum Thread
Die Thread-Basis ist damit nicht nur erreicht, sondern deutlich ueberholt. Der passende Schluss daraus ist nicht "mehr von Grund auf bauen", sondern "die bestehende App sauber halten und gezielt ausbauen".

## Konkrete Luecken zum Thread-Baseline
Es gibt keine harte funktionale Luecke mehr bei den initialen Anforderungen.

Die einzige relevante Differenz ist die Produktform:
- Der Thread dachte an eine kleine Starter-App.
- Das Repo enthaelt bereits eine echte Arbeits-App mit Fachlogik.

## Wo das aktuelle `app.py` bewusst weiter geht
- Die App trennt Filterung und Entscheidungslogik.
- Sie hat ein versionierbares Kriterien-Set.
- Sie speichert Zustand und referenzielle Daten.
- Sie bietet Vergleich und Fachauswertung.

Das ist fachlich sinnvoll, aber deutlich komplexer als der Thread-Start.

## Risiken beim weiteren Ausbau
- Sidebar-Logik kann schnell unuebersichtlich werden.
- Widget-Keys und Session-State sind empfindlich.
- Zu viele gleichzeitig aktive Modus-Wechsel erzeugen schwer nachvollziehbare UI-Zustaende.

## Empfehlung fuer die naechste Bearbeitungsreihenfolge
1. Keine neue Kernlogik, bevor die vorhandenen Zustaende dokumentiert sind.
2. Nur eine Erweiterung pro Arbeitsschritt.
3. Erst dann an Filter-UX, Tabellenformatierung oder weitere Prio-Regeln gehen.

## Technische Einordnung der bestehenden Struktur
Die aktuelle App ist bereits modular genug, um als Ausgangsbasis fuer den Thread zu dienen:
- `apply_filters(...)`
- `apply_criteria_rules(...)`
- `render_view(...)`
- `render_criteria_editor(...)`
- `render_phase_plan(...)`
- `render_criteria_comparison(...)`

Das ist der richtige Ort fuer weitere Anpassungen, nicht ein kompletter Neubau.
