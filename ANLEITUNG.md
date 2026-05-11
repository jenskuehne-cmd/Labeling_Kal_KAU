# Anleitung: Labeling Prio Festlegung

## Start
```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```

## Datenquelle
In der Sidebar unter `Datenquelle`:
- vorhandene CSV aus `data/` wählen
- oder neue CSV hochladen
- optional dauerhaft speichern

## Grundprinzip
Die App trennt drei Ebenen:
1. Interaktive Filter (A/B/C) für Exploration.
2. Kriterien-Sets (versionierbare Regeln) für nachvollziehbare Entscheidungen.
3. Referenzlisten (z. B. QC-Selbstlabeling) für externe Ausnahmen.

Damit ist dokumentierbar:
- was ausgeschlossen wurde
- warum ausgeschlossen wurde
- mit welcher Kriterien-Version entschieden wurde

## Tabs
- `Filteransicht A`
- `Filteransicht B`
- `Filteransicht C`
- `Standort-Analyse`
- `Kriterien`

## Filteransichten A/B/C
Jede Ansicht hat eigenen Zustand:
- Regex-/Kategorien-/Numerik-/Datumsfilter
- Rule Engine
- Preset (Name, Bemerkung, Laden/Speichern)
- manueller Override

Zusätzliche Transparenz:
- Tabelle `Ausgeschlossen durch QC-Regel`
- Tabelle `Ausgeschlossen durch Kriterien-Set`

In den Ausschlusstabellen sind Regel-IDs und Gründe sichtbar.

## Presets
Pro Ansicht:
- `Filteransicht speichern (JSON)`
- `Filteransicht laden (JSON)`

Hinweis:
- Presets speichern Filter-/UI-Zustände.
- Formale Ausschlusslogik gehört ins Kriterien-Set.

## QC-Referenzliste (Excel grün)
In der Sidebar unter `Referenzlisten > QC-Grünliste (Excel)`:
1. Excel hochladen
2. ggf. Farbcode/Asset-Spalte anpassen
3. `QC-Liste einlesen`

Die grün markierten Zeilen werden über `Asset ID` extrahiert und gespeichert als:
- `reference_lists/qc_self_labeled_assets.csv`

Diese Liste wird beim nächsten App-Start automatisch geladen.

## Kriterien-Sets (versionierbar)
In der Sidebar unter `Kriterien`:
- aktives Kriterien-Set auswählen
- im Block `Live-Kriterien` die wichtigsten Regeln direkt live anpassen
  - `Intervall > (Monate)`
  - `Asset ID Länge <`
  - `QC-Selbstlabeling`
  - inkl. sichtbarer Trefferzahl pro Regel

Im Reiter `Kriterien`:
1. `Regel-Builder (visuell)` nutzen (1-8 Bedingungen + AND/OR) und Regel hinzufügen
2. optional Regeln in der Tabelle feinbearbeiten
3. neue Version speichern (`Kriterien-Version speichern`)
4. Datei wird in `criteria_sets/` abgelegt

Unterstützte Regeltypen (Legacy):
- `numeric_gt`
- `string_length_lt`
- `regex_match`
- `date_between`
- `ref_list_match`

Zusätzlich: Bedingungsblöcke mit `when`:
- `op: AND` / `OR` / `NOT`
- Vergleiche: `>`, `>=`, `<`, `<=`, `==`, `!=`
- Felder: `column`, `value` oder `value_col`

Unterstützte Aktionen:
- `exclude_from_prio`
- `assign_prio` mit `priority: P1|P2|P3`

Beispiel (dein Use-Case):
- Bedingung 1: `Interval > 7`
- Bedingung 2: `months_until_due <= interval_half_months`
- Verknüpfung: `AND`
- Aktion: `assign_prio`, Prio `P2`

## Prioritätsstufen
Wenn Kriterien mit `assign_prio` aktiv sind:
- Ergebnis enthält Spalte `prio_stage` (`P1`, `P2`, `P3`)
- In jeder Filteransicht wird `Prioritätsstufen (aus Kriterien-Set)` als Summen-Tabelle angezeigt
- Exkludierte Zeilen bleiben in `Ausgeschlossen durch Kriterien-Set` mit Begründung sichtbar

## Summen und Details
In jeder Filteransicht:
- Summenanzeige nach Gruppe
- Gruppendetails mit auswählbaren Spalten

In `Standort-Analyse`:
- Zugänglichkeit per Regex/Leer filtern
- Standort-Basis bilden
- Detail- und Rohdatenansicht

## Export
Je Filteransicht:
- `CSV exportieren`
- `CSV + Filter (ZIP)` mit Filterzustand als JSON

In `Standort-Analyse`:
- CSV-Export
- ZIP-Export mit Filterzustand

- Im Builder steuerst du über `Anzahl Bedingungen`, wie viele Einzelkriterien die Regel enthält.

## Go-Live Bezug
- In der Sidebar unter `Kriterien` setzt du `Go-Live Referenzdatum` (z. B. `2026-08-10`).
- Die App erzeugt daraus die Kennzahl `months_from_golive` auf Basis der Due-Date-Spalte.
- Positiv bedeutet: Due-Date liegt nach Go-Live.

Für den Fall „Intervall < 12 und Due-Date kommt nach Go-Live innerhalb 5 Monate“:
- Bedingung 1: `Interval < 12`
- Bedingung 2: `months_from_golive >= 0`
- Bedingung 3: `months_from_golive <= 5`
- Verknüpfung: `AND`

