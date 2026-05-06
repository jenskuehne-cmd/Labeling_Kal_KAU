# Anleitung: Labeling Prio Festlegung

## Start
```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```

## Datenquelle
In der linken Sidebar unter **Datenquelle**:
- vorhandene CSV aus `data/` wählen
- oder neue CSV hochladen
- optional dauerhaft speichern

## Tabs / Filteransichten
Es gibt 3 Tabs:
- Filteransicht A
- Filteransicht B
- Filteransicht C

Jeder Tab hat eigene:
- Filter
- Rule-Engine-Einstellungen
- Preset-Name/Bemerkung
- manuelle Overrides

## Filter pro Tab
Die Filter sind im jeweiligen Tab:
- Regex-Filter
- Kategorienfilter
- numerische Filter
- Datumsfilter
- Rule Engine

## Presets pro Tab
Im Tab:
- `Filteransicht speichern (JSON)`
- `Filteransicht laden (JSON)`

Du kannst ein Preset aus A auch in B oder C laden und dort weiterarbeiten.

## Editierbare Ansicht
- `Manuellen Override aktivieren`
- editierbare Spalten:
  - `prio_manual` (`Low`, `Medium`, `High`)
  - `prio_note`

Zusätzlich entstehen:
- `prio_final_score`
- `prio_final_label`
- `prio_final_match`

## Summenanzeigen
Pro Gruppierung (z. B. `Bereich`):
- `anzahl_messstellen`
- `anzahl_prio_treffer`
- `summe_prio_score`

## Export
Je Tab kann die aktuelle Ansicht als CSV exportiert werden.
