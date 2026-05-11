# Anleitung: Labeling Prio Festlegung

## 1) Start
```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```

## 2) Grundprinzip (wichtig)
Die App hat 3 Ebenen, die getrennt gedacht werden müssen:

1. Filteransichten A/B/C
- Für interaktive Sicht auf Daten (Regex, Kategorien, Zahlen, Datum).
- Eher für Analyse und Kommunikation.

2. Kriterien-Set (versioniert)
- Das ist die offizielle Entscheidungslogik.
- Regeln sind nachvollziehbar und wiederholbar.

3. Referenzlisten (z. B. QC)
- Externe Ausnahmen, die bewusst und dokumentiert wirken.

## 3) Was ist „nachvollziehbar“?
Nachvollziehbar heißt:
- Welche Version war aktiv (`criteria_sets/*.json`)?
- Welche Regel hat gegriffen (`matched_rule_ids`)?
- Warum wurde ausgeschlossen/markiert (`exclude_reasons`, `unterbruch_erforderlich`)?
- Wie verändert sich das Ergebnis zwischen Versionen?

## 4) Standard-Vorgehen pro Runde

### Schritt A: Daten laden
- In `Datenquelle` aktuelle CSV wählen oder hochladen.
- Optional im Projekt speichern (`data/`).

### Schritt B: Kriterien-Version wählen
- Sidebar `Kriterien > Aktives Kriterien-Set` auswählen.
- Für neue Runde: auf bestehender Version aufbauen.

### Schritt C: Regeln anpassen
Im Tab `Kriterien`:
- `Übersteuerung: Unterbruch erforderlich` setzen
- `Regel-Builder (visuell)` für AND/OR-Ketten nutzen
- ggf. Tabelle unten für Feinjustierung

### Schritt D: Version speichern
- `Kriterien-Version speichern`
- sprechender Dateiname, z. B.:
  - `baseline_v1.json`
  - `v2_unterbruch_strikt.json`
  - `v3_due_window_5m.json`

### Schritt E: Ergebnis lesen
In A/B/C:
- `x von y Zeilen`
- `Ausgeschlossen durch QC-Regel`
- `Ausgeschlossen durch Kriterien-Set`
- `Prioritätsstufen (aus Kriterien-Set)`
- Spalten: `prio_stage`, `prio_substage`, `unterbruch_erforderlich`

## 5) Konkrete Regelbeispiele

### Beispiel 1: Go-Live-Fenster
Ziel: Intervall < 12 und Due-Date in 0..5 Monaten nach Go-Live.

Vorher:
- Sidebar `Kriterien > Go-Live Referenzdatum` setzen (z. B. `2026-08-10`).

Regel im Builder:
- Bedingung 1: `Interval < 12`
- Bedingung 2: `months_from_golive >= 0`
- Bedingung 3: `months_from_golive <= 5`
- Verknüpfung: `AND`
- Aktion: `assign_prio`
- Prio: z. B. `P2`

### Beispiel 2: Unterbruchspflicht (dein Fall)
Ziel: Alles außer `einfach`/`mittel` soll Unterbruch bekommen.

In `Übersteuerung: Unterbruch erforderlich`:
- `Kein Unterbruch für diese Werte` = `einfach`, `mittel`
- `Mit Asset-ID-Länge` = aus (wenn Länge egal sein soll)
- Prio = `P1`
- Sub-Prio = `1A`
- `Übersteuerungs-Regel hinzufügen/aktualisieren`

Ergebnis:
- `unterbruch_erforderlich = True`
- `prio_stage = P1`
- `prio_substage = 1A`

### Beispiel 3: Unterbruch nur bei langen IDs (optional)
Wie Beispiel 2, aber:
- `Mit Asset-ID-Länge` = an
- Schwellwert setzen (z. B. `>34`)

## 6) A/B/C sinnvoll für 3 Teams/Varianten nutzen
Empfehlung:
- A = Baseline
- B = Variante „streng“
- C = Variante „operativ“

Pro Ansicht:
- Filter nur für diese Ansicht setzen
- Preset speichern (`Filteransicht speichern (JSON)`)
- CSV + ZIP exportieren

Wichtig:
- Filter-Presets sind UI/Analyse-Zustand.
- Entscheidungslogik liegt im Kriterien-Set.

## 7) Unterschiede zwischen Versionen nachvollziehen
Praktischer Ablauf:

1. Version 1 aktivieren (`baseline_v1.json`), Ergebnis exportieren.
2. Version 2 aktivieren, Ergebnis exportieren.
3. Vergleich über:
- Anzahl Zeilen in Prio
- Tabelle `Ausgeschlossen durch Kriterien-Set`
- Verteilung `prio_stage` / `prio_substage`
- Summenanzeigen nach Bereich/Standort

Zusatzempfehlung:
- Im Dateinamen Ziel und Annahme kodieren, z. B.:
  - `v4_no_easy_medium_shutdown.json`
  - `v5_shutdown_with_len34.json`

## 8) Gedankengang dokumentieren (für Dritte)
Für jede neue Kriterien-Version kurz festhalten:
- Hypothese: Was soll diese Version verbessern?
- Regeländerung: Was wurde konkret geändert?
- Erwarteter Effekt: Mehr/weniger Zeilen, welche Gruppen betroffen?
- Ergebnis: Was ist tatsächlich passiert?
- Entscheidung: Übernehmen / Verwerfen / Nächste Iteration

Das kann in:
- `To Do und vorgehen Logik.md`
- oder separater `runs/` Notiz je Runde

## 9) Typische Stolperfallen
- Oben im Builder wird nicht automatisch die komplette Tabelle unten „rückwärts geladen“: Builder ist primär Erzeugung/Update.
- Nach Regeländerung immer klicken:
  - `Übersteuerungs-Regel hinzufügen/aktualisieren` oder `Regel ... hinzufügen`
  - danach `Kriterien-Version speichern`
- Wenn Wirkung fehlt: prüfen, ob richtiges `Aktives Kriterien-Set` gewählt ist.

## 10) Export
Pro Filteransicht:
- CSV
- ZIP (CSV + Filterzustand)

Standort-Analyse:
- CSV
- ZIP

Für Freigabe/Review immer beilegen:
- verwendete CSV-Version
- Kriterien-Set-Dateiname
- Exportdatum

## 11) Neuer Tab: Vergleich
Im Tab `Vergleich` kannst du 2 Kriterien-Sets direkt gegeneinander rechnen:
- `Version A` wählen
- `Version B` wählen

Du siehst:
- Kennzahlenvergleich mit Delta (`in_prio`, `excluded`, `P1`, `P2`, `P3`, `unterbruch`)
- Unterschiedsliste auf Zeilenebene (pro `Asset ID`):
  - Status A/B
  - Prio A/B
  - Sub-Prio A/B
  - Unterbruch A/B

Damit kannst du für Dritte transparent zeigen:
- was sich geändert hat
- wie stark es sich geändert hat
- welche konkreten Messstellen betroffen sind


## 12) Bedeutet `assign_prio` Ausschluss?
Nein.

- `assign_prio`:
  - Zeile bleibt in der Liste (`in_prio`).
  - Es werden nur Markierungen gesetzt, z. B.:
    - `prio_stage` (P1/P2/P3)
    - `prio_substage` (z. B. 1A)
    - `unterbruch_erforderlich`

- `exclude_from_prio`:
  - Zeile fliegt aus der aktiven Prio-Liste raus.
  - Zeile erscheint in `Ausgeschlossen durch Kriterien-Set` mit Grund.

Merksatz:
- `assign_prio` = "drin lassen und einstufen"
- `exclude_from_prio` = "rausnehmen"

Prüfhinweis:
Wenn sich nur Prio-Markierungen ändern, bleibt die Zeilenanzahl gleich. Das ist korrekt.
