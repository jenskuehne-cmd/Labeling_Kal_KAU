# Klickanleitung: KAU Relabeling Priorisierung

## Ziel
Diese Anleitung beschreibt den konkreten Klickpfad in der App, damit auch neue Nutzer reproduzierbar arbeiten können.

## 1) App starten
1. Terminal öffnen.
2. Ausführen:
```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```
3. Browser mit der Streamlit-App öffnen.

## 2) Daten laden
1. Links `Datenquelle` aufklappen.
2. Entweder vorhandene CSV unter `Aktuelle CSV` auswählen.
3. Oder `Neue CSV hochladen` nutzen.
4. Wenn der Datenstand reproduzierbar bleiben soll:
   - Checkbox `Hochgeladene CSV im Projekt speichern` aktiv lassen.
   - Dateiname vergeben (mit Datum/Version), z. B. `Labeling_KAU_2026-05-26_v1.csv`.
   - `CSV speichern und verwenden` klicken.

## 3) Analyse-Modus wählen
1. Links unter `Entscheidungsbaum` bei `Aktiver Entscheidungsbaum` auswählen:
   - `'(kein Entscheidungsbaum)'` = nur manuelle Filter, keine Entscheidungslogik.
   - JSON-Datei = offizielle Entscheidungsbaum-Engine aktiv.
2. `Go-Live Referenzdatum` prüfen/setzen.

## 4) Filteransichten A/B/C sauber trennen
Empfohlene Nutzung:
- `Filteransicht A` = Baseline
- `Filteransicht B` = strenge Variante
- `Filteransicht C` = operative Variante

Pro Ansicht:
1. Tab öffnen (`Filteransicht A/B/C`).
2. Optional `Diese Ansicht links bearbeiten` klicken, damit Sidebar-Filter zu dieser Ansicht gehören.
3. Feld `Ansichtszweck` ausfüllen (1 Satz, klarer Zweck).
4. Regex/Kategorien/Numerik/Datum setzen.
5. `Filter-Zusammenfassung (auto)` prüfen.

## 5) Entscheidungsbaumregeln bauen (offizielle Logik)
1. Tab `Entscheidungsbaum` öffnen.
2. Im `Entscheidungsbaum-Builder`:
   - `Aktion` wählen (`assign_prio` oder `exclude_from_prio`).
   - `Anzahl Bedingungen` setzen.
   - Pro Bedingung Spalte/Operator/Wert setzen.
   - Bei Textmustern `Operator = REGEX` verwenden.
   - Bei Negation `NOT` aktivieren.
3. `Regel-ID` und `Begründung` setzen.
4. `Regel zum aktiven Entscheidungsbaum hinzufügen` klicken.

## 6) Häufiger Anwendungsfall (kurze IDs)
Ziel: kurze IDs raus, außer schwer zugänglich (dann P1A + Unterbruch)

Variante mit Rezept:
1. Im Abschnitt `Rezept: P1 (Länge) + P1A (Unterbruch)`:
   - Schwelle = `34`
   - Regex = `schwer|sehr\\s*schwer`
   - Sub-Prio = `1A`
2. `Rezept-Regeln hinzufügen/aktualisieren` klicken.

Zusatzregel (kurz + nicht schwer => Ausschluss):
1. Builder: `Aktion = exclude_from_prio`
2. Bedingung 1: `asset_id_length <= 34`
3. Bedingung 2: `REGEX schwer|sehr\\s*schwer` + `NOT` aktiv
4. Regel hinzufügen.

## 7) Entscheidungsbaum-Version speichern
1. Im Tab `Entscheidungsbaum` unten `Dateiname für neue Version` setzen, z. B.:
   - `v04_shortid_shutdown_logic.json`
2. `Entscheidungsbaum speichern` klicken.
3. Kontrolle links: `Aktiver Entscheidungsbaum` ist die gerade gespeicherte Datei.

## 8) Wirkung prüfen
In Filteransicht A/B/C prüfen:
1. `x von y Zeilen` (wurde wirklich ausgeschlossen?).
2. Tabelle `Ausgeschlossen durch Entscheidungsbaum`.
3. Spalten in Detailtabelle einblenden:
   - `prio_stage`
   - `prio_substage`
   - `unterbruch_erforderlich`
   - `decision_reason`

Wichtig:
- `assign_prio` = bleibt in Liste, nur Markierung.
- `exclude_from_prio` = wird aus Liste entfernt.

## 9) Versionen vergleichen
1. Tab `Vergleich` öffnen.
2. `Version A` und `Version B` wählen.
3. Prüfen:
   - Delta-Kennzahlen (`in_prio`, `excluded`, P-Stufen, Unterbruch)
   - Unterschiedsliste auf Zeilenebene (`Asset ID`).

## 10) Kapazität planen
1. Links unter `Kapazität` Werte setzen:
   - Minuten normal/mittel
   - Puffer %
   - Personenanzahl
   - Stunden/Person/Tag
2. Tab `Phasenplan` öffnen.
3. Aufwand, Personentage und Wochenbedarf prüfen.

## 11) Exporte für Review
Pro Filteransicht:
1. `CSV exportieren`
2. `CSV + Filter (ZIP)` exportieren

Für Freigabe immer mitgeben:
- CSV-Dateiname
- Entscheidungsbaum-Dateiname
- Exportdatum

## 12) Troubleshooting (Kurz)
- Nichts ändert sich:
  - Prüfen, ob `'(kein Entscheidungsbaum)'` aktiv ist.
- Regel scheint nicht zu greifen:
  - Operator/Spalte prüfen (bei Länge immer `asset_id_length`, nicht `Asset ID <= 34`).
- Regel nicht sichtbar:
  - Wurde `Regel ... hinzufügen` geklickt?
  - Wurde `Entscheidungsbaum speichern` geklickt?
- QC ohne Wirkung:
  - In Ansicht unter `QC-Ausnahme` die Checkbox aktivieren.


## 13) Szenario bauen: Annahme ändern (Nicht-Profi Ablauf)
Beispielannahme: "`>30 Zeichen` soll nicht mehr prioritätsrelevant sein".

1. Links `Entscheidungsbaum > Aktiver Entscheidungsbaum` auf das bisherige Set stellen.
2. Tab `Entscheidungsbaum` öffnen.
3. In der Regeltabelle alle Regeln identifizieren, die `ms_legacy_class` verwenden.
4. Für das Szenario entweder:
   - `active` auf `false` setzen (temporär aus), oder
   - Regel löschen (`Regel entfernen`), oder
   - `when` anpassen und die Legacy-Bedingung entfernen.
5. Unten neuen Dateinamen setzen, z. B. `v08_no_len30_assumption.json`.
6. `Entscheidungsbaum speichern` klicken.
7. Tab `Vergleich` öffnen:
   - `Version A` = alte Version
   - `Version B` = neue Szenario-Version
8. Deltas prüfen:
   - `in_prio` / `excluded`
   - Prio-Stufen
   - Unterschiedsliste der Assets
9. Tab `Phasenplan` öffnen und Aufwand prüfen.
10. Ergebnis notieren (kurz):
   - Was geändert?
   - Was hat sich in Zahlen verändert?
   - Entscheidung: übernehmen oder verwerfen?


## 14) Manuelle Override-Liste laden (Einzelfälle)
1. Links `Referenzlisten` aufklappen.
2. `Manuelle Overrides (CSV)` öffnen.
3. CSV hochladen.
4. Entweder `Override-Liste einlesen` (bei Upload) oder direkt Tabelle bearbeiten.
5. `Override-Liste speichern` klicken.
6. Kontrolle: `Aktive Override-Liste: X Assets`.

### Minimales CSV-Beispiel
```csv
Asset ID,override_action,override_prio,override_substage,override_reason,override_shutdown
235.A01.34.300.QT.184A2,assign_prio,P1,1A,Besonderer Einzelfall,true
250-VF090-BK000-YY018,exclude_from_prio,,,Nicht im Scope,false
```
