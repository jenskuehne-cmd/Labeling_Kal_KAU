# Anleitung: KAU Relabeling Priorisierung (Einsteiger + Power-User)

## 1) Ziel der App
Die App hilft, Messstellen systematisch zu priorisieren:
- Was muss sofort relabelt werden?
- Was kann später erfolgen?
- Was wird bewusst ausgeschlossen (z. B. QC/ITOT)?

Wichtig: Die App trennt **Exploration** (Filteransichten) von **Entscheidungslogik** (Entscheidungsbaum).

## 2) Start
```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```

## 3) Grundlogik in 30 Sekunden
Es gibt 3 Ebenen:

1. Filteransichten A/B/C
- zum Sichten und Analysieren (Regex, Kategorien, numerisch, Datum)
- schnell, flexibel, aber nicht die „offizielle“ Logik

2. Entscheidungsbaum (JSON)
- versionierbare Entscheidungslogik
- regelt Ausschluss, Prio-Stufe, Unterbruch, Phase

3. Referenzlisten (QC)
- externe Ausschlüsse
- können in Regeln berücksichtigt werden

Merksatz:
- **Filteransicht** = „Wie schaue ich auf die Daten?“
- **Entscheidungsbaum** = „Wie entscheide ich verbindlich?“

## 4) Datenquelle richtig nutzen
In der Sidebar unter `Datenquelle`:
- vorhandene CSV wählen
- oder neue CSV hochladen

### Was bedeutet „optional im Projekt speichern“?
Wenn aktiv:
- die hochgeladene Datei wird nach `data/` geschrieben
- du kannst sie später wieder auswählen
- sinnvoll für reproduzierbare Analysen

Wenn nicht aktiv:
- Datei ist nur temporär für diese Session
- beim Neustart ist sie weg

Empfehlung für Teamarbeit:
- wichtige Datenstände immer speichern
- Dateiname mit Datum/Version, z. B.:
  - `Labeling_KAU_2026-05-26_v1.csv`
  - `Labeling_KAU_2026-05-26_v2_qcfix.csv`

## 5) Unterschiedliche Analysen sauber trennen
Verwende feste Rollen für die Tabs:
- A = Baseline
- B = Strenge Variante
- C = Operative Variante

Pro Ansicht:
- `Ansichtszweck` ausfüllen (kurz, 1 Satz)
- `Filter-Zusammenfassung (auto)` als Schnellcheck nutzen
- bei Bedarf Preset speichern

### Empfohlenes Benennungsschema
Für Entscheidungsbaum-Versionen:
- `v01_baseline.json`
- `v02_streng_interval.json`
- `v03_shutdown_focus.json`

Für Presets:
- `A_baseline_fokus.json`
- `B_interval_11_cut.json`
- `C_operativ_peaks.json`

## 6) Was ist nachvollziehbar dokumentiert?
Nachvollziehbar heißt:
- Welcher CSV-Datenstand?
- Welcher aktive Entscheidungsbaum?
- Welche Regeln haben gegriffen?
- Welche Zeilen wurden ausgeschlossen und warum?

In der App sichtbar über:
- `matched_rule_ids`
- `exclude_reasons`
- `prio_stage`, `prio_substage`
- `unterbruch_erforderlich`
- Tab `Vergleich`

## 7) Standard-Workflow pro Analyse-Runde

### Schritt A: Ausgangslage festhalten
- Datenquelle wählen
- aktiven Entscheidungsbaum notieren
- Ziel der Runde in `Ansichtszweck` eintragen

### Schritt B: Baseline prüfen
- in A ohne Änderungen prüfen: Anzahl, Summen, Ausschlüsse

### Schritt C: Eine Änderung je Iteration
- nur einen Regelblock ändern
- sonst weiß man später nicht, was den Effekt verursacht hat

### Schritt D: Speichern
- `Entscheidungsbaum speichern`
- sprechender Dateiname

### Schritt E: Wirkung prüfen
- in A/B/C: Zeilenzahl, Prio-Verteilung, Ausschlüsse
- im Tab `Vergleich`: alte vs neue Version

### Schritt F: Entscheidung dokumentieren
- Hypothese
- tatsächlicher Effekt
- Entscheidung (behalten/verwerfen)

## 8) `assign_prio` vs `exclude_from_prio`

- `assign_prio`
  - Zeile bleibt in `in_prio`
  - bekommt Markierungen (z. B. `P1`, `1A`, Unterbruch)

- `exclude_from_prio`
  - Zeile wird aus aktiver Prio-Liste entfernt
  - erscheint in `Ausgeschlossen durch Entscheidungsbaum`

Deshalb kann die Zeilenzahl gleich bleiben, obwohl Regeln wirken (wenn nur `assign_prio` genutzt wird).

## 9) Regelbeispiele

### Beispiel 1: Go-Live-Logik
- `Interval < 12`
- `months_from_golive >= 0`
- `months_from_golive <= 5`
- Aktion: `assign_prio`

### Beispiel 2: Kurze IDs raus, außer schwer zugänglich
- Regel A: `asset_id_length <= 34 AND REGEX schwer|sehr\s*schwer` -> `assign_prio P1`, `substage=1A`
- Regel B: `asset_id_length <= 34 AND NOT(REGEX schwer|sehr\s*schwer)` -> `exclude_from_prio`

## 10) Typische Stolperfallen
- `(kein Entscheidungsbaum)` aktiv: dann greifen keine Entscheidungsregeln
- Regel hinzugefügt, aber nicht gespeichert: nach Neustart weg
- Falsche Spalte in Regel (z. B. `Asset ID <= 34` statt `asset_id_length <= 34`)
- `assign_prio` erwartet Zeilenreduktion (falsch)

## 11) Vergleich-Tab richtig nutzen
Im Tab `Vergleich`:
- Version A und B wählen
- Kennzahlen-Deltas prüfen
- Unterschiedsliste pro `Asset ID` prüfen

Damit siehst du:
- welche konkrete Regeländerung welchen Effekt hatte
- welche Messstellen Status/Prio geändert haben

## 12) Phasenplan und Kapazität
Sidebar `Kapazität`:
- Minuten/Messstelle normal
- Minuten/Messstelle mittel
- Puffer %
- Personenanzahl
- Stunden/Person/Tag

Tab `Phasenplan` zeigt:
- Anzahl pro Phase/Prio
- Aufwand
- Personentage
- Wochenbedarf

## 13) QC-Ausschlüsse: zwei Wege

### Weg A (in der App)
- QC-Excel laden
- QC-Ausschlussregeln anwenden

### Weg B (stabiler für Team)
- Ausschlussspalten direkt in Basis-CSV pflegen, z. B.:
  - `exclude_qc_self` (TRUE/FALSE)
  - `exclude_reason`
- dann per Kriterienregel direkt auf diese Spalten gehen

Vorteil: kein separater Upload nötig, reproduzierbarer Datenstand.

## 14) Mini-Protokollvorlage pro Version
Für jede neue Entscheidungsbaum-Version kurz notieren:
- Ziel:
- Änderung:
- Erwartung:
- Ergebnis:
- Entscheidung:

Beispiel:
- Ziel: kurze IDs ohne harte Zugänglichkeit ausschließen
- Änderung: Regel `exclude_short_not_hard` hinzugefügt
- Erwartung: weniger in_prio, mehr in Ausschlussliste
- Ergebnis: in_prio -1200, P1A +280
- Entscheidung: behalten

## 15) Schneller Check für Neulinge (vor Freigabe)
1. Ist der richtige Entscheidungsbaum aktiv?
2. Ist `(kein Entscheidungsbaum)` sicher nicht ausgewählt?
3. Sind Änderungen gespeichert (Datei in `criteria_sets/`)?
4. Wurden A/B über `Vergleich` gegengeprüft?
5. Sind Ausschlüsse mit Grund sichtbar?


## 16) Szenario-Logik ändern (einfach, auch für Nicht-Profis)

### Wann brauche ich ein Szenario?
Wenn sich eine fachliche Annahme ändert, z. B.:
- "Kriterium `>30 Zeichen` ist doch nicht relevant"
- "Zugänglichkeit muss stärker gewichtet werden"
- "Intervall-Ausnahme soll gelockert werden"

Ziel: nicht den alten Entscheidungsbaum kaputt machen, sondern eine neue Vergleichsversion erzeugen.

### Schritt-für-Schritt (Klickpfad)
1. Sidebar `Entscheidungsbaum > Aktiver Entscheidungsbaum` auf das aktuelle Referenz-Set stellen.
2. Tab `Entscheidungsbaum` öffnen.
3. Die relevante Regel suchen:
   - entweder über `Regel entfernen` (wenn komplett weg)
   - oder in der Tabelle in `when` / `active` anpassen (wenn nur ändern)
4. Unten bei `Dateiname für neue Version` neuen Namen vergeben:
   - Beispiel: `v08_no_len30_assumption.json`
5. `Entscheidungsbaum speichern` klicken.
6. Sidebar prüfen: neues Set aktiv.
7. Tab `Vergleich` öffnen:
   - `Version A` = vorherige Version
   - `Version B` = neue Szenario-Version
8. Wirkung prüfen:
   - Delta `in_prio`, `excluded`
   - Delta P-Stufen
   - geänderte Zeilenliste
9. Tab `Phasenplan` prüfen:
   - Personentage / Wochenbedarf
10. Ergebnis kurz dokumentieren (siehe Punkt 14 Vorlage).

### Beispiel: "`>30 Zeichen` nicht mehr relevant"
Auswirkung in der Logik:
- Bedingungen mit `ms_legacy_class == MS_legacy_krit` entfernen oder deaktivieren.

Praktisch:
- Entweder alle betroffenen Regeln löschen und mit neuer Logik neu anlegen
- Oder in jeder betroffenen Regel das `when` so ändern, dass die Legacy-Bedingung entfällt.

Danach:
- als neue Version speichern
- im Vergleich gegen alte Version rechnen

### Einfache Entscheidungsregel für Teams
Pro Szenario genau 1 Änderungskategorie:
- nur Legacy-Logik
- oder nur Zugänglichkeits-Logik
- oder nur Intervall-/Zeitlogik

Warum?
- Sonst ist später nicht klar, welche Änderung den Effekt verursacht hat.

### Minimaler Qualitätscheck vor Freigabe
- richtige CSV geladen?
- richtiger Entscheidungsbaum aktiv?
- Version gespeichert?
- Vergleich A/B durchgeführt?
- Ausschlussgründe plausibel?
- Kapazitätsauswirkung geprüft?


## 17) Manuelle Overrides (Einzelfälle)
Wenn einzelne Messstellen nicht sauber über allgemeine Regeln abbildbar sind:

Lade eine separate CSV unter:
- Sidebar `Referenzlisten` -> `Manuelle Overrides (CSV)`
- Entweder CSV hochladen + `Override-Liste einlesen`
- Oder direkt im eingebauten Tabelleneditor pflegen und `Override-Liste speichern`

Datei wird im Projekt gespeichert unter:
- `reference_lists/manual_overrides.csv`

### Format der Override-CSV
Pflichtspalte:
- `Asset ID`

Optionale Spalten:
- `override_action` (`assign_prio` oder `exclude_from_prio`)
- `override_prio` (z. B. `P1`, `P2A`, `P4`)
- `override_substage` (z. B. `1A`)
- `override_reason` (Begründung)
- `override_shutdown` (`true/false`)

### Wirkung
- Overrides werden nach den normalen Entscheidungsbaumregeln angewendet (haben Vorrang).
- Damit sind gezielte Ausnahmen nachvollziehbar und wiederholbar.

## 18) Automatische Wiederherstellung des letzten Arbeitsstands
Die App speichert den Arbeitsstand automatisch in:

- `runs/last_session_state.json`

Beim nächsten Start wird dieser Stand automatisch geladen, bevor die Sidebar-Widgets aufgebaut werden.

### Was wird automatisch wiederhergestellt?
- zuletzt gewählte CSV aus `data/`
- aktive Filteransicht A/B/C
- Filterwerte der Ansichten A/B/C
- Standort-Analyse-Filter
- aktiver Entscheidungsbaum
- Go-Live Referenzdatum
- Kapazitätswerte

### Was wird bewusst nicht automatisch wiederhergestellt?
- hochgeladene Dateien im Upload-Feld
- Download-Buttons
- interne Tabelleneditor-Zustände
- temporäre DataFrames
- berechnete Ausschlusslisten

Diese Dinge werden nicht gespeichert, weil Streamlit sonst beim Neustart Fehler erzeugen kann.

### Wo sehe ich das in der App?
In der Sidebar gibt es den Abschnitt:

- `Arbeitsstand`

Dort siehst du:
- wann zuletzt ein Arbeitsstand geladen wurde
- Button `Jetzt speichern`
- Button `Zurücksetzen`

### Typischer Workflow
1. App starten.
2. Die App lädt automatisch die letzte CSV, Filter und Kriterienauswahl.
3. Kurz prüfen:
   - richtige CSV?
   - richtiger Entscheidungsbaum?
   - richtige Filteransicht aktiv?
4. Weiterarbeiten.

### Wenn etwas komisch aussieht
Nutze in der Sidebar:

- `Arbeitsstand` -> `Zurücksetzen`

Danach die App neu starten oder neu laden. Die aktuellen gespeicherten Filter werden dann nicht mehr automatisch eingelesen.

### Wichtig zur Nachvollziehbarkeit
Der automatische Arbeitsstand ist nur ein Komfortspeicher. Für fachliche Nachweise weiterhin separat speichern:

- Filteransichten als JSON, wenn du sie später teilen oder belegen willst
- Entscheidungsbaum-Versionen als JSON-Version, wenn sich die Entscheidungslogik ändert
- Override-Liste als `reference_lists/manual_overrides.csv`, wenn Einzelfälle dauerhaft übersteuert werden

Kurzregel:
- `last_session_state.json` = bequem weiterarbeiten
- Filter-Preset JSON = konkrete Ansicht dokumentieren
- Entscheidungsbaum JSON = fachliche Logik versionieren
- Override-CSV = dauerhafte Einzelfallentscheidung dokumentieren

## 19) Sidebar-Arbeitsmodus
Die linke Sidebar ist bewusst nicht automatisch an den aktiven Tab gekoppelt. Stattdessen gibt es oben den stabilen `Arbeitsmodus`.

Warum?
- Streamlit-Tabs sind keine echten Seiten.
- Eine automatische Tab-Erkennung kann zu Widget- und Session-State-Konflikten führen.
- Der Arbeitsmodus ist stabiler und macht klar, welche Bedienfelder gerade relevant sind.

### Modi
- `Filter bearbeiten`: zeigt nur Auswahl der Filteransicht und die Filterfelder für A/B/C.
- `Daten & Referenzen`: zeigt CSV-Auswahl, Upload, QC-Grünliste, Overrides und Arbeitsstand.
- `Entscheidungsbaum & Szenarien`: zeigt Go-Live, Entscheidungsbaum und Live-Entscheidungsbaum.
- `Kapazität`: zeigt Aufwand- und Teamannahmen.
- `Hilfe`: zeigt Hinweise zur Anleitung.

### Wichtig
Wenn du z. B. im Tab `Phasenplan` bist, brauchst du links meistens keine Filter. Stelle dann den Arbeitsmodus auf `Kapazität` oder `Entscheidungsbaum & Szenarien`. Die Filter bleiben im Hintergrund erhalten und werden nicht gelöscht.
