# KAU Relabeling Priorisierung

Lokale Streamlit-App zur Priorisierung von Relabeling-Messstellen fuer KAU.  
Die App trennt:
- Exploration und Filterung
- fachliche Entscheidungslogik
- Referenzlisten und manuelle Overrides
- Auswertung, Vergleich und Export

## Schnellstart

```bash
cd "/Users/kuehnej/Documents/Codex/Labeling_Prio_Festlegung"
source "/Users/kuehnej/role-dashboard/.venv/bin/activate"
streamlit run app.py
```

## So gehst du vor

1. App starten.
2. Links unter `Arbeitsmodus` die passende Sidebar-Ansicht waehlen.
3. Unter `Daten & Referenzen` die CSV auswaehlen oder hochladen.
4. Unter `Entscheidungsbaum & Szenarien` den aktiven Entscheidungsbaum pruefen.
5. In den Tabs rechts filtern, vergleichen, pruefen und exportieren.

## Was die App kann

- CSV laden, anzeigen und exportieren
- Filteransichten A/B/C mit Regex, Kategorien, numerischen und Datumsfiltern
- versionierte Entscheidungsbaum-Versionen als JSON
- QC-Referenzliste und manuelle Overrides
- Entscheidungsbaum-Vergleich zwischen Versionen
- Phasenplan und Kapazitaetsauswertung
- Hilfe- und Erklaerungstabs fuer Einsteiger

## Wichtige Dokumentation

- [ANLEITUNG.md](ANLEITUNG.md)
- [KLICKANLEITUNG.md](KLICKANLEITUNG.md)
- [To Do und vorgehen Logik.md](To%20Do%20und%20vorgehen%20Logik.md)
- [THREAD_UMSETZUNGSPLAN.md](THREAD_UMSETZUNGSPLAN.md)
- [THREAD_APP_ABGLEICH.md](THREAD_APP_ABGLEICH.md)
- [STREAMLIT_RELABELING_THREAD_HANDBUCH.md](STREAMLIT_RELABELING_THREAD_HANDBUCH.md)

## Projektstruktur

- `app.py` - Streamlit-App
- `data/` - lokale CSV-Daten
- `criteria_sets/` - versionierte Entscheidungsbaum-Versionen
- `reference_lists/` - QC-Listen und Overrides
- `runs/` - automatisch gespeicherter Arbeitsstand
- `ANLEITUNG.md` - fachliche Anleitung
- `KLICKANLEITUNG.md` - Klickpfad fuer den Alltag

## Arbeitslogik in kurz

- `Filteransichten` sind fuer Sichtung und Analyse.
- `Entscheidungsbaum` ist die offizielle Entscheidungslogik.
- `QC` und `Overrides` sind gezielte Fachausnahmen.
- `Vergleich` zeigt die Wirkung zwischen zwei Entscheidungsbaum-Versionen.

## Typische Reihenfolge fuer neue Analysen

1. CSV auswaehlen.
2. Aktiven Entscheidungsbaum pruefen.
3. Filteransicht A als Baseline nutzen.
4. Eine Aenderung je Iteration machen.
5. Entscheidungsbaum-Version speichern.
6. Mit `Vergleich` gegen die vorherige Version pruefen.
7. Ergebnis dokumentieren.

## Hinweise

- Wenn kein Entscheidungsbaum aktiv ist, greift nur die manuelle Filterung.
- Geaenderte Entscheidungsbaum-Versionen sollten immer als neue JSON-Version gespeichert werden.
- Der automatische Arbeitsstand wird in `runs/last_session_state.json` abgelegt und ist nicht fuer das manuelle Editieren gedacht.
