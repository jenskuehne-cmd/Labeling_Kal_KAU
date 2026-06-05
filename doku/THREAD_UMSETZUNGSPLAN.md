# Umsetzungsplan: Baue Streamlit-Relabeling-App

## Zielbild
Die App soll eine lokale Streamlit-Anwendung fuer die Priorisierung von Relabeling-Messstellen sein. Der Thread hatte als Einstieg eine schlanke CSV-Filter-App vorgesehen. Das aktuelle Projekt ist bereits weiter und enthaelt schon Kriterienlogik, QC-Listen, Overrides und mehrere Ansichten.

## Leitprinzipien
- Erst sauber sichten, dann erweitern.
- Filterung und Entscheidungslogik getrennt halten.
- Neue Logik nur dann in die App ziehen, wenn sie reproduzierbar und versionierbar ist.
- Keine Widgets oder Keys leichtfertig umbauen, wenn sie bereits gespeicherten Zustand tragen.

## Reihenfolge fuer die weitere Arbeit

### Phase 1: Basis stabil halten
- Aktuelle CSV-Auswahl und Upload-Logik beibehalten.
- Bestehende Filteransichten A/B/C nicht umbauen, solange die Benutzerfuehrung noch funktioniert.
- Vor jeder funktionalen Erweiterung einen kurzen Syntax- und State-Check machen.

### Phase 2: Filter-Schicht sauber strukturieren
- Textfilter, Regex Include/Exclude, Kategoriefilter, numerische Filter und Datumsfilter klar als Filterebene behandeln.
- Filterzustand pro Ansicht separat speichern.
- Export der gefilterten Daten als CSV und optional als ZIP mit Filterzustand beibehalten.

### Phase 3: Entscheidungslogik ausbauen
- Entscheidungsbaum-Versionen als versionierte JSON-Dateien weiterpflegen.
- Regeltypen klar trennen:
  - Ausschlussregeln
  - Prio-Zuweisung
  - Unterbruch / Shutdown-Markierungen
  - Referenzlisten-Regeln
- Entscheidungsbaum-Builder nur dann erweitern, wenn die Fachlogik dafuer wirklich gebraucht wird.

### Phase 4: Fachliche Sonderfaelle integrieren
- QC-Ausschluesse weiter ueber separate Referenzlisten oder Spalten abbilden.
- Manuelle Overrides als bewusstes Ausnahmeinstrument behalten.
- Szenarien / Varianten nicht im gleichen Mechanismus wie Filteransichten vermischen.

### Phase 5: Auswertung und Nachvollziehbarkeit
- Vergleich von Entscheidungsbaum-Versionen weiter ausbauen, wenn fachlich noetig.
- Ausschlusslisten, Trefferlisten und Ergebnislisten transparent halten.
- In der Anleitung klar dokumentieren, welche Ebene was tut.

## Was aus dem Thread noch als Minimalziel gilt
Wenn man den urspruenglichen Thread woechentlich neu schneiden wuerde, dann waeren diese Punkte die Kernbasis:
- CSV hochladen oder aus lokalem Ordner waehlen
- Spalten anzeigen
- einfache Textfilter
- Regex Include / Regex Exclude
- gefilterte Daten exportieren

## Was jetzt nicht als naechster Schritt sinnvoll ist
- Kein kompletter Umbau der App auf einen anderen State-Mechanismus.
- Keine automatische Tab-Erkennung als Steuerung der Sidebar.
- Keine Vermischung von UI-Filterung und fachlicher Prio-Logik.

## Praktische Regel fuer naechste Aenderungen
Jede neue Aenderung sollte vorab in eine dieser drei Schubladen passen:
1. Anzeige / Navigation
2. Filterung / Exploration
3. Kriterien / Entscheidung

Wenn eine Aenderung in mehr als eine Schublade faellt, sollte sie kleiner geschnitten werden.
