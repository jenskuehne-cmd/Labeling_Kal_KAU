# Phasenplan-Logik

## Zweck
Diese Logik dient dazu, Messstellen aus einer CSV-Datei automatisiert in Prioritaeten und Phasen zu ueberfuehren. Das Ergebnis ist ein Phasenplan, der fuer die Personalplanung, die Priorisierung und die Kommunikation des Ablaufplans im Zeitverlauf verwendet werden kann.

Die Logik ist so aufgebaut, dass sie:
- die relevante Grundgesamtheit bestimmt,
- Messstellen nach fachlichen Regeln priorisiert,
- den Aufwand je Phase schaetzt,
- die Ergebnisse exportierbar macht,
- und so dokumentiert, dass eine KI daraus eine saubere Zusammenstellung, Management-Zusammenfassung oder Praesentation ableiten kann.

## Eingaben
Die App arbeitet mit einer CSV-Datei mit Messstellen und optionalen Zusatzdateien oder Presets.

Wesentliche Eingabespalten sind typischerweise:
- `Asset ID`
- `Due Date BMRAM / End Datum SAP`
- `Start Datum SAP Auftrag`
- `Interval`
- `Messstellen-Beschreibung`
- `Zugänglichkeit`
- `Gebäude / MU`
- `Status`
- `Kategorie`

Zusatzinputs:
- QC-Gruenliste fuer bereits durch QC selbst gelabelte Messstellen
- manuelle Overrides
- gespeicherte Filter-Presets je Ansicht
- Entscheidungbaum-Konfiguration

## Grundlogik
Die Priorisierung laeuft in einer festen Reihenfolge:

1. QC-/ITOT-Ausnahmen werden zuerst entfernt.
2. Zeitliche Einordnung erfolgt ueber das Due Date.
3. Legacy-Klasse wird ueber die Laenge der Asset-ID abgeleitet.
4. Zugaenglichkeit wird ueber Textmuster klassifiziert.
5. Daraus wird die Prioritaet `P1` bis `P6` bestimmt.
6. Aus den Prioritaeten werden Phasen und Zeitfenster abgeleitet.

## Zeitlogik
Die Zeitlogik orientiert sich am Due Date und nicht mehr nur an einer reinen Monatsnaeherung.

Standardmaessig gilt:
- `MS_Time_iScope`: Due Date liegt zwischen `2026-08-10` und `2027-01-31`
- `T0_Immediate`: Due Date liegt vor dem Start des iScope-Fensters und das Intervall ist kleiner als der Immediate-Schwellenwert
- `MS_Time_Long`: Due Date liegt nach dem Ende des iScope-Fensters

Das Zeitfenster kann in der Entscheidungsbaum-Konfiguration angepasst werden.

## Legacy-Logik
Die Legacy-Klasse wird ueber die Asset-ID-Laenge gebildet:
- `MS_legacy_krit`: Asset ID laenger als der konfigurierte Schwellwert, standardmaessig `> 30`
- `MS_legacy_iO`: Asset ID bis zum Schwellwert

Diese Klassifikation ist wichtig, weil sie trennt zwischen:
- komplexeren, eher kritischen Messstellen
- und handhabbaren Messstellen mit kuerzerer ID

## Zugaenglichkeits-Logik
Die Zugaenglichkeit wird textbasiert klassifiziert:
- `ABC` = schwer zugaenglich
- `easy` = einfach zugaenglich
- `unknown` = unklar oder nicht eindeutig zuordenbar

Die Zuordnung erfolgt ueber definierte Textmuster in der Beschreibung bzw. Zugaenglichkeitsspalte.

## Prioritaeten
Die Zielprioritaeten sind:

- `P1`: lange Legacy-ID, im iScope-Fenster, schwer zugaenglich
- `P2`: lange Legacy-ID, im iScope-Fenster, einfach oder unklar zugaenglich
- `P2A`: lange Legacy-ID, nach dem iScope-Fenster, einfach zugaenglich
- `P3`: lange Legacy-ID, nach dem iScope-Fenster, schwer zugaenglich
- `P4`: kuerzere Legacy-ID, schwer zugaenglich
- `P5`: kuerzere Legacy-ID, einfach zugaenglich
- `P6`: Rest / nicht eindeutig zugeordnet

Die genaue Regelreihenfolge kann im Entscheidungsbaum angepasst werden.

## Phasen
Aus der Prioritaet werden Relabeling-Phasen gebildet, zum Beispiel:
- `Shutdown_26`
- `Klaerung_Zugaenglichkeit`
- `PostgL_bis_Oktober`
- `PostgL_bis_Januar`
- `Shutdown_27`
- `Opportunistisch_im_Shutdown`
- `Backlog_optional`

Die Phasen sind fuer die operative Planung gedacht und stellen die Reihenfolge dar, in der Messstellen bearbeitet werden sollen.

## Aufwand und Kapazitaet
Der Phasenplan rechnet die Messstellen in Aufwand um.

Typische Kennzahlen:
- Anzahl Messstellen pro Phase und Prioritaet
- Aufwand mit 7 Minuten pro Messstelle
- Aufwand mit 12 Minuten pro Messstelle
- Aufwand plus Puffer
- Personentage
- Wochenbedarf

Die Kapazitaetsannahmen koennen in der Sidebar angepasst werden:
- Minuten pro Messstelle
- Puffer in Prozent
- Anzahl Personen
- Arbeitsstunden pro Person und Tag

## Exportierte Ergebnisse
Der Excel-Export enthaelt mehrere relevante Reiter:

- `Start`: Zusammenfassung
- `phase_plan_aggregated`: aggregierter Phasenplan
- `raw_classified`: Rohdaten mit abgeleiteten Klassifikationen
- `metadata`: technische Metadaten
- `decision_tree_rules`: aktive Regeln
- `decision_tree_steps`: Entscheidungsbaum-Zwischenschritte
- `decision_tree_step_detail`: Ausschluss-/Move-Log
- `decision_tree_path_detail`: kompletter Pfad je Messstelle
- `decision_tree_funnel`: komprimierte Wasserfall-Zusammenfassung
- `decision_tree_funnel_steps`: Wasserfall ueber Zwischenschritte

## Wie die KI den Export nutzen sollte
Wenn du die Excel-Datei an eine KI gibst, sollte sie daraus drei Ebenen erzeugen:

### 1. Executive Summary
Die KI soll in wenigen Saetzen beantworten:
- wie viele Messstellen in Scope sind,
- welche Prioritaeten dominieren,
- welche Phasen den hoechsten Aufwand erzeugen,
- wo die kritischen Engpaesse liegen,
- und was das fuer den Zeitplan bedeutet.

### 2. Management-Praesentation
Die KI soll daraus eine Praesentation aufbauen mit:
- Ziel und Methode
- Datengrundlage
- Entscheidungslogik
- Prioritaetenverteilung
- Phasen- und Aufwandssicht
- Zeitlicher Ablauf
- Risiken und Abhaengigkeiten
- naechste Schritte

### 3. Personalplanung
Die KI soll aus den Phasen und dem Aufwand ableiten:
- welche Phasen zuerst besetzt werden muessen,
- wo hohe Prioritaeten eine fruehe Bearbeitung brauchen,
- wie sich die Last ueber den Zeitraum verteilt,
- und welche Kapazitaet pro Zeitraum benoetigt wird.

## Empfohlene Analysefragen fuer eine KI
Du kannst der KI zum Beispiel folgende Fragen geben:

1. Welche Messstellen sind in P1 bis P3 und warum?
2. Wie verteilt sich der Aufwand auf die Phasen?
3. Welche Phasen sind zeitkritisch und muessen zuerst bearbeitet werden?
4. Welche Datenqualitaets- oder Klassifikationsrisiken gibt es?
5. Wie sollte die Personalplanung ueber den Zeitverlauf aussehen?
6. Welche Kernaussagen eignen sich fuer eine Management-Praesentation?

## Wichtiger Hinweis
Die Logik ist konfigurierbar. Wenn sich die fachlichen Regeln aendern, kann und soll der Entscheidungsbaum angepasst werden. Die Doku beschreibt den aktuellen Stand der Logik, nicht eine unveraenderliche Regel.

