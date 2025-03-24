# Beschreibung: test_rename_handling.py

## Übersicht

Das Modul `test_rename_handling.py` führt umfangreiche Funktionstests zur Umbenennung von MSG-Dateien durch. Ziel ist die Validierung der automatisierten Umbenennung basierend auf Metadaten wie Absender, Betreff und Versanddatum sowie die strukturierte Protokollierung dieser Vorgänge.

---

## Hauptfunktionen

- Initialisierung von Testverzeichnissen mit Beispieldaten
- Prüfung des Datei-Zugriffs (lesen/schreiben)
- Generierung eines neuen Dateinamens basierend auf Metadaten
- Kürzen des Dateinamens bei Überschreitung der Pfadlänge
- Umbenennung der Datei mit Konflikterkennung (Doubletten)
- Setzen von Erstellungs- und Änderungsdatum
- Logging aller Vorgänge in eine Excel-Logdatei

---

## Konfigurierbare Optionen

| Variable           | Funktion                                                |
|--------------------|---------------------------------------------------------|
| `DEBUG_MODE`       | Aktiviert detailliertes Logging                         |
| `TEST_RUN`         | Führt den Testlauf ohne tatsächliche Dateioperationen aus |
| `SET_FILEDATE`     | Steuert, ob Datei-Zeitstempel gesetzt werden sollen     |
| `INIT_TESTDATA`    | Steuert, ob Testdaten initialisiert werden              |
| `max_path_length`  | Maximale Pfadlänge (Standard: 260)                      |

---

## Protokollierung

- Konsolen- und Datei-Logging (`test_rename_prog_log.log`)
- Excel-Log mit folgenden Spalten:
  - Nummer
  - Verzeichnis
  - Original- & neuer Dateiname
  - Betreff (roh & bereinigt)
  - Versanddatum (roh & formatiert)
  - Absendername & E-Mail
  - Flag zur Dateinamenskürzung

---

## Abhängigkeiten

- `msg_generate_new_filename`
- `msg_handling`
- `file_handling`
- `testset_preparation`

---

## Beispielausgabe

```plaintext
Aktuelle MSG-Datei: projektstatus.msg
Neuer Dateiname: 20240324-12uhr00_max.mustermann@example.com_Besprechung.msg
Neues Erstellungsdatum erfolgreich gesetzt.
Neues Änderungsdatum erfolgreich gesetzt.
```

---

## Abschlussstatistik

Nach dem Lauf werden unter anderem folgende Kennzahlen ausgegeben:

- Anzahl umbenannter Dateien
- Anzahl gekürzter Dateinamen
- Anzahl gefundener Doubletten
- Anzahl erfolgreich gesetzter Datei-Zeitstempel

---

## Ausführung

```bash
python test_rename_handling.py
```

---

Erstellt aus dem Quellcode `test_rename_handling.py`.
