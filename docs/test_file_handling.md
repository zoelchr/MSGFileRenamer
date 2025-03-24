# Beschreibung: test_file_handling.py

## Übersicht

Das Modul `test_file_handling.py` enthält umfangreiche Funktionstests für das Modul `file_handling.py`. Es simuliert realistische Arbeitsabläufe zur Prüfung von Dateioperationen wie Umbenennen, Erstell-/Änderungsdatum setzen und Datei-Zugriff.

---

## Zweck der Tests

- **Validierung** von Funktionen wie:
  - `rename_file`
  - `set_file_creation_date`
  - `set_file_modification_date`
  - `test_file_access`
  - `format_datetime_stamp`
- **Dateizugriff** (lesend, schreibend, gesperrt)
- **Umbenennung** von MSG-Dateien auf Basis von Metadaten
- **Setzen von Datei-Zeitstempeln** entsprechend dem Versanddatum einer Nachricht
- **Logging** von Programmverlauf und Fehlern

---

## Aufbau des Tests

- Verwendung eines **Quellverzeichnisses** mit MSG-Dateien
- Kopie der Dateien in ein **temporäres Testverzeichnis**
- Durchführung der Prüfungen auf jedem MSG-File im Zielverzeichnis

---

## Abhängigkeiten

- `file_handling` (aus `utils`):
  - Dateioperationen
- `msg_handling` (aus `modules`):
  - Extraktion von Versanddatum
- `testset_preparation`:
  - Vorbereitung des Zielverzeichnisses

---

## Logging

- Logdatei:  
  `D:/Dev/pycharm/MSGFileRenamer/logs/test_file_handling_prog_log.log`
- Log-Level über `DEBUG_MODE = True` steuerbar
- Ausgaben sowohl in Konsole als auch in Logdatei

---

## Beispielausgabe

```plaintext
MSG-Datei: beispiel.msg
Ergebnis Zugriffsprüfung: ['Readable', 'Writable']
Versandzeitpunkt: 2024-03-22 15:00:00
Neuer Dateiname: 20240322-15uhr00_beispiel.msg
Neues Erstellungsdatum gesetzt.
Neues Änderungsdatum gesetzt.
```

---

## Ausführung

Das Skript kann direkt ausgeführt werden:

```bash
python test_file_handling.py
```

oder via Unittest:

```bash
python -m unittest test_file_handling.py
```

---

Erstellt aus dem Quellcode `test_file_handling.py`.
