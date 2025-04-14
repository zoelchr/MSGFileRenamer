# Dokumentation zur ausgewählten Datei

## Übersicht
Die Datei ist ein Python-Skript, das für die Verarbeitung und Verwaltung von MSG-Dateien (Microsoft Outlook Mail-Dateien) verwendet wird. Das Skript bietet eine Vielzahl von Funktionen, darunter:
- Testdatenverarbeitung
- Umbenennung von Dateien basierend auf bestimmten Kriterien
- Anpassung der Zeitstempel von MSG-Dateien
- Generierung von Excel- und Debug-Logs
- Möglichkeit, PDFs aus MSG-Dateien zu generieren

## Importierte Module
Das Skript verwendet sowohl interne als auch benutzerdefinierte Module. Hier sind die wichtigsten Importe:
- **Standardmodule:**
    - `os`
    - `logging`
    - `datetime`
    - `argparse`
    - `pathlib.Path`
- **Benutzerdefinierte Module:**
    - `modules.msg_generate_new_filename`:
        - **Funktion:** `generate_new_msg_filename`
    - `utils.file_handling`:
        - **Funktionen:** `rename_file`, `test_file_access`, `set_file_creation_date`, `set_file_modification_date`
        - **Enums:** `FileAccessStatus`, `FileOperationResult`
    - `modules.msg_handling`:
        - **Funktionen:** `create_log_file`, `log_entry`
    - `utils.testset_preparation`:
        - **Funktion:** `prepare_test_directory`
    - `utils.pdf_generation`:
        - **Funktion:** `generate_pdf_from_msg`

## Globale Variablen
### Verzeichnisse
- **SOURCE_DIRECTORY_TEST_DATA:** Basispfad für Sample-Testdaten.
- **TARGET_DIRECTORY_TEST_DATA:** Zielverzeichnis für die Tests.
- **TARGET_DIRECTORY:** Hauptzielverzeichnis – initial leer.
- **MAX_PATH_LENGTH:** Maximale erlaubte Pfadlänge (Standard: **260**).

### Logging und Excel-Log-Verzeichnisse
- **excel_log_directory:** Verzeichnis für Excel-Logs.
- **excel_log_basename:** Basisname der Excel-Logdatei.
- **excel_log_file_path:** Kombinierter Pfad und Dateiname der Excel-Logdatei.
- **debug_log_directory:** Verzeichnis für Debug-Logs.
- **prog_log_file_path:** Kombinierter Pfad und Dateiname für das Programm-Log.

### Weitere Variablen
- **LOG_TABLE_HEADER:** Spaltennamen für die Generierung der Excel-Logdatei.

## Wichtige Funktionen
### `setup_logging`
- **Beschreibung:** Konfiguriert das Logging-System basierend auf Debug-Modus und einer angegebenen Logdatei.
- **Parameter:**
    - `file`: Pfad zur Logdatei.
    - `debug`: Debug-Modus (Boolean).

### Hauptprogramm
Das Hauptprogramm wird mit einem `if __name__ == '__main__':` Block gestartet. Es enthält folgende Kernbereiche:
1. **Argumentenparser:**
    - Verarbeitet Kommandozeilenargumente für verschiedene Funktionen wie `--no_test_run`, `--init_testdata`, `--set_filedate`, und viele weitere.
2. **Initialisierung von Verzeichnissen:**
    - Prüft und erstellt Ziel- und Testdaten-Verzeichnisse.
3. **Dateiverarbeitung:**
    - Durchläuft die angegebenen Verzeichnisse rekursiv (optional: rekursive Suche abschaltbar).
    - Überprüft Dateizugriff (Lesen/Schreiben).
    - Führt ggf. verschiedene Operationen durch, darunter:
        - Umbenennen von MSG-Dateien.
        - Löschen von doppelten Dateien.
        - Setzen von Zeitstempeln auf das Versanddatum der E-Mail.
        - Generieren eines PDFs (optional).

### Kernelemente der Verarbeitung
#### Verarbeitung der Dateien
- Überprüfung der Dateiendung `.msg`.
- Prüfung des Zugriffs (Lesen/Schreiben) mit `test_file_access`.
- Generieren eines neuen Dateinamens mit `generate_new_msg_filename`.
- Umbenennen der Datei mit `rename_file` (abhängig vom Testlauf-Modus).
- Optionale Anpassung von Erstellungs- und Änderungsdatum mit `set_file_creation_date` und `set_file_modification_date`.

#### Logging
- Protokolliert die Verarbeitungsergebnisse in einer Excel-Datei sowie einer Debug-Logdatei.

### Schleifensteuerung
Nach Abschluss der rekursiven Suche wird die Schleife beendet.

## Kommandozeilenparameter
Das Skript unterstützt zahlreiche Parameter, die über die Kommandozeile übergeben werden können:

| Parameter                     | Beschreibung                                                                                       | Beispielwert          |
|-------------------------------|---------------------------------------------------------------------------------------------------|-----------------------|
| `--no_test_run` / `-ntr`      | Testmodus aktivieren (ohne Dateioperationen).                                                    | `False`              |
| `--init_testdata` / `-it`     | Zielverzeichnis initialisieren und mit Testdaten füllen.                                         | `True`               |
| `--set_filedate` / `-fd`      | Zeitstempel der Dateien anpassen.                                                                | `False`              |
| `--debug_mode` / `-db`        | Debug-Modus (zusätzliche Logs).                                                                  | `True`               |
| `--search_directory` / `-sd`  | Verzeichnis, das durchsucht werden soll.                                                          | `./data/sample_files`|
| `--excel_log_basename` / `-elb` | Basisname für die Logdatei im Excel-Format.                                                     | `log_file`           |
| `--excel_log_directory` / `-elf` | Zielverzeichnis der Excel-Logdatei.                                                           | `./`                 |
| `--debug_log_directory` / `-dlf` | Zielverzeichnis der Debug-Logdatei.                                                           | `./logs`             |
| `--no_shorten_path_name` / `-spn` | Pfadlängenbegrenzung deaktivieren.                                                           | `False`              |
| `--generate_pdf` / `-pdf`     | Aus MSG-Dateien PDFs generieren.                                                                | `True`               |
| `--overwrite_pdf` / `-opdf`   | Bereits existierende PDFs überschreiben.                                                        | `False`              |
| `--recursive_search` / `-rs`  | Verzeichnis rekursiv nach MSG-Dateien durchsuchen.                                              | `False`              |

## Ergebnisse
Am Ende der Verarbeitung erstellt das Skript eine Auswertung in der Konsole sowie in den Logs. Die wichtigsten Kennzahlen umfassen:
- Anzahl gefundener MSG-Dateien
- Anzahl erfolgreich umbenannter Dateien
- Anzahl aufgetretener Probleme
- Anzahl gekürzter Dateinamen
- Statistiken über doppelte Dateien und gelöschte Duplikate

## Benutzungsbeispiel
```bash
python script.py --init_testdata --set_filedate --debug_mode --generate_pdf --recursive_search
```

Dieses Kommando initialisiert die Testdaten, passt Zeitstempel an, aktiviert den Debug-Modus, generiert PDFs und führt eine rekursive Suche durch.

---

## Fazit
Dieses Skript ist äußerst flexibel und bietet eine Vielzahl von Funktionen zur Verarbeitung und Verwaltung von MSG-Dateien. Die saubere Struktur ermöglicht eine einfache Anpassung an individuelle Anforderungen, wie z.B. die Integration zusätzlicher Parameter oder die Erweiterung der Dateiverarbeitung.