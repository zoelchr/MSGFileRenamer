# Dokumentation für das Modul `msg_directory_scanner.py`

## Übersicht

Das Modul `msg_directory_scanner.py` durchsucht ein angegebenes Verzeichnis und alle seine Unterverzeichnisse nach MSG-Dateien. Es erstellt eine Übersicht in Form einer Excel-Datei, die Informationen über jede gefundene MSG-Datei enthält, einschließlich einer fortlaufenden Nummer, dem Dateinamen, dem Pfadnamen und der Länge des vollständigen Pfades.

## Funktionen

- **get_msg_files_from_directory(directory)**: Durchsucht das angegebene Verzeichnis nach MSG-Dateien und gibt eine Liste der gefundenen Dateien zurück.
- **create_excel_list(msg_files)**: Erstellt ein DataFrame mit Informationen über die gefundenen MSG-Dateien.
- **save_excel_file(excel_list, output_file)**: Speichert die Excel-Liste in einer angegebenen Datei.

## Verwendung

Um das Modul auszuführen, kann es direkt über die Kommandozeile aufgerufen werden. Es können folgende optionale Parameter übergeben werden:

- `-d "Pfad/zum/startverzeichnis"`: Das Verzeichnis, das durchsucht werden soll (Standard: `D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long`).
- `-l`: Das aktuelle Verzeichnis wird durchsucht und die Excel-Datei wird dort abgelegt.
- `-o "Pfad/zum/ausgabeverzeichnis"`: Der Pfad, an dem die Excel-Datei gespeichert werden soll (Standard: `D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx`).

### Beispielaufruf

```bash
python msg_directory_scanner.py -d "C:/path/to/search" -o "C:/path/to/output.xlsx"
