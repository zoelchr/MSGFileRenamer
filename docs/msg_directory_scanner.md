# msg_directory_scanner.py

Dieses Modul durchsucht ein angegebenes Verzeichnis und alle seine Unterverzeichnisse nach MSG-Dateien.
Es erstellt eine Übersicht in Form einer Excel-Datei, die Informationen über jede gefundene MSG-Datei enthält,
einschließlich einer fortlaufenden Nummer, dem Dateinamen, dem Pfadnamen und der Länge des vollständigen Pfades.

## Inhaltsverzeichnis

- [Einführung](#einführung)
- [Verwendung](#verwendung)
- [Funktionen](#funktionen)
- [Beispielaufruf](#beispielaufruf)

## Einführung

Das `msg_directory_scanner.py`-Modul ermöglicht es, MSG-Dateien in einem bestimmten Verzeichnis zu finden und deren Metadaten in einer Excel-Datei zu speichern. Es bietet eine einfache Möglichkeit, Informationen über MSG-Dateien zu sammeln und zu organisieren.

## Verwendung

Um das Modul auszuführen, kann es direkt über die Kommandozeile aufgerufen werden. Es können zwei optionale Parameter übergeben werden:

- **-d "Pfad/zum/startverzeichnis"**: Das Verzeichnis, das durchsucht werden soll (Standard: `D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long`).
- **-l**: Das aktuelle Verzeichnis wird durchsucht und die Excel-Datei wird dort abgelegt.
- **-o "Pfad/zum/ausgabeverzeichnis"**: Der Pfad, an dem die Excel-Datei gespeichert werden soll (Standard: `D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx`).

## Funktionen

### 1. `get_msg_files_from_directory(directory)`
Durchsucht das angegebene Verzeichnis und alle Unterverzeichnisse nach MSG-Dateien.

- **Parameter**: 
  - `directory` (str): Das Verzeichnis, das durchsucht werden soll.
- **Rückgabewert**: 
  - list: Eine Liste der gefundenen MSG-Dateien (vollständige Pfade).

### Beispielaufruf

Um das Skript auszuführen, verwenden Sie den folgenden Befehl in der Kommandozeile:

- Beispiel für die Suche in einem bestimmten Verzeichnis:
  - `python msg_directory_scanner.py -d "C:/path/to/search" -o "C:/path/to/output.xlsx"`
  
- Beispiel für die Verwendung des aktuellen Verzeichnisses:
  - `python msg_directory_scanner.py -l`

## Ergebnisse

Nach Abschluss der Ausführung gibt das Skript die Anzahl der gefundenen MSG-Dateien aus. Die gesammelten Informationen werden in einer Excel-Datei gespeichert, die im angegebenen Verzeichnis abgelegt wird.
