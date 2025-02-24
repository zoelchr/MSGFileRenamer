# test_file_handling.py

Dieses Modul enthält Tests für die Funktionen im Modul `file_handling.py`. Es stellt sicher, dass die Funktionen korrekt arbeiten und die erwarteten Ergebnisse liefern.

## Inhaltsverzeichnis

- [Einführung](#einführung)
- [Verwendung](#verwendung)
- [Funktionen](#funktionen)
- [Testablauf](#testablauf)
- [Ergebnisse](#ergebnisse)

## Einführung

Das `test_file_handling.py`-Modul führt eine Reihe von Tests durch, um die Funktionalität der Dateioperationen zu überprüfen, die im `file_handling.py`-Modul definiert sind. Die Tests umfassen das Löschen, Kopieren und Umbenennen von Dateien sowie das Setzen von Erstellungs- und Änderungsdaten.

## Verwendung

Um die Tests auszuführen, verwenden Sie den folgenden Befehl in der Kommandozeile:

```bash
python -m unittest test_file_handling.py
```

## Funktionen

Die wichtigsten Funktionen, die in diesem Modul getestet werden, sind:

- **delete_directory_contents**: Löscht den Inhalt eines angegebenen Verzeichnisses.
- **copy_directory_contents**: Kopiert den Inhalt eines Quellverzeichnisses in ein Zielverzeichnis.
- **rename_file**: Benennt eine Datei um.
- **delete_file**: Löscht eine angegebene Datei.
- **format_datetime_stamp**: Formatiert einen Zeitstempel in ein bestimmtes Format.
- **set_file_creation_date**: Setzt das Erstelldatum einer Datei.
- **set_file_date**: Setzt das Änderungsdatum einer Datei.
- **test_file_access**: Überprüft den Zugriff auf eine Datei.
- **get_date_sent_msg_file**: Ruft das Versanddatum aus einer MSG-Datei ab.

## Testablauf

1. **Verzeichnisse definieren**: Das Quell- und Zielverzeichnis für die Tests werden festgelegt.
2. **Zielverzeichnis erstellen**: Wenn das Zielverzeichnis nicht existiert, wird es erstellt.
3. **Inhalt löschen**: Der Inhalt des Zielverzeichnisses wird gelöscht.
4. **Inhalt kopieren**: Der Inhalt des Quellverzeichnisses wird in das Zielverzeichnis kopiert.
5. **Durchsuchen nach MSG-Dateien**: Das Zielverzeichnis wird nach Dateien mit der Endung `.msg` durchsucht.
6. **Dateizugriff prüfen**: Der Zugriff auf jede gefundene Datei wird überprüft.
7. **Versanddatum abrufen**: Das Versanddatum wird aus der MSG-Datei abgerufen.
8. **Datei umbenennen**: Die Datei wird basierend auf dem Versanddatum umbenannt.
9. **Daten setzen**: Das Erstelldatum und das Änderungsdatum der Datei werden gesetzt.
10. **Ergebnisse ausgeben**: Am Ende werden die Anzahl der umbenannten Dateien und der Dateien mit Problemen ausgegeben.

## Ergebnisse

Nach Abschluss der Tests gibt das Skript die Anzahl der erfolgreich umbenannten Dateien und die Anzahl der Dateien mit Problemen aus. Dies ermöglicht eine schnelle Überprüfung der Funktionalität des `file_handling.py`-Moduls.
