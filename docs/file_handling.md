# Beschreibung: file_handling.py

## Übersicht

Das Modul `file_handling.py` enthält eine Vielzahl von Funktionen für Datei- und Verzeichnisoperationen unter Windows. Es unterstützt insbesondere die Manipulation von Datei-Metadaten (Erstell-/Änderungsdatum) sowie das sichere Umbenennen, Löschen und Kopieren von Dateien und Verzeichnissen.

---

## Enthaltene Funktionen

| Funktion | Beschreibung |
|----------|--------------|
| `test_file_access(file_path)` | Prüft, ob eine Datei lesbar, schreibbar, ausführbar oder gesperrt ist. |
| `rename_file(current_name, new_name)` | Benennt eine Datei um, mit Wiederholversuchen bei Fehlern. |
| `delete_file(file_path)` | Löscht eine Datei mit optionalen Wiederholversuchen. |
| `sanitize_filename(filename)` | Entfernt ungültige Zeichen aus Dateinamen. |
| `format_datetime_stamp(datetime_stamp, format_string)` | Wandelt einen Zeitstempel in einen formatierten String um. |
| `set_file_modification_date(file_path, new_date)` | Setzt das Änderungsdatum einer Datei. |
| `set_file_creation_date(file_path, new_creation_date)` | Setzt das Erstellungsdatum einer Datei (Windows-spezifisch). |
| `delete_directory_contents(directory_path)` | Löscht den gesamten Inhalt eines Verzeichnisses. |
| `copy_directory_contents(source_directory_path, target_directory_path)` | Kopiert Inhalte eines Verzeichnisses rekursiv in ein anderes. |

---

## Verwendete Klassen und Enums

### `FileAccessStatus (Enum)`
- READABLE, WRITABLE, EXECUTABLE, NOT_FOUND, NO_PERMISSION, LOCKED, UNKNOWN_ERROR, UNKNOWN

### `FileOperationResult (Enum)`
- SUCCESS, FILE_NOT_FOUND, DESTINATION_EXISTS, PERMISSION_DENIED, INVALID_FILENAME, UNKNOWN_ERROR, etc.

### `FileHandle (Class)`
Ein Kontextmanager zur Verwendung nativer Windows-Handles für Dateizugriff über `win32file`.

---

## Abhängigkeiten

- `os`, `time`, `shutil`, `re`, `datetime`, `logging`
- Windows-spezifisch:
  - `pywintypes`
  - `win32file`
  - `win32con`

---

## Plattformhinweis

Dieses Modul ist speziell für **Windows** konzipiert und verwendet APIs aus `pywin32`. Es funktioniert nicht unter Linux oder macOS ohne Modifikationen.

---

## Anwendungsbeispiel

```python
from file_handling import delete_file, FileOperationResult

result = delete_file("beispiel.txt")
if result == FileOperationResult.SUCCESS:
    print("Datei erfolgreich gelöscht.")
```

---

Erstellt aus dem Quellcode `file_handling.py`.
