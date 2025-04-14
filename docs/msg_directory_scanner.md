# Modul-Dokumentation: `msg_directory_scanner.py`

## Übersicht
Das Modul `msg_directory_scanner.py` durchsucht ein angegebenes Verzeichnis sowie dessen Unterverzeichnisse nach `.msg`-Dateien. Es erstellt eine Excel-Datei, die eine Übersicht über die gefundenen Dateien bietet, inklusive:
- Fortlaufender Nummer,
- Dateiname,
- Pfadname und
- Länge des vollständigen Pfades.

Das Modul kann direkt über die Kommandozeile aufgerufen werden und bietet verschiedene Optionen zur Anpassung der Verzeichnis- und Dateiausgabe.

---

## Funktionsweise

### Kernfunktion
```python
get_msg_files_from_directory(directory)
```
- **Zweck:** Rekursive Suche nach MSG-Dateien in einem angegebenen Verzeichnis.
- **Parameter:**
  - `directory` (str): Das Verzeichnis, das durchsucht werden soll.
- **Rückgabewert:**
  - Eine Liste mit den vollständigen Pfaden der gefundenen MSG-Dateien.
- **Verwendet:** `os.walk`, um rekursiv alle Dateien mit der Endung `.msg` zu suchen.

---

## Kommandozeilen-Argumente

`msg_directory_scanner.py` akzeptiert beim Aufruf folgende optionale Argumente:

| Argument/Flag               | Beschreibung                                                                                     | Standardwert                                                               |
|-----------------------------|-------------------------------------------------------------------------------------------------|---------------------------------------------------------------------------|
| `-d "Pfad/zum/startverzeichnis"` | Startverzeichnis für die MSG-Dateisuche.                                                     | `D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long`            |
| `-l`                        | Setzt das aktuelle Verzeichnis als Startverzeichnis für die Suche und legt die Excel-Datei dort ab. |                                                                              |
| `-o "Pfad/zum/output.xlsx"` | Speicherpfad der Excel-Datei mit der Übersicht der MSG-Dateien.                                  | `D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx` |

---

## Ablaufbeschreibung

### Initialisierung
1. Setzt Standardwerte für:
  - Startverzeichnis (`default_start_directory`)
  - Speicherpfad der Excel-Datei (`default_output_excel_file`).

### Verarbeitung
1. **Kommandozeilen-Argumente prüfen**:
  - Überprüft die übergebenen Argumente, um benutzerdefinierte Verzeichnisse oder Ausgabepfade zu setzen.
2. **Rekursive Verzeichnissuche**:
  - Methode `get_msg_files_from_directory` sucht rekursiv nach `.msg`-Dateien im angegebenen Startverzeichnis.
3. **Excel-Liste erstellen**:
  - Ruft die Hilfsmethode `create_excel_list` aus dem Modul `utils.excel_handling` auf, um eine Excel-Liste basierend auf den gefundenen Dateien zu erstellen.
4. **Excel-Datei speichern**:
  - Speichert die Liste mit der Methode `save_excel_file` am spezifizierten Ausgabepfad.
5. **Ergebnis ausgeben**:
  - Gibt die Gesamtanzahl der gefundenen `.msg`-Dateien in der Konsole aus.

---

## Beispielaufrufe

### Standard-Suche
```bash
python msg_directory_scanner.py
```
- Durchsucht das Standardverzeichnis:
  ```plaintext
  D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long
  ```
- Speichert die Excel-Datei unter:
  ```plaintext
  D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx
  ```

### Benutzerdefiniertes Startverzeichnis
```bash
python msg_directory_scanner.py -d "C:/Benutzerdaten/Dokumente/Mails"
```
- Durchsucht das Verzeichnis `C:/Benutzerdaten/Dokumente/Mails`.

### Aktuelles Verzeichnis verwenden
```bash
python msg_directory_scanner.py -l
```
- Setzt das aktuelle Verzeichnis als Startverzeichnis.
- Speichert die Excel-Datei im aktuellen Verzeichnis mit dem Namen `msg_files_overview.xlsx`.

### Benutzerdefinierte Ausgabe
```bash
python msg_directory_scanner.py -d "C:/Mails" -o "C:/Ausgaben/scan_ergebnis.xlsx"
```
- Startverzeichnis: `C:/Mails`
- Excel-Datei wird unter `C:/Ausgaben/scan_ergebnis.xlsx` gespeichert.

---

## Abhängigkeiten
### Eingesetzte Module
- **Standardmodule:**
  - `os`: Für Dateisystemoperationen.
  - `sys`: Zur Verarbeitung von Kommandozeilenargumenten.
- **Hilfsmodule:**
  - `utils.excel_handling`:
    - `create_excel_list`: Erstellt eine Excel-kompatible Liste aus den gefundenen MSG-Dateien.
    - `save_excel_file`: Speichert die erstellte Liste als Excel-Datei.

---

## Erweiterbarkeit

Das Modul kann leicht erweitert werden, um zusätzliche Funktionen hinzuzufügen:
1. **Filterkriterien:**
  - Erweiterung der Datei-Suche auf andere Dateitypen durch Änderung des Filters im `get_msg_files_from_directory`.
2. **Zusätzliche Informationen:**
  - Erweiterung der Excel-Liste mit weiteren Eigenschaften der MSG-Dateien (z. B. Datum, Größe).
3. **Verbesserte Argumentverarbeitung:**
  - Nutzung von Bibliotheken wie `argparse` für eine flexiblere und sicherere Verarbeitung der Kommandozeilenargumente.

---

## Fazit
Das Modul `msg_directory_scanner.py` bietet eine einfache aber effektive Möglichkeit, MSG-Dateien in einem Verzeichnisbaum zu suchen und deren Übersicht in einer Excel-Datei zu speichern. Es ist flexibel und kann individuell durch Kommandozeilenargumente angepasst werden.