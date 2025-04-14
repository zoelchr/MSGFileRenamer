# MSG Directory Scanner

Ein einfaches, aber leistungsstarkes Python-Skript, das ein Verzeichnis und alle darin enthaltenen Unterverzeichnisse rekursiv nach `.msg`-Dateien durchsucht. Das Ergebnis wird als Excel-Übersicht gespeichert, die nützliche Informationen wie Dateinamen, Pfade und Pfadlängen enthält.

---

## Features

- **Rekursive Verzeichnissuche**: Durchsucht ein angegebenes Verzeichnis und alle Unterverzeichnisse nach `.msg`-Dateien.
- **Excel-Export**: Erstellt eine übersichtliche Excel-Datei für die gefundenen Results.
- **Flexible Konfiguration**: Start- und Ausgabe-Verzeichnisse können individuell angegeben werden.
- **Einfach zu benutzen**: Klares CLI-Interface für einfache Bedienung.

---

## Voraussetzungen

Stelle vor der Nutzung sicher, dass folgende Voraussetzungen erfüllt sind:

### Software
- Python 3.7 oder neuer

### Abhängigkeiten
Das Modul benötigt externe Python-Bibliotheken, die wie folgt installiert werden können:
```bash
pip install openpyxl
```

---

## Installation

1. **Repository klonen**:
   ```bash
   git clone https://github.com/<dein-benutzername>/msg-directory-scanner.git
   cd msg-directory-scanner
   ```

2. **Benötigte Abhängigkeiten installieren**:
   ```bash
   pip install -r requirements.txt
   ```

---

## Verwendung

Das Skript kann einfach über die Kommandozeile ausgeführt werden:

### Basisaufruf
```bash
python msg_directory_scanner.py
```
- Startverzeichnis: Standardpfad (`D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long`)
- Ergebnis: Excel-Datei unter `D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx`

### Optionen
| Argument                     | Beschreibung                                                                                          | Standardwert                                    |
|------------------------------|------------------------------------------------------------------------------------------------------|------------------------------------------------|
| `-d "Pfad/zum/startverzeichnis"` | Gibt das Verzeichnis an, das durchsucht werden soll.                                                  | `D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long` |
| `-l`                         | Setzt das aktuelle Verzeichnis als Startverzeichnis und legt die Excel-Datei dort ab.                 |                                                |
| `-o "Pfad/zum/output.xlsx"`  | Gibt das Verzeichnis und den Dateinamen an, unter dem die Excel-Ausgabe gespeichert werden soll.       | `D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx` |

### Beispiele
1. **Startverzeichnis angeben**:
   ```bash
   python msg_directory_scanner.py -d "/path/to/search"
   ```

2. **Excel-Datei in einem benutzerdefinierten Verzeichnis speichern**:
   ```bash
   python msg_directory_scanner.py -d "/path/to/search" -o "/path/to/output.xlsx"
   ```

3. **Aktuelles Verzeichnis verwenden**:
   ```bash
   python msg_directory_scanner.py -l
   ```

---

## Ausgabe

Das Skript speichert die Ergebnisse in einer Excel-Datei mit folgenden Spalten:
- **Fortlaufende Nummer**: Identifiziert jede gefundene MSG-Datei eindeutig.
- **Dateiname**: Name der gefundenen Datei.
- **Pfad**: Vollständiger Pfad zur Datei.
- **Pfadlänge**: Die Länge des vollständigen Pfades.

---

## Beispielausgabe

Nach dem Durchsuchen eines Verzeichnisses wird eine Excel-Datei generiert, z. B.:

| Fortlaufende Nummer | Dateiname          | Pfad                                  | Pfadlänge |
|---------------------|--------------------|---------------------------------------|-----------|
| 1                   | Beispiel1.msg     | C:/Mails/Beispiel1.msg               | 25        |
| 2                   | Beispiel2.msg     | C:/Mails/Unterverzeichnis/Beispiel2.msg | 45        |

---

## Ordnerstruktur

Die Ordnerstruktur des Projekts:

```plaintext
msg-directory-scanner/
├── modules/
│   ├── msg_directory_scanner.py  # Hauptmodul
│   ├── utils/
│       ├── excel_handling.py     # Hilfsmodul für die Excel-Verarbeitung
├── logs/                         # Standardausgabeordner der Excel-Datei
├── data/sample_files/            # Beispiel-Testdaten (MSG-Dateien)
├── README.md                     # Diese Dokumentation
├── requirements.txt              # Abhängigkeiten
```

---

## Beiträge

Beiträge und Verbesserungsvorschläge sind willkommen! Sei es durch das Melden von Fehlern, Vorschläge für neue Features oder direkte Beiträge zum Code.

1. Erstelle einen neuen Branch:
   ```bash
   git checkout -b feature/neues-feature
   ```

2. Nimm deine Änderungen vor und committe diese:
   ```bash
   git commit -m "Beschreibung der Änderungen"
   ```

3. Push deine Änderungen:
   ```bash
   git push origin feature/neues-feature
   ```

4. Erstelle einen Pull Request.

---

## Lizenz

Dieses Projekt ist unter der **MIT-Lizenz** lizenziert – siehe die [LICENSE](LICENSE) Datei für Details.

---

## Autor

**Rüdiger Zölch**  
**Email:** ruediger@zoelch.me   
**GitHub:** zoelchr

---

## Danksagung
Ein besonderer Dank geht an die Entwickler und Community, die zur Entwicklung dieses Tools beigetragen haben.