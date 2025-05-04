# MSGFileRenamer

Ein Python-Tool zur automatisierten Verarbeitung und Verwaltung von MSG-Dateien (Microsoft Outlook E-Mails).

## Features

- Automatische Umbenennung von MSG-Dateien basierend auf Metadaten
- Anpassung der Zeitstempel auf das E-Mail-Versanddatum
- PDF-Generierung aus MSG-Dateien (optional)
- Detaillierte Excel-Protokollierung aller Operationen
- Unterstützung für bekannte Absender via CSV-Datei
- Umfangreiche Konfigurationsmöglichkeiten über Kommandozeile

## Kurzanleitung

Diese Anleitung richtet sich an alle, die das Programm einfach ausprobieren möchten – ganz ohne Python-Kenntnisse oder zusätzliche Installationen.

### 1. Programmverzeichnis kopieren
- Navigiere in das Unterverzeichnis `.\dist`
- Wähle das gewünschte Release aus, z.B.:
  ```
  .\dist\MSGFileRenamer 1.1
  ```
- Kopiere dieses Verzeichnis auf deinen Rechner

### 2. Programm starten
- Das Programm wird über die Batch-Datei gestartet:
  ```
  .\MSGFileRenamer 1.1\msg_file_renamer.bat
  ```
- Einfach Doppelklick – fertig!

### 3. Optionale Anpassung: bekannte Absender
- Bei Bedarf kannst du die Datei anpassen:
  ```
  .\MSGFileRenamer 1.1\config\known_senders.csv
  ```
- Das ist nicht zwingend erforderlich für die Funktion des Programms

### 4. (Optional) Testdaten nutzen
- Du kannst eigene MSG-Dateien in folgenden Ordner kopieren:
  ```
  .\MSGFileRenamer 1.1\tests\functional\testdir
  ```
- Auch dieser Schritt ist nicht notwendig, kann aber beim Testen helfen

### 5. Ergebnisse nach dem Programmlauf
Nach dem Ausführen des Programms findest du:
- Eine Excel-Datei mit den durchgeführten Umbenennungen
- Eine Debug-Logdatei (für normalen Gebrauch nicht relevant)

## Entwickler-Dokumentation

### Technische Voraussetzungen
- Python 3.10.6
- Virtualenv als Package Manager
- Installierte Pakete:
  - numpy
  - openpyxl
  - pandas
  - pillow
  - pyparsing
  - pytz
  - six
  - wheel

### Kommandozeilenparameter

| Parameter | Beschreibung | Standard |
|-----------|--------------|----------|
| `-ntr`, `--no_test_run` | Testmodus (keine Dateioperationen) | `False` |
| `-fd`, `--set_filedate` | Zeitstempel anpassen | `False` |
| `-pdf`, `--generate_pdf` | PDFs generieren | `False` |
| `-rs`, `--recursive_search` | Rekursive Verzeichnissuche | `False` |
| `-sd`, `--search_directory` | Suchverzeichnis | `./tests/functional/testdir` |

### Build-Prozess
Release-Build mit PyInstaller erstellen:
```bash
pyinstaller --onefile --console msg_file_renamer.py
```
Optionen:
- `--onefile`: Erstellt eine einzelne EXE-Datei
- `--console`: Aktiviert Konsolenausgabe

---

## 🗂️ Projektstruktur

```plaintext
MSGFileRenamer/
├── msg_file_renamer.bat             # Komfortabler Einstiegspunkt (Batch)
├── msg_file_renamer.py              # Hauptskript
├── requirements.txt                 # Abhängigkeiten
├── modules/
│   ├── msg_generate_new_filename.py # Logik zur Dateinamengenerierung
│   └── msg_handling.py              # Metadaten-Extraktion
├── utils/
│   ├── excel_handling.py            # Excel-Export
│   ├── file_handling.py             # Umbenennen, löschen, Zeit setzen
│   └── pdf_generation.py            # PDF-Erstellung
├── dist/
│   └── MSGFielRenamer 1.0/          # Distributable
├── tests/
│   └── functional/                  
│       └── testdir/                 # Testdaten
└── config/
    └── known_senders_private.csv   # Liste mit bekannten Absendern
```

---

## 💡 Einsatzszenarien

- Langzeitarchivierung von E-Mails
- DSGVO-konforme Dublettenbereinigung
- Automatisierte Datei-Umbenennung in Mailarchiven
- PDF-Export für rechtssichere Ablage

---

## 📝 Lizenz

Dieses Projekt steht unter der **MIT-Lizenz**. Details siehe [LICENSE](LICENSE).

---

## 👤 Autor

**Rüdiger Zölch**  
📧 ruediger@zoelch.me  
🔗 [GitHub: zoelchr](https://github.com/zoelchr)
