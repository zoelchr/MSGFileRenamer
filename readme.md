# 📁 MSG File Renamer

Ein vielseitiges Python-Skript zur automatisierten Verarbeitung von Outlook-`.msg`-Dateien mit Fokus auf **Standardisierung**, **Archivierung** und **Dublettenkontrolle**.

---

## 🚀 Funktionen

- 🏷️ **Einheitliche Dateibenennung**  
  Erzeugt normierte Dateinamen auf Basis von Absender, Versanddatum, Betreff etc.

- ♻️ **Doubletten-Erkennung & -Löschung**  
  Erkennt Duplikate anhand von Hashwerten und entfernt sie automatisch.

- 🔍 **Rekursive Dateisuche**  
  Durchsucht auf Wunsch auch alle Unterverzeichnisse.

- 📊 **Excel-Export**  
  Protokolliert alle Ergebnisse übersichtlich in einer `.xlsx`-Datei.

- 🕒 **Dateizeitstempel anpassen (optional)**  
  Setzt Erstellungs-/Änderungsdatum auf den Versandzeitpunkt der E-Mail.

- 🧾 **PDF-Zusammenfassung (optional)**  
  Erzeugt für jede E-Mail eine kompakte PDF mit allen Kerndaten.

- 🧪 **Simulationsmodus**  
  Führt Testlauf ohne Dateiveränderungen durch – ideal zur Überprüfung.

---

## ⚙️ Konfiguration & CLI-Optionen

Das Skript kann direkt über die Kommandozeile mit Argumenten oder komfortabel über die Batch-Datei `msg_file_renamer.bat` gestartet werden.

### 🧾 Häufig verwendete CLI-Argumente

| Argument                      | Beschreibung                                                                |
|-------------------------------|-----------------------------------------------------------------------------|
| `--no_test_run` (`-ntr`)      | Führt **echte Änderungen** durch (standardmäßig ist Testlauf aktiv).        |
| `--set_filedate` (`-fd`)      | Aktiviert die Anpassung des Dateidatums auf das Versanddatum.              |
| `--search_directory` (`-sd`)  | Pfad zum Eingangsverzeichnis.                                              |
| `--recursive_search` (`-rs`)  | Rekursive Suche in Unterverzeichnissen aktivieren.                         |
| `--generate_pdf` (`-pdf`)     | Erstellt PDFs aus den `.msg`-Dateien.                                      |
| `--overwrite_pdf` (`-opdf`)   | Überschreibt bereits vorhandene PDFs.                                      |
| `--use_knownsender_file` (`-ucf`) | Aktiviert Filterung anhand bekannter Absender.                         |
| `--knownsender_file` (`-cf`)  | Pfad zur CSV-Datei mit bekannten Absendern.                                |
| `--no_shorten_path_name` (`-spn`) | Verhindert Kürzung langer Pfadnamen.                                 |

---

## 📦 Installation

### 🔧 Voraussetzungen

- **Python** ≥ 3.7

### 📥 Schritte

1. **Repository klonen**  
   ```bash
   git clone https://github.com/zoelchr/MSGFileRenamer.git
   cd MSGFileRenamer
   ```

2. **(Optional) Virtuelle Umgebung erstellen**  
   ```bash
   python -m venv venv
   venv\Scripts\activate       # Windows
   ```

3. **Abhängigkeiten installieren**  
   ```bash
   pip install -r requirements.txt
   ```

---

## 🛠 Verwendung

### 📁 Start per Batch-Datei
```bash
msg_file_renamer.bat
```

### 🔄 Direkter Aufruf mit Argumenten
```bash
python msg_file_renamer.py --search_directory "D:/Mails" --recursive_search --generate_pdf --no_test_run
```

---

## 📤 Ausgabe

Die Ergebnisse werden in einer Excel-Datei gespeichert mit u. a. folgenden Spalten:

- **Dateiname**
- **Pfad**
- **Pfadlänge**
- **Versanddatum**
- **Betreff**
- **Absender**
- **PDF erstellt (Ja/Nein)**
- ...

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
