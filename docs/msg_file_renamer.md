# msg_file_renamer.py

Dieses Modul wurde entwickelt, um **MSG-Dateien** (z. B. E-Mail-Dateien) auf Basis ihrer Metadaten in ein einheitliches Namensschema umzubenennen. Ziel ist es, Dateien besser organisieren, durchsuchen und strukturieren zu können. Zusätzlich werden umfassende Logdaten erstellt, um die Verarbeitung nachzuvollziehen und Fehler zu dokumentieren.

---

## Funktionen und Prozesse

### 1. **Extraktion von Metadaten**
- Die relevanten Metadaten umfassen:
    - **Betreff** der E-Mail.
    - **Absender** und **Empfänger**.
    - **Datum** der E-Mail.
- Sonderzeichen und ungültige Zeichen in den Metadaten werden automatisch bereinigt, um gültige Dateinamen zu erstellen.

### 2. **Umbenennung von Dateien**
- Neue Dateinamen werden gemäß eines vordefinierten Schemas erstellt.
- Das Modul überprüft Namenskonflikte und sorgt für eindeutige Benennungen.
- Unterstützt das Speichern in spezifizierte Zielverzeichnisse.

### 3. **Fehlerbehandlung**
- Erkennt fehlerhafte oder beschädigte MSG-Dateien und dokumentiert diese im Log.
- Vermeidet Konflikte durch vorhandene Dateien mit ähnlichen Namen.
- Gibt Warnungen und Fehler zu verarbeiteten Dateien aus.

### 4. **Protokollierung**
- Erstellt ausführliche Protokolle aller umbenannten Dateien.
- Dokumentiert auftretende Fehler und problematische Dateien.
- Logdaten können für weitere Analysen exportiert werden.

### 5. **Zusätzliche Features**
- **Testmodus**:
    - Führt den kompletten Umbenennungsprozess durch, ohne tatsächlich Dateien zu ändern.
    - Ideal, um die Konfiguration vor der endgültigen Ausführung zu prüfen.
- Unterstützt flexible Zielverzeichnisse und konfigurierte Umbenennungsschemata.

---

## Verwendung

Das Modul kann in Szenarien eingesetzt werden, in denen E-Mail-Dateien (MSG-Dateien) in großer Menge verarbeitet und organisiert werden müssen. Es eignet sich beispielsweise für:

- Elektronische Archivierung von E-Mails.
- Vorgaben für einheitliche Dateinamen in Unternehmensstrukturen.
- Schnellere und durchsuchbare Organisation von E-Mails.

---

## Hinweise zur Nutzung

1. **Vorbereitung**:
    - Stellen Sie sicher, dass alle Abhängigkeiten für das Lesen von MSG-Dateien installiert sind.
    - Konfigurieren Sie das gewünschte Namensschema vor der Nutzung.

2. **Konfigurierbarer Testmodus**:
    - Nutzen Sie den Testmodus, um den Umbenennungsprozess zu simulieren, ohne Dateien zu verändern.
    - Der Testmodus hilft, potenzielle Probleme frühzeitig zu identifizieren.

3. **Flexibilität**:
    - Definieren Sie Zielverzeichnisse und andere Parameter entsprechend Ihren Anforderungen.

---

## Beispiel für den Prozessfluss

1. Das Modul durchsucht ein Zielverzeichnis nach `.msg`-Dateien.
2. Relevante Metadaten werden aus den Dateien extrahiert:
    - Betreff: `Projektstatus-Update`
    - Datum: `2023-11-15`
    - Absender: `max.mustermann@example.com`
3. Basierend auf diesen Daten wird ein neuer Dateiname erstellt, z. B.:
    - `2023-11-15_Projektstatus-Update.msg`
4. Die Datei wird (optional) umbenannt und im Zielverzeichnis gespeichert.
5. Protokolle werden erstellt, um die Verarbeitung zu dokumentieren.

---

## Anforderungen

- Unterstützung für MSG-Dateien (z. B. durch Abhängigkeiten wie `extract-msg` oder ähnliche Bibliotheken).
- Python-Umgebung mit den erforderlichen Modulen für Datei- und Metadatenoperationen.
- Konfigurierbare Details wie Zielverzeichnisse, Logformate und das gewünschte Namensschema.

---

## Fazit

Das Modul `msg_file_renamer.py` ist ein leistungsstarkes Werkzeug zur effizienten Organisation und Umbenennung von MSG-Dateien auf Basis ihrer Metadaten. Es unterstützt sowohl echte Dateioperationen als auch simulierte Testläufe, um den Prozess vor der Ausführung umfassend prüfen zu können. Dank Protokollierung und Fehlerbehandlung gewährleistet es Zuverlässigkeit und Transparenz in der Verarbeitung.

---