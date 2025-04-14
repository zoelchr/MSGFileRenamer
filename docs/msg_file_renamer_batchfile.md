# Dokumentation: msg_file_renamer.bat

## Übersicht
Die `msg_file_renamer.bat` ist ein Batch-Skript, das dazu dient, das Modul `msg_file_renamer` zu konfigurieren und auszuführen. Es erlaubt die flexible Steuerung verschiedener Optionen, wie Testläufe, Anpassung der Dateidaten, Generierung von PDF-Dateien und rekursive Verarbeitung von MSG-Dateien in Verzeichnissen.

Das Skript kann sowohl mit vordefinierten Einstellungen als auch interaktiv mit benutzerdefinierten Parametern verwendet werden.

---

## Funktionsweise

### Hauptaufgaben
1. **Benutzerinteraktion und Konfiguration**:
    - Fragt den Benutzer interaktiv, ob Defaultwerte verwendet werden sollen oder eine manuelle Konfiguration gewünscht ist.
2. **Ausführung des `msg_file_renamer` Moduls**:
    - Startet den Verarbeitungsprozess für MSG-Dateien mit den angegebenen Optionen.
3. **Flexible Konfiguration**:
    - Bietet verschiedene Parameter, die während des Batch-Laufs angepasst werden können.

### Typische Anwendungsfälle
- **Testlauf ohne Veränderung von Dateien** (Simulationslauf).
- **Anpassung von Dateizeiten**: Setzt den Erstellungs- oder Änderungszeitstempel basierend auf dem MSG-Inhalt.
- **PDF-Generierung** aus MSG-Dateien (optional mit Überschreibung existierender PDFs).
- **Verarbeitung von Verzeichnissen**: Sowohl rekursiv als auch nicht-rekursiv.

---

## Optionen und Parameter

| Option                       | Beschreibung                                                                                   | Standardwert                      |
|------------------------------|-----------------------------------------------------------------------------------------------|----------------------------------|
| `-ntr / --no_test_run`       | Aktiviert den Testlauf: Keine Änderungen an Dateien (Standard: Änderungen werden durchgeführt). | `Nicht gesetzt`                  |
| `-fd / --set_filedate`       | Passt Dateizeitstempel an: Setzt das Dateidatum aus MSG-Inhalt.                                | `Nicht gesetzt`                  |
| `-sd / --search_directory`   | Pfad, der nach MSG-Dateien durchsucht wird.                                                   | `DEFAULT_ZIELPFAD`               |
| `-spn / --no_shorten_path_name` | Unterdrückt das Kürzen langer Dateipfade (für Windows).                                      | `Nicht gesetzt`                  |
| `-pdf / --generate_pdf`      | Aktiviert die PDF-Erstellung aus MSG-Dateien.                                                 | `Nicht gesetzt`                  |
| `-opdf / --overwrite_pdf`    | Überschreibt vorhandene PDF-Dateien bei Bedarf.                                               | `Nicht gesetzt`                  |
| `-rs / --recursive_search`   | Aktiviert die rekursive Suche in Verzeichnissen.                                              | `Nicht gesetzt`                  |

---

## Benutzerinteraktion

### Interaktive Konfiguration
Das Skript bietet interaktive Eingabeaufforderungen, um wichtige Konfigurationen festzulegen. Hier eine Übersicht der Schritte:

1. **Standardwerte verwenden oder manuell konfigurieren**:
    - Benutzer entscheidet, ob die Standardkonfiguration übernommen wird oder eine interaktive Eingabe erfolgt.
2. **Testlauf aktivieren**:
    - Option, um den Testlauf zu aktivieren (`--no_test_run`).
3. **Weitere Optionen**:
    - Änderungsoptionen wie:
        - Anpassung des Dateidatums (`--set_filedate`).
        - Generierung von PDFs (`--generate_pdf`, `--overwrite_pdf`).
        - Rekursive Verarbeitung (`--recursive_search`).

---

## Ablauf des Skripts

### Konfiguration
1. **Initialisierung**:
    - Festlegen von Standardwerten für:
        - Zielpfad (`DEFAULT_ZIELPFAD`)
        - Testlauf (`NOTESTLAUF`)
        - PDF-Erstellung (`GENERATEPDF`)
    - Ermitteln des Verzeichnisses der Batch-Datei.

2. **Interaktive Konfigurationsabfragen**:
    - Fragen nach der Verwendung der Standardkonfiguration.
    - Aktivieren oder Deaktivieren von Optionen wie Testlauf, PDF-Erstellung oder rekursive Suche.

### Ausführung
- Generiert die vollständigen Befehle anhand der Nutzerkonfiguration.
- Weist die Konfiguration dem `msg_file_renamer`-Modul zu und startet es.

---

## Beispielprozess

1. Das Skript begrüßt den Benutzer und bietet die Möglichkeit, die Standardkonfiguration zu übernehmen oder selbst anzupassen:
   ```bash
   Soll die Default-Konfiguration verwendet oder abgebrochen werden? (J/N/A):
   ```

2. Bei Auswahl von `N` wird interaktiv nach weiteren Einstellungen gefragt:
   ```bash
   Möchtest du einen Testlauf durchführen? (J/N/A):
   ```

3. Wenn alle Eingaben gemacht wurden, wird das `msg_file_renamer`-Modul mit den konfigurierten Parametern gestartet.

---

## Beispielaufruf
Hier ein Beispiel, wie das Skript manuell aufgerufen werden kann (anstatt interaktiv):
```batch
msg_file_renamer.bat -fd --generate_pdf -rs -sd "C:\Users\Beispiel\MSG-Files"
```
### Beschreibung:
- `-fd`: Passt Dateidaten der Dateien an.
- `--generate_pdf`: Erzeugt PDFs aus den gefundenen MSG-Dateien.
- `-rs`: Führt eine rekursive Suche durch.
- `-sd "C:\Users\Beispiel\MSG-Files"`: Setzt das Suchverzeichnis.

---

## Wartung und Erweiterung

### Anpassungsoptionen
Das Skript ermöglicht einfache Änderungen und Erweiterungen:
- **Voreingestellte Variablen**:
    - Das Zielverzeichnis kann mit `DEFAULT_ZIELPFAD` angepasst werden.
- **Hinzufügen neuer Optionen**:
    - Neue Parameter können leicht eingebaut werden, indem passende Variablen in der Batch-Datei ergänzt werden.

### Fehlersuche
- Bei falscher Eingabe wird der Benutzer aufgefordert, gültige Werte (z. B. `J`, `N`, `A`) einzugeben.
- Das Skript bricht automatisch ab, wenn die Eingabe `A` (Abbrechen) gewählt wird.

---

## Fazit
Die `msg_file_renamer.bat` ist ein vielseitiges Tool, das die Verarbeitung von MSG-Dateien effizient unterstützt. Durch die interaktive Konfiguration und die umfangreichen Anpassungsmöglichkeiten kann sie flexibel auf verschiedene Arbeitsanforderungen angepasst werden.