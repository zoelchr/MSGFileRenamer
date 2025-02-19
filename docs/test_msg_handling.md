# Dokumentation für das Modul `test_msg_handling.py`

## Übersicht

Das Modul `test_msg_handling.py` enthält Tests für die Funktionen im Modul `msg_handling.py`. Es stellt sicher, dass die Funktionen korrekt arbeiten und die erwarteten Ergebnisse liefern. Die Tests decken verschiedene Aspekte der Funktionalität ab, einschließlich der Verarbeitung von MSG-Dateien und der Protokollierung.

## Testfunktionen

Die Tests umfassen:

- **Überprüfung der Rückgabewerte**:
  - `get_sender_msg_file`: Überprüft, ob die Absender-E-Mail-Adresse korrekt zurückgegeben wird.
  - `get_subject_msg_file`: Validiert, dass der Betreff der Nachricht korrekt extrahiert wird.
  - `get_date_sent_msg_file`: Stellt sicher, dass das Versanddatum korrekt abgerufen wird.

- **Validierung der Funktionalität**:
  - `parse_sender_msg_file`: Testet die korrekte Extraktion von Senderinformationen aus dem Sender-String.
  
- **Protokollierung**:
  - `create_log_file`: Überprüft, ob Logdateien erfolgreich erstellt werden.
  - `log_entry`: Stellt sicher, dass Einträge korrekt in die Logdatei geschrieben werden.

- **Textverarbeitung**:
  - `custom_sanitize_text`: Testet die Bereinigung von Texten, um unerwünschte Zeichen zu entfernen.
  - `truncate_filename_if_needed`: Überprüft, ob Dateinamen korrekt gekürzt werden, um die maximal zulässige Pfadlänge nicht zu überschreiten.

## Verwendung

Um alle Tests durchzuführen und die Ergebnisse zu überprüfen, führen Sie dieses Skript aus. Stellen Sie sicher, dass die erforderlichen Module installiert sind.

### Vorbedingungen

- Die Datei `known_senders.csv` muss befüllt werden, und der Pfad muss korrekt eingetragen sein.
- Zusätzlich benötigt das Programm geeignete Testdaten (MSG-Dateien) in einem Verzeichnis, bei Bedarf auch mit weiteren Unterverzeichnissen. Auch dieses Verzeichnis muss im Modul korrekt eingetragen sein.

### Ausführen der Tests

Führen Sie den folgenden Befehl in der Kommandozeile aus:

```bash
python -m unittest test_msg_handling.py
