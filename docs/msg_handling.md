# Dokumentation für das Modul `msg_handling.py`

## Übersicht

Das Modul `msg_handling.py` bietet Funktionen zum Abrufen und Verarbeiten von Metadaten aus MSG-Dateien. Es ermöglicht das Extrahieren von Informationen wie Absender, Empfänger, Betreff und das Datum der gesendeten Nachricht. Darüber hinaus enthält es Funktionen zur Protokollierung und zur Handhabung von Dateinamen.

## Funktionen

### 1. `get_sender_msg_file(file_path)`
Gibt die E-Mail-Adresse des Absenders aus einer MSG-Datei zurück.

- **Parameter**: 
  - `file_path` (str): Der Pfad zur MSG-Datei.
- **Rückgabewert**: 
  - str: Die E-Mail-Adresse des Absenders oder "Unbekannt".

### 2. `parse_sender_msg_file(sender: str)`
Analysiert den Sender-String und extrahiert Name und E-Mail-Adresse.

- **Parameter**: 
  - `sender` (str): Der Sender-String.
- **Rückgabewert**: 
  - dict: Ein Dictionary mit 'sender_name', 'sender_email' und 'contains_sender_email'.

### 3. `get_subject_msg_file(file_path)`
Gibt den Betreff der Nachricht aus einer MSG-Datei zurück.

- **Parameter**: 
  - `file_path` (str): Der Pfad zur MSG-Datei.
- **Rückgabewert**: 
  - str: Der Betreff der Nachricht oder "Unbekannt".

### 4. `load_known_senders(file_path)`
Lädt bekannte Sender aus einer CSV-Datei.

- **Parameter**: 
  - `file_path` (str): Der Pfad zur CSV-Datei.
- **Rückgabewert**: 
  - DataFrame: Ein DataFrame mit den bekannten Sendern.

### 5. `get_date_sent_msg_file(file_path)`
Gibt das Datum der gesendeten Nachricht aus einer MSG-Datei zurück.

- **Parameter**: 
  - `file_path` (str): Der Pfad zur MSG-Datei.
- **Rückgabewert**: 
  - str: Das Datum der gesendeten Nachricht oder "Unbekannt".

### 6. `create_log_file(base_name, directory)`
Erstellt ein Logfile im Excel-Format mit einem Zeitstempel im Namen.

- **Parameter**: 
  - `base_name` (str): Der Basisname des Logfiles.
  - `directory` (str): Das Verzeichnis, in dem das Logfile gespeichert werden soll.
- **Rückgabewert**: 
  - str: Der Pfad zur erstellten Logdatei.

### 7. `log_entry(log_file_path, entry)`
Fügt einen neuen Eintrag in das Logfile hinzu.

- **Parameter**: 
  - `log_file_path` (str): Der Pfad zur Logdatei.
  - `entry` (dict): Ein Dictionary mit den Werten für die Logzeile.

### 8. `convert_to_utc_naive(datetime_stamp)`
Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.

- **Parameter**: 
  - `datetime_stamp` (datetime): Der Zeitstempel, der konvertiert werden soll.
- **Rückgabewert**: 
  - datetime: Ein UTC-naives Datetime-Objekt.

### 9. `format_datetime(datetime_stamp, format_string)`
Formatiert einen Zeitstempel in das angegebene Format.

- **Parameter**: 
  - `datetime_stamp` (datetime): Der Zeitstempel, der formatiert werden soll.
  - `format_string` (str): Das gewünschte Format für den Zeitstempel.
- **Rückgabewert**: 
  - str: Der formatierte Zeitstempel als String.

### 10. `custom_sanitize_text(encoded_textstring)`
Bereinigt einen Textstring von unerwünschten Zeichen.

- **Parameter**: 
  - `encoded_textstring` (str): Der ursprüngliche Textstring, der bereinigt werden soll.
- **Rückgabewert**: 
  - str: Der bereinigte Textstring, der als gültiger Dateiname verwendet werden kann.

### 11. `truncate_filename_if_needed(file_path, max_length, truncation_marker)`
Kürzt den Dateinamen, wenn der gesamte Pfad die maximal zulässige Länge überschreitet.

- **Parameter**: 
  - `file_path` (str): Der vollständige Dateipfad, der überprüft und möglicherweise gekürzt werden soll.
  - `max_length` (int): Die maximal zulässige Länge des gesamten Dateipfads.
  - `truncation_marker` (str): Die Zeichenkette, die verwendet wird, um das Kürzen anzuzeigen.
- **Rückgabewert**: 
  - str: Der möglicherweise gekürzte Dateipfad.

## Verwendung

Importiere die Funktionen aus diesem Modul in deinem Hauptprogramm oder anderen Modulen, um auf die Metadaten von MSG-Dateien zuzugreifen.

### Beispiel

```python
from modules.msg_handling import get_sender_msg_file
sender_email = get_sender_msg_file('example.msg')
