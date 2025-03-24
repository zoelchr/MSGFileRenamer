# Beschreibung: msg_handling.py

## Übersicht

Das Modul `msg_handling.py` enthält Funktionen zur Verarbeitung und Analyse von MSG-Dateien (Microsoft Outlook-Nachrichtenformate). Es ermöglicht das Extrahieren relevanter Metadaten, Logging, Formatierung sowie Validierung von Absenderinformationen. Es unterstützt zudem das Logging in Excel und die Reduktion von E-Mail-Threads.

---

## Hauptfunktionen

### `get_msg_object(msg_file: str) -> dict`
Liest eine MSG-Datei aus und extrahiert Metadaten wie Betreff, Absender, Datum, Text, Anhänge, Signatur- und Verschlüsselungsstatus.

---

### `create_log_file(base_name, directory, table_header) -> str`
Erstellt eine neue Excel-Logdatei mit Zeitstempel im Dateinamen.

---

### `log_entry(log_file_path, entry)`
Fügt einen neuen Eintrag (Dictionary) in die angegebene Excel-Logdatei ein.

---

### `convert_to_utc_naive(datetime_stamp)`
Entfernt die Zeitzonen-Information von `datetime`-Objekten.

---

### `format_datetime(datetime_stamp, format_string)`
Formatiert einen Zeitstempel in ein definiertes Format (z. B. "%Y-%m-%d %H:%M:%S").

---

### `custom_sanitize_text(encoded_textstring)`
Bereinigt Textzeichenfolgen von Sonderzeichen – besonders geeignet zur Vorbereitung auf Dateinamen.

---

### `truncate_filename_if_needed(file_path, max_length, truncation_marker)`
Kürzt Dateinamen, falls der vollständige Pfad zu lang ist, und fügt ein Trunkierungszeichen ein.

---

### `parse_sender_msg_file(msg_absender_str: str)`
Extrahiert Namen und E-Mail-Adresse aus einem MSG-Absenderstring.

---

### `load_known_senders(file_path)`
Lädt bekannte Absender aus einer CSV-Datei in ein Pandas-DataFrame.

---

### `reduce_thread_in_msg_message(email_text, max_older_emails=2)`
Reduziert ältere E-Mail-Inhalte im E-Mail-Text auf eine maximale Anzahl, um E-Mail-Ketten zu kürzen.

---

## Enum-Klassen

### `MsgAccessStatus`
Detaillierte Statuscodes zur Kennzeichnung von Erfolgen und Fehlern beim Zugriff auf MSG-Dateien:
- `SUCCESS`, `DATA_NOT_FOUND`, `FILE_NOT_FOUND`, `PERMISSION_ERROR`, `ATTRIBUTE_ERROR`, `VALUE_ERROR`, u. v. m.

---

## Abhängigkeiten

- `extract_msg`
- `pandas`
- `datetime`, `re`, `os`, `logging`
- `enum`

---

## Anwendungsbeispiel

```python
from msg_handling import get_msg_object

msg_info = get_msg_object("beispiel.msg")
print(msg_info["subject"], msg_info["sender"])
```

---

Erstellt aus dem Modul `msg_handling.py`.
