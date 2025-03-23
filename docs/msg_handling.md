# msg_handling.py

Dieses Modul enthält Funktionen zum Abrufen und Verarbeiten von Metadaten aus `.msg`-Dateien (Outlook-Nachrichten). Es unterstützt das Extrahieren von Absenderinformationen, Betreff, Textinhalt, Anhängen sowie das Logging und weitere Hilfsfunktionen zur Bearbeitung und Bereinigung von E-Mail-Inhalten.

## 📦 Inhalt

### Hauptfunktionen

- **`get_msg_object(msg_file)`**  
  Öffnet eine `.msg`-Datei und extrahiert Informationen wie Betreff, Absender, Datum, Nachrichtentext, Anhänge etc.

- **`create_log_file(base_name, directory, table_header)`**  
  Erstellt ein neues Excel-Logfile mit einem Zeitstempel im Dateinamen.

- **`log_entry(log_file_path, entry)`**  
  Fügt einen neuen Datensatz (Dictionary) als Zeile in ein bestehendes Logfile ein.

- **`convert_to_utc_naive(datetime_stamp)`**  
  Entfernt Zeitzoneninformationen von einem `datetime`-Objekt.

- **`format_datetime(datetime_stamp, format_string)`**  
  Gibt einen formatierten Zeitstempel zurück (z. B. `"2024-12-31 12:00:00"`).

- **`custom_sanitize_text(encoded_textstring)`**  
  Entfernt oder ersetzt ungültige Zeichen (z. B. für Dateinamen) aus einem Textstring.

- **`truncate_filename_if_needed(file_path, max_length, truncation_marker)`**  
  Kürzt Dateinamen, wenn der vollständige Pfad eine bestimmte Länge überschreitet.

- **`parse_sender_msg_file(msg_absender_str)`**  
  Trennt den Namen und die E-Mail-Adresse aus einem Absenderstring (z. B. `"Max Mustermann <max@example.com>"`).

- **`load_known_senders(file_path)`**  
  Lädt bekannte Absender aus einer CSV-Datei in ein `pandas.DataFrame`.

- **`reduce_thread_in_msg_message(email_text, max_older_emails=2)`**  
  Entfernt ältere Nachrichten aus E-Mail-Threads nach konfigurierbarem Limit.

## 📁 Anforderungen

- `extract_msg`
- `pandas`
- `openpyxl`
- `re`
- `logging`
- `datetime`

## 🔍 Beispiel zur Verwendung

```python
from msg_handling import get_msg_object

msg_data = get_msg_object("example.msg")
print(msg_data["subject"])
```

## 🧪 Testdaten

Für Unit-Tests können `.msg`-Dateien mit verschiedenen Attributen (fehlender Absender, kein Body, keine Anhänge etc.) verwendet werden, um den Umgang mit Fehlerszenarien zu prüfen.

## 🧰 Status-Codes

Das Modul verwendet die `MsgAccessStatus`-Enum zur Beschreibung möglicher Zugriffsergebnisse, darunter:

- `SUCCESS`
- `FILE_NOT_FOUND`
- `PERMISSION_ERROR`
- `SUBJECT_MISSING`
- `BODY_MISSING`
- u. v. m.

## 👤 Autor

Dieses Modul wurde entwickelt zur automatisierten Verarbeitung von MSG-Dateien, z. B. zur Dokumentation von Behördenkommunikation, für E-Mail-Analysen oder forensische Zwecke.

---

> Hinweis: Dieses Modul ist besonders nützlich im Zusammenspiel mit Projektstrukturen, in denen E-Mail-Metadaten protokolliert und archiviert werden müssen.
