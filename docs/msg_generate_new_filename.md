# Beschreibung: msg_generate_new_filename.py

## Übersicht

Das Modul `msg_generate_new_filename.py` dient der automatisierten Erzeugung sprechender Dateinamen für MSG-Dateien. Es nutzt dabei die in der Datei enthaltenen Metadaten wie Versanddatum, Absender und Betreff, um eindeutige, strukturierte und bereinigte Dateinamen zu generieren – optional gekürzt, falls sie die maximale Pfadlänge überschreiten.

---

## Hauptfunktion

### `generate_new_msg_filename(msg_path_and_filename, max_path_length=260)`

Erzeugt einen neuen Dateinamen für eine MSG-Datei, bestehend aus:
- Versanddatum (formatiert)
- Absender-E-Mail
- Betreff (bereinigt von Sonderzeichen)

**Parameter:**
- `msg_path_and_filename` (str): Pfad zur Originaldatei
- `max_path_length` (int): Maximale Pfadlänge (Standard: 260 Zeichen)

**Rückgabe:**  
Ein Objekt der Datenklasse `MsgFilenameResult`, das folgende Attribute enthält:
- `datetime_stamp`: Versandzeitpunkt als datetime-Objekt
- `formatted_timestamp`: Formatierter Versandzeitpunkt als String
- `sender_name`, `sender_email`: Extrahierte Absenderinformationen
- `msg_subject`: Original-Betreff
- `msg_subject_sanitized`: Bereinigter Betreff
- `new_msg_filename`: Generierter Dateiname
- `new_truncated_msg_filename`: Gekürzter Dateiname, falls nötig
- `is_msg_filename_truncated`: True/False, je nachdem ob gekürzt wurde

---

## Interne Hilfsfunktionen / Abhängigkeiten

Das Modul nutzt Funktionen aus dem Modul `msg_handling`, z. B.:
- `get_msg_object()`
- `parse_sender_msg_file()`
- `load_known_senders()`
- `custom_sanitize_text()`
- `format_datetime()`
- `truncate_filename_if_needed()`

---

## Datenquelle

- Liste bekannter Absender wird aus einer CSV-Datei geladen:  
  `D:/Dev/pycharm/MSGFileRenamer/config/known_senders_private.csv`

---

## Beispiel für generierten Dateinamen

```
20240324-19uhr15_max.mustermann@example.com_Projektmeeting_Statusupdate.msg
```

---

## Anwendungsbeispiel

```python
from msg_generate_new_filename import generate_new_msg_filename

result = generate_new_msg_filename("C:/mails/mail1.msg")
print(result.new_msg_filename)
```

---

Erstellt aus dem Quellcode `msg_generate_new_filename.py`.
