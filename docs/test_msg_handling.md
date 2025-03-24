# Beschreibung: test_msg_handling.py

## Übersicht

Das Modul `test_msg_handling.py` testet die Funktionen des Moduls `msg_handling.py`. Es prüft die Extraktion und Verarbeitung von Metadaten aus MSG-Dateien sowie die Formatierung, Validierung und Reduktion von E-Mail-Inhalten. Zusätzlich wird ein strukturierter Logeintrag für jede verarbeitete Datei erzeugt.

---

## Ziele der Tests

- Validierung der Funktionalität von:
  - `get_msg_object()`
  - `parse_sender_msg_file()`
  - `load_known_senders()`
  - `format_datetime()`
  - `custom_sanitize_text()`
  - `truncate_filename_if_needed()`
  - `reduce_thread_in_msg_message()`
- Sicherstellen der Metadaten-Extraktion:
  - Absendername und E-Mail
  - Versandzeitpunkt (formatiert)
  - Betreff (roh und bereinigt)
  - Nachrichtentext (reduziert)
- Logging in strukturierter Excel-Datei

---

## Aufbau

1. Testdaten werden in ein Zielverzeichnis kopiert.
2. Jede MSG-Datei wird eingelesen und analysiert.
3. Ein neuer, sprechender Dateiname wird erzeugt.
4. Optional wird der Name bei zu großer Länge gekürzt.
5. Die Ergebnisse werden protokolliert.

---

## Wichtige Einstellungen

- `DEBUG_MODE`: Aktiviert detailliertes Logging
- `LIST_OF_KNOWN_SENDERS`: CSV-Datei mit bekannten Absendern
- `format_timestamp_string`: Format für Datumsangabe im Dateinamen
- Logdatei: `msg_log.xlsx`

---

## Beispiel für erzeugten Dateinamen

```
20240324-16uhr10_jane.doe@example.com_Projektstatus.msg
```

---

## Logeinträge

In der Excel-Logdatei werden folgende Felder festgehalten:

- Fortlaufende Nummer
- Verzeichnisname
- Original-Dateiname
- Absendername & E-Mail
- Zeitstempel (roh & formatiert)
- Betreff (roh & bereinigt)
- Neuer Dateiname & gekürzte Version
- Gekürzter Pfadname
- Nachricht & gekürzte Nachricht
- Anzahl reduzierter E-Mails

---

## Ausführung

```bash
python test_msg_handling.py
```

oder mit `unittest`:

```bash
python -m unittest test_msg_handling.py
```

---

Erstellt aus dem Quellcode `test_msg_handling.py`.
