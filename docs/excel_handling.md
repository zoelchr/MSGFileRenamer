# Beschreibung: excel_handling.py

## Übersicht

Das Modul `excel_handling.py` stellt Funktionen zur Verfügung, mit denen aus einer Liste von MSG-Dateien eine strukturierte Excel-Datei erstellt und gespeichert werden kann. Es wird typischerweise in Kombination mit einem Verzeichnis-Scanner für MSG-Dateien verwendet.

---

## Enthaltene Funktionen

### `create_excel_list(msg_files)`
Erzeugt eine Excel-kompatible Liste (Pandas DataFrame) aus den Pfaden zu MSG-Dateien.

**Parameter:**
- `msg_files`: Liste von vollständigen Pfadnamen zu MSG-Dateien

**Rückgabe:**
- Pandas DataFrame mit den Spalten:
  - `Nummer`: fortlaufende Nummer
  - `Dateiname`: Name der Datei
  - `Pfadname`: Verzeichnis, in dem sich die Datei befindet
  - `Pfadlänge`: Länge des vollständigen Pfades

---

### `save_excel_file(excel_list, output_file)`
Speichert ein DataFrame als `.xlsx`-Datei.

**Parameter:**
- `excel_list`: Pandas DataFrame mit MSG-Dateiinformationen
- `output_file`: Zielpfad für die Excel-Datei

**Effekt:**
- Schreibt die Excel-Datei an den angegebenen Speicherort.

---

## Abhängigkeiten

- `os` (Standardbibliothek)
- `pandas` (für die Datenstruktur und Excel-Export)

---

## Beispielverwendung

```python
from excel_handling import create_excel_list, save_excel_file

msg_files = [
    "C:/Mails/mail1.msg",
    "C:/Mails/mail2.msg"
]

df = create_excel_list(msg_files)
save_excel_file(df, "C:/Ausgabe/msg_liste.xlsx")
```

---

Erstellt aus dem Quellcode `excel_handling.py`.
