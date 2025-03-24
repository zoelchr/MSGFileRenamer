# Beschreibung: testset_preparation.py

## Übersicht

Das Modul `testset_preparation.py` stellt eine Funktion zur Verfügung, um automatisiert Testverzeichnisse für die Verarbeitung von MSG-Dateien vorzubereiten. Dies ist insbesondere in Testumgebungen nützlich, in denen reproduzierbare Bedingungen geschaffen werden sollen.

---

## Funktion

### `prepare_test_directory(source_dir, target_dir)`

Bereitet ein Zielverzeichnis für Tests vor, indem es:
1. Das Zielverzeichnis erstellt (falls es nicht existiert),
2. Dessen bestehenden Inhalt löscht,
3. Und alle Inhalte aus dem Quellverzeichnis in das Zielverzeichnis kopiert.

**Parameter:**
- `source_dir` *(str)*: Quellverzeichnis mit Dateien für den Test
- `target_dir` *(str)*: Zielverzeichnis, das neu vorbereitet wird

**Rückgabewert:**
- `True`, wenn die Vorbereitung erfolgreich war
- `False`, wenn ein Fehler auftritt

**Beispiel:**
```python
success = prepare_test_directory("D:/quelle", "D:/ziel")
if success:
    print("Testverzeichnis erfolgreich vorbereitet.")
```

---

## Interne Abhängigkeiten

Dieses Modul verwendet folgende Funktionen aus dem Modul `utils.file_handling`:
- `delete_directory_contents()`: Löscht den Inhalt eines Verzeichnisses
- `copy_directory_contents()`: Kopiert den Inhalt eines Verzeichnisses

---

## Abhängigkeiten

- `os`
- `utils.file_handling`

---

## Anwendungsfall

Ideal zur Vorbereitung automatisierter Testläufe, z. B. für Unit-Tests oder manuelle Prüfungen von Verarbeitungsskripten für MSG-Dateien.

---

Erstellt aus dem Quellcode `testset_preparation.py`.
