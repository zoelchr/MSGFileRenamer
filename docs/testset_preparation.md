# testset_preparation.py

Dieses Modul enthält Funktionen zur Vorbereitung von Testverzeichnissen für die Verarbeitung von MSG-Dateien.
Es ermöglicht das Erstellen eines Zielverzeichnisses, das Löschen seines Inhalts und das Kopieren von Dateien
aus einem Quellverzeichnis. Diese Funktionalität ist nützlich für automatisierte Tests, bei denen eine saubere
Umgebung erforderlich ist.

## Inhaltsverzeichnis

- [Einführung](#einführung)
- [Funktionen](#funktionen)
- [Verwendung](#verwendung)

## Einführung

Das `testset_preparation.py`-Modul bietet eine einfache Möglichkeit, Testverzeichnisse für die Verarbeitung von MSG-Dateien vorzubereiten. Es stellt sicher, dass das Zielverzeichnis existiert, löscht dessen Inhalt und kopiert die erforderlichen Dateien aus einem Quellverzeichnis.

## Funktionen

### 1. `prepare_test_directory(source_dir, target_dir)`
Bereitet das Zielverzeichnis für Tests vor, indem es erstellt, den Inhalt löscht und die Dateien aus dem Quellverzeichnis kopiert.

- **Parameter**:
  - `source_dir` (str): Der Pfad zum Quellverzeichnis, aus dem die Dateien kopiert werden.
  - `target_dir` (str): Der Pfad zum Zielverzeichnis, das vorbereitet werden soll.

- **Rückgabewert**:
  - bool: True, wenn die Operation erfolgreich war, andernfalls False.

- **Beispiel**:
    ```python
    success = prepare_test_directory("D:/source_directory", "D:/target_directory")
    if success:
        print("Das Testverzeichnis wurde erfolgreich vorbereitet.")
    ```
  
## Verwendung
Um die Funktion `prepare_test_directory` zu verwenden, importiere sie in dein Hauptprogramm oder andere Module und rufe sie mit den entsprechenden Verzeichnispfaden auf. Die Funktion gibt einen Status zurück, der angibt, ob die Vorbereitung des Testverzeichnisses erfolgreich war.