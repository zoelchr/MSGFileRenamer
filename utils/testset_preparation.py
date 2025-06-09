"""
testset_preparation.py

Dieses Modul enthält Funktionen zur Vorbereitung von Testverzeichnissen für die Verarbeitung von MSG-Dateien.
Es ermöglicht das Erstellen eines Zielverzeichnisses, das Löschen seines Inhalts und das Kopieren von Dateien
aus einem Quellverzeichnis. Diese Funktionalität ist nützlich für automatisierte Tests, bei denen eine saubere
Umgebung erforderlich ist.

Funktionen:
- prepare_test_directory(source_dir, target_dir): Erstellt das Zielverzeichnis, löscht dessen Inhalt und kopiert Dateien aus dem Quellverzeichnis.
"""

import os
from utils.file_handling import delete_directory_contents, copy_directory_contents
from logger import initialize_logger

# In der Log-Datei wird als Quelle der Modulname "__main__" verwendet
app_logger = initialize_logger(__name__)
app_logger.debug("Debug-Logging im Modul 'msg_to_pdf' aktiviert.")

def prepare_test_directory(source_dir, target_dir):
    """
    Bereitet das Zielverzeichnis für Tests vor, indem es erstellt, den Inhalt löscht und die Dateien aus dem Quellverzeichnis kopiert.

    Diese Funktion führt folgende Schritte durch:
    1. Überprüft, ob das Zielverzeichnis existiert; falls nicht, wird es erstellt.
    2. Löscht den gesamten Inhalt des Zielverzeichnisses.
    3. Kopiert alle Dateien aus dem Quellverzeichnis in das Zielverzeichnis.

    Parameter:
    source_dir (str): Der Pfad zum Quellverzeichnis, aus dem die Dateien kopiert werden.
    target_dir (str): Der Pfad zum Zielverzeichnis, das vorbereitet werden soll.

    Rückgabewert:
    bool: True, wenn die Operation erfolgreich war, andernfalls False.

    Beispiel:
        success = prepare_test_directory("D:/source_directory", "D:/target_directory")
        if success:
            print("Das Testverzeichnis wurde erfolgreich vorbereitet.")
    """
    try:
        # Wenn das Zielverzeichnis noch nicht existiert, dann erstellen
        if not os.path.isdir(target_dir):
            os.makedirs(target_dir, exist_ok=True)

        # Löschen des Inhalts des Zielverzeichnisses
        print(f"\tLösche Inhalt von: {target_dir}")
        delete_directory_contents(target_dir)

        # Kopieren des Quellverzeichnisses in das Zielverzeichnis
        print(f"\tKopiere Inhalt von: {source_dir} nach: {target_dir}")
        copy_directory_contents(source_dir, target_dir)

        return True  # Operation war erfolgreich
    except Exception as e:
        print(f"\tFehler bei der Vorbereitung des Testverzeichnisses: {str(e)}")
        return False  # Operation war nicht erfolgreich


