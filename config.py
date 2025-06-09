"""
config.py

Dieses Modul liest Konfigurationswerte aus einer `.env`-Datei ein und stellt sie zentral
für andere Anwendungsteile zur Verfügung.

Unterstützt werden:
- einstellbare Debug-Stufen,
- dynamische Definition des Log-Dateinamens,
- Exportverzeichnis für Ausgabedateien,
- eine Liste von Postfächern, die in der GUI nicht angezeigt werden sollen,
- eine Liste von Ordnernamen, die beim Laden der Outlook-Ordnerstruktur ignoriert werden sollen.
"""
import os
from dotenv import load_dotenv

# .env-Datei laden (sofern vorhanden)
load_dotenv()

# Allgemeine Einstellungen
DEBUG_LEVEL = int(os.getenv("DEBUG_LEVEL", "0"))
LOG_FILE_DIRECTORY = os.getenv("LOG_FILE_DIRECTORY", ".")
EXPORT_PATH = os.getenv("EXPORT_PATH", "./export")

# Maximale Anzahl an gespeicherten Log-Dateien
MAX_DEBUG_LOG_FILE_COUNT = int(os.getenv("MAX_DEBUG_LOG_FILE_COUNT", "10"))
MAX_EXCEL_LOG_FILE_COUNT = int(os.getenv("MAX_EXCEL_LOG_FILE_COUNT", "10"))

# Konsolen-Logging zusätzlich aktivieren
LOG_TO_CONSOLE = os.getenv("LOG_TO_CONSOLE", "false").lower() == "true"

LIST_OF_KNOWN_SENDERS = os.getenv("KNOWNSENDER_FILE", "./known_senders_private.csv")

# Testmodus (kein Email-Export)
TEST_MODE = os.getenv("TEST_MODE", "false").lower() in ["true", "1", "yes", "y"]

# Testdaten-Verzeichnis (für funktionale Tests)
SOURCE_DIRECTORY_TEST_DATA = os.getenv("SOURCE_DIRECTORY_TEST_DATA", "./data/sample_files/testset-public")
TARGET_DIRECTORY_TEST_DATA = os.getenv("TARGET_DIRECTORY_TEST_DATA", "./tests/functional/testdir")

# Verzeichnis, in dem MSG-Dateien gesucht werden sollen (in einem Testlauf entspricht das TARGET_DIRECTORY_TEST_DATA
TARGET_DIRECTORY = os.getenv("TARGET_DIRECTORY", "./tests/functional/testdir")

# Maximale Länge des Dateipfades (Windows-Limitierung)
MAX_PATH_LENGTH = int(os.getenv("MAX_PATH_LENGTH", "255"))