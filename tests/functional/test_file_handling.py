"""
test_file_handling.py

Dieses Modul enthält Tests für die Funktionen im Modul file_handling.py.
Es stellt sicher, dass die Funktionen korrekt arbeiten und die erwarteten Ergebnisse liefern.

Verwendung:
Führen Sie dieses Skript aus, um alle Tests durchzuführen und die Ergebnisse zu überprüfen.

Beispiel:
    python -m unittest test_file_handling.py
"""

import os
import logging
import datetime
from utils.file_handling import (
    rename_file,
    format_datetime_stamp,
    set_file_creation_date,
    set_file_date,
    test_file_access,
    test_file_access2,
    FileAccessStatus
)
from modules.msg_handling import get_date_sent_msg_file
from utils.testset_preparation import prepare_test_directory

# Log-Verzeichnis und Basisname für Logdateien festlegen
LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
PROG_LOG_NAME = 'test_file_handling_prog_log.log'
prog_log_file_path = os.path.join(LOG_DIRECTORY, PROG_LOG_NAME)

# Logging für Test und Debugging konfigurieren
DEBUG_MODE=True # Debugging-Modus aktivieren

def setup_logging(debug=False):
    log_level=logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(funcName)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        filename=prog_log_file_path
    )

if __name__ == '__main__':
    # Verzeichnisse für die Tests definieren
    SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-long'
    TARGET_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\tests\functional\testdir'

    # Logging für Test und Debugging konfigurieren
    setup_logging(DEBUG_MODE)
    logger = logging.getLogger(prog_log_file_path)
    # Logging für die Programm-Ausführung
    logger.info("Programm gestartet")
    logger.info(f"Debug-Modus: {DEBUG_MODE}")
    logging.info(f"Programm Logdatei: {prog_log_file_path}")

    # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
    if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
        os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)

    # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
    success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
    if not success:
        print("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Console
        logging.error("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Log-File
        exit(1)  # Programm abbrechen

    # Zähler für umbenannte Dateien und Probleme initialisieren
    renamed_count = 0
    problem_count = 0

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    for root, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        for file in files:
            # Überprüfen, ob die Datei die Endung .msg hat
            if file.lower().endswith('.msg'):
                file_path = os.path.join(root, file)

                print(f"\nAktuelle MSG-Date: '{file}' im Verzeichnis '{root}'")  # Debugging-Ausgabe: Console
                logging.debug(f"\nAktuelle MSG-Date: '{file}' im Verzeichnis '{root}'")  # Debugging-Ausgabe: Log-File

                # Überprüfen des Lesezugriffs auf die Datei
                access_result = test_file_access2(file_path)
                print(f"Ergebnis Zugriffsprüfung: '{[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console

                if FileAccessStatus.WRITABLE in access_result:
                    print(f"Lesender und schreibender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Console
                    logging.debug(f"Lesender und schreibender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Log-File

                    # Versanddatum abrufen
                    datetime_stamp = get_date_sent_msg_file(file_path)
                    if datetime_stamp != "Unbekannt":
                        # Formatieren des Versanddatums
                        format_string = "%Y%m%d-%Huhr%M"
                        formatted_date = format_datetime_stamp(datetime_stamp, format_string)  # Call the new function
                        new_filename = f"{formatted_date}_{file}"  # Neuer Dateiname basierend auf dem formatierten Versanddatum
                        new_file_path = os.path.join(root, new_filename)

                        # Umbenennen der Datei
                        logging.debug(f"\nUmbenennung Datei '{file}' in '{new_filename}'")  # Debugging-Ausgabe: Log-File
                        print(f"Aktueller Dateiname: {file_path}") # Debugging-Ausgabe: Console
                        print(f"Neuer Dateiname: {new_file_path}") # Debugging-Ausgabe: Console
                        rename_result = rename_file(file_path, new_file_path)
                        logging.debug(f"Ergebnis der Umbenennung Datei: {rename_result}")  # Debugging-Ausgabe: Log-File
                        print(rename_result) # Debugging-Ausgabe: Console

                        # Wenn das Umbenennen erfolgreich war, alte Datei löschen
                        if "erfolgreich umbenannt" in rename_result:
                            renamed_count += 1  # Zähler erhöhen

                            # Setze das Erstelldatum auf das Versanddatum
                            if datetime_stamp != "Unbekannt":
                                # Überprüfen, ob datetime_stamp ein datetime-Objekt ist
                                if isinstance(datetime_stamp, datetime.datetime):
                                    # Konvertiere das datetime-Objekt in einen String im richtigen Format
                                    datetime_stamp_str = datetime_stamp.strftime("%Y-%m-%d %H:%M:%S")
                                else:
                                    datetime_stamp_str = datetime_stamp  # Annehmen, dass es bereits ein String ist

                                set_creation_result = set_file_creation_date(new_file_path, datetime_stamp_str)
                                print(set_creation_result)  # Ausgabe des Ergebnisses

                                # Setze das Änderungsdatum auf das gleiche Datum
                                set_modification_result = set_file_date(new_file_path, datetime_stamp_str)
                                print(set_modification_result)  # Ausgabe des Ergebnisses für das Änderungsdatum
                    else:
                        print(f"Versanddatum für '{file}' konnte nicht abgerufen werden.") # Debugging-Ausgabe: Console
                        logging.error("Versanddatum für '{file}' konnte nicht abgerufen werden.") # Debugging-Ausgabe: Log-File
                        problem_count += 1  # Problemzähler erhöhen

                elif FileAccessStatus.READABLE in access_result:
                    print(f"Lesender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Console
                    logging.debug(f"Lesender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Log-FileDatei '{file}' erfolgreich.")
                    problem_count += 1  # Problemzähler erhöhen

                else:
                    print(f"Zugriff auf Datei '{file}' fehlgeschlagen: '{[s.value for s in access_result]}'") # Debugging-Ausgabe: Console
                    logging.error(f"Zugriff auf Datei '{file}' fehlgeschlagen: '{[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File
                    problem_count += 1  # Problemzähler erhöhen

    # Ausgabe der Ergebnisse
    print(f"\nAnzahl der umbenannten Dateien: {renamed_count}")
    print(f"Anzahl der Dateien mit Problemen: {problem_count}")
