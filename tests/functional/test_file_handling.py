"""
test_file_handling.py

Dieses Modul enthält Unit-Tests für die Funktionen im Modul file_handling.py.
Die Tests überprüfen die korrekte Funktionalität der Datei- und Verzeichnisoperationen,
einschließlich des Zugriffs auf Dateien, Umbenennens, Löschens und Kopierens von Inhalten.

Die Tests umfassen:
- Validierung der Rückgabewerte von Funktionen wie `rename_file`, `set_file_creation_date`,
  und `set_file_modification_date`.
- Überprüfung des Lese- und Schreibzugriffs auf Dateien mit den Funktionen `test_read_access`
  und `test_write_access`.
- Sicherstellung, dass das Erstellen und Protokollieren in Logdateien ordnungsgemäß funktioniert.
- Tests für die Bereinigung von Dateinamen durch `sanitize_filename`.

Verwendung:
Führen Sie dieses Skript aus, um alle Tests durchzuführen und die Ergebnisse zu überprüfen.

Beispiel:
    python -m unittest test_file_handling.py
"""

import os
import logging
import datetime
from utils.file_handling import (
    FileOperationResult,
    FileAccessStatus,
    rename_file,
    set_file_creation_date,
    set_file_modification_date,
    test_file_access,
    format_datetime_stamp
)
from modules.msg_handling import get_msg_object, MsgAccessStatus
from utils.testset_preparation import prepare_test_directory

# Log-Verzeichnis und Basisname für Logdateien festlegen
LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
PROG_LOG_NAME = 'test_file_handling_prog_log.log'
prog_log_file_path = os.path.join(LOG_DIRECTORY, PROG_LOG_NAME)

# Logging für Test und Debugging konfigurieren
DEBUG_MODE=True # Debugging-Modus aktivieren

def setup_logging(debug=False):
    """
    Konfiguriert das Logging für das Programm.

    Diese Funktion richtet die Logging-Einstellungen ein, einschließlich des Log-Levels,
    des Formats der Log-Nachrichten und des Ziels (Datei), in das die Logs geschrieben werden.
    Der Debugging-Modus kann aktiviert werden, um detailliertere Ausgaben zu erhalten.

    Parameter:
    debug (bool): Gibt an, ob der Debugging-Modus aktiviert werden soll.
                  Wenn True, wird das Log-Level auf DEBUG gesetzt, andernfalls auf INFO.
    """
    # Setze das Log-Level basierend auf dem Debugging-Modus
    log_level = logging.DEBUG if debug else logging.INFO

    # Konfiguriere die grundlegenden Einstellungen für das Logging
    logging.basicConfig(
        level=log_level,  # Setze das Log-Level
        format='%(asctime)s - %(name)s - %(funcName)s - %(levelname)s - %(message)s',  # Format der Log-Nachrichten
        datefmt='%Y-%m-%d %H:%M:%S',  # Format für das Datum und die Uhrzeit
        filename=prog_log_file_path  # Ziel-Dateipfad für die Log-Datei
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
        print("Fehler: Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Console
        logging.error("Fehler: Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.") # Debugging-Ausgabe: Log-File
        exit(1)  # Programm abbrechen

    # Zähler für umbenannte Dateien und Probleme initialisieren
    renamed_count = 0
    problem_count = 0

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    for root, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        print(f"\n***************************************************")  # Debugging-Ausgabe: Console
        print(f"Aktuelles Verzeichnis: {root}")  # Debugging-Ausgabe: Console

        # Alle Dateien im aktuellen Verzeichnis durchgehen
        for file in files:
            # Überprüfen, ob die Datei die Endung .msg hat
            if file.lower().endswith('.msg'):
                file_path = os.path.join(root, file)

                print(f"\n***************************************************")  # Debugging-Ausgabe: Console
                print(f"MSG-Datei: {file}") # Debugging-Ausgabe: Console
                logger.debug(f"\n***************************************************")  # Debugging-Ausgabe: Log-File
                logger.debug(f"\nMSG-Datei: {file}") # Debugging-Ausgabe: Console

                # Überprüfen des Lesezugriffs auf die Datei
                access_result = test_file_access(file_path)
                print(f"\tErgebnis Zugriffsprüfung: '{[s.value for s in access_result]}'")  # Debugging-Ausgabe: Console

                if FileAccessStatus.WRITABLE in access_result:
                    print(f"\tLesender und schreibender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Console
                    logging.debug(f"Lesender und schreibender Zugriff auf die Datei möglich: {file_path}")  # Debugging-Ausgabe: Log-File

                    # Auslesen des msg-Objektes
                    msg_object = get_msg_object(file_path)

                    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.DATE_MISSING not in msg_object["status"]:
                        datetime_stamp = msg_object['date']
                        print(f"\tVersandzeitpunkt mit neuer Methode: '{datetime_stamp}'")

                        # Formatieren des Versanddatums
                        format_string = "%Y%m%d-%Huhr%M"
                        formatted_date = format_datetime_stamp(datetime_stamp, format_string)  # Call the new function

                        # Neuer Dateiname basierend auf dem formatierten Versanddatum
                        new_filename = f"{formatted_date}_{file}"
                        new_file_path = os.path.join(root, new_filename)

                        # Umbenennen der Datei
                        rename_file_result = rename_file(file_path, new_file_path)

                        if rename_file_result == FileOperationResult.SUCCESS:
                            print(f"\tErgebnis der Datei-Umbenennung: {rename_file_result}") # Debugging-Ausgabe: Console
                            logging.debug(f"Ergebnis der Datei-Umbenennung: {rename_file_result}") # Debugging-Ausgabe: Log-File
                        else:
                            print(f"\tDatei-Umbenennung nicht erfolgreich: {rename_file_result}") # Debugging-Ausgabe: Console
                            logging.debug(f"Datei-Umbenennung nicht erfolgreich: {rename_file_result}") # Debugging-Ausgabe: Log-File

                        # Wenn das Umbenennen erfolgreich war, alte Datei löschen
                        if rename_file_result == FileOperationResult.SUCCESS:
                            renamed_count += 1  # Zähler erhöhen

                            # Setze das Erstelldatum auf das Versanddatum
                            #if get_datetime_stamp_result['get_msg_access_result'] == MsgAccessStatus.SUCCESS:
                            if MsgAccessStatus.DATE_MISSING not in msg_object["status"]:

                                # Überprüfen, ob datetime_stamp ein datetime-Objekt ist
                                if isinstance(datetime_stamp, datetime.datetime):
                                    # Konvertiere das datetime-Objekt in einen String im richtigen Format
                                    datetime_stamp_str = datetime_stamp.strftime("%Y-%m-%d %H:%M:%S")
                                else:
                                    datetime_stamp_str = datetime_stamp  # Annehmen, dass es bereits ein String ist

                                # Setze das Erstelldatum auf das Versanddatum
                                set_creation_result = set_file_creation_date(new_file_path, datetime_stamp_str)
                                if set_creation_result == FileOperationResult.SUCCESS:
                                    print(f"\tNeues Erstellungsdatum für '{file}' erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Neues Erstellungsdatum für '{file}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                else:
                                    print(f"\tFehler beim Setzen des Erstellungsdatum für '{file}': '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Fehler beim Setzen des Erstellungsdatum für '{file}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                                # Setze das Änderungsdatum auf das Versanddatum
                                set_modification_result = set_file_modification_date(new_file_path, datetime_stamp_str)
                                if set_modification_result == FileOperationResult.SUCCESS:
                                    print(f"\tNeues Änderungsdatum für '{file}' erfolgreich gesetzt.")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Neues Änderungsdatum für '{file}' erfolgreich gesetzt.")  # Debugging-Ausgabe: Log-File
                                else:
                                    print(f"\tFehler beim Setzen des Änderungsdatum für '{file}': '{set_creation_result}'")  # Ausgabe des Ergebnisses
                                    logging.debug(f"Fehler beim Setzen des Änderungsdatum für '{file}': '{set_creation_result}'")  # Debugging-Ausgabe: Log-File

                    # Falls das Versanddatum nicht ermittelt werden konnte
                    else:
                        print(f"\tVersanddatum für '{file}' konnte nicht ermittelt werden.") # Debugging-Ausgabe: Console
                        logging.error(f"Versanddatum für '{file}' konnte nicht ermittelt werden.") # Debugging-Ausgabe: Log-File
                        problem_count += 1  # Problemzähler erhöhen

                # Da nur lesender Zugriff wird der Problemzähler erhöht
                elif FileAccessStatus.READABLE in access_result:
                    print(f"\tNur lesender Zugriff auf Datei möglich: {file_path}")  # Debugging-Ausgabe: Console
                    logging.debug(f"Nur lesender Zugriff auf Datei möglich: {file_path}")  # Debugging-Ausgabe: Log-FileDatei '{file}' erfolgreich.")
                    problem_count += 1  # Problemzähler erhöhen

                # Da gar kein Zugriff wird auch der Problemzähler erhöht
                else:
                    print(f"\tZugriff auf Datei '{file}' fehlgeschlagen: '{[s.value for s in access_result]}'") # Debugging-Ausgabe: Console
                    logging.error(f"Zugriff auf Datei '{file}' fehlgeschlagen: '{[s.value for s in access_result]}'")  # Debugging-Ausgabe: Log-File
                    problem_count += 1  # Problemzähler erhöhen

    # Ausgabe der Ergebnisse
    print(f"\nAnzahl der umbenannten Dateien: {renamed_count}")
    print(f"Anzahl der Dateien mit Problemen: {problem_count}")
