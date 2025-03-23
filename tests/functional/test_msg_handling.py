"""
test_msg_handling_old.py

Dieses Modul enthält Tests für die Funktionen im Modul msg_handling.py.
Es stellt sicher, dass die Funktionen korrekt arbeiten und die erwarteten Ergebnisse liefern.

Die Tests umfassen:
- Überprüfung der Rückgabewerte von Funktionen wie get_sender_msg_file, get_subject_msg_file,
  und get_date_sent_msg_file.
- Validierung der Funktionalität von parse_sender_msg_file zur korrekten Extraktion von Senderinformationen.
- Sicherstellung, dass das Erstellen und Protokollieren in Logdateien ordnungsgemäß funktioniert.
- Tests für die Bereinigung von Texten durch custom_sanitize_text und die Kürzung von Dateinamen
  durch truncate_filename_if_needed.

Verwendung:
Führen Sie dieses Skript aus, um alle Tests durchzuführen und die Ergebnisse zu überprüfen.

Beispiel:
    python -m unittest test_msg_handling_old.py
"""

import os
import logging

# Importieren von Funktionen aus den Modulen msg_handling
from modules.msg_handling import (
    parse_sender_msg_file,
    create_log_file,
    log_entry,
    load_known_senders,
    convert_to_utc_naive,
    format_datetime,
    custom_sanitize_text,
    truncate_filename_if_needed,
    reduce_thread_in_msg_message,
    MsgAccessStatus,
    get_msg_object
)

# Importieren von Funktionen aus meinem Modul testset_preparation
from utils.testset_preparation import prepare_test_directory

# Log-Verzeichnis und Basisname für Logdateien festlegen
LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
PROG_LOG_NAME = 'test_msg_handling_prog_log.log'
prog_log_file_path = os.path.join(LOG_DIRECTORY, PROG_LOG_NAME)

# Logging für Test und Debugging konfigurieren
DEBUG_MODE=True # Aktiviert den Debugging-Modus für detaillierte Ausgaben

def setup_logging(debug=False):
    """
    Konfiguriert das Logging für das Programm.

    Parameter:
    debug (bool): Gibt an, ob der Debugging-Modus aktiviert werden soll.
    """
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

    # Maximal zulässige Pfadlänge für Windows 11
    max_path_length = 260

    # Logging für MSG-Bearbeitung konfigurieren (Excel-Datei)
    MSG_LOG_NAME = 'msg_log'
    MSG_LOG_TABLE_HEADER= columns=["Fortlaufende Nummer", "Verzeichnisname", "Filename", "Sendername", "Senderemail",
                               "Contains Senderemail", "Timestamp", "Formatierter Timestamp", "Betreff",
                               "Bereinigter Betreff", "Neuer Filename", "Neuer gekürzter Filename", "Neuer Dateipfad"]

    # Liste der bekannten Sender aus einer CSV-Datei
    LIST_OF_KNOWN_SENDERS = r'D:\Dev\pycharm\MSGFileRenamer\config\known_senders_private.csv'

    # Beispiel für das gewünschte Format für Zeitstempel
    format_timestamp_string = "%Y%m%d-%Huhr%M"  # Beispiel: JJJJMMTT-HHuhrMM

    # Logging für Test und Debugging konfigurieren
    setup_logging(DEBUG_MODE)
    logger = logging.getLogger(prog_log_file_path)
    # Logging für die Programm-Ausführung
    logger.info("Programm gestartet")
    logger.info(f"Debug-Modus: {DEBUG_MODE}")
    logging.info(f"Programm Logdatei: {prog_log_file_path}")

    # MSG-Logdatei erstellen und den Pfad ausgeben
    msg_log_file_path = create_log_file(MSG_LOG_NAME, LOG_DIRECTORY, MSG_LOG_TABLE_HEADER)
    print(f"MSG Logdatei erstellt: {msg_log_file_path}") # Debugging-Ausgabe: Console

    # Wenn TARGET_DIRECTORY_TEST_DATA noch nicht existiert, dann erstellen
    if not os.path.isdir(TARGET_DIRECTORY_TEST_DATA):
        os.makedirs(TARGET_DIRECTORY_TEST_DATA, exist_ok=True)

    # Bereite das Zielverzeichnis vor und überprüfe den Erfolg
    success = prepare_test_directory(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA)
    if not success:
        print("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")
        logging.error("Fehler: Die Vorbereitung des Testverzeichnisses ist fehlgeschlagen. Das Programm wird abgebrochen.")
        exit(1)  # Programm abbrechen

    # Laden der bekannten Sender aus der CSV-Datei
    known_senders_df = load_known_senders(LIST_OF_KNOWN_SENDERS)

    # Zähler initialisieren
    msg_file_count = 0
    msg_sender_found = 0
    msg_date_found = 0
    msg_subject_found = 0
    msg_body_found = 0

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    for root, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        print(f"\n***************************************************")  # Debugging-Ausgabe: Console
        print(f"Aktuelles Verzeichnis: {root}")  # Debugging-Ausgabe: Console

        # Alle Dateien im aktuellen Verzeichnis durchgehen
        for file in files:
            file_path = os.path.join(root, file)

            if file.endswith(".msg"):
                msg_file_count += 1 # Zähler für die Anzahl der MSG-Dateien erhöhen

                # Absender-String aus der MSG-Datei abrufen
                print(f"\n***************************************************")  # Debugging-Ausgabe: Console
                print(f"MSG-Datei: {file}") # Debugging-Ausgabe: Console
                logger.debug(f"\n***************************************************")  # Debugging-Ausgabe: Log-File
                logger.debug(f"\nMSG-Datei: {file}") # Debugging-Ausgabe: Console

                # Auslesen des msg-Objektes
                msg_object = get_msg_object(file_path)

                # Prüfen, ob die Datei erfolgreich verarbeitet wurde und der Sender existiert
                if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object[
                    "status"]:
                    sender = msg_object["sender"] # Absender extrahieren
                    msg_sender_found += 1 # Zähler für gefundene Absender erhöhen
                else:
                    sender = "Absender unbekannt" # Standardwert, wenn kein Absender gefunden wurde

                # Ausgabe des extrahierten Absenders
                print(f"\tExtrahierter Sender: {sender}")  # Debugging-Ausgabe: Console
                logger.debug(f"\nExtrahierter Sender: {sender}")  # Debugging-Ausgabe: Log-File

                # Aufruf der Funktion zur Analyse der Absender-Email
                parsed_sender_email = parse_sender_msg_file(sender)
                if parsed_sender_email["contains_sender_email"]:
                    print(f"\tParsed Sender-Email: {parsed_sender_email['sender_email']}")  # Debugging-Ausgabe: Console
                    logger.debug(f"Parsed Sender-Email: {parsed_sender_email['sender_email']}")  # Debugging-Ausgabe: Log-File
                else:
                    print(f"\tKeine Absender-Email gefunden.") # Debugging-Ausgabe: Console
                    logger.debug(f"Keine Absender-Email gefunden.")  # Debugging-Ausgabe: Log-File

                # Überprüfen, ob eine Absender-Email vorhanden ist
                if not parsed_sender_email["sender_email"]:  # Nur ausführen, wenn keine Absender-Email gefunden wurde

                    # Überprüfen, ob der Sendername in der Tabelle der bekannten Sender vorhanden ist
                    known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]
                    if not known_sender_row.empty:
                        # Wenn der Sendername bekannt ist, die Email-Adresse hinzufügen
                        parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
                        parsed_sender_email["contains_sender_email"] = True
                        print(f"\tSender-Email aus Tabelle: {parsed_sender_email['sender_email']}")  # Ausgabe der Absender-Email aus der Tabelle
                        logger.debug(f"Sender-Email aus Tabelle: {parsed_sender_email['sender_email']}")  # Debugging-Ausgabe: Log-File
                    else:
                        parsed_sender_email["contains_sender_email"] = False # Debugging-Ausgabe: Console
                        logger.debug(f"Keine Sender-Email in Tabelle gefunden.")  # Debugging-Ausgabe: Log-File
                else:
                    parsed_sender_email["contains_sender_email"] = True  # Wenn eine Email gefunden wurde

                # Versandzeitpunkt abrufen und konvertieren
                if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.DATE_MISSING not in msg_object["status"]:
                    datetime_stamp = msg_object["date"] # Datum extrahieren
                    datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
                    msg_date_found += 1 # Zähler für gefundene Daten erhöhen
                else:
                    datetime_stamp = "Datum unbekannt" # Standardwert, wenn kein Datum gefunden wurde

                # Ausgabe des extrahierten Versandzeitpunkts
                print(f"\tExtrahiertes Versandzeitpunkt: {datetime_stamp}") # Debugging-Ausgabe: Console
                logger.debug(f"\nExtrahiertes Versandzeitpunkt: {datetime_stamp}")  # Debugging-Ausgabe: Log-File

                # Formatieren des Zeitstempels
                formatted_timestamp = format_datetime(datetime_stamp, format_timestamp_string) # Formatieren des Zeitstempels
                print(f"\tFormatierter Versandzeitpunkt: {formatted_timestamp}") # Debugging-Ausgabe: Console
                logger.debug(f"\nFormatierter Versandzeitpunkt: {formatted_timestamp}")  # Debugging-Ausgabe: Log-File

                # Betreff auslesen
                if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SUBJECT_MISSING not in msg_object["status"]:
                    msg_subject = msg_object["subject"] # Betreff extrahieren
                    msg_subject_found += 1 # Zähler für gefundene Betreffs erhöhen
                else:
                    msg_subject = "Betreff unbekannt" # Standardwert, wenn kein Betreff gefunden wurde

                # Betreff bereinigen
                sanitized_subject = custom_sanitize_text(msg_subject) # Bereinigung des Betreffs
                print(f"\tBereinigter Betreff: {sanitized_subject}") # Debugging-Ausgabe: Console
                logger.debug(f"Bereinigter Betreff: {sanitized_subject}")  # Debugging-Ausgabe: Log-File

                # Neuen Namen der Datei festlegen
                new_filename = f"{formatted_timestamp}_{parsed_sender_email['sender_email']}_{sanitized_subject}.msg"  # Beispiel für den neuen Dateinamen
                new_file_path = os.path.join(root, new_filename) # Vollständiger Pfad für die neue Datei
                print(f"\tNeuer Dateiname: {new_filename}") # Debugging-Ausgabe: Console
                logger.debug(f"Neuer Dateiname: {new_filename}")  # Debugging-Ausgabe: Log-File

                # Kürzen des Dateinamens, falls nötig
                new_truncated_path = truncate_filename_if_needed(new_file_path, max_path_length, "<>.msg") # Kürzen des Dateinamens

                # Gekürzter Dateiname ohne Pfad
                new_truncated_filename = os.path.basename(new_truncated_path)
                print(f"\tNeuer gekürzter Dateiname: {new_truncated_filename}") # Debugging-Ausgabe: Console
                logger.debug(f"Neuer gekürzter Dateiname: {new_truncated_filename}")  # Debugging-Ausgabe: Log-File

                # Eigentliche Nachricht
                if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.BODY_MISSING not in msg_object["status"]:
                    msg_body = msg_object["body"] # Nachricht extrahieren
                    msg_body_found += 1 # Zähler für gefundene Nachrichten erhöhen
                    print(f"\tNachricht erfolgreich ausgelesen.")  # Debugging-Ausgabe: Console
                    logger.debug(f"Nachricht erfolgreich ausgelesen.")  # Debugging-Ausgabe: Log-File
                else:
                    msg_body = ""
                    print(f"\tNachricht nicht erfolgreich ausgelesen.")  # Debugging-Ausgabe: Console
                    logger.debug(f"Nachricht nicht erfolgreich ausgelesen.")  # Debugging-Ausgabe: Log-File

                # Die Nachricht durch Entfernen von älteren Emails reduzieren
                reduce_result = reduce_thread_in_msg_message(msg_body,2) # Reduzieren der Anzahl der älteren E-Mails
                print(f"\tAnzahl der reduzierten Emails: {reduce_result['deleted_count']}") # Ausgabe der Anzahl reduzierter E-Mails

                # Logeintrag erstellen
                entry = {
                    "Fortlaufende Nummer": msg_file_count,
                    "Verzeichnisname": root,
                    "Filename": file,
                    "Sendername": parsed_sender_email["sender_name"],
                    "Senderemail": parsed_sender_email["sender_email"],
                    "Contains Senderemail": parsed_sender_email["contains_sender_email"],
                    "Timestamp": datetime_stamp,
                    "Formatierter Timestamp": formatted_timestamp,
                    "Betreff": msg_subject,
                    "Bereinigter Betreff": sanitized_subject,
                    "Neuer Filename": new_filename,
                    "Neuer gekürzter Filename": new_truncated_path,
                    "Neuer Dateipfad": new_file_path,
                    "Nachricht": msg_body,
                    "Gekürzte Nachricht": reduce_result['new_email_text'],
                    "Anzahl der reduzierten Emails": reduce_result['deleted_count']
                }

                # Eintrag ins Logfile hinzufügen
                log_entry(msg_log_file_path, entry)

    # Zusammenfassung der gefundenen Informationen
    print(f"Gefundene MSG-Files: {msg_file_count}")
    print(f"Sender gefunden: {msg_sender_found}")
    print(f"Datum gefunden: {msg_date_found}")
    print(f"Betreff gefunden: {msg_subject_found}")
    print(f"Nachricht gefunden: {msg_body_found}")