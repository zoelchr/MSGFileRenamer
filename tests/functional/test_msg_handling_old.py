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

# Importieren von Funktionen aus den Modulen msg_handling und testset_preparation
from modules.msg_handling import (
    get_sender_msg_file,
    parse_sender_msg_file,
    create_log_file,
    log_entry,
    load_known_senders,
    get_date_sent_msg_file,
    convert_to_utc_naive,
    format_datetime,
    get_subject_msg_file,
    custom_sanitize_text,
    truncate_filename_if_needed,
    reduce_thread_in_msg_message,
    get_message_msg_file,
    get_message_msg_file2,
    get_sender_msg_file2,
    get_subject_msg_file2,
    MsgAccessStatus,
    get_date_sent_msg_file2,
    get_msg_object2,
    check_email_security,
    print_msg_to_pdf
)
# Importieren von Funktionen aus meinem Modul testset_preparation
from utils.testset_preparation import prepare_test_directory

# Log-Verzeichnis und Basisname für Logdateien festlegen
LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
PROG_LOG_NAME = 'test_msg_handling_prog_log.log'
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
    print(f"MSG Logdatei erstellt: {msg_log_file_path}")

    # Zähler für die fortlaufende Nummer initialisieren
    counter = 1

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
    msg_sender_not_found = 0

    # Durchsuchen des Zielverzeichnisses nach MSG-Dateien
    for root, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        for file in files:
            file_path = os.path.join(root, file)
            print(f"Aktuelles Verzeichnis: {root}")
            if file.endswith(".msg"):
                msg_file_count += 1

                # Auslesen des msg-Objektes mit der neuen Methode
                msg_object2 = get_msg_object2(file_path)

                # Absender-String aus der MSG-Datei abrufen
                sender = get_sender_msg_file(file_path)
                print(f"\tDatei: {file}")
                print(f"\tExtrahierter Sender (get_sender_msg_file): {sender}")
                logger.debug(f"\n***************************************************")  # Debugging-Ausgabe: Log-File
                logger.debug(f"\nExtrahierter Sender (get_sender_msg_file): {sender}")  # Debugging-Ausgabe: Log-File

                # Absender-String aus der MSG-Datei abrufen mit neuer Methode
                get_sender_result = get_sender_msg_file2(file_path)
                if get_sender_result["get_msg_sender_result"] == MsgAccessStatus.SUCCESS:
                    print(f"\tExtrahierter Sender (get_sender_msg_file2): {get_sender_result['msg_sender']}")
                else:
                    print(f"\tSender konnte nicht extrahiert werden, da {get_sender_result['get_msg_sender_result']} auftrat.")
                    logger.error(f"Sender konnte nicht extrahiert werden, da {get_sender_result['get_msg_sender_result']} auftrat.")  # Debugging-Ausgabe: Log-File
                    continue  # Nächste Datei verarbeiten

                # Absender-String aus dem MSG-Objekt msg_object2 abrufen
                # Prüfen, ob die Datei erfolgreich verarbeitet wurde und der Sender existiert
                if MsgAccessStatus.SUCCESS in msg_object2["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object2[
                    "status"]:
                    sender3 = msg_object2["sender"]
                else:
                    sender3 = "Absender unbekannt"
                print(f"\tExtrahierter Sender (msg_object2): {sender3}")

                # Aufruf der Funktion zur Analyse der Absender-Email
                parsed_sender_email = parse_sender_msg_file(sender)
                if parsed_sender_email["contains_sender_email"]:
                    print(f"\tParsed Sender-Email: {parsed_sender_email['sender_email']}")  # Ausgabe der analysierten Sender-Email
                    logger.debug(f"Parsed Sender-Email: {parsed_sender_email['sender_email']}")  # Debugging-Ausgabe: Log-File
                else:
                    print(f"\tKeine Absender-Email gefunden.")
                    logger.debug(f"Keine Absender-Email gefunden.")  # Debugging-Ausgabe: Log-File

                # Überprüfen, ob eine Sender-Email vorhanden ist
                if not parsed_sender_email["sender_email"]:  # Nur ausführen, wenn keine Sender-Email gefunden wurde
                    # Überprüfen, ob der Sendername in der Tabelle der bekannten Sender vorhanden ist
                    known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]
                    if not known_sender_row.empty:
                        # Wenn der Sendername bekannt ist, die Email-Adresse hinzufügen
                        parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
                        parsed_sender_email["contains_sender_email"] = True
                        print(f"\tSender-Email aus Tabelle: {parsed_sender_email['sender_email']}")  # Ausgabe der Sender-Email aus der Tabelle
                        logger.debug(f"Sender-Email aus Tabelle: {parsed_sender_email['sender_email']}")  # Debugging-Ausgabe: Log-File
                    else:
                        parsed_sender_email["contains_sender_email"] = False
                        logger.debug(f"Keine Sender-Email in Tabelle gefunden.")  # Debugging-Ausgabe: Log-File
                else:
                    parsed_sender_email["contains_sender_email"] = True  # Wenn eine Email gefunden wurde

                # Versanddatum abrufen und konvertieren
                datetime_stamp = get_date_sent_msg_file(os.path.join(root, file))  # Versanddatum abrufen
                datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
                print(f"\tExtrahierter Versandzeitpunkt: {datetime_stamp}")  # Ausgabe des Versanddatums

                # Versanddatum abrufen und konvertieren mit neuer Methode
                get_date_result = get_date_sent_msg_file2(os.path.join(root, file))
                print(f"\tErgebnis des Zugriffs auf das Versanddatum: {get_date_result}")
                if get_date_result["get_msg_access_result"] == MsgAccessStatus.SUCCESS:
                    datetime_stamp2 = get_date_result["msg_date"]
                    datetime_stamp2 = convert_to_utc_naive(datetime_stamp2)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
                    if datetime_stamp2 == datetime_stamp:
                        print(f"\tExtrahiertes Datum von 'get_date_msg_file' und 'get_date_msg_file2' sind gleich.")

                if MsgAccessStatus.SUCCESS in msg_object2["status"] and MsgAccessStatus.DATE_MISSING not in msg_object2[
                    "status"]:
                    datetime_stamp3 = msg_object2["date"]
                else:
                    datetime_stamp3 = "Datum unbekannt"
                print(f"\tExtrahierter Sender (msg_object2): {datetime_stamp3}")

                # Formatieren des Zeitstempels
                formatted_timestamp = format_datetime(datetime_stamp, format_timestamp_string)
                print(f"\tFormatierter Versandzeitpunkt: {formatted_timestamp}")

                # Betreff auslesen
                msg_subject = get_subject_msg_file(file_path)
                print(f"\tExtrahierter Betreff von 'get_subject_msg_file': {msg_subject}")

                # Betreff auslesen mit neuer Methode
                get_subject_result = get_subject_msg_file2(file_path)
                if get_subject_result["get_msg_access_result"] == MsgAccessStatus.SUCCESS:
                    msg_subject2 = get_subject_result["msg_subject"]
                    if msg_subject2 == msg_subject:
                        print(f"\tExtrahierter Betreff von 'get_subject_msg_file' und 'get_subject_msg_file2' sind gleich.")

                # Betreff auslesen mit dritter Methode
                if MsgAccessStatus.SUCCESS in msg_object2["status"] and MsgAccessStatus.SUBJECT_MISSING not in msg_object2["status"]:
                    msg_subject3 = msg_object2["subject"]
                else:
                    msg_subject3 = "Betreff unbekannt"
                    if msg_subject3 == msg_subject:
                        print(
                            f"\tExtrahierter Betreff von 'get_subject_msg_file' und 'get_msg_object2' sind gleich.")

                # Betreff bereinigen
                sanitized_subject = custom_sanitize_text(msg_subject)
                print(f"\tBereinigter Betreff: {sanitized_subject}")

                # Neuen Namen der Datei festlegen
                new_filename = f"{formatted_timestamp}_{parsed_sender_email['sender_email']}_{sanitized_subject}.msg"  # Beispiel für den neuen Dateinamen
                new_file_path = os.path.join(root, new_filename)
                print(f"\tNeuer Dateiname: {new_filename}")

                # Kürzen des Dateinamens, falls nötig
                new_truncated_path = truncate_filename_if_needed(new_file_path, max_path_length, "<>.msg")

                # Gekürzter Dateiname ohne Pfad
                new_truncated_filename = os.path.basename(new_truncated_path)
                print(f"\tNeuer gekürzter Dateiname: {new_truncated_filename}")

                # Eigentliche Nachricht auslesen
                get_message_result = get_message_msg_file(file_path)
                message_text = get_message_result["msg_message"]
                formatted_message_text = message_text.replace('\n', ' ').replace('\r', ' ')
                #print(f"\tExtrahierte Nachricht: {formatted_message_text[:50]}...")

                # Eigentliche Nachricht auslesen mit neuer Methode
                get_message_result = get_message_msg_file2(file_path)
                if get_message_result["get_msg_access_result"] == MsgAccessStatus.SUCCESS:
                    message_text2 = get_message_result["msg_message"]
                    formatted_message_text2 = message_text2.replace('\n', ' ').replace('\r', ' ')
                    if message_text2 == message_text:
                        print(f"\tExtrahierte Nachricht von 'get_message_msg_file' und 'get_message_msg_file2' sind gleich.")

                # Eigentliche Nachricht auslesen mit dritter Methode
                if MsgAccessStatus.SUCCESS in msg_object2["status"] and MsgAccessStatus.BODY_MISSING not in msg_object2["status"]:
                    msg_body3 = msg_object2["body"]
                else:
                    msg_body3 = ""
                if msg_body3 == message_text:
                    print(f"\tExtrahierte Nachricht von 'get_message_msg_file' und 'get_msg_object2' sind gleich.")
                else:
                    print(f"\tExtrahierte Nachricht von 'get_message_msg_file' und 'get_message_msg_file2' sind unterschiedlich.")

                # Die Nachricht durch Entfernen von älteren Emails reduzieren
                reduce_result = reduce_thread_in_msg_message(message_text,2)
                print(f"\tAnzahl der reduzierten Emails: {reduce_result['deleted_count']}")

                # print_msg_to_pdf(file_path, file_path + ".pdf")

                # Logeintrag erstellen
                entry = {
                    "Fortlaufende Nummer": counter,
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
                    "Neuer gekürzter Fielname": new_truncated_path,
                    "Neuer Dateipfad": new_file_path,
                    "Nachricht": message_text,
                    "Gekürzte Nachricht": reduce_result['new_email_text'],
                    "Anzahl der reduzierten Emails": reduce_result['deleted_count']
                }

                # Eintrag ins Logfile hinzufügen
                log_entry(msg_log_file_path, entry)

                # Zähler erhöhen
                counter += 1

                # Nutzung
                security_info = check_email_security(file_path)

                print(f"Signiert: {security_info['is_signed']}")
                print(f"Verschlüsselt: {security_info['is_encrypted']}")