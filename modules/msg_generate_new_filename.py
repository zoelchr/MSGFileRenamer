# Das Modul behinhaltet eine Funktion, der ein Filename übergeben wird und dann mit Hilfe der MSG-Metadaten einen neuen Dateinamen generiert und diesen als Ergbenis zurückliefter

import os
import logging
from datetime import datetime  # Stellen Sie sicher, dass nur die Klasse datetime importiert wird
from modules.msg_handling import parse_sender_msg_file, \
    load_known_senders, convert_to_utc_naive, format_datetime, \
    custom_sanitize_text, truncate_filename_if_needed, MsgAccessStatus, get_msg_object2
from enum import Enum
from dataclasses import dataclass

class ContainEmailSenderInfo(Enum):
    FOUND_SENDERSTRING = "Name of email sender found"
    FOUND_SENDEREMAIL = "Sender email found"
    FOUND_TABELBASED_SENDEREMAIL = "Sender email found in table"
    NOT_FOUND = "Sender email not found"
    UNKNOWN = "Status is unkown"

@dataclass
class MsgFilenameResult:
    datetime_stamp: datetime
    formatted_timestamp: str
    sender_name: str
    sender_email: str
    msg_subject: str
    msg_subject_sanitized: str
    new_msg_filename: str
    new_truncated_msg_filename: str
    is_msg_filename_truncated: bool

logger = logging.getLogger(__name__)

# Liste der bekannten Email-Absender aus einer CSV-Datei
LIST_OF_KNOWN_SENDERS = r'D:\Dev\pycharm\MSGFileRenamer\config\known_senders_private.csv'

PRINT_RESULT = False

def generate_new_msg_filename(msg_path_and_filename, max_path_length=260):
    format_string = "%Y%m%d-%Huhr%M"  # Beispiel für das gewünschte Format für Zeitstempel

    # 0. Schritt: Laden der bekannten Sender aus der CSV-Datei
    logger.debug(f"\t\tVersuche Einlesen Liste bekannter Email-Absender aus CSV-Datei: {LIST_OF_KNOWN_SENDERS}'")  # Debugging-Ausgabe: Log-File
    known_senders_df = load_known_senders(LIST_OF_KNOWN_SENDERS)  # Laden der bekannten Sender aus der CSV-Datei als Dataframe
    logger.debug(f"Liste der bekannten Email-Absender: {known_senders_df}'")  # Debugging-Ausgabe: Log-File

    # Auslesen des msg-Objektes
    msg_object = get_msg_object2(msg_path_and_filename)

    # 1. Schritt: Absender-String aus der MSG-Datei abrufen mit alternativer Methode
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object["status"]:
        found_msg_sender_string = msg_object["sender"]  # Absender extrahieren
        print(f"\t\tSchritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'") # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'")  # Debugging-Ausgabe: Log-File
    else:
        found_msg_sender_string = ""
        print(f"\t\tSchritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Console
        logger.warning(f"Schritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Log-File

    # 2. Schritt: Absender-Email aus dem gefundenen Absender-String mit Hilfe einer Regex-Methode extrahieren
    parsed_sender_email = {"sender_name": "", "sender_email": "", "contains_sender_email": False} # Defaultwerte für parsed_sender_email setzen

    # Wenn der Absender-String aus der MSG-Datei erfolgreich ausgelesen wurde, dann wird die Absender-Email aus dem Absender-String extrahiert
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object["status"]:
        parsed_sender_email = parse_sender_msg_file(found_msg_sender_string)
        print(f"\t\tSchritt 2: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 2: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Log-File
    else:
        print(f"\tSchritt 2: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt.")
        logger.debug(f"Schritt 2: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt.")

    #  3. Schritt: Wenn im 2. Schritt keine Absender-Email gefunden wurde, dann in der Liste der bekannten Email-Absender nachsehen, ob der Absendername enthalten ist
    if not parsed_sender_email["contains_sender_email"]:

        known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]

        if not known_sender_row.empty:
            parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
            parsed_sender_email["contains_sender_email"] = True
            print(f"\tSchritt 3: In Tabelle gefundene Absender-Email: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Console
            logger.debug(f"Schritt 3: In Tabelle gefundene Absender-Email: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Log-File
        else:
            parsed_sender_email["contains_sender_email"] = False
            print(f"\t\tSchritt 3: In Tabelle keine Absender-Email für folgenden Absender-String gefunden: '{found_msg_sender_string}'")  # Debugging-Ausgabe: Console
            logger.warning(f"Schritt 3: In Tabelle keine Absender-Email für folgenden Absender-String gefunden: '{found_msg_sender_string}'")  # Debugging-Ausgabe: Log-File
    else:
        print(f"\t\tSchritt 3: Kein Nachschlagen in der Tabelle der bekannten Email-Absender erforderlich.")  # Debugging-Ausgabe: Console
        logger.warning(f"Kein Nachschlagen in der Tabelle der bekannten Email-Absender erforderlich.")

    # 4. Schritt: Versanddatum abrufen und konvertieren
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.DATE_MISSING not in msg_object["status"]:
        datetime_stamp = msg_object['date']
        datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
        print(f"\t\tSchritt 4: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Log-File

        # 4a. Schritt: Formatiertes Versanddatum ermitteln
        formatted_timestamp = format_datetime(datetime_stamp, format_string)  # Formatieren des Zeitstempels
        print(f"\t\tSchritt 4a: Formatiertes Versanddatum: '{formatted_timestamp}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4a: Formatiertes Versanddatum: '{formatted_timestamp}'")  # Debugging-Ausgabe: Log-File
    else:
        datetime_stamp = ""
        formatted_timestamp = ""
        print(f"\t\tSchritt 4: Kein Versanddatum gefunden: '{msg_object['status']}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4: Kein Versanddatum gefunden: '{msg_object['status']}'")  # Debugging-Ausgabe: Log-File

    # 5. Schritt: Betreff ermitteln mit neuer Methode
    #msg_object = get_msg_object2(msg_path_and_filename)  # Auslesen des msg-Objektes

    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SUBJECT_MISSING not in msg_object["status"]:
        msg_subject = msg_object["subject"]
        print(f"\t\tSchritt 5: Ermittelter Betreff: '{msg_subject}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 5: Betreff ermitteln: '{msg_subject}'")  # Debugging-Ausgabe: Log-File

        # 6. Schritt: Betreff bereinigen
        msg_subject_sanitized = custom_sanitize_text(msg_subject)  # Betreff bereinigen
        print(f"\t\tSchritt 6: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 6: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Log-File

    else:
        msg_subject = ""
        msg_subject_sanitized = ""

    # 7. Schritt: Neuen Namen der Datei festlegen
    new_msg_filename = f"{formatted_timestamp}_{parsed_sender_email['sender_email']}_{msg_subject_sanitized}.msg"
    msg_pathname = os.path.dirname(msg_path_and_filename)  # Verzeichnisname der MSG-Datei
    print(f"\t\tSchritt 7: Neuer Dateiname: '{new_msg_filename}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 7: Neuer Dateiname: '{new_msg_filename}'")  # Debugging-Ausgabe: Log-File

    new_msg_path_and_filename = os.path.join(msg_pathname, new_msg_filename)  # Neuer absoluter Dateipfad
    print(f"\t\tSchritt 8: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 8: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Log-File

    # 9. Schritt: Kürzen des Dateinamens, falls nötig
    if len(new_msg_path_and_filename) > max_path_length:
        new_truncated_msg_path_and_filename = truncate_filename_if_needed(new_msg_path_and_filename, max_path_length, "...msg")
        new_truncated_msg_filename = os.path.basename(new_truncated_msg_path_and_filename)
        is_msg_filename_truncated = True
        print(f"\t\tSchritt 9: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 9: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Log-File
    else:
        new_truncated_msg_path_and_filename = new_msg_path_and_filename
        new_truncated_msg_filename = new_msg_filename
        is_msg_filename_truncated = False
        print(f"\t\tSchritt 9: Kein Kürzen des Dateinames erforderlich.")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 9: Kein Kürzen des Dateinames erforderlich.")  # Debugging-Ausgabe: Log-File

    # Ausgabe aller Informationen
    if PRINT_RESULT:
        print(f"\n\t\t'generate_new_msg_filename' - Ausgabe aller Informationen")
        print(f"\t\tVollständiger Pfad der MSG-Datei: {msg_path_and_filename}")
        print(f"\t\tEnthält Email-Absender: {parsed_sender_email['contains_sender_email']}")
        print(f"\t\tGefundener Email-Absender: {parsed_sender_email['sender_email']}")
        print(f"\t\tExtrahierter Versandzeitpunkt: {datetime_stamp}")
        print(f"\t\tFormatierter Versandzeitpunkt: {formatted_timestamp}")
        print(f"\t\tExtrahierter Betreff: {msg_subject}")
        print(f"\t\tBereinigter Betreff: {msg_subject_sanitized}")
        print(f"\t\tNeuer Dateiname: {new_msg_filename}")
        print(f"\t\tKürzung Dateiname erforderlich: {is_msg_filename_truncated}")
        print(f"\t\tNeuer gekürzter Dateiname: {new_truncated_msg_filename}\n")

    # Rückgabe der gewünschten Informationen als MsgFilenameResult
    return MsgFilenameResult(
        datetime_stamp=datetime_stamp,
        formatted_timestamp=formatted_timestamp,
        sender_name=parsed_sender_email["sender_name"],
        sender_email=parsed_sender_email["sender_email"],
        msg_subject=msg_subject,
        msg_subject_sanitized=msg_subject_sanitized,
        new_msg_filename=new_msg_filename,
        new_truncated_msg_filename=new_truncated_msg_filename,
        is_msg_filename_truncated=is_msg_filename_truncated
    )