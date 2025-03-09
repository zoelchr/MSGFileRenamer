# Das Modul behinhaltet eine Funktion, der ein Filename übergeben wird und dann mit Hilfe der MSG-Metadaten einen neuen Dateinamen generiert und diesen als Ergbenis zurückliefter

import os
import logging
from datetime import datetime  # Stellen Sie sicher, dass nur die Klasse datetime importiert wird
from operator import truediv
from modules.msg_handling import get_date_sent_msg_file, parse_sender_msg_file, \
    get_sender_msg_file, load_known_senders, convert_to_utc_naive, format_datetime, get_subject_msg_file, \
    custom_sanitize_text, truncate_filename_if_needed
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

PRINT_RESULT = True

def generate_new_msg_filename(msg_path_and_filename, max_path_length=260):
    format_string = "%Y%m%d-%Huhr%M"  # Beispiel für das gewünschte Format für Zeitstempel

    # 0. Schritt: Laden der bekannten Sender aus der CSV-Datei
    logger.debug(f"\tVersuche Einlesen Liste bekannter Email-Absender aus CSV-Datei: {LIST_OF_KNOWN_SENDERS}'")  # Debugging-Ausgabe: Log-File
    known_senders_df = load_known_senders(LIST_OF_KNOWN_SENDERS)  # Laden der bekannten Sender aus der CSV-Datei als Dataframe
    logger.debug(f"Liste der bekannten Email-Absender: {known_senders_df}'")  # Debugging-Ausgabe: Log-File

    # 1. Schritt: Absender-String aus der MSG-Datei abrufen
    found_msg_sender_string = get_sender_msg_file(msg_path_and_filename)
    if not found_msg_sender_string.startswith("Fehler beim Auslesen des Senders"):
        print(f"\tSchritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'") # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'")  # Debugging-Ausgabe: Log-File
    else:
        print(f"\tSchritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Log-File

    # 2. Schritt: Absender-Email aus dem gefundenen Absender-String mit Hilfe einer Regex-Methode extrahieren

    # Defaultwerte für parsed_sender_email setzen
    parsed_sender_email = {"sender_name": "", "sender_email": "", "contains_sender_email": False}

    if not (found_msg_sender_string.startswith("Fehler beim Auslesen des Senders") or "Unbekannt" in found_msg_sender_string):
        parsed_sender_email = parse_sender_msg_file(found_msg_sender_string)
        print(f"\tSchritt 2a: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 2a: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Log-File
    else:
        print(f"\tSchritt 2a: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt")
        logger.debug(f"Schritt 2a: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt")

    print(f"\tSchritt 2b: Inhalt von 'parsed_sender_email': '{parsed_sender_email}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 2b: Inhalt von 'parsed_sender_email': '{parsed_sender_email}'")  # Debugging-Ausgabe: Log-File

    #  3. Schritt: Wenn im 2. Schritt keine Absender-Email gefunden wurde, dann in der Liste der bekannten Email-Absender nachsehen, ob der Absendername enthalten ist
    if not parsed_sender_email["contains_sender_email"]:

        known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]

        if not known_sender_row.empty:
             parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
             parsed_sender_email["contains_sender_email"] = True
        else:
             parsed_sender_email["contains_sender_email"] = False

    print(f"\tSchritt 3: Inhalt von 'parsed_sender_email': '{parsed_sender_email}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 3: Inhalt von 'parsed_sender_email': '{parsed_sender_email}'")  # Debugging-Ausgabe: Log-File

    # 4. Schritt: Versanddatum abrufen und konvertieren
    datetime_stamp = get_date_sent_msg_file(msg_path_and_filename)

    datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
    print(f"\tSchritt 4a: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 4a: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Log-File

    formatted_timestamp = format_datetime(datetime_stamp, format_string)  # Formatieren des Zeitstempels
    print(f"\tSchritt 4b: Versanddatum abrufen und konvertieren: '{formatted_timestamp}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 4b: Versanddatum abrufen und konvertieren: '{formatted_timestamp}'")  # Debugging-Ausgabe: Log-File

    # 5. Schritt: Bereinigten Betreff ermitteln
    msg_subject = get_subject_msg_file(msg_path_and_filename)  # Betreff auslesen
    print(f"\tSchritt 5a: Betreff ermitteln: '{msg_subject}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 5a: Betreff ermitteln: '{msg_subject}'")  # Debugging-Ausgabe: Log-File

    msg_subject_sanitized = custom_sanitize_text(msg_subject)  # Betreff bereinigen
    print(f"\tSchritt 5b: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 5b: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Log-File

    # 6. Schritt: Neuen Namen der Datei festlegen
    new_msg_filename = f"{formatted_timestamp}_{parsed_sender_email['sender_email']}_{msg_subject_sanitized}.msg"
    msg_pathname = os.path.dirname(msg_path_and_filename)  # Verzeichnisname der MSG-Datei
    print(f"\tSchritt 6a: Neuer Dateiname: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 6a: Neuer Dateiname: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Log-File

    new_msg_path_and_filename = os.path.join(msg_pathname, new_msg_filename)  # Neuer absoluter Dateipfad
    print(f"\tSchritt 6b: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 6b: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Log-File

    # 7. Schritt: Kürzen des Dateinamens, falls nötig
    if len(new_msg_path_and_filename) > max_path_length:
        new_truncated_msg_path_and_filename = truncate_filename_if_needed(new_msg_path_and_filename, max_path_length, "...msg")
        new_truncated_msg_filename = os.path.basename(new_truncated_msg_path_and_filename)
        is_msg_filename_truncated = True
        print(f"\tSchritt 7: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 7: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Log-File
    else:
        new_truncated_msg_path_and_filename = new_msg_path_and_filename
        new_truncated_msg_filename = new_msg_filename
        is_msg_filename_truncated = False
        print(f"\tSchritt 7: Kein Kürzen des Dateinames erforderlich.")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 7: Kein Kürzen des Dateinames erforderlich.")  # Debugging-Ausgabe: Log-File

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