# Dieses Skript benötigt die folgenden Bibliotheken:
# - pandas
# - openpyxl
# Installiere sie mit: pip install -r requirements.txt
import os
import shutil
from modules.msg_handling import get_sender_msg_file, parse_sender_msg_file, create_log_file, log_entry, load_known_senders, get_date_sent_msg_file, convert_to_utc_naive, format_datetime, get_subject_msg_file, custom_sanitize_text, truncate_filename_if_needed

import pandas as pd
from openpyxl import load_workbook

def delete_directory_contents(directory_path):
    """
    Löscht den gesamten Inhalt des angegebenen Verzeichnisses.

    Parameter:
    directory_path (str): Der Pfad des Verzeichnisses, dessen Inhalt gelöscht werden soll.

    Gibt:
    str: Eine Bestätigung, dass der Inhalt erfolgreich gelöscht wurde.

    Wirft:
    OSError: Wenn das Löschen des Inhalts nicht erfolgreich ist.
    """
    if not os.path.isdir(directory_path):
        raise OSError(f"{directory_path} ist kein gültiges Verzeichnis.")

    try:
        for item in os.listdir(directory_path):
            item_path = os.path.join(directory_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            else:
                shutil.rmtree(item_path)
        return "Inhalt erfolgreich gelöscht."
    except Exception as e:
        raise OSError(f"Fehler beim Löschen des Inhalts: {e}")

def copy_directory_contents(source_directory_path, target_directory_path):
    """
    Kopiert den gesamten Inhalt des angegebenen Quellverzeichnisses in das Zielverzeichnis.

    Parameter:
    source_directory_path (str): Der Pfad des Quellverzeichnisses.
    target_directory_path (str): Der Pfad des Zielverzeichnisses.

    Gibt:
    str: Eine Bestätigung, dass der Inhalt erfolgreich kopiert wurde.

    Wirft:
    OSError: Wenn das Kopieren des Inhalts nicht erfolgreich ist.
    """
    if not os.path.isdir(source_directory_path):
        raise OSError(f"{source_directory_path} ist kein gültiges Quellverzeichnis.")

    os.makedirs(target_directory_path, exist_ok=True)

    try:
        for item in os.listdir(source_directory_path):
            s = os.path.join(source_directory_path, item)
            d = os.path.join(target_directory_path, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, False, None)
            else:
                shutil.copy2(s, d)
        return "Inhalt erfolgreich kopiert."
    except Exception as e:
        raise OSError(f"Fehler beim Kopieren des Inhalts: {e}")

if __name__ == '__main__':
    # Verzeichnisse für die Tests definieren
    # SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-short'
    SOURCE_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\data\sample_files\testset-long'
    TARGET_DIRECTORY_TEST_DATA = r'D:\Dev\pycharm\MSGFileRenamer\tests\functional\testdir'

    # Maximal zulässige Pfadlänge für Windows 11
    max_path_length = 260

    LOG_DIRECTORY = r'D:\Dev\pycharm\MSGFileRenamer\logs'
    BASE_LOG_NAME = 'msg_log'

    LIST_OF_KNOWN_SENDERS = r'D:\Dev\pycharm\MSGFileRenamer\config\known_senders.csv'

    # Beispiel für das gewünschte Format
    format_string = "%Y%m%d-%Huhr%M"  # Beispiel: JJJJMMTT-HHMM

    # Zähler für die fortlaufende Nummer
    counter = 1

    # Zeitstempel initialisieren
    datetime_stamp = ""

    # Logfile erstellen
    log_file_path = create_log_file(BASE_LOG_NAME, LOG_DIRECTORY)
    print(f"Logdatei erstellt: {log_file_path}")

    # Löschen des Zielverzeichnisses und Ausgabe des Verzeichnisnamens
    print(f"Lösche Inhalt von: {TARGET_DIRECTORY_TEST_DATA}")
    print(delete_directory_contents(TARGET_DIRECTORY_TEST_DATA))

    # Kopieren des Quellverzeichnisses und Ausgabe der Verzeichnisnamen
    print(f"Kopiere Inhalt von: {SOURCE_DIRECTORY_TEST_DATA} nach: {TARGET_DIRECTORY_TEST_DATA}")
    print(copy_directory_contents(SOURCE_DIRECTORY_TEST_DATA, TARGET_DIRECTORY_TEST_DATA))

    # Im Verzeichnis TARGET_DIRECTORY_TEST_DATA und allen Unterverzeichnissen alle MSG-Files suchen und bearbeiten
    known_senders_df = load_known_senders(LIST_OF_KNOWN_SENDERS)

    # Im Verzeichnis TARGET_DIRECTORY_TEST_DATA und allen Unterverzeichnissen alle MSG-Files suchen und bearbeiten
    for root, dirs, files in os.walk(TARGET_DIRECTORY_TEST_DATA):
        for file in files:
            file_path = os.path.join(root, file)
            print(f"Verzeichnis: {root}")
            if file.endswith(".msg"):
                file_path = os.path.join(root, file)
                sender = get_sender_msg_file(file_path)
                print(f"\tDatei: {file}")
                print(f"\tExtrahierter Sender: {sender}")

                # Aufruf der parse_sender_msg_file-Funktion
                parsed_sender_email = parse_sender_msg_file(sender)
                if parsed_sender_email["sender_email"]:
                    print(f"\tParsed Sender-Email: {parsed_sender_email['sender_email']}")  # Ausgabe des analysierten Senders

                # Überprüfen, ob eine Senderemail vorhanden ist
                if not parsed_sender_email["sender_email"]:  # Nur ausführen, wenn keine Senderemail gefunden wurde
                    # Überprüfen, ob der Sendername in der Tabelle der bekannten Sender vorhanden ist
                    known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]
                    if not known_sender_row.empty:
                        # Wenn der Sendername bekannt ist, die Email-Adresse hinzufügen
                        parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
                        parsed_sender_email["contains_sender_email"] = True
                        print(f"\tSender-Email aus Tabelle: {parsed_sender_email['sender_email']}")  # Ausgabe des analysierten Senders
                    else:
                        parsed_sender_email["contains_sender_email"] = False
                else:
                    parsed_sender_email["contains_sender_email"] = True  # Wenn eine Email gefunden wurde

                datetime_stamp = get_date_sent_msg_file(os.path.join(root, file))
                datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
                print(f"\tExtrahierter Versandzeitpunkt: {datetime_stamp}")  # Ausgabe des analysierten Datums

                # Im Hauptteil des Codes vor dem Logeintrag
                formatted_timestamp = format_datetime(datetime_stamp, format_string)  # Formatieren des Zeitstempels
                print(f"\tFormatierter Versandzeitpunkt: {formatted_timestamp}")

                # Betreff auslesen
                subject = get_subject_msg_file(file_path)
                print(f"\tExtrahierter Betreff: {subject}")

                # Betreff bereinigen
                sanitized_subject = custom_sanitize_text(subject)
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
                    "Betreff": subject,
                    "Bereinigter Betreff": sanitized_subject,
                    "Neuer Filename": new_filename,
                    "Neuer gekürzter Fielname": new_truncated_path,
                    "Neuer Dateipfad": new_file_path
                }

                # Eintrag ins Logfile hinzufügen
                log_entry(log_file_path, entry)

                # Zähler erhöhen
                counter += 1

