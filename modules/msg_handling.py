"""
msg_handling.py

Dieses Modul enthält Funktionen zum Abrufen von Metadaten aus MSG-Dateien.
Es bietet Routinen, um verschiedene Informationen wie Absender, Empfänger,
Betreff und andere relevante Daten zu extrahieren.

Funktionen:
- get_sender_msg_file(file_path): Gibt die E-Mail-Adresse des Absenders aus einer MSG-Datei zurück.
- parse_sender_msg_file(sender: str): Analysiert den Sender-String und extrahiert Name und E-Mail.
- get_subject_msg_file(file_path): Gibt den Betreff der Nachricht aus einer MSG-Datei zurück.
- get_subject_msg_file2(msg_file): Ruft den Betreff aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.
- load_known_senders(file_path): Lädt bekannte Sender aus einer CSV-Datei.
- get_date_sent_msg_file(file_path): Gibt das Datum der gesendeten Nachricht aus einer MSG-Datei zurück.
- get_date_sent_msg_file2(msg_file): Ruft das Datum der gesendeten Nachricht aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.
- create_log_file(base_name, directory, table_header): Erstellt ein Logfile im Excel-Format mit einem Zeitstempel im Namen.
- log_entry(log_file_path, entry): Fügt einen neuen Eintrag in das Logfile hinzu.
- convert_to_utc_naive(datetime_stamp): Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.
- format_datetime(datetime_stamp, format_string): Formatiert einen Zeitstempel in das angegebene Format.
- custom_sanitize_text(encoded_textstring): Bereinigt einen Textstring von unerwünschten Zeichen.
- truncate_filename_if_needed(file_path, max_length, truncation_marker): Kürzt den Dateinamen, wenn nötig.
- reduce_thread_in_msg_message(email_text, max_older_emails): Reduziert die Anzahl der angehängten älteren E-Mails auf max_older_emails.

Verwendung:
Importiere die Funktionen aus diesem Modul in deinem Hauptprogramm oder anderen Modulen,
um auf die Metadaten von MSG-Dateien zuzugreifen.

Beispiel:
    from modules.msg_handling import get_sender_msg_file
    sender_email = get_sender_msg_file('example.msg')
"""
from distutils.dep_util import newer

import extract_msg
import re
import os
import pandas as pd
from datetime import datetime
import logging
from enum import Enum
import win32com.client
import olefile
from typing import List
from fpdf import FPDF
import copy

from cryptography.x509.ocsp import load_der_ocsp_response, load_der_ocsp_request
from pandas.core.construction import ensure_wrapped_if_datetimelike

logger = logging.getLogger(__name__)

class MsgAccessStatus(Enum):
    SUCCESS = "Success"
    DATA_NOT_FOUND = "Data not found"
    FILE_NOT_FOUND = "File not found"
    PERMISSION_ERROR = "No permission to access"
    FILE_LOCKED = "File is locked by another process"
    ATTRIBUTE_ERROR = "Attribute error"
    UNICODE_DECODE_ERROR = "Unicode decode error"
    UNICODE_ENCODE_ERROR = "Unicode encode error"
    TYPE_ERROR = "Type error"
    VALUE_ERROR = "Value error"
    OTHER_ERROR = "Other error"
    UNKNOWN = "Unknown"
    NO_MESSAGE_FOUND = "No message found"
    NO_SENDER_FOUND = "No sender found"
    # Neue Statuscodes für fehlende Daten
    SUBJECT_MISSING = "Subject missing"
    SENDER_MISSING = "Sender missing"
    DATE_MISSING = "Date missing"
    BODY_MISSING = "Body missing"
    ATTACHMENTS_MISSING = "Attachments missing"

def get_msg_object(msg_file: str) -> dict:
    """
    Öffnet eine MSG-Datei und gibt ein Dictionary mit dem MSG-Objekt und dem Status zurück.

    Diese Funktion versucht, ein MSG-Objekt aus einer angegebenen MSG-Datei zu erstellen.
    Bei Erfolg wird das MSG-Objekt zusammen mit dem Status zurückgegeben.
    Bei Fehlern wird der entsprechende Status in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, die geöffnet werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den Schlüsseln:
        - "msg_object": Das MSG-Objekt oder None, wenn nicht erfolgreich.
        - "get_msg_object_result": Ein Enum-Wert des Typs MsgAccessStatus, der den Status des Zugriffs beschreibt.

    Beispiel:
        result = get_msg_object('example.msg')
        if result['get_msg_object_result'] == MsgAccessStatus.SUCCESS:
            # Arbeiten mit result['msg_object']
    """
    # Rückgabewert initialisieren, aber eventuell auch überflüssig
    msg_object = None  # Standardwert
    get_msg_object_result = MsgAccessStatus.UNKNOWN

    # Versuche, die MSG-Datei zu öffnen und MSG-Object erzeugen auszulesen
    try:
        logger.debug(f"Versuche mit extract_msg ein msg-Objekt zu erzeugen: {msg_file}")  # Debugging-Ausgabe: Log-File

        # MSG-Objekt erzeugen und automatisches Schließen der Datei durch Verwenden eines with-Blocks
        with extract_msg.Message(msg_file) as msg:

            # War die Erzeugung des MSG-Objekts aus der MSG-Datei erfolgreich?
            if msg is not None:
                msg_object = msg  # Kopie des MSG-Objekts erzeugen
                get_msg_object_result = MsgAccessStatus.SUCCESS
                logger.debug(
                    f"MSG-Objekt aus MSG-Datei erfolgreich erzeugt: '{msg_file}'")  # Debugging-Ausgabe: Log-File

    except FileNotFoundError:
        get_msg_object_result = MsgAccessStatus.FILE_NOT_FOUND
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da Datei nicht gefunden wurde: {msg_file}")  # Debugging-Ausgabe: Log-File

    except AttributeError:
        get_msg_object_result = MsgAccessStatus.ATTRIBUTE_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da Attribut nicht gefunden wurde: {msg_file}")  # Debugging-Ausgabe: Log-File

    except UnicodeEncodeError:
        get_msg_object_result = MsgAccessStatus.UNICODE_ENCODE_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da UnicodeEncodeError auftrat: {msg_file}")  # Debugging-Ausgabe: Log-File

    except UnicodeDecodeError:
        get_msg_object_result = MsgAccessStatus.UNICODE_DECODE_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da UnicodeDecodeError auftrat: {msg_file}")  # Debugging-Ausgabe: Log-File

    except TypeError:
        get_msg_object_result = MsgAccessStatus.TYPE_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da TypeError auftrat: {msg_file}")  # Debugging-Ausgabe: Log-File

    except ValueError:
        get_msg_object_result = MsgAccessStatus.VALUE_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da ValueError auftrat: {msg_file}")  # Debugging-Ausgabe: Log-File

    except PermissionError:
        get_msg_object_result = MsgAccessStatus.PERMISSION_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, da PermissionError auftrat: {msg_file}")  # Debugging-Ausgabe: Log-File

    except Exception as e:
        get_msg_object_result = MsgAccessStatus.OTHER_ERROR
        logger.error(
            f"MSG-Objekt aus MSG-Datei nicht erzeugt, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Log-File

    # Rückgabe des Ergebnisses als Dictionary
    return {
            "msg_object": msg_object,  # msg_object wird hier zurückgegeben
            "get_msg_object_result": get_msg_object_result
    }

def get_msg_object2(msg_file: str) -> dict:
    """
    Öffnet eine MSG-Datei, extrahiert relevante Daten und gibt sie als Dictionary zurück.

    Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
    Bei Erfolg werden die relevanten Informationen wie Betreff, Absender, Datum,
    Nachrichtentext und Anhänge extrahiert. Bei Fehlern wird der entsprechende Status
    in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, die geöffnet werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den extrahierten Daten und einer Liste von Statuscodes:
        - "subject": Der Betreff der Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "sender": Der Absender der Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "date": Das Datum der gesendeten Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "body": Der Inhalt der Nachricht oder "Kein Inhalt verfügbar", wenn nicht vorhanden.
        - "attachments": Eine Liste der Dateinamen der Anhänge oder eine leere Liste, wenn keine vorhanden sind.
        - "status": Eine Liste von Statuscodes, die den Erfolg oder Fehler des Zugriffs beschreiben.
        - "signed": Boolean, ob die Nachricht signiert ist.
        - "encrypted": Boolean, ob die Nachricht verschlüsselt ist.
        - "reply_count": Anzahl der Antworten oder Weiterleitungen.
        - "has_defects": Boolean, ob die Nachricht Hinweise auf Defekte enthält.
    """

    # Initialisierung des Rückgabewerts mit Standardwerten
    msg_data = {
        "subject": "Unbekannt",
        "sender": "Unbekannt",
        "date": "Unbekannt",
        "body": "Kein Inhalt verfügbar",
        "attachments": [],
        "status": [MsgAccessStatus.UNKNOWN],  # Liste mit Statuscodes
        "signed": False,
        "encrypted": False,
        "reply_count": 0,
        "has_defects": False
    }

    try:
        logger.debug(f"Öffne MSG-Datei: {msg_file}")  # Debugging-Ausgabe

        # Sicherstellen, dass die Datei mit `with` geöffnet und automatisch geschlossen wird
        with extract_msg.Message(msg_file) as msg_object:

            # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
            if msg_object is None:
                msg_data["status"].append(MsgAccessStatus.DATA_NOT_FOUND)
                logger.error(f"MSG-Datei konnte nicht verarbeitet werden: {msg_file}")
                return msg_data  # Sofort zurückgeben

            # Jedes Attribut separat absichern
            try:
                if msg_object.subject:
                    msg_data["subject"] = msg_object.subject
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Betreff erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.SUBJECT_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.sender:
                    msg_data["sender"] = msg_object.sender
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Sender erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.SENDER_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.date:
                    msg_data["date"] = msg_object.date
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Datum erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.DATE_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.body:
                    msg_data["body"] = msg_object.body
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Body erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.BODY_MISSING)
            except UnicodeDecodeError:
                msg_data["status"].append(MsgAccessStatus.UNICODE_DECODE_ERROR)
            except UnicodeEncodeError:
                msg_data["status"].append(MsgAccessStatus.UNICODE_ENCODE_ERROR)

            try:
                if msg_object.attachments:
                    msg_data["attachments"] = [att.filename for att in msg_object.attachments]
                else:
                    msg_data["status"].append(MsgAccessStatus.ATTACHMENTS_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            # Zusätzliche Informationen extrahieren
            msg_data["signed"] = hasattr(msg_object, 'signed') and msg_object.signed
            msg_data["encrypted"] = hasattr(msg_object, 'encrypted') and msg_object.encrypted
            msg_data["reply_count"] = getattr(msg_object, 'reply_count', 0)
            msg_data["has_defects"] = hasattr(msg_object, 'has_defects') and msg_object.has_defects

        logger.debug(f"MSG-Daten erfolgreich extrahiert: {msg_data}")  # Debugging-Ausgabe

    except FileNotFoundError:
        msg_data["status"] = [MsgAccessStatus.FILE_NOT_FOUND]
        logger.error(f"MSG-Datei nicht gefunden: {msg_file}")

    except PermissionError:
        msg_data["status"] = [MsgAccessStatus.PERMISSION_ERROR]
        logger.error(f"Keine Berechtigung, um die MSG-Datei zu öffnen: {msg_file}")

    except TypeError:
        msg_data["status"] = [MsgAccessStatus.TYPE_ERROR]
        logger.error(f"Falscher Datentyp in MSG-Datei: {msg_file}")

    except ValueError:
        msg_data["status"] = [MsgAccessStatus.VALUE_ERROR]
        logger.error(f"Ungültiger Wert in MSG-Datei: {msg_file}")

    except Exception as e:
        msg_data["status"] = [MsgAccessStatus.OTHER_ERROR]
        logger.error(f"Allgemeiner Fehler beim Öffnen der MSG-Datei: {str(e)}")

    return msg_data

def get_sender_msg_file(file_path) -> str:
    """
    Ruft den Absender aus einer MSG-Datei ab.
    Der Absender besteht meistens aus einem Namen sowie der zugehörigen Emailadresse.
    Der Absender einer MSG-Datei ist nicht immer vorhanden.
    In diesem Fall wird "Unbekannt" zurückgegeben.

    Parameter:
    file_path (str): Der Pfad zur MSG-Datei.

    Rückgabewert:
    str: Der Absender der MSG-Datei.
    str = "Unbekannt", wenn kein Sender vorhanden ist.
    str = "Datei nicht gefunden", wenn auf die MSG-Datei nicht zugegriffen werden konnte.
    str = "Fehler beim Auslesen des Senders: " und ein zusätzlicher Fehlercode, bei sonstigen Problemen.
    str = ... es werden diverse weitere Fehlerunterschieden

    Beispiel:
    from modules.msg_handling import get_sender_msg_file
    sender_email = get_sender_msg_file('example.msg')
    print(sender_email)
    # Ausgabe: "Max Mustermann <max.mustermann@example.com>"
    """
    msg_sender = "Unbekannt"  # Standardwert

    try:
        print(f"\tVersuche den Sender der MSG-Datei auzulesen: {file_path}")  # Debugging-Ausgabe: Console
        logger.debug(f"Versuche den Sender der MSG-Datei auzulesen: {file_path}")  # Debugging-Ausgabe: Log-File

        # MSG-Objekt erzeugen und automatisches Schließen der Datei durch Verwenden eines with-Blocks
        with extract_msg.Message(file_path) as msg:

            # War die Erzeugung des MSG-Objekts aus der MSG-Datei erfolgreich?
            if msg.sender is not None:
                msg_sender=msg.sender # Absender aus dem MSG-Objekt auslesen
                print(f"\tAus MSG-Datei extrahierter Absender: '{msg_sender}'")
                logger.debug(f"Aus MSG-Datei extrahierter Absender: '{msg_sender}'") # Debugging-Ausgabe: Log-File
            else:
                print(f"\tIn MSG-Datei wurde kein Absender im extrahierten MSG-Objekt gefunden.")  # Debugging-Ausgabe: Console
                logger.debug(f"In MSG-Datei wurde kein Absender im extrahierten MSG-Objekt gefunden.")  # Debugging-Ausgabe: Log-File

    except FileNotFoundError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da Datei nicht gefunden wurde: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da Datei nicht gefunden wurde: {file_path}") # Debugging-Ausgabe: Log-File

    except AttributeError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da Attribut nicht gefunden wurde: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da Attribut nicht gefunden wurde: {file_path}") # Debugging-Ausgabe: Log-File

    except UnicodeEncodeError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da UnicodeEncodeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da UnicodeEncodeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except UnicodeDecodeError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da UnicodeDecodeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da UnicodeDecodeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except TypeError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da TypeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da TypeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except ValueError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da ValueError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da ValueError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except PermissionError:
        msg_sender = "Fehler beim Auslesen des Senders"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, da PermissionError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, da PermissionError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except Exception as e:
        msg_sender=f"Fehler beim Auslesen des Senders: {str(e)}"
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'") # Debugging-Ausgabe: Log-File

    finally:
        msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig

    return msg_sender


def get_sender_msg_file2(msg_file: str) -> dict:
    """
    Ruft den Absender aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.

    Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
    Wenn das MSG-Objekt erfolgreich erstellt wurde, wird der Absender extrahiert.
    Bei Fehlern wird der entsprechende Status in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, aus der der Absender abgerufen werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den Schlüsseln:
        - "msg_sender": Die E-Mail-Adresse des Absenders oder ein leerer String, wenn nicht erfolgreich.
        - "get_msg_access_result": Ein Enum-Wert des Typs MsgAccessStatus, der den Status des Zugriffs beschreibt.

    Beispiel:
        result = get_sender_msg_file2('example.msg')
        if result['get_msg_access_result'] == MsgAccessStatus.SUCCESS:
            print(f"Absender: {result['msg_sender']}")
    """

    # Initialisierung der Rückgabewerte
    msg_sender = ""  # Standardwert für den Absender
    get_msg_sender_result = MsgAccessStatus.UNKNOWN  # Initialer Status
    msg_object = None  # Initialisierung von msg_object

    try:
        # Versuche, das MSG-Objekt zu erhalten
        msg_object = get_msg_object(msg_file)

        # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
        if (msg_object["get_msg_object_result"] == MsgAccessStatus.SUCCESS) & (msg_object["msg_object"].sender is not None) :
            msg_sender = msg_object["msg_object"].sender  # Absender extrahieren
            get_msg_sender_result = msg_object["get_msg_object_result"]  # Status aktualisieren
        else:
            # Fehler beim Auslesen des Absenders
            get_msg_sender_result = MsgAccessStatus.NO_SENDER_FOUND
            print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {msg_object['get_msg_object_result']}")  # Debugging-Ausgabe: Console
            logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {msg_object['get_msg_object_result']}")  # Debugging-Ausgabe: Log-File

    except Exception as e:
        # Allgemeiner Fehlerfall
        get_msg_sender_result = MsgAccessStatus.OTHER_ERROR
        print(f"\tSender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
        logger.error(f"Sender der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Log-File

    # Rückgabe des Ergebnisses als Dictionary
    return {
        "msg_sender": msg_sender,
        "get_msg_sender_result": get_msg_sender_result
    }


def get_message_msg_file(file_path) -> dict:
    """
    Ruft den eigentlichen Nachrichttext aus einer MSG-Datei ab.

    Parameter:
    file_path (str): Der Pfad zur MSG-Datei.

    Rückgabewert:


    Beispiel:

    """
    msg_message = ""  # Standardwert
    get_msg_access_result = MsgAccessStatus.UNKNOWN
    get_message_result = {
        "msg_message": msg_message,
        "get_msg_access_result": get_msg_access_result
        }

    try:
        print(f"\tVersuche den Nachrichtentext aus der MSG-Datei auzulesen: {file_path}")  # Debugging-Ausgabe: Console
        logger.debug(f"Versuche den Nachrichtentext aus der MSG-Datei auzulesen: {file_path}")  # Debugging-Ausgabe: Log-File

        # MSG-Objekt erzeugen und automatisches Schließen der Datei durch Verwenden eines with-Blocks
        with extract_msg.Message(file_path) as msg:

            # War die Erzeugung des MSG-Objekts aus der MSG-Datei erfolgreich?
            if msg.body is not None:
                msg_message = msg.body # Absender aus dem MSG-Objekt auslesen
                get_msg_access_result = MsgAccessStatus.SUCCESS
                print(f"\tAus MSG-Datei den Nachrichtentext erfolgreich extrahiert: '{msg_message[:30]}'")  # Debugging-Ausgabe: Console
                logger.debug(f"Aus MSG-Datei den Nachrichtentext erfolgreich extrahiert: '{msg_message[:30]}'")  # Debugging-Ausgabe: Log-File
            else:
                get_msg_access_result = MsgAccessStatus.NO_MESSAGE_FOUND
                print(f"\tIn MSG-Datei wurde kein Nachrichtentext im extrahierten MSG-Objekt gefunden.")  # Debugging-Ausgabe: Console
                logger.debug(f"In MSG-Datei wurde kein Nachrichtentext im extrahierten MSG-Objekt gefunden.")  # Debugging-Ausgabe: Log-File

    except FileNotFoundError:
        get_msg_access_result = MsgAccessStatus.FILE_NOT_FOUND
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da Datei nicht gefunden wurde: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da Datei nicht gefunden wurde: {file_path}") # Debugging-Ausgabe: Log-File

    except AttributeError:
        get_msg_access_result = MsgAccessStatus.ATTRIBUTE_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da Attribut nicht gefunden wurde: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da Attribut nicht gefunden wurde: {file_path}") # Debugging-Ausgabe: Log-File

    except UnicodeEncodeError:
        get_msg_access_result = MsgAccessStatus.UNICODE_ENCODE_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da UnicodeEncodeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da UnicodeEncodeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except UnicodeDecodeError:
        get_msg_access_result = MsgAccessStatus.UNICODE_DECODE_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da UnicodeDecodeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da UnicodeDecodeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except TypeError:
        get_msg_access_result = MsgAccessStatus.TYPE_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da TypeError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da TypeError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except ValueError:
        get_msg_access_result = MsgAccessStatus.TYPE_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da ValueError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da ValueError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except PermissionError:
        get_msg_access_result = MsgAccessStatus.PERMISSION_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da PermissionError auftrat: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, da PermissionError auftrat: {file_path}") # Debugging-Ausgabe: Log-File

    except Exception as e:
        get_msg_access_result = MsgAccessStatus.OTHER_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'") # Debugging-Ausgabe: Log-File

    finally:
        msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig

    get_message_result = {
        "msg_message": msg_message,
        "get_msg_access_result": get_msg_access_result
    }

    return get_message_result


def get_message_msg_file2(msg_file: str) -> dict:
    """
    Ruft den Nachrichtentext aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.

    Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
    Wenn das MSG-Objekt erfolgreich erstellt wurde, wird der Nachrichtentext extrahiert.
    Bei Fehlern wird der entsprechende Status in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, aus der der Nachrichtentext abgerufen werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den Schlüsseln:
        - "msg_message": Der Nachrichtentext oder ein leerer String, wenn nicht erfolgreich.
        - "get_msg_access_result": Ein Enum-Wert des Typs MsgAccessStatus, der den Status des Zugriffs beschreibt.
    """
    try:
        # Versuche, das MSG-Objekt zu erhalten
        msg_object = get_msg_object(msg_file)

        # Zuweisung des Rückgabewerts für Debuggingzwecke
        get_msg_access_result = msg_object["get_msg_object_result"]

        # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
        if (get_msg_access_result == MsgAccessStatus.SUCCESS) & (msg_object["msg_object"].body is not None):
            msg_message = msg_object["msg_object"].body  # Nachricht extrahieren
            print(f"\tNachrichtentext erfolgreich extrahiert.")  # Debugging-Ausgabe: Console
            logger.debug(f"Nachrichtentext erfolgreich extrahiert.")  # Debugging-Ausgabe: Log-File
        else:
            # Fehler beim Auslesen des Nachrichtentextes
            msg_message = ""
            print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Console
            logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Log-File

    except Exception as e:
        # Allgemeiner Fehlerfall
        msg_message = ""
        get_msg_access_result = MsgAccessStatus.OTHER_ERROR
        print(f"\tNachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
        logger.error(f"Nachrichtentext der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Log-File

    # Rückgabe des Ergebnisses als Dictionary
    return {
        "msg_message": msg_message,
        "get_msg_access_result": get_msg_access_result
    }


def get_subject_msg_file(file_path):
    """
    Gibt den Betreff der Nachricht aus einer MSG-Datei zurück.

    Parameter:
    file_path (str): Der Pfad zur MSG-Datei.

    Rückgabewert:
    str: Der Betreff der Nachricht oder "Unbekannt", wenn kein Betreff vorhanden ist.
    """
    try:
        msg = extract_msg.Message(file_path)
        #print(f"Betreff erfolgreich extrahiert: {msg.subject}")  # Debugging-Ausgabe: Console
        logger.debug(f"Betreff erfolgreich extrahiert: {msg.subject}")  # Debugging-Ausgabe: Log-File
        #msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig
        return msg.subject if msg.subject else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Betreff vorhanden ist
    except FileNotFoundError:
        #print(f"Datei nicht gefunden: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Datei nicht gefunden: {file_path}")  # Debugging-Ausgabe: Log-File
        return "Datei nicht gefunden"
    except Exception as e:
        #print(f"Fehler beim Auslesen des Betreffs: {str(e)}")  # Debugging-Ausgabe: Console
        logger.error(f"Fehler beim Auslesen des Betreffs: {str(e)}")  # Debugging-Ausgabe: Log-File
        msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig
        return f"Fehler beim Auslesen des Betreffs: {str(e)}"


def get_subject_msg_file2(msg_file: str) -> dict:
        """
        Ruft den Betreff aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.

        Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
        Wenn das MSG-Objekt erfolgreich erstellt wurde, wird der Betreff extrahiert.
        Bei Fehlern wird der entsprechende Status in der Rückgabe angezeigt.

        Parameter:
        msg_file (str): Der Pfad zur MSG-Datei, aus der der Betreff abgerufen werden soll.

        Rückgabewert:
        dict: Ein Dictionary mit den Schlüsseln:
            - "msg_subject": Der Betreff der Nachricht oder ein leerer String, wenn nicht erfolgreich.
            - "get_msg_access_result": Ein Enum-Wert des Typs MsgAccessStatus, der den Status des Zugriffs beschreibt.

        Beispiel:
            result = get_subject_msg_file2('example.msg')
            if result['get_msg_access_result'] == MsgAccessStatus.SUCCESS:
                print(f"Betreff: {result['msg_subject']}")
        """

        #msg_subject = ""  # Standardwert für den Betreff
        #get_msg_access_result = MsgAccessStatus.UNKNOWN  # Initialer Status

        try:
            # Versuche, das MSG-Objekt zu erhalten
            msg_object = get_msg_object(msg_file)

            # Zuweisung des Rückgabewerts für Debuggingzwecke
            get_msg_access_result = msg_object["get_msg_object_result"]

            print(f"\n\tAccess Result von 'get_subject_msg_file2': {msg_object['get_msg_object_result'].name}")
            print(f"\tExtrahierter Betreff von 'get_subject_msg_file2': {msg_object['msg_object'].subject}")

            # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
            if (msg_object["get_msg_object_result"] == MsgAccessStatus.SUCCESS) & (msg_object["msg_object"].subject is not None):
                msg_subject = msg_object["msg_object"].subject  # Betreff extrahieren
                print(f"\tBetreff der MSG-Datei erfolgreich extrahiert: '{msg_subject}'")  # Debugging-Ausgabe: Console
                logger.debug(
                    f"Betreff der MSG-Datei erfolgreich extrahiert: '{msg_subject}'")  # Debugging-Ausgabe: Log-File
            else:
                # Fehler beim Auslesen des Betreffs
                print(
                    f"\tBetreff der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Console
                logger.error(
                    f"Betreff der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Log-File

        except Exception as e:
            # Allgemeiner Fehlerfall
            msg_subject = ""
            get_msg_access_result = MsgAccessStatus.OTHER_ERROR
            print(
                f"\tBetreff der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
            logger.error(
                f"Betreff der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Log-File

        # Rückgabe des Ergebnisses als Dictionary
        return {
            "msg_subject": msg_subject,
            "get_msg_access_result": get_msg_access_result
        }


def get_date_sent_msg_file(file_path):
    """
    Gibt das Datum der gesendeten Nachricht aus einer MSG-Datei zurück.

    Parameter:
    file_path (str): Der Pfad zur MSG-Datei.

    Rückgabewert:
    str: Das Datum der gesendeten Nachricht oder "Unbekannt", wenn kein Datum vorhanden ist.
    """
    try:
        with extract_msg.Message(file_path) as msg:  # Hier wird die Datei geöffnet
            #print(f"Datum erfolgreich extrahiert: {msg.date}")  # Debugging-Ausgabe: Console
            logger.debug(f"Datum erfolgreich extrahiert: {msg.date}")  # Debugging-Ausgabe: Log-File
            msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig
            return msg.date if msg.date else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Datum vorhanden ist
    except FileNotFoundError:
        #print(f"Datei nicht gefunden: {file_path}")  # Debugging-Ausgabe: Console
        logger.error(f"Datei nicht gefunden: {file_path}")  # Debugging-Ausgabe: Log-File
        return "Datei nicht gefunden"
    except Exception as e:
        #print(f"Fehler beim Auslesen des Datums: {str(e)}")  # Debugging-Ausgabe: Console
        logger.error(f"Fehler beim Auslesen des Datums: {str(e)}")  # Debugging-Ausgabe: Log-File
        msg.close()  # MSG-Datei zur Sicherheit schließen, wenn auch durch With-Block nicht zwingend nötig
        return f"Fehler beim Auslesen des Datums: {str(e)}"


def get_date_sent_msg_file2(msg_file) -> dict:
    """
    Ruft das Datum der gesendeten Nachricht aus einer MSG-Datei ab und gibt das Ergebnis als Dictionary zurück.

    Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
    Wenn das MSG-Objekt erfolgreich erstellt wurde, wird das Datum der gesendeten Nachricht extrahiert.
    Bei Fehlern wird der entsprechende Status in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, aus der das Datum abgerufen werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den Schlüsseln:
        - "msg_date": Das Datum der gesendeten Nachricht oder None, wenn nicht erfolgreich.
        - "get_msg_access_result": Ein Enum-Wert des Typs MsgAccessStatus, der den Status des Zugriffs beschreibt.

    Beispiel:
        result = get_date_sent_msg_file2('example.msg')
        if result['get_msg_access_result'] == MsgAccessStatus.SUCCESS:
            print(f"Gesendetes Datum: {result['msg_date']}")
    """
    msg_date = None  # Standardwert für das Datum
    get_msg_access_result = MsgAccessStatus.UNKNOWN  # Initialer Status

    try:
        # Versuche, das MSG-Objekt zu erhalten
        msg_object = get_msg_object(msg_file)

        # Zuweisung des Rückgabewerts für Debuggingzwecke
        get_msg_access_result = msg_object["get_msg_object_result"]

        # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
        if get_msg_access_result == MsgAccessStatus.SUCCESS and msg_object["msg_object"].date is not None:
            msg_date = msg_object["msg_object"].date  # Datum extrahieren
            print(f"\tDatum erfolgreich extrahiert: {msg_date}")  # Debugging-Ausgabe: Console
            logger.debug(f"Datum erfolgreich extrahiert: {msg_date}")  # Debugging-Ausgabe: Log-File
        else:
            # Fehler beim Auslesen des Datums
            print(f"\tDatum der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Console
            logger.error(f"Datum der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: {get_msg_access_result}")  # Debugging-Ausgabe: Log-File

    except Exception as e:
        # Allgemeiner Fehlerfall
        get_msg_access_result = MsgAccessStatus.OTHER_ERROR
        print(f"\tDatum der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Console
        logger.error(f"Datum der MSG-Datei konnte nicht ausgelesen werden, wegen folgenden Fehlers: '{str(e)}'")  # Debugging-Ausgabe: Log-File

    # Rückgabe des Ergebnisses als Dictionary
    return {
        "msg_date": msg_date,
        "get_msg_access_result": get_msg_access_result
    }


def create_log_file(base_name, directory, table_header):
    """
    Erstellt ein Logfile im Excel-Format mit einem Zeitstempel im Namen.

    Parameter:
    base_name (str): Der Basisname des Logfiles.
    directory (str): Das Verzeichnis, in dem das Logfile gespeichert werden soll.

    Rückgabewert:
    str: Der Pfad zur erstellten Logdatei.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{base_name}_{timestamp}.xlsx"
    log_file_path = os.path.join(directory, log_file_name)

    # Leeres DataFrame mit den gewünschten Spalten erstellen
    #df = pd.DataFrame(columns=["Fortlaufende Nummer", "Verzeichnisname", "Filename", "Sendername", "Senderemail",
    #                           "Contains Senderemail", "Timestamp", "Formatierter Timestamp", "Betreff",
    #                           "Bereinigter Betreff", "Neuer Filename", "Neuer gekürzter Fielname", "Neuer Dateipfad"])
    df = pd.DataFrame(columns=table_header)

    try:
        df.to_excel(log_file_path, index=False)
        logger.debug(f"Logging Excel-Datei erfolgreich erstellt: {log_file_path}")  # Debugging-Ausgabe: Log-File
        return log_file_path
    except Exception as e:
        logger.error(f"Fehler beim Erstellen der Logdatei: {e}")  # Debugging-Ausgabe: Log-File
        raise OSError(f"Fehler beim Erstellen der Logdatei: {e}")


def log_entry(log_file_path, entry):
    """
    Fügt einen neuen Eintrag in das Logfile hinzu.

    Parameter:
    log_file_path (str): Der Pfad zur Logdatei.
    entry (dict): Ein Dictionary mit den Werten für die Logzeile.

    Rückgabewert:
    None
    """
    df = pd.read_excel(log_file_path)

    # Erstellen eines DataFrames aus dem Eintrag
    new_entry_df = pd.DataFrame([entry])

    # Überprüfen, ob new_entry_df leer ist oder nur NA-Werte enthält
    if not new_entry_df.empty and not new_entry_df.isnull().all(axis=1).any():
        # Logdatei laden oder erstellen
        if os.path.exists(log_file_path):
            with pd.ExcelFile(log_file_path) as xls:  # Verwende einen with-Block
                df = pd.read_excel(xls)
                # Nur nicht-leere DataFrames zusammenführen
                if not df.empty:
                    df = pd.concat([df, new_entry_df], ignore_index=True)  # Eintrag hinzufügen
                else:
                    df = new_entry_df  # Neue Logdatei erstellen, wenn df leer ist
        else:
            df = new_entry_df  # Neue Logdatei erstellen

        # Speichern des aktualisierten DataFrames in die Logdatei
        df.to_excel(log_file_path, index=False)


def convert_to_utc_naive(datetime_stamp):
    """
    Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der konvertiert werden soll.

    Rückgabewert:
    datetime: Ein UTC-naives Datetime-Objekt.
    """
    try:
        if datetime_stamp.tzinfo is not None:
            new_datetime_stamp = datetime_stamp.replace(tzinfo=None)
            logger.debug(f"Konvertierter Zeitstempel in ein UTC-naives Datetime-Objekt: {new_datetime_stamp}")  # Debugging-Ausgabe: Log-File
            return datetime_stamp.replace(tzinfo=None)  # Entfernen der Zeitzone
            logger.debug(f"Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt: {msg.date}")  # Debugging-Ausgabe: Log-File
        return datetime_stamp
    except Exception as e:
        print(f"Fehler beim Konvertieren des Zeitstempels: {str(e)}") # Debugging-Ausgabe: Console
        logger.error(f"Fehler beim Konvertieren des Zeitstempels: {str(e)}") # Debugging-Ausgabe: Log-File
        return datetime_stamp


def format_datetime(datetime_stamp, format_string):
    """
    Formatiert einen Zeitstempel in das angegebene Format.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der formatiert werden soll.
    format_string (str): Das gewünschte Format für den Zeitstempel.

    Rückgabewert:
    str: Der formatierte Zeitstempel als String.

    Beispiel:
        formatted_time = format_datetime(datetime.now(), "%Y-%m-%d %H:%M:%S")
    """
    if datetime_stamp is None:
        logger.error(f"Der Zeitstempel darf nicht None sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("Der Zeitstempel darf nicht None sein.")

    if not isinstance(datetime_stamp, datetime):
        logger.error(f"Ungültiger Zeitstempel: erwartet datetime, erhalten {type(datetime_stamp).__name__}.")  # Debugging-Ausgabe: Log-File
        raise ValueError(f"Ungültiger Zeitstempel: erwartet datetime, erhalten {type(datetime_stamp).__name__}.")

    return datetime_stamp.strftime(format_string)


def custom_sanitize_text(encoded_textstring):
    """
    Bereinigt einen Textstring, indem unerwünschte Zeichen ersetzt und Formatierungen angepasst werden.

    Diese Funktion führt mehrere Schritte zur Bereinigung des Eingabetextes durch:
    1. Ersetzt mehrere aufeinanderfolgende Leerzeichen durch ein einzelnes Leerzeichen.
    2. Entfernt Leerzeichen am Ende des Textes.
    3. Ersetzt spezifische unerwünschte Zeichenfolgen durch definierte Alternativen.
    4. Ersetzt unerwünschte Zeichen durch sichere Alternativen, um sicherzustellen, dass der Text als Dateiname verwendet werden kann.

    Parameter:
    encoded_textstring (str): Der ursprüngliche Textstring, der bereinigt werden soll.

    Rückgabewert:
    str: Der bereinigte Textstring, der als gültiger Dateiname verwendet werden kann,
         wobei unerwünschte Zeichen entfernt oder ersetzt wurden.

    Beispiel:
        sanitized_string = custom_sanitize_text("Beispiel: ungültige Zeichen / \\ * ? < > |")
    """
    # Ersetze mehrere aufeinanderfolgende Leerzeichen durch ein einzelnes Leerzeichen
    encoded_textstring = re.sub(r'\s+', ' ', encoded_textstring)
    # Entferne Leerzeichen am Ende
    encoded_textstring = encoded_textstring.rstrip()

    # Ersetze spezielle Zeichenfolgen
    encoded_textstring = encoded_textstring.replace("_-_", "-")
    encoded_textstring = encoded_textstring.replace(" - ", "-")
    encoded_textstring = encoded_textstring.replace("._", "_")
    encoded_textstring = encoded_textstring.replace("_.", "_")
    encoded_textstring = encoded_textstring.replace(" .", "_")
    encoded_textstring = encoded_textstring.replace(". ", "_")
    encoded_textstring = encoded_textstring.replace(" / ", "_")
    encoded_textstring = encoded_textstring.replace(" & ", "_")
    encoded_textstring = encoded_textstring.replace("; ", "_")
    encoded_textstring = encoded_textstring.replace("/ ", "_")
    encoded_textstring = encoded_textstring.replace(" | ", "_")

    replacements = {
        " ": "_",
        "#": "_",
        "%": "_",
        "&": "_",
        "*": "-",
        "{": "-",
        "}": "-",
        "\\": "-",
        ":": "",
        "<": "-",
        ">": "-",
        "?": "-",
        "/": "_",
        "|": "_",
        "\"": "",
        "ä": "ae",
        "Ä": "Ae",
        "ö": "oe",
        "Ö": "Oe",
        "ü": "ue",
        "Ü": "Ue",
        "ß": "ss",
        "é": "e",
        ",": "",
        "!": "",
        "'": "_",
        ";": "_",
        "“": "",
        "„": ""
    }

    for old_char, new_char in replacements.items():
        encoded_textstring = encoded_textstring.replace(old_char, new_char)

    return encoded_textstring


def truncate_filename_if_needed(file_path, max_length, truncation_marker):
    """
    Kürzt den Dateinamen, wenn der gesamte Pfad die maximal zulässige Länge überschreitet.

    Diese Funktion überprüft die Länge des angegebenen Dateipfads und vergleicht sie mit der maximalen
    zulässigen Länge. Wenn der Pfad diese Länge überschreitet, wird der Dateiname so gekürzt, dass er
    zusammen mit dem Verzeichnispfad die maximale Länge nicht überschreitet. Am Ende des gekürzten
    Dateinamens wird ein Truncation Marker hinzugefügt, um anzuzeigen, dass der Name gekürzt wurde.

    Parameter:
    file_path (str): Der vollständige Dateipfad, der überprüft und möglicherweise gekürzt werden soll.
    max_length (int): Die maximal zulässige Länge des gesamten Dateipfads.
    truncation_marker (str): Die Zeichenkette, die verwendet wird, um das Kürzen anzuzeigen (z.B. "<>").

    Rückgabewert:
    str: Der möglicherweise gekürzte Dateipfad, der die maximal zulässige Länge nicht überschreitet.

    Beispiel:
        truncated_path = truncate_filename_if_needed("D:/Dev/pycharm/MSGFileRenamer/modules/very_long_filename_that_exceeds_the_limit.txt", 50, "...")
    """
    if file_path is None:
        logger.error(f"file_path darf nicht None sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("file_path darf nicht None sein.")

    if not isinstance(max_length, int) or max_length <= 0:
        logger.error(f"max_length muss eine positive Ganzzahl sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("max_length muss eine positive Ganzzahl sein.")

    if not truncation_marker:
        logger.error(f"truncation_marker darf nicht leer sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("truncation_marker darf nicht leer sein.")

    if len(file_path) > max_length:
        # Berechne die maximale Länge für den Dateinamen
        path_length = len(os.path.dirname(file_path))
        max_filename_length = max_length - path_length - len(truncation_marker) - 1  # -1 für den Schrägstrich

        # Extrahiere den Dateinamen
        filename = os.path.basename(file_path)

        if len(filename) > max_filename_length:
            # Kürze den Dateinamen und füge den Truncation Marker hinzu
            truncated_filename = filename[:max_filename_length] + truncation_marker
            return os.path.join(os.path.dirname(file_path), truncated_filename)

    return file_path

def parse_sender_msg_file(msg_absender_str: str) -> dict:
    """
    Analysiert den Sender-String eines MSG-Files und extrahiert den Namen und die E-Mail-Adresse.

    Parameter:
    sender (str): Der Sender-String.

    Rückgabewert:
    dict: Ein Dictionary mit 'sender_name', 'sender_email' und 'contains_sender_email'.
    """
    sender_name = ""
    sender_email = ""
    contains_sender_email = False

    # Regulärer Ausdruck für die E-Mail-Adresse
    email_pattern = r'<(.*?)>'
    email_match = re.search(email_pattern, msg_absender_str)

    if email_match:
        sender_email = email_match.group(1)
        contains_sender_email = True
        logger.debug(f"Im Absender der MSG-Datei ist folgende Email enthalten: {sender_email}")  # Debugging-Ausgabe: Log-File
    else:
        sender_email = ""
        contains_sender_email = False
        logger.debug(f"Im Absender der MSG-Datei ist keine Email enthalten: {msg_absender_str}")  # Debugging-Ausgabe: Log-File

    # Entferne die E-Mail-Adresse aus dem Sender-String
    sender_name = re.sub(email_pattern, '', msg_absender_str).strip()
    # Entferne Anführungszeichen aus dem Sender-String
    sender_name = sender_name.replace("\"", '')

    return {
        "sender_name": sender_name,
        "sender_email": sender_email,
        "contains_sender_email": contains_sender_email
    }


def load_known_senders(file_path):
    """
    Lädt bekannte Sender aus einer CSV-Datei.

    Parameter:
    file_path (str): Der Pfad zur CSV-Datei.

    Rückgabewert:
    DataFrame: Ein DataFrame mit den bekannten Sendern.
    """
    return pd.read_csv(file_path)


def reduce_thread_in_msg_message(email_text, max_older_emails=2) -> dict:
    """
    Reduziert die Anzahl der angehängten älteren E-Mails auf max_older_emails.
    Ältere E-Mails werden anhand der typischen Kopfzeilen (Von, Gesendet, An, Cc, Betreff) erkannt.

    :param email_text: Der vollständige Text der E-Mail
    :param max_older_emails: Die maximale Anzahl an beizubehaltenden alten E-Mails
    :return: Ein Dictionary mit dem bereinigten E-Mail-Text und der Anzahl der gelöschten alten E-Mails
    """
    new_email_text = ""

    # Regulärer Ausdruck für eine typische E-Mail-Kopfzeile
    email_header_pattern = re.compile(
        r"(Von: .+?Betreff: .+?)(\n\n|\r\n\r\n)", re.DOTALL | re.IGNORECASE
    )

    # Alle gefundenen älteren E-Mails identifizieren
    older_emails_text = email_header_pattern.split(email_text)

    # Anzahl der gefundenen älteren E-Mails
    total_older_emails = (len(older_emails_text) // 3) - 1  # Abzüglich der ursprünglichen Nachricht

    if total_older_emails <= max_older_emails:
        return {"new_email_text": email_text, "deleted_count": 0}  # Keine Kürzung nötig

    # Behalten der neuesten E-Mail + der maximal erlaubten Anzahl alter E-Mails
    new_email_text = older_emails_text[0]  # Original-Nachricht ohne Anhang
    for i in range(1, min((max_older_emails * 3) + 1, len(older_emails_text)), 3):
        new_email_text += older_emails_text[i] + older_emails_text[i + 1]

    # Berechnung der Anzahl der gelöschten E-Mails
    deleted_count = total_older_emails - max_older_emails

    # Hinzufügen eines Hinweises für entfernte ältere E-Mails
    new_email_text += "\n\n--- deleted_count ältere E-Mails wurden entfernt. Vollständige E-Mail-Kette im Projektpostfach einsehbar. ---\n"

    return {"new_email_text": new_email_text, "deleted_count": deleted_count}

def print_msg_to_pdf(msg_file, pdf_file):
    print(f"Versuche, die MSG-Datei '{msg_file}' zu öffnen...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    msg = outlook.Session.OpenSharedItem(msg_file)

    print("Öffne Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Word im Hintergrund ausführen

    print(f"Öffne das Dokument '{msg_file}' in Word...")
    doc = word.Documents.Open(msg_file)
    print(f"Exportiere '{msg_file}' als PDF...")
    doc.ExportAsFixedFormat(pdf_file, 17)  # 17 = PDF
    doc.Close()

    word.Quit()
    print(f"Gespeichert als {pdf_file}")


def check_email_security(msg_file: str):
    """
    Prüft, ob eine MSG-Datei signiert oder verschlüsselt wurde.
    """
    msg = extract_msg.Message(msg_file)

    # Standardwerte für Sicherheitseinstellungen
    is_signed = False
    is_encrypted = False

    try:
        # Öffne die MSG-Datei als OLE-Container
        ole = olefile.OleFileIO(msg_file)

        # PR_SECURITY Property (0x6E010003) prüfen
        if ole.exists('__properties_version1.0'):
            prop_stream = ole.openstream('__properties_version1.0')
            data = prop_stream.read()

            # Die PR_SECURITY Property befindet sich an Offset 0x6E01 (je nach MSG-Struktur variabel)
            security_value = int.from_bytes(data[0x6E01:0x6E05], "little", signed=False)

            # Prüfen, ob Verschlüsselung oder Signatur aktiviert ist
            if security_value & 0x0001:
                is_signed = True
            if security_value & 0x0002:
                is_encrypted = True

        # Alternativer Weg: PR_MESSAGE_CLASS prüfen
        if msg.messageClass in ["IPM.Note.SMIME", "IPM.Note.SMIME.MultipartSigned"]:
            is_signed = True
        if msg.messageClass in ["IPM.Note.SMIME", "IPM.Note.SMIME.MultipartEncrypted"]:
            is_encrypted = True

    except Exception as e:
        print(f"Fehler beim Prüfen der Sicherheitseinstellungen: {e}")

    finally:
        msg.close()
        ole.close()

    return {
        "is_signed": is_signed,
        "is_encrypted": is_encrypted
    }
