"""
msg_handling.py

Dieses Modul enthält Funktionen zum Abrufen von Metadaten aus MSG-Dateien.
Es bietet Routinen, um verschiedene Informationen wie Absender, Empfänger,
Betreff und andere relevante Daten zu extrahieren.

Funktionen:
- get_sender_msg_file(file_path): Gibt die E-Mail-Adresse des Absenders zurück.
- parse_sender_msg_file(sender: str): Analysiert den Sender-String und extrahiert Name und E-Mail.
- get_subject_msg_file(file_path): Gibt den Betreff der Nachricht zurück.
- load_known_senders(file_path): Lädt bekannte Sender aus einer CSV-Datei.
- get_date_sent_msg_file(file_path): Gibt das Datum der gesendeten Nachricht zurück.
- create_log_file(base_name, directory): Erstellt ein Logfile im Excel-Format.
- log_entry(log_file_path, entry): Fügt einen neuen Eintrag in das Logfile hinzu.
- convert_to_utc_naive(datetime_stamp): Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.
- format_datetime(datetime_stamp, format_string): Formatiert einen Zeitstempel in das angegebene Format.
- custom_sanitize_text(encoded_textstring): Bereinigt einen Textstring von unerwünschten Zeichen.
- truncate_filename_if_needed(file_path, max_length, truncation_marker): Kürzt den Dateinamen, wenn nötig.

Verwendung:
Importiere die Funktionen aus diesem Modul in deinem Hauptprogramm oder anderen Modulen,
um auf die Metadaten von MSG-Dateien zuzugreifen.

Beispiel:
    from modules.msg_handling import get_sender_msg_file
    sender_email = get_sender_msg_file('example.msg')
"""

import extract_msg
import re
import os
import pandas as pd
from datetime import datetime


def get_sender_msg_file(file_path):
    """
    Ruft die E-Mail-Adresse des Absenders aus einer MSG-Datei ab.

    Parameter:
    file_path (str): Der Pfad zur MSG-Datei.

    Rückgabewert:
    str: Die E-Mail-Adresse des Absenders oder "Unbekannt", wenn kein Sender vorhanden ist.
    """
    try:
        msg = extract_msg.Message(file_path)
        return msg.sender if msg.sender else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Sender vorhanden ist
    except FileNotFoundError:
        return "Datei nicht gefunden"
    except Exception as e:
        return f"Fehler beim Auslesen des Senders: {str(e)}"


def parse_sender_msg_file(sender: str):
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
    email_match = re.search(email_pattern, sender)

    if email_match:
        sender_email = email_match.group(1)
        contains_sender_email = True

    # Entferne die E-Mail-Adresse aus dem Sender-String
    sender_name = re.sub(email_pattern, '', sender).strip()
    # Entferne Anführungszeichen aus dem Sender-String
    sender_name = sender_name.replace("\"", '')

    return {
        "sender_name": sender_name,
        "sender_email": sender_email,
        "contains_sender_email": contains_sender_email
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
        return msg.subject if msg.subject else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Betreff vorhanden ist
    except FileNotFoundError:
        return "Datei nicht gefunden"
    except Exception as e:
        return f"Fehler beim Auslesen des Betreffs: {str(e)}"


def load_known_senders(file_path):
    """
    Lädt bekannte Sender aus einer CSV-Datei.

    Parameter:
    file_path (str): Der Pfad zur CSV-Datei.

    Rückgabewert:
    DataFrame: Ein DataFrame mit den bekannten Sendern.
    """
    return pd.read_csv(file_path)


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
            return msg.date if msg.date else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Datum vorhanden ist
    except FileNotFoundError:
        return "Datei nicht gefunden"
    except Exception as e:
        return f"Fehler beim Auslesen des Datums: {str(e)}"


def create_log_file(base_name, directory):
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
    df = pd.DataFrame(columns=["Fortlaufende Nummer", "Verzeichnisname", "Filename", "Sendername", "Senderemail",
                               "Contains Senderemail", "Timestamp", "Formatierter Timestamp", "Betreff",
                               "Bereinigter Betreff", "Neuer Filename", "Neuer gekürzter Fielname", "Neuer Dateipfad"])

    try:
        df.to_excel(log_file_path, index=False)
        return log_file_path
    except Exception as e:
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
    if datetime_stamp.tzinfo is not None:
        return datetime_stamp.replace(tzinfo=None)  # Entfernen der Zeitzone
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
    if isinstance(datetime_stamp, datetime):
        return datetime_stamp.strftime(format_string)
    raise ValueError("Ungültiger Zeitstempel.")


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

