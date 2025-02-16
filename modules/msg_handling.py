"""
msg_handling.py

Dieses Modul enthält Funktionen zum Abrufen von Metadaten aus MSG-Dateien.
Es bietet Routinen, um verschiedene Informationen wie Absender, Empfänger,
Betreff und andere relevante Daten zu extrahieren.

Funktionen:
- get_sender_email_msg_file(file_path): Gibt die E-Mail-Adresse des Absenders zurück.
- get_subject_msg_file(file_path): Gibt den Betreff der Nachricht zurück.
- get_date_sent_msg_file(file_path): Gibt das Datum der gesendeten Nachricht zurück.

Verwendung:
Importiere die Funktionen aus diesem Modul in deinem Hauptprogramm oder anderen Modulen,
um auf die Metadaten von MSG-Dateien zuzugreifen.

Beispiel:
    from modules.msg_handling import get_sender_email
    sender_email = get_sender_email('example.msg')
"""
import extract_msg
import re
import os
import pandas as pd
from datetime import datetime


def get_sender_msg_file(file_path):
    # Abrufen des Senders aus einem MSG-File
    try:
        msg = extract_msg.Message(file_path)
        return msg.sender if msg.sender else "Unbekannt"  # Rückgabe eines Standardwerts, wenn kein Sender vorhanden ist
    except FileNotFoundError:
        return "Datei nicht gefunden"
    except Exception as e:
        return f"Fehler beim Auslesen des Senders: {str(e)}"

def parse_sender_msg_file(sender: str):
    """
    Analysiert den Sender-String eines MSG-Files und extrahiert den Namen und die Email-Adresse.

    Parameter:
    sender (str): Der Sender-String.

    Gibt:
    dict: Ein Dictionary mit 'sender_name', 'sender_email' und 'contains_sender_email'.
    """
    sender_name = ""
    sender_email = ""
    contains_sender_email = False

    # Regulärer Ausdruck für die Email-Adresse
    email_pattern = r'<(.*?)>'
    email_match = re.search(email_pattern, sender)

    if email_match:
        sender_email = email_match.group(1)
        contains_sender_email = True

    # Entferne die Email-Adresse aus dem Sender-String
    sender_name = re.sub(email_pattern, '', sender).strip()
    # Entferne Anführungszeichen aus dem Sender-String
    sender_name = sender_name.replace("\"", '')

    return {
        "sender_name": sender_name,
        "sender_email": sender_email,
        "contains_sender_email": contains_sender_email
    }

# Funktion zum Laden der bekannten Sender
def load_known_senders(file_path):
    """
    Lädt die bekannten Sender aus einer CSV-Datei.

    Parameter:
    file_path (str): Der Pfad zur CSV-Datei.

    Gibt:
    DataFrame: Ein DataFrame mit den bekannten Sendern.
    """
    return pd.read_csv(file_path)

def get_date_sent_msg_file(file_path):
    # Abrufen des Datums aus einem MSG-File
    try:
        msg = extract_msg.Message(file_path)
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

    Gibt:
    str: Der Pfad zur erstellten Logdatei.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{base_name}_{timestamp}.xlsx"
    log_file_path = os.path.join(directory, log_file_name)

    # Leeres DataFrame mit den gewünschten Spalten erstellen
    df = pd.DataFrame(columns=["Fortlaufende Nummer", "Verzeichnisname", "Filename", "Sendername", "Senderemail", "Contains Senderemail", "Timestamp", "Formatierter Timestamp"])

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
    """
    df = pd.read_excel(log_file_path)

    # Erstellen eines DataFrames aus dem Eintrag
    new_entry_df = pd.DataFrame([entry])

    # Überprüfen, ob new_entry_df leer ist oder nur NA-Werte enthält
    if not new_entry_df.empty and not new_entry_df.isnull().all(axis=1).any():
        # Logdatei laden oder erstellen
        if os.path.exists(log_file_path):
            df = pd.read_excel(log_file_path)
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

    Gibt:
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

    Gibt:
    str: Der formatierte Zeitstempel als String.
    """
    if isinstance(datetime_stamp, datetime):
        return datetime_stamp.strftime(format_string)
    raise ValueError("Ungültiger Zeitstempel.")