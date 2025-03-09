"""
file_handling.py

Dieses Modul enthält Funktionen zur Handhabung von Dateien und Verzeichnissen.
Es bietet Routinen zum Testen des Zugriffs auf Dateien, Umbenennen von Dateien,
Löschen von Dateien und Verzeichnissen sowie zum Kopieren von Inhalten zwischen Verzeichnissen.
Zusätzlich ermöglicht es das Setzen von Erstellungs- und Änderungsdaten für Dateien.

Funktionen:
- test_file_access(file_path): Testet den Lese- und Schreibzugriff auf eine Datei.
- rename_file(current_name, new_name, retries=3, delay_ms=1000): Benennt eine Datei um und prüft die erfolgreiche Umbenennung.
- delete_file(file_path, retries=3, delay_ms=1000): Löscht eine Datei und prüft die erfolgreiche Löschung.
- sanitize_filename(filename): Ersetzt ungültige Zeichen durch Unterstriche.
- format_datetime_stamp(datetime_stamp, format_string): Formatiert einen Zeitstempel in das angegebene Format.
- set_file_date(file_path, new_date): Setzt das Änderungsdatum einer Datei auf einen vorgegebenen Wert.
- set_file_creation_date(file_path, new_creation_date): Setzt das Erstelldatum einer Datei auf einen vorgegebenen Wert.
- delete_directory_contents(directory_path): Löscht den gesamten Inhalt eines angegebenen Verzeichnisses.
- copy_directory_contents(source_directory_path, target_directory_path): Kopiert den gesamten Inhalt eines Quellverzeichnisses in ein Zielverzeichnis.
"""

import os
import time
import pywintypes
import win32file
import win32con
import shutil
import re
import datetime
import logging
from enum import Enum

class FileAccessStatus(Enum):
    READABLE = "File is readable"
    WRITABLE = "File is writable"
    EXECUTABLE = "File is executable"
    NOT_FOUND = "File not found"
    NO_PERMISSION = "No permission to access"
    LOCKED = "File is locked by another process"
    UNKNOWN_ERROR = "Unknown error"
    UNKNOWN = "Status is unkown"

class RenameResult(Enum):
    SUCCESS = "Success"
    FILE_NOT_FOUND = "File not found"
    DESTINATION_EXISTS = "Destination file already exists"
    PERMISSION_DENIED = "Permission denied"
    INVALID_FILENAME = "Invalid filename"
    INVALID_FILENAME1 = "Source is a file and destination is a directory"
    INVALID_FILENAME2 = "Part of the path is not a directory"
    UNKNOWN_ERROR = "Unknown error"

logger = logging.getLogger(__name__)

class FileHandle:
    def __init__(self, file_path):
        self.file_path = file_path
        self.handle = None

    def __enter__(self):
        self.handle = win32file.CreateFile(
            self.file_path,
            win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        return self.handle

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.handle is not None:
            win32file.CloseHandle(self.handle)

def test_read_access(file_path):
    try:
        with open(file_path, 'r'):
            logger.debug(f"Auf die Datei ist Lesezugriff möglich: {file_path}")
            return True
    except Exception as e:
        logger.debug(f"Auf die Datei ist kein Lesezugriff möglich: {file_path} (Exception-Objekt: {e})")
        return False

def test_write_access(file_path):
    try:
        with open(file_path, 'a'):
            return True
    except Exception:
        return False

def test_file_access(file_path):
    """
    Test ob der lesende und schreibende Zugriff auf eine Datei möglich ist.

    Rückgabewert:
    dict: Ein Dictionary mit dem Zugriffsergebnis und einer Detailinformation.
    """
    access_result = {}
    try:
        # Teste Lesezugriff
        with open(file_path, 'r'):
            try:
                # Teste Schreibzugriff
                with open(file_path, 'a'):
                    access_result['status'] = "Zugriff: Lesen und Schreiben möglich"
            except Exception:
                logger.debug(f"Auf die Datei ist nur lesender Zugriff möglich: {file_path}")  # Debugging-Ausgabe: Log-File
                access_result['status'] = "Zugriff: Nur Lesen möglich"
    except Exception as e:
        logger.debug(f"Auf die Datei ist kein Zugriff möglich: {file_path} (Exception-Objekt: {e})")  # Debugging-Ausgabe: Log-File
        access_result['status'] = "Zugriff: Nicht möglich"
        access_result['detail'] = str(e)

    return access_result

def test_file_access2(file_path: str) -> list[FileAccessStatus]:
    """Überprüft den Datei-Zugriffsstatus und gibt eine Liste von FileAccessStatus-Enums zurück.

    Beispiel:
    status = check_file_access(file_path)
    print(f"Status: {[s.value for s in status]}")
    """
    access_status = []

    try:
        # Prüfen, ob die Datei existiert
        if not os.path.exists(file_path):
            logging.warning(f"Datei nicht gefunden: {file_path}")
            return [FileAccessStatus.NOT_FOUND]

        # Prüfen, ob die Datei gesperrt ist (Windows-typisch)
        try:
            with open(file_path, "a"):  # Testweise öffnen zum Schreiben
                pass
        except PermissionError:
            logging.warning(f"Datei ist gesperrt oder nicht schreibbar: {file_path}")
            access_status.append(FileAccessStatus.LOCKED)

        # Prüfen, welche Berechtigungen vorliegen
        if os.access(file_path, os.R_OK):
            access_status.append(FileAccessStatus.READABLE)

        if os.access(file_path, os.W_OK):
            access_status.append(FileAccessStatus.WRITABLE)

        if os.access(file_path, os.X_OK):
            access_status.append(FileAccessStatus.EXECUTABLE)

        # Falls keine der Berechtigungen vorhanden ist
        if not access_status:
            logging.warning(f"Kein Zugriff auf Datei: {file_path}")
            access_status.append(FileAccessStatus.NO_PERMISSION)

        return access_status

    except FileNotFoundError:
        logging.error(f"Datei wurde nicht gefunden: {file_path}")
        return [FileAccessStatus.NOT_FOUND]

    except Exception as e:
        logging.exception(f"Unerwarteter Fehler beim Zugriff auf {file_path}: {e}")
        return [FileAccessStatus.UNKNOWN_ERROR]


def rename_file(current_name, new_name, retries=1, delay_ms=200):
    """
    Benennt eine Datei um und prüft die erfolgreiche Umbenennung.

    Parameter:
    current_name (str): Der aktuelle Dateiname.
    new_name (str): Der neue Dateiname.
    retries (int): Anzahl der Wiederholungen bei Misserfolg (Standard: 1).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 200).

    Rückgabewert:
    str: Erfolgsmeldung oder Fehlermeldung.
    """
    attempt = 0
    last_error = None  # Variable für den letzten Fehler

    while attempt < retries:
        try:
            logging.debug(f"Aktueller Dateiname: {current_name}")  # Debugging-Ausgabe: Log-File
            logging.debug(f"Neuer Dateiname: {new_name}")  # Debugging-Ausgabe: Log-File
            os.rename(current_name, new_name)
            return "Datei erfolgreich umbenannt."
        except FileNotFoundError:
            print(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Console
            logging.debug(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Log-File
            last_error = "Datei nicht gefunden."
        except PermissionError:
            print(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Console
            logging.debug(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Log-File
            last_error = "Berechtigungsfehler."
        except Exception as e:
            last_error = f"Fehler bei Umbenennung: {str(e)}"  # Speichere den Fehler

        attempt += 1
        time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    return f"{last_error}: Nach {retries} Versuchen fehlgeschlagen."

def rename_file2(current_name, new_name, retries=1, delay_ms=200) -> RenameResult:
    """
    Benennt eine Datei um und prüft die erfolgreiche Umbenennung.
    Neue verbesserte Version

    Parameter:
    current_name (str): Der aktuelle Dateiname.
    new_name (str): Der neue Dateiname.
    retries (int): Anzahl der Wiederholungen bei Misserfolg (Standard: 1).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 200).

    Rückgabewert:
    RenameResult: Enum-Wert, der den Erfolg oder Fehler beschreibt.
    """
    attempt = 0
    rename_file_result = None  # Variable für den Rückgabewert

    while attempt < retries:
        try:
            logging.debug(f"Aktueller Dateiname: {current_name}")  # Debugging-Ausgabe: Log-File
            logging.debug(f"Neuer Dateiname: {new_name}")  # Debugging-Ausgabe: Log-File
            os.rename(current_name, new_name)
            rename_file_result = RenameResult.SUCCESS  # Erfolgreiche Umbenennung
        except FileNotFoundError:
            print(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Console
            logging.error(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Log-File
            rename_file_result = RenameResult.FILE_NOT_FOUND  # Datei nicht gefunden
        except PermissionError:
            print(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Console
            logging.error(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Log-File
            rename_file_result = RenameResult.PERMISSION_DENIED  # Berechtigungsfehler
        except FileExistsError:
            print(f"Zieldatei existiert bereits: {new_name}")  # Debugging-Ausgabe: Console
            logging.error(f"Zieldatei existiert bereits: {new_name}")  # Debugging-Ausgabe: Log-File
            return RenameResult.DESTINATION_EXISTS
        except IsADirectoryError:
            print(f"Kann nicht umbenennen, da die Quelle eine Datei und das Ziel ein Verzeichnis ist.")  # Debugging-Ausgabe: Console
            logging.error(f"Kann nicht umbenennen, da die Quelle eine Datei und das Ziel ein Verzeichnis ist.")  # Debugging-Ausgabe: Log-File
            return RenameResult.INVALID_FILENAME1
        except NotADirectoryError:
            print(f"Ein Teil des Pfades ist kein Verzeichnis: {current_name} oder {new_name}")  # Debugging-Ausgabe: Console
            logging.error(f"Ein Teil des Pfades ist kein Verzeichnis: {current_name} oder {new_name}")  # Debugging-Ausgabe: Log-File
            return RenameResult.INVALID_FILENAME2
        except Exception as e:
            rename_file_result = RenameResult.UNKNOWN_ERROR  # Unbekannter Fehler
            logging.error(f"Fehler bei Umbenennung: {str(e)}")  # Debugging-Ausgabe: Log-File

        attempt += 1
        time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    return rename_file_result

def delete_file(file_path, retries=1, delay_ms=200):
    """
    Löscht eine Datei und prüft die erfolgreiche Löschung.

    Parameter:
    file_path (str): Der Pfad zur Datei, die gelöscht werden soll.
    retries (int): Anzahl der Wiederholungen bei Misserfolg (Standard: 3).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 1000).

    Rückgabewert:
    str: Erfolgsmeldung oder Fehlermeldung.
    """
    attempt = 0
    last_error = None  # Variable für den letzten Fehler

    while attempt < retries:
        try:
            os.remove(file_path)
            # Überprüfen, ob die Datei tatsächlich gelöscht wurde
            if not os.path.exists(file_path):
                return f"Datei erfolgreich gelöscht: {file_path}"
        except Exception as e:
            last_error = e  # Speichere den Fehler
            attempt += 1
            time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    return f"Fehler: Löschen der Datei '{file_path}' nach {retries} Versuchen fehlgeschlagen: {str(last_error)}"


def sanitize_filename(filename):
    """
    Ersetzt ungültige Zeichen im Dateinamen durch Unterstriche.

    Parameter:
    filename (str): Der ursprüngliche Dateiname.

    Rückgabewert:
    str: Der bereinigte Dateiname, der nur gültige Zeichen enthält.
    """
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def format_datetime_stamp(datetime_stamp, format_string):
    """
    Formatiert einen Zeitstempel in das angegebene Format.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der formatiert werden soll.
    format_string (str): Das gewünschte Format für den Zeitstempel.

    Rückgabewert:
    str: Der formatierte Zeitstempel als String.
    """
    # Überprüfen, ob datetime_stamp ein datetime-Objekt ist
    if isinstance(datetime_stamp, datetime.datetime):
        dt = datetime_stamp  # Direkt verwenden, wenn es ein datetime-Objekt ist
    else:
        # Annehmen, dass es ein String ist und in ein datetime-Objekt umwandeln
        dt = datetime.datetime.strptime(datetime_stamp, "%Y-%m-%d %H:%M:%S")  # Format anpassen, falls nötig

    # Formatieren des datetime-Objekts gemäß dem bereitgestellten Format
    return dt.strftime(format_string)

def set_file_date(file_path, new_date):
    """
    Setzt das Änderungsdatum einer Datei auf einen vorgegebenen Wert.

    Parameter:
    file_path (str): Der Pfad zur Datei.
    new_date (str): Das neue Datum im Format 'YYYY-MM-DD HH:MM:SS'.

    Rückgabewert:
    str: Erfolgsmeldung oder Fehlermeldung.
    """
    try:
        # Konvertiere das Datum in einen Zeitstempel
        timestamp = time.mktime(datetime.datetime.strptime(new_date, '%Y-%m-%d %H:%M:%S').timetuple())
        os.utime(file_path, (timestamp, timestamp))
        return f"Das Datum der Datei '{file_path}' wurde erfolgreich auf {new_date} gesetzt."
    except Exception as e:
        return f"Fehler beim Setzen des Datums für '{file_path}': {str(e)}"

def set_file_creation_date(file_path, new_creation_date):
    """
    Setzt das Erstelldatum einer Datei auf einen vorgegebenen Wert.

    Parameter:
    file_path (str): Der Pfad zur Datei.
    new_creation_date (str): Das neue Erstelldatum im Format 'YYYY-MM-DD HH:MM:SS'.

    Rückgabewert:
    str: Erfolgsmeldung oder Fehlermeldung.
    """
    try:
        # Konvertiere das Datum in einen Zeitstempel
        timestamp = time.mktime(time.strptime(new_creation_date, '%Y-%m-%d %H:%M:%S'))
        creation_time = pywintypes.Time(timestamp)

        with FileHandle(file_path) as handle:
            # Setze das Erstelldatum
            win32file.SetFileTime(handle, creation_time, None, None)

        return f"Das Erstelldatum der Datei '{file_path}' wurde erfolgreich auf {new_creation_date} gesetzt."
    except Exception as e:

        return f"Fehler beim Setzen des Erstelldatums für '{file_path}': {str(e)}"


def delete_directory_contents(directory_path):
    """
    Löscht den gesamten Inhalt des angegebenen Verzeichnisses.

    Parameter:
    directory_path (str): Der Pfad des Verzeichnisses, dessen Inhalt gelöscht werden soll.

    Rückgabewert:
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
                os.remove(item_path)  # Löscht die Datei
            else:
                shutil.rmtree(item_path)  # Löscht das Verzeichnis rekursiv
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

    os.makedirs(target_directory_path, exist_ok=True)  # Erstellt das Zielverzeichnis, falls es nicht existiert

    try:
        for item in os.listdir(source_directory_path):
            s = os.path.join(source_directory_path, item)
            d = os.path.join(target_directory_path, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, False, None)  # Kopiert Verzeichnisse rekursiv
            else:
                shutil.copy2(s, d)  # Kopiert Dateien
        return "Inhalt erfolgreich kopiert."
    except Exception as e:
        raise OSError(f"Fehler beim Kopieren des Inhalts: {e}")

