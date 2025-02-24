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
                access_result['status'] = "Zugriff: Nur Lesen möglich"
    except Exception as e:
        access_result['status'] = "Zugriff: Nicht möglich"
        access_result['detail'] = str(e)

    return access_result

def rename_file(current_name, new_name, retries=3, delay_ms=1000):
    """
    Benennt eine Datei um und prüft die erfolgreiche Umbenennung.

    Parameter:
    current_name (str): Der aktuelle Dateiname.
    new_name (str): Der neue Dateiname.
    retries (int): Anzahl der Wiederholungen bei Misserfolg (Standard: 3).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 1000).

    Rückgabewert:
    str: Erfolgsmeldung oder Fehlermeldung.
    """
    attempt = 0
    last_error = None  # Variable für den letzten Fehler

    while attempt < retries:
        try:
            os.rename(current_name, new_name)
            return f"Datei erfolgreich umbenannt in: {new_name}"
        except Exception as e:
            last_error = e  # Speichere den Fehler
            attempt += 1
            time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    return f"Fehler: Umbenennung von '{current_name}' in '{new_name}' nach {retries} Versuchen fehlgeschlagen: {str(last_error)}"

def delete_file(file_path, retries=3, delay_ms=1000):
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

        # Öffne die Datei ohne with-Block
        handle = win32file.CreateFile(
            file_path,
            win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        try:
            # Setze das Erstelldatum
            win32file.SetFileTime(handle, creation_time, None, None)
        finally:
            # Stelle sicher, dass der Handle immer geschlossen wird
            win32file.CloseHandle(handle)

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

