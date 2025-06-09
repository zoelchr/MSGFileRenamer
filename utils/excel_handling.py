# -*- coding: utf-8 -*-
"""
excel_handling.py

Dieses Modul enthält Funktionen zur Erstellung und Speicherung von Excel-Listen aus MSG-Dateien.
Es ermöglicht das Sammeln von Metadaten über MSG-Dateien in einem strukturierten Format, das leicht
in Excel exportiert werden kann. Die Hauptfunktionen umfassen das Erstellen einer Excel-Liste aus
einer Liste von MSG-Dateien sowie das Speichern dieser Liste in einer Excel-Datei.

Funktionen:
- create_excel_list(msg_files): Erstellt eine Excel-Liste aus den gefundenen MSG-Dateien.
- save_excel_file(excel_list, output_file): Speichert die Excel-Liste in einer angegebenen Datei.
"""
import os
import pandas as pd
from logger import initialize_logger #, DEBUG_LEVEL_TEXT, prog_log_file_path

# In der Log-Datei wird als Quelle der Modulname "__main__" verwendet
app_logger = initialize_logger(__name__)
app_logger.debug("Debug-Logging im Modul 'excel_handling' aktiviert.")


def create_excel_list(msg_files):
    """
    Erstellt eine Excel-Liste aus den gefundenen MSG-Dateien.

    Diese Funktion erstellt ein DataFrame mit Informationen über jede gefundene MSG-Datei,
    einschließlich einer fortlaufenden Nummer, dem Dateinamen, dem Pfadnamen und der Länge des Pfades.

    :param msg_files: Eine Liste der gefundenen MSG-Dateien (vollständige Pfade).
    :return: Ein DataFrame, das die Informationen über die MSG-Dateien enthält.
    """
    # Initialisiere ein leeres DataFrame mit den entsprechenden Spalten
    excel_list = pd.DataFrame(columns=["Nummer", "Dateiname", "Pfadname", "Pfadlänge"])

    for i, file in enumerate(msg_files):
        filename = os.path.basename(file)  # Extrahiere den Dateinamen
        path = os.path.dirname(file)        # Extrahiere den Pfad
        entry = {
            "Nummer": i + 1,                # Fortlaufende Nummer
            "Dateiname": filename,           # Name der Datei
            "Pfadname": path,                # Verzeichnis der Datei
            "Pfadlänge": len(file)           # Länge des vollständigen Pfades
        }
        # Füge den neuen Eintrag zum DataFrame hinzu
        excel_list = pd.concat([excel_list, pd.DataFrame([entry])], ignore_index=True)
    return excel_list

def save_excel_file(excel_list, output_file):
    """
    Speichert die Excel-Liste in einer angegebenen Datei.

    Diese Funktion speichert das übergebene DataFrame als Excel-Datei an dem angegebenen Speicherort.

    :param excel_list: Das DataFrame, das gespeichert werden soll.
    :param output_file: Der Pfad zur Ausgabedatei, in der die Excel-Liste gespeichert wird.
    """
    # Speichere das DataFrame als Excel-Datei
    excel_list.to_excel(output_file, index=False)
    print(f"Excel-Liste erfolgreich gespeichert unter: {output_file}")


def clean_old_excel_files(directory: str, max_file_count: int, name_contains: str):
    """
    Entfernt ältere Excel-Log-Dateien im Verzeichnis, wenn die maximale Anzahl überschritten ist.
    Berücksichtigt nur Dateien, deren Namen einen bestimmten Teilstring enthalten.

    :param directory: Verzeichnis mit den Excel-Logdateien.
    :param max_file_count: Maximale Anzahl von Excel-Logdateien, die aufbewahrt werden.
    :param name_contains: Ein Teilstring, der im Namen der Excel-Dateien enthalten sein muss.
    :return: Anzahl der gelöschten Excel-Logdateien.
    """
    try:
        # Finde alle .xlsx-Dateien im angegebenen Verzeichnis, die den Teilstring enthalten
        excel_files = [
            os.path.join(directory, f)
            for f in os.listdir(directory)
            if f.endswith(".xlsx") and name_contains in f
        ]

        # Dateien nach Änderungsdatum sortieren (älteste zuerst)
        excel_files.sort(key=os.path.getmtime)

        deleted_files_count = 0

        # Prüfen, ob die maximale Anzahl überschritten ist
        if len(excel_files) > max_file_count:
            files_to_delete = len(excel_files) - max_file_count
            for i in range(files_to_delete):
                os.remove(excel_files[i])  # Lösche die älteste Datei
                deleted_files_count = deleted_files_count + 1

        app_logger.info(f"{deleted_files_count} alte Excel-Log-Datei(en) wurde(n) gelöscht.")
        return deleted_files_count

    except Exception as e:
        # Fehlerprotokollierung für Debugging
        print(f"Fehler beim Bereinigen der Excel-Logdateien: {e}")
        app_logger.error(f"Fehler beim Bereinigen der Excel-Logdateien: {e}")
        return 0
    return 0
