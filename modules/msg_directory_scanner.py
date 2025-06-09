"""
msg_directory_scanner.py

Dieses Modul durchsucht ein angegebenes Verzeichnis und alle seine Unterverzeichnisse nach MSG-Dateien.
Es erstellt eine Übersicht in Form einer Excel-Datei, die Informationen über jede gefundene MSG-Datei enthält,
einschließlich einer fortlaufenden Nummer, dem Dateinamen, dem Pfadnamen und der Länge des vollständigen Pfades.

Verwendung:
Um das Modul auszuführen, kann es direkt über die Kommandozeile aufgerufen werden.
Es können zwei optionale Parameter übergeben werden:
- -d "Pfad/zum/startverzeichnis": Das Verzeichnis, das durchsucht werden soll (Standard: D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long).
- -l: Das aktuelle Verzeichnis wird durchsucht und die Excel-Datei wird dort abgelegt.
- -o "Pfad/zum/ausgabeverzeichnis": Der Pfad, an dem die Excel-Datei gespeichert werden soll (Standard: D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx).

Beispielaufruf:
python msg_directory_scanner.py -d "C:/path/to/search" -o "C:/path/to/output.xlsx"
oder
python msg_directory_scanner.py -l
"""

import os
import sys
from utils.excel_handling import create_excel_list, save_excel_file
from logger import initialize_logger

# In der Log-Datei wird als Quelle der Modulname "__main__" verwendet
app_logger = initialize_logger(__name__)
app_logger.debug("Debug-Logging im Modul 'msg_directory_scanner' gestartet.")

def get_msg_files_from_directory(directory):
    """
    Durchsucht das angegebene Verzeichnis und alle Unterverzeichnisse nach MSG-Dateien.

    Diese Funktion verwendet os.walk, um rekursiv durch das Verzeichnis zu navigieren
    und alle Dateien mit der Endung '.msg' zu finden.

    :param directory: Das Verzeichnis, das durchsucht werden soll.
    :return: Eine Liste der gefundenen MSG-Dateien (vollständige Pfade).
    """
    msg_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".msg"):
                # Füge den vollständigen Pfad der gefundenen MSG-Datei zur Liste hinzu
                msg_files.append(os.path.join(root, file))
    return msg_files

if __name__ == "__main__":
    # Standardverzeichnisse für die Suche nach MSG-Dateien und die Ausgabe der Excel-Datei
    default_start_directory = r"D:/Dev/pycharm/MSGFileRenamer/data/sample_files/testset-long"
    default_output_excel_file = r"D:/Dev/pycharm/MSGFileRenamer/logs/msg_files_overview.xlsx"

    # Initialisierung der Parameter mit Standardwerten
    start_directory = default_start_directory
    output_excel_file = default_output_excel_file

    # Überprüfen der übergebenen Argumente für benutzerdefinierte Verzeichnisse
    args = sys.argv[1:]  # Alle Argumente außer dem Skriptnamen
    for i in range(len(args)):
        if args[i] == '-d' and i + 1 < len(args):
            start_directory = args[i + 1]  # Setze das Startverzeichnis
        elif args[i] == '-l':
            start_directory = os.getcwd()  # Setze das aktuelle Verzeichnis als Startverzeichnis
            output_excel_file = os.path.join(start_directory, "msg_files_overview.xlsx")  # Speichere die Excel-Datei im aktuellen Verzeichnis
        elif args[i] == '-o' and i + 1 < len(args):
            output_excel_file = args[i + 1]  # Setze den Pfad für die Excel-Datei

    # Suche nach MSG-Dateien im angegebenen Verzeichnis
    msg_files = get_msg_files_from_directory(start_directory)

    # Erstelle die Excel-Liste aus den gefundenen MSG-Dateien
    excel_list = create_excel_list(msg_files)

    # Speichere die Excel-Liste in einer Datei
    save_excel_file(excel_list, output_excel_file)

    # Ausgabe der Anzahl der gefundenen MSG-Dateien
    print(f"Anzahl der gefundenen MSG-Dateien: {len(msg_files)}")

