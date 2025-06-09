import logging
import os
from datetime import datetime
from config import LOG_FILE_DIRECTORY, DEBUG_LEVEL, LOG_TO_CONSOLE, MAX_DEBUG_LOG_FILE_COUNT

# DEBUG_LEVEL_TEXT hinzufügen
DEBUG_LEVEL_TEXT = {0: "ERROR", 1: "WARNING", 2: "INFO", 3: "DEBUG", 4: "TRACE"}

# Globale Variable, um sicherzustellen, dass Clean-Up nur einmal ausgeführt wird
_is_logger_initialized = False

# Logdatei: Hole das Verzeichnis aus der .env-Datei und prüfe, ob es existiert. Wenn nicht, erstelle es.
log_dir = os.path.abspath(LOG_FILE_DIRECTORY)
if not os.path.exists(log_dir):
    os.makedirs(log_dir, exist_ok=True)

# Log-Datei: Erstelle den Dateinamen mit Zeitstempel und erstelle den vollständigen Pfad.
current_time = datetime.now()
debug_log_file_name = current_time.strftime("debug_log_file_%Y-%m-%d_%HUhr%M_%Ss.txt")
prog_log_file_path = os.path.join(log_dir, debug_log_file_name)


def clean_old_log_files(directory: str, max_file_count: int):
    """
    Entfernt ältere Log-Dateien im Verzeichnis, wenn die maximale Anzahl überschritten ist.
    Gibt die Anzahl der gelöschten Dateien zurück.

    :param directory: Verzeichnis mit den Log-Dateien.
    :param max_file_count: Maximale Anzahl von Log-Dateien, die aufbewahrt werden.
    :return: Anzahl der gelöschten Log-Dateien.
    """
    try:
        # Log-Dateien im Verzeichnis suchen
        log_files = [
            os.path.join(directory, f)
            for f in os.listdir(directory)
            if f.startswith("debug_log_file_") and f.endswith(".txt")
        ]

        # Log-Dateien nach Erstellungsdatum sortieren
        log_files.sort(key=os.path.getmtime)

        # Soviele Log-Dateien löschen, bis die maximale Anzahl wieder eingehalten wird
        if len(log_files) > max_file_count:
            files_to_delete = len(log_files) - max_file_count
            for i in range(files_to_delete):
                os.remove(log_files[i])
            return files_to_delete

    except Exception as e:
        # Fehler beim Bereinigen protokollieren
        print(f"Fehler beim Bereinigen der Logdateien: {e}")
        return 0
    return 0


def initialize_logger(module_name: str):
    """
    Initialisiert den Logger für das angegebene Modul und richtet zentrale Logging-Funktionen ein.
    Diese Funktion erstellt einen benutzerdefinierten Logger, der unter anderem eine TRACE-Ebene
    (unterhalb von DEBUG) und Zeitstempel mit Millisekunden unterstützt. Zusätzlich wird sichergestellt,
    dass Initialisierungsaufgaben wie das Bereinigen der Logger-Konfiguration nur einmal umgesetzt werden.

    :param module_name: Name des Moduls, das den Logger verwendet (wird im Log angezeigt).
    :return: Ein initialisiertes Logger-Objekt.
        - Beim ersten Aufruf: Ein vollständig konfiguriertes Logger-Objekt für das angegebene Modul.
        - Bei wiederholten Aufrufen: Das bereits existierende (gleich konfiguriertes) Logger-Objekt.
    """

    # Globale Variable zur Überprüfung, ob die Logger-Konfiguration bereits initialisiert wurde.
    #global _is_logger_initialized
    #if _is_logger_initialized:
        # Wenn der Logger bereits initialisiert ist, wird der angeforderte Logger direkt zurückgegeben.
     #  return logging.getLogger(module_name)

    # Definiere ein benutzerdefiniertes TRACE-Level für detaillierte Debug-Informationen.
    TRACE_LEVEL = 9 # TRACE liegt zwischen NOTSET (0) und DEBUG (10).
    logging.addLevelName(TRACE_LEVEL, "TRACE") # Fügt das Level dem Logging-System hinzu.

    # Füge TRACE-Level als Methode für alle Logger hinzu.
    def trace(self, message, *args, **kwargs):
        if self.isEnabledFor(TRACE_LEVEL):
            self._log(TRACE_LEVEL, message, args, **kwargs)
    logging.Logger.trace = trace # Methode zu `logging.Logger` hinzufügen.

    # Mapping von konfigurierbaren Log-Ebenen (z. B. aus einer `.env`-Datei).
    LEVEL_NAME_MAP = {
        0: logging.ERROR,   # Fehler (ERROR)
        1: logging.WARNING, # Warnungen (WARNING)
        2: logging.INFO,    # Informationen (INFO)
        3: logging.DEBUG,   # Debug-Details (DEBUG)
        4: TRACE_LEVEL      # Erweiterte Debug-Details (TRACE)
    }

    # Versuche, den Log-Level basierend auf DEBUG_LEVEL zu setzen. Fallback ist INFO.
    log_level = LEVEL_NAME_MAP.get(DEBUG_LEVEL, logging.INFO)
    if log_level is None:
        log_level = logging.INFO

    # Erstelle oder hole den Logger für das aktuelle Modul.
    logger = logging.getLogger(module_name)

    # Entferne bestehende StreamHandler, falls vorhanden (z. B. um doppelte Konsolenausgabe zu vermeiden).
    for handler in list(logger.handlers):
        if isinstance(handler, logging.StreamHandler):
            logger.removeHandler(handler)

    # Konfiguriere Logger nur, wenn er noch keine Handlers besitzt.
    if not logger.hasHandlers():
        # Setze die globale Log-Level-Konfiguration für den Logger.
        logger.setLevel(log_level)

        # Richte den FileHandler ein, um Log-Einträge in eine Datei zu schreiben.
        file_handler = logging.FileHandler(prog_log_file_path, mode="a", encoding="utf-8")
        file_handler.setLevel(log_level) # Derselbe Level, wie im Logger konfiguriert.

        # Formatierungsregeln für das Log: inklusive Zeitstempel, Modulname und Zeilennummer.
        datefmt = "%Y-%m-%d %H:%M:%S"  # Ohne Millisekunden!
        # MillisecondFormatter ergänzt Zeitstempel um Millisekunden
        formatter = MillisecondFormatter(
            fmt="[%(asctime)s] [%(levelname)s] [%(name)s:%(lineno)d] %(message)s",
            datefmt=datefmt
        )
        file_handler.setFormatter(formatter) # Füge den benutzerdefinierten Formatter hinzu.

        # Füge den FileHandler dem Logger hinzu.
        logger.addHandler(file_handler)

    # Markiere die Logger-Initialisierung als abgeschlossen.
    #_is_logger_initialized = True

    return logger # Rückgabe des initialisierten Logger-Objekts.


def clean_logs_and_initialize():
    """
    Führt die Bereinigung alter Log-Dateien aus und initialisiert den Logger für die Anwendung.
    Diese Funktion wird idealerweise nur einmal vom Hauptprogramm (main) aufgerufen.
    """
    global _is_logger_initialized

    if not _is_logger_initialized:
        deleted_files_count = clean_old_log_files(log_dir, MAX_DEBUG_LOG_FILE_COUNT)
        app_logger = initialize_logger("__main__")  # Logger für das Hauptprogramm
        app_logger.info(f"{deleted_files_count} alte Debug-Log-Datei(en) wurde(n) gelöscht.")
        _is_logger_initialized = True


class MillisecondFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        from datetime import datetime
        ct = datetime.fromtimestamp(record.created)
        if datefmt:
            # Standardzeit (+ Millisekunden korrekt formatiert)
            s = ct.strftime(datefmt)  # Datefmt bleibt bis Sekunden korrekt
            return f"{s}.{ct.microsecond // 1000:03d}"  # Millisekunden hinzufügen
        else:
            # Der Standardwert hilft, wenn kein Format explizit angegeben ist
            return super().formatTime(record, datefmt)


def flush_log(logger):
    """
    Erzwingt das Leeren des Log-Buffers aller FileHandler eines angegebenen Loggers.

    :param logger: Ein Logger-Objekt, dessen Buffer geleert werden soll.
    """
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            handler.flush()

