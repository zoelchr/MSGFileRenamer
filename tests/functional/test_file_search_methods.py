import os
from pathlib import Path
import winreg

def normalize_path(path):
    """Stellt sicher, dass der Pfad mit \\?\ beginnt."""
    abs_path = os.path.abspath(path)
    if not abs_path.startswith("\\\\?\\"):  # Kein Raw-String nötig
        return "\\\\?\\" + abs_path  # Korrekte Escape-Sequenz
    return abs_path

def search_with_os_walk(target_directory):
    """Suche Dateien mit os.walk."""
    target_directory = normalize_path(target_directory)
    found_files = set()
    for root, _, files in os.walk(target_directory):
        for file in files:
            found_files.add(os.path.join(root, file))
    return found_files

def search_with_rglob(target_directory):
    """Suche Dateien mit pathlib.Path.rglob."""
    target_directory = normalize_path(target_directory)
    found_files = set(Path(target_directory).rglob("*"))
    return {str(file) for file in found_files}

def compare_results(os_walk_files, rglob_files):
    """Vergleiche die Ergebnisse der beiden Suchmethoden."""
    only_in_os_walk = os_walk_files - rglob_files
    only_in_rglob = rglob_files - os_walk_files

    print("Vergleich der Suchmethoden:\n")
    print(f"Gesamtanzahl gefundener Dateien mit os.walk: {len(os_walk_files)}")
    print(f"Gesamtanzahl gefundener Dateien mit rglob: {len(rglob_files)}\n")

    if only_in_os_walk:
        print("Dateien, die nur mit os.walk gefunden wurden:")
        print("\n".join(only_in_os_walk))

    if only_in_rglob:
        print("\nDateien, die nur mit rglob gefunden wurden:")
        print("\n".join(only_in_rglob))

def check_long_paths_enabled():
    """Überprüft, ob die LongPathsEnabled-Einstellung aktiviert ist."""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem", 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, "LongPathsEnabled")
            return value == 1
    except FileNotFoundError:
        return False

if __name__ == "__main__":
    if check_long_paths_enabled():
        print("✅ Lange Pfade sind aktiviert!")
    else:
        print("❌ Lange Pfade sind nicht aktiviert!")

    TARGET_DIRECTORY = r"D:\@MyWorkLocal\lgl-email-beispiele"

    os_walk_files = search_with_os_walk(TARGET_DIRECTORY)
    rglob_files = search_with_rglob(TARGET_DIRECTORY)

    compare_results(os_walk_files, rglob_files)