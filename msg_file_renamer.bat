@echo off

:: Aktiviert die virtuelle Umgebung
call venv\Scripts\activate

:: Führt das Python-Skript in der aktivierten Umgebung aus
::python msg_file_renamer.py --no_test_run --init_testdata --set_filedate --debug_mode --excel_log_directory "./" --excel_log_basename "my_excel_log" --debug_log_directory "./"
::python msg_file_renamer.py --init_testdata --set_filedate --debug_mode --excel_log_directory "./" --excel_log_basename "my_excel_log" --debug_log_directory "./"

:: Nur ein Testlauf mit Erstellung von Testdaten
::python msg_file_renamer.py --init_testdata --set_filedate --debug_mode --excel_log_directory "./" --excel_log_basename "my_excel_log" --debug_log_directory "./"

:: Kein Testlauf mit Erstellung von Testdaten sowie aktivierten Debugmode, aber ohne Änderung des Filedatums (set_filedate)
:: python msg_file_renamer.py --no_test_run --init_testdata --debug_mode --excel_log_directory "./" --excel_log_basename "my_excel_log" --debug_log_directory "./"

:: Kein Testlauf mit Erstellung von Testdaten sowie aktivierten Debugmode, aber ohne Änderung des Filedatums (set_filedate)
:: python msg_file_renamer.py --no_test_run --init_testdata --debug_mode --set_filedate --excel_log_directory "./" --excel_log_basename "my_excel_log" --debug_log_directory "./"

:: Kein Testlauf ohne Erstellung von Testdaten sowie aktivierten Debugmode mit Änderung des Filedatums und Angabe des Suchpfades
:: python msg_file_renamer.py --no_test_run --debug_mode --set_filedate --search_directory "C:\Users\ruedi\Downloads" --excel_log_directory "./" --excel_log_basename "my_excel_log"

:: Testlauf ohne Erstellung von Testdaten sowie aktivierten Debugmode mit Änderung des Filedatums und Angabe des Suchpfades
:: python msg_file_renamer.py --debug_mode --set_filedate --search_directory "C:\Users\ruedi\Downloads" --excel_log_directory "./" --excel_log_basename "my_excel_log"

:: Testlauf ohne Erstellung von Testdaten sowie aktivierten Debugmode mit Änderung des Filedatums und Angabe des Suchpfades
python msg_file_renamer.py --no_test_run --debug_mode --set_filedate --search_directory "C:\Users\ruedi\Downloads" --excel_log_directory "./" --excel_log_basename "my_excel_log"


:: Hält die Konsole offen
pause

