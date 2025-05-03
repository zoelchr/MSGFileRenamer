@echo off
:: SHAPTH - Dokumentenordner UAG
echo ***************************************************
echo *********** Dokumentenordner UAG ******************
echo ***************************************************
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "N:\11_Work"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "N:\10_Sonstige Besprechungen"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "N:\08_Veroeffentlichungen"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory  "N:\05_Kommunikationsunterlagen"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "N:\09_Foerderantraege"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "N:\08_Veroeffentlichungen"
echo ***************************************************
echo *********** Dokumentenordner Kernprojekt***********
echo ***************************************************
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "M:\8_Besprechungen"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "M:\0_Projektmanagement"
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "M:\3_Konzept"
echo ***************************************************
echo *********** Dokumentenordner LGL-Intern************
echo ***************************************************
python msg_file_renamer.py --no_test_run --set_filedate --generate_pdf  --recursive_search --excel_log_directory "./" --use_knownsender_file --knownsender_file ".\config\known_senders.csv" --search_directory "O:\4_Vergaben"

pause