# -*- coding: utf-8 -*-
"""
Modul: pdf_generation.py

Dieses Modul stellt Funktionen bereit, um aus einer MSG-Datei ein PDF-Dokument zu erzeugen.
Es umfasst Funktionen zur Bereinigung und Formatierung des E-Mail-Inhalts, zur Extraktion
von relevanten Daten (wie Datum, Absender, Empfänger, Betreff, Textinhalt und Anhänge) sowie
zur Erzeugung eines kompakten PDF-Ausdrucks der E-Mail. Der erzeugte PDF-Ausdruck wird
als Zusammenfassung der Inhalte der MSG-Datei angezeigt und berücksichtigt maximale
Zeichenbegrenzungen sowie spezielle Formatierungen (z.B. Reduzierung mehrfacher Leerzeilen
und Einfügen von Leerzeilen vor typischen deutschen Grußformeln).

Verwendungszweck:
- Extraktion und Bereinigung des E-Mail-Textes aus MSG-Dateien.
- Erstellung eines PDF-Dokuments mit formatiertem E-Mail-Inhalt und Anhängen.
- Unterstützung der Darstellung auch bei längeren Inhalten und komplexen Formatierungen.

Abhängigkeiten:
- fpdf (zur PDF-Erzeugung)
- Standardbibliotheken wie os, logging, re
- Modul 'modules.msg_handling' für den Zugriff auf MSG-Dateien und Statusabfragen.
"""

import os
import re
import unicodedata
from fpdf import FPDF
from modules.msg_handling import MsgAccessStatus, get_msg_object

from logger import initialize_logger
app_logger = initialize_logger(__name__)
app_logger.debug("Debug-Logging im Modul 'msg_to_pdf' aktiviert.")

ALLOWED_CONTROL_CHARACTERS = ['\n', '\t', '\r', '\f', '\v']


def clean_email_text(text):
    """
    Bereinigt den Text einer E-Mail, um die Lesbarkeit zu verbessern und unnötige Leerzeilen zu entfernen.

    :param text: Der ursprüngliche E-Mail-Text.
    :return: Der bereinigte E-Mail-Text.
    """

    # 1. Decode eventuell vorhandene HTML-Entities (falls notwendig)
    import html
    text = html.unescape(text)

    # 2. Tabulator durch Leerzeichen
    text = text.replace('\t', '    ')

    # 3a. Mehrfache Leerzeilen Fall 1
    text = re.sub('\r\n\r\n \r\n\r\n \r\n\r\n  \r\n\r\n  \r\n\r\n', '\n\n\n', text)

    # 3b. Mehrfache Leerzeilen Fall 2
    text = re.sub('\r\n\r\n \r\n\r\n \r\n\r\n  \r\n\r\n', '\n\n', text)

    # 3c. Mehrfache Leerzeilen Fall 3
    text = re.sub('\r\n\r\n \r\n\r\n \r\n\r\n', '\n\n', text)

    # 3d. Mehrfache Leerzeilen Fall 4
    text = re.sub('\r\n\r\n \r\n\r\n', '\n\n', text)

    # 3e. Mehrfache Leerzeilen Fall 6
    text = re.sub('\r\n\r\n', '\n', text)

    # 4. Normalize Windows-style line endings and remove extra spaces around newlines
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'[ \t]+\n', '\n', text)
    text = re.sub(r'\n[ \t]+', '\n', text)

    # 5. Reduce multiple newlines (with optional spaces in between) to a maximum of two
    text = re.sub(r'(\n\s*){3,}', '\n\n', text)

    # 6. Remove non-printable characters
    text = ''.join(c for c in text if c.isprintable() or c in '\n')

    # 7. Ensure at least one blank line before typical German closing phrases
    closing_phrases = [
        "Mit freundlichen Grüßen",
        "Viele Grüße",
        "Herzliche Grüße",
        "Beste Grüße",
        "Mit besten Grüßen",
        "Mit herzlichen Grüßen"
    ]

    for phrase in closing_phrases:
        text = re.sub(rf"(?<!\n)\n({phrase})", r"\n\n\1", text)

    # 8. Spezialfall für eine spezifische E-Mail-Adresse
    text = re.sub(rf'E-Mail: ruediger.zoelch@lgl.bayern.de <mailto:ruediger.zoelch@lgl.bayern.de>', rf"E-Mail: ruediger.zoelch@lgl.bayern.de <mailto:ruediger.zoelch@lgl.bayern.de>\n", text)

    # 9. Spezielle Zeichen ersetzen, z.B. Emojis
    text = text.replace("\U0001f60a", ":)")
    text = text.replace("\U0001f603", ":)")
    text = text.replace("\U0001f614", ":)")
    text = text.replace("\U0001f622", ":)")

    return text.strip()


def remove_unsupported_chars(text):
    """
    Entfernt ungültige Zeichen und erlaubt nur explizit definierte Steuerzeichen.
    """
    return ''.join(
        c for c in text
        if c in ALLOWED_CONTROL_CHARACTERS or c.isprintable()
    )


def generate_pdf_from_msg(msg_path_and_filename:str, MAX_LENGTH_SENDERLIST: int):
    """
    Erzeugt ein PDF-Dokument aus einer MSG-Datei.

    :param msg_path_and_filename: Der Dateiname der MSG-Datei.
    :param MAX_LENGTH_SENDERLIST: Maximale Länge der Empfängerliste im PDF.
    :return: Der Pfad zur erzeugten PDF-Datei.
    """

    # Vorbelegung der Rückgabewerte
    is_generate_pdf_successful = False
    pdf_path_and_filename = ""

    # PDF erstellen
    pdf = FPDF()
    pdf.add_page()

    # Schriftart NotoSans Regular laden
    pdf.add_font("NotoSans", style="", fname="./font/NotoSans-Regular.ttf", uni=True)

    # Schriftart NotoSans Bold laden
    pdf.add_font("NotoSans", style="B", fname="./font/NotoSans-Bold.ttf", uni=True)

    # Segoe UI TTF von Windows einbinden (Unicode-fähig)
    # pdf.add_font("Segoe", "", "C:\\Windows\\Fonts\\segoeui.ttf", uni=True)
    # pdf.add_font("Segoeb", "", "C:\\Windows\\Fonts\\segoeuib.ttf", uni=True)

    pdf.set_font("NotoSans", size=4)
    pdf.write(5, f"Dieser PDF-Ausdruck der Email ist eventuell gekürzt (max 6000 Zeichen). Zusätzlich können Beeinträchtigungen bei der Formatierung auftreten, z.B. Darstellung von Tabellen. Die vollständige Email findet sich in der zugehörigen MSG-Datei.\n")
    pdf.set_font("NotoSans", size=8)

    # Schritt 1: Überprüfen, ob der Pfad zu einer existierenden Datei führt
    msg_object = {"status": [MsgAccessStatus.FILE_NOT_FOUND]} # Vorbelegung der Rückgabewerte, auch wenn kein msg_object erzeugt werden kann
    if os.path.isfile(msg_path_and_filename):
        app_logger.debug(f"Schritt 1: Die Datei '{msg_path_and_filename}' existiert.")  # Debugging-Ausgabe: Log-File

        # Schritt 2: PDF-Dateiname festlegen
        pdf_path_and_filename = os.path.splitext(msg_path_and_filename)[0] + ".pdf"
        app_logger.debug(f"Schritt 2: Dateiname für PDF-Dokument '{pdf_path_and_filename}'.")  # Debugging-Ausgabe: Log-File

        # Schritt 3: Auslesen des msg-Objektes und bei Fehler abbrechen
        try:
            msg_object = get_msg_object(msg_path_and_filename)
        except Exception as e:
            app_logger.warning(f"Schritt 3: Fehler bei der PDF-Erstellung da kein Zugriff auf msg_object: {e}")
            is_generate_pdf_successful = False
            return is_generate_pdf_successful, pdf_path_and_filename

        # Schritt 4: Zeitstempel ausgeben
        if not MsgAccessStatus.DATE_MISSING in msg_object["status"]:
            msg_date = msg_object["date"]
            app_logger.debug(f"Schritt 4: Zeitstempel für PDF-Ausgabe '{msg_date}'.")  # Debugging-Ausgabe: Log-File

            try:
                pdf.set_font("NotoSans", style="B", size=8)
                pdf.write(5, f"Versandzeitpunkt: ")
                pdf.set_font("NotoSans", size=8)
                pdf.write(5, f"{msg_date}\n")
            except Exception as e:
                app_logger.warning(f"Schritt 4: Fehler bei der PDF-Erstellung: {e}")

        # Schritt 5: Sender ausgeben
        if not MsgAccessStatus.SENDER_MISSING in msg_object["status"]:
            msg_sender = msg_object["sender"]
            msg_sender = remove_unsupported_chars(msg_sender)
            app_logger.debug(f"Schritt 5: Sender für PDF-Ausgabe '{msg_sender}'.")  # Debugging-Ausgabe: Log-File

            try:
                pdf.set_font("NotoSans", style="B", size=8)
                pdf.write(5, f"Absender: ")
                pdf.set_font("NotoSans", size=8)
                pdf.write(5, f"{msg_sender}\n")
            except Exception as e:
                app_logger.warning(f"Schritt 5: Fehler bei der PDF-Erstellung: {e}")

        # Schritt 6: Empfänger ausgeben
        if not MsgAccessStatus.NO_RECIPIENT_FOUND in msg_object["status"]:
            msg_recipient = msg_object["recipient"]

            # Extract email addresses using regex
            emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', msg_recipient)
            email_list = ', '.join(emails)

            # Truncate the email list if it exceeds the maximum length
            if len(email_list) > MAX_LENGTH_SENDERLIST:
                email_list = email_list[:MAX_LENGTH_SENDERLIST] + '<GEKÜRZT>'

            # Remove unsupported characters
            cleaned_text = remove_unsupported_chars(email_list)

            try:
                pdf.set_font("NotoSans", style="B", size=8)
                pdf.write(5, f"Empfänger: ")
                pdf.set_font("NotoSans", size=8)
                pdf.write(5, f"{cleaned_text}\n")
            except Exception as e:
                app_logger.warning(f"Schritt 6: Fehler bei der PDF-Erstellung: {e}")

        # Schritt 7: Betreff ausgeben
        if not MsgAccessStatus.SUBJECT_MISSING in msg_object["status"]:
            cleaned_text = msg_object["subject"]
            cleaned_text = remove_unsupported_chars(cleaned_text)
            app_logger.debug(f"Schritt 7: Betreff für PDF-Ausgabe '{cleaned_text}'.")  # Debugging-Ausgabe: Log-File

            # PDF-Inhalt hinzufügen
            try:
                pdf.set_font("NotoSans", style="B", size=8)
                pdf.write(5, f"Betreff: ")
                pdf.set_font("NotoSans", size=8)
                pdf.write(5, f"{cleaned_text}\n")
            except Exception as e:
                app_logger.warning(f"Schritt 7:  Fehler bei der PDF-Erstellung: {e}")

        # Schritt 8: Inhalt ausgeben
        if not MsgAccessStatus.BODY_MISSING in msg_object["status"]:
            msg_body = msg_object["body"]

            # Truncate the body if it exceeds certain number of characters
            MAX_BODY_LENGTH = 6000
            if len(msg_body) > MAX_BODY_LENGTH:
                pdf.set_font("NotoSans", size=8)
                msg_body = msg_body[:MAX_BODY_LENGTH] + '\n<HINWEIS: NACHRICHT WURDE AUF 6000 ZEICHEN GEKÜRZT! VOLLSTÄNDIGE NACHRICHT SIEHE GLEICHNAMIGES MSG-FILE>'

            # Body ausgeben
            pdf.set_font("NotoSans", style="B", size=10)
            pdf.write(5, f"\nInhalt:\n")
            pdf.set_font("NotoSans", size=8)

            # Text bereinigen, damit kompakte Darstellung möglich
            cleaned_text = clean_email_text(msg_body)  # <- dein ausgelesener Text
            cleaned_text = remove_unsupported_chars(cleaned_text)
            app_logger.debug(f"Schritt 8: Body-Text wurde bereinigt und besitzt nach Kürzung eine Länge von {len(cleaned_text)}")  # Debugging-Ausgabe: Log-File

            # Text ausgeben
            try:
                pdf.write(5, f"{cleaned_text}")
            except Exception as e:
                app_logger.warning(f"Schritt 8:  Fehler bei der PDF-Erstellung: {e}")

        else:
            pdf.set_font("NotoSans", style="B", size=10)
            pdf.write(5, f"\nInhalt:\n")
            pdf.set_font("NotoSans", size=8)
            pdf.write(5, f"\nHINWEIS: NACHRICHT OHNE INHALT ODER KANN NICHT GELESEN WERDEN (z.B. SIGNATURPRÜFUNG)!")

        # Schritt 9: Anhänge ausgeben
        if not MsgAccessStatus.ATTACHMENTS_MISSING in msg_object["status"]:
            msg_attachments = msg_object["attachments"]
            pdf.set_font("NotoSans", style="B", size=8)
            pdf.write(5, f"\n\nAnhänge:\n")
            pdf.set_font("NotoSans", size=8)

            for filename in msg_object["attachments"]:
               pdf.write(5, f"- {filename}\n")

        # Speichern der PDF-Datei
        pdf.output(pdf_path_and_filename)

        is_generate_pdf_successful = True

        app_logger.debug(f"\tDie PDF wurde unter '{os.path.abspath(pdf_path_and_filename)}' gespeichert.")

    else:
        #print(f"\tDie MSG-Datei '{msg_path_and_filename}' konnte nicht gelesen werden.")
        app_logger.warning(f"Die MSG-Datei '{msg_path_and_filename}' konnte nicht gelesen werden.")
        is_generate_pdf_successful = False

    return is_generate_pdf_successful, pdf_path_and_filename