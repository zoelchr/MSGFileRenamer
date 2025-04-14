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
import logging
import re
from fpdf import FPDF
from modules.msg_handling import MsgAccessStatus, get_msg_object

logger = logging.getLogger(__name__)

def clean_email_text(text):
    """
    Bereinigt den Text einer E-Mail, um die Lesbarkeit zu verbessern und unnötige Leerzeilen zu entfernen.

    :param text: Der ursprüngliche E-Mail-Text.
    :return: Der bereinigte E-Mail-Text.
    """

    # 1. Decode eventuell vorhandene HTML-Entities (falls notwendig)
    import html
    text = html.unescape(text)

    # 2. Tabulator entfernen
    text = text.replace('\t', ' ')

    # 3. Kompaktere Lösung für alle Fälle
    # text = re.sub(r'(\r\n\s*){3,}', '\n\n', text)

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

    return text.strip()


def generate_pdf_from_msg(msg_path_and_filename:str, MAX_LENGTH_SENDERLIST: int):
    """
    Erzeugt ein PDF-Dokument aus einer MSG-Datei.

    :param msg_path_and_filename: Der Dateiname der MSG-Datei.
    :param MAX_LENGTH_SENDERLIST: Maximale Länge der Empfängerliste im PDF.
    :return: Der Pfad zur erzeugten PDF-Datei.
    """

    # PDF erstellen
    pdf = FPDF()
    pdf.add_page()

    # Segoe UI TTF von Windows einbinden (Unicode-fähig)
    pdf.add_font("Segoe", "", "C:\\Windows\\Fonts\\segoeui.ttf", uni=True)
    pdf.add_font("Segoeb", "", "C:\\Windows\\Fonts\\segoeuib.ttf", uni=True)


    pdf.set_font("Segoe", size=5)
    pdf.write(5, f"Dieser PDF-Ausdruck der Email ist eventuell gekürzt (max 6000 Zeichen). Zusätzlich können Beeinträchtigungen bei der Formatierung auftreten, z.B. Darstellung von Tabellen. Die vollständige Email findet sich in der zugehörigen MSG-Datei.\n")
    pdf.set_font("Segoe", size=8)

    # Überprüfen, ob der Pfad zu einer existierenden Datei führt
    if os.path.isfile(msg_path_and_filename):
        logger.debug(f"Schritt 1: Die Datei '{msg_path_and_filename}' existiert.")  # Debugging-Ausgabe: Log-File

        # PDF-Dateiname festlegen
        pdf_path_and_filename = os.path.splitext(msg_path_and_filename)[0] + ".pdf"
        logger.debug(f"Schritt 2: Dateiname für PDF-Dokument '{pdf_path_and_filename}'.")  # Debugging-Ausgabe: Log-File

        # Auslesen des msg-Objektes
        msg_object = get_msg_object(msg_path_and_filename)

        # Zeitstempel ausgeben
        if not MsgAccessStatus.DATE_MISSING in msg_object["status"]:
            msg_date = msg_object["date"]

            #pdf.set_font("Segoe", style="B", size=8)
            pdf.set_font("Segoeb", size=8)
            pdf.write(5, f"Versandzeitpunkt: ")
            pdf.set_font("Segoe", size=8)
            pdf.write(5, f"{msg_date}\n")

        # Sender ausgeben
        if not MsgAccessStatus.SENDER_MISSING in msg_object["status"]:
            msg_sender = msg_object["sender"]
            pdf.set_font("Segoeb", size=8)
            pdf.write(5, f"Absender: ")
            pdf.set_font("Segoe", size=8)
            pdf.write(5, f"{msg_sender}\n")

        # Empfänger ausgeben
        if not MsgAccessStatus.NO_RECIPIENT_FOUND in msg_object["status"]:
            msg_recipient = msg_object["recipient"]

            # Extract email addresses using regex
            emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', msg_recipient)
            email_list = ', '.join(emails)

            # Truncate the email list if it exceeds the maximum length
            if len(email_list) > MAX_LENGTH_SENDERLIST:
                email_list = email_list[:MAX_LENGTH_SENDERLIST] + '<GEKÜRZT>'

            pdf.set_font("Segoeb", size=8)
            pdf.write(5, f"Empfänger: ")
            pdf.set_font("Segoe", size=8)
            pdf.write(5, f"{email_list}\n")

        # Betreff ausgeben
        if not MsgAccessStatus.SUBJECT_MISSING in msg_object["status"]:
            msg_subject = msg_object["subject"]
            # PDF-Inhalt hinzufügen
            pdf.set_font("Segoeb", size=8)
            pdf.write(5, f"Betreff: ")
            pdf.set_font("Segoe", size=8)
            pdf.write(5, f"{msg_subject}\n")

        # Inhalt ausgeben
        if not MsgAccessStatus.BODY_MISSING in msg_object["status"]:
            msg_body = msg_object["body"]

            # Truncate the body if it exceeds certain number of characters
            MAX_BODY_LENGTH = 6000
            if len(msg_body) > MAX_BODY_LENGTH:
                pdf.set_font("Segoeb", size=8)
                msg_body = msg_body[:MAX_BODY_LENGTH] + '\n<HINWEIS: NACHRICHT WURDE AUF 6000 ZEICHEN GEKÜRZT! VOLLSTÄNDIGE NACHRICHT SIEHE GLEICHNAMIGES MSG-FILE>'

            # Body ausgeben
            pdf.set_font("Segoeb", size=10)
            pdf.write(5, f"\nInhalt:\n")
            pdf.set_font("Segoe", size=8)

            # Text bereinigen, damit kompakte Darstellung möglich
            cleaned_text = clean_email_text(msg_body)  # <- dein ausgelesener Text
            logging.debug(f"msg_body nach Bereinigung: {cleaned_text}")  # Debugging-Ausgabe: Log-File

            # Text ausgeben
            pdf.write(5, f"{cleaned_text}")

        else:
            pdf.set_font("Segoeb", size=10)
            pdf.write(5, f"\nInhalt:\n")
            pdf.set_font("Segoe", size=8)
            pdf.write(5, f"\nHINWEIS: NACHRICHT OHNE INHALT ODER KANN NICHT GELESEN WERDEN (z.B. SIGNATURPRÜFUNG)!")

        # Anhänge ausgeben
        if not MsgAccessStatus.ATTACHMENTS_MISSING in msg_object["status"]:
            msg_attachments = msg_object["attachments"]
            pdf.set_font("Segoeb", size=8)
            pdf.write(5, f"\n\nAnhänge:\n")
            pdf.set_font("Segoe", size=8)

            for filename in msg_object["attachments"]:
               pdf.write(5, f"- {filename}\n")

        # Speichern der PDF-Datei
        pdf.output(pdf_path_and_filename)

        print(f"\tDie PDF wurde unter '{os.path.abspath(pdf_path_and_filename)}' gespeichert.")
