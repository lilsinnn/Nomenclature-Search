import io
import logging
import imaplib
import email
from email.header import decode_header

import pytesseract
from PIL import Image
from bs4 import BeautifulSoup
from PyPDF2 import PdfReader
import docx2txt
import openpyxl

# --- Настройка OCR ---
# Выносим настройку Tesseract сюда.
# ВНИМАНИЕ: Этот путь был жестко задан в main.py.
# В будущем его лучше вынести в config.json или убедиться, что Tesseract в системном PATH.
try:
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  #
    logging.info(f"Путь к Tesseract установлен: {pytesseract.pytesseract.tesseract_cmd}")
except Exception as e:
    logging.warning(f"Не удалось указать путь к Tesseract 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'. "
                    f"Убедитесь, что Tesseract в системном PATH или путь указан верно. Ошибка: {e}")


def _read_attachment(part, decoded_filename: str) -> str:
    """Вспомогательная функция для чтения содержимого одного вложения."""
    payload = part.get_payload(decode=True)
    if not payload:
        logging.warning(f"Вложение {decoded_filename} не имеет данных (пустое).")
        return ""

    text = ""
    lower_filename = decoded_filename.lower()

    try:
        if lower_filename.endswith((".txt", ".csv")):
            text = payload.decode("utf-8-sig", errors="ignore")

        elif lower_filename.endswith(".pdf"):
            pdf_text = "".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(payload)).pages)  #
            text = pdf_text

        elif lower_filename.endswith(".docx"):
            text = docx2txt.process(io.BytesIO(payload))  #

        elif lower_filename.endswith(".xlsx"):
            workbook = openpyxl.load_workbook(io.BytesIO(payload), data_only=True)  #
            for sheet in workbook.worksheets:
                text += f"\nЛист: {sheet.title}\n"
                for row in sheet.iter_rows(values_only=True):
                    text += "\t".join([str(cell) if cell is not None else "" for cell in row]) + "\n"

        elif lower_filename.endswith((".png", ".jpg", "jpeg")):
            ocr_text = pytesseract.image_to_string(Image.open(io.BytesIO(payload)), lang='rus+eng')  #
            text = f"[Распознанный текст с изображения]:\n{ocr_text}"

        else:
            text = f"[Формат файла '{decoded_filename}' не поддерживается для чтения]"
            logging.warning(f"Вложение '{decoded_filename}' имеет неподдерживаемый тип.")

    except Exception as e:
        logging.error(f"Не удалось прочитать вложение {decoded_filename}: {e}", exc_info=True)
        text = f"[ОШИБКА ЧТЕНИЯ ФАЙЛА {decoded_filename}]"

    return text


def get_email_text_with_attachments(config: dict) -> tuple:
    """
    Подключается к IMAP, извлекает последнее письмо, читает тело и все вложения.
    Возвравращает (combined_text, msg_object) или ("", None) в случае неудачи.
    """
    logging.info("Подключаюсь к IMAP для получения последнего письма...")
    try:
        mail = imaplib.IMAP4_SSL(config["IMAP_SERVER"])  #
        mail.login(config["MAIL_USER"], config["MAIL_PASSWORD"])  #
        mail.select("INBOX")

        # Используем ту же логику, что и в main.py: ищем ВСЕ письма
        _, data = mail.search(None, "ALL")  #
        mail_ids = data[0].split()

        if not mail_ids:
            mail.logout()
            logging.info("Почтовый ящик пуст.")
            return "", None

        # Берем самое последнее письмо
        latest_id = mail_ids[-1]  #
        _, msg_data = mail.fetch(latest_id, '(RFC822)')  #
        mail.logout()

    except Exception as e:
        logging.error(f"Ошибка подключения или чтения почты: {e}", exc_info=True)
        return "", None

    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)  #

    main_text_plain = ""
    main_text_html = ""
    attachments_text = ""

    for part in msg.walk():  #
        if part.get_content_type() == "multipart/alternative":
            continue

        filename = part.get_filename()

        # --- 1. ОБРАБОТКА ВЛОЖЕНИЙ ---
        if filename:
            decoded_filename = filename
            try:
                # Корректная обработка имен файлов в любой кодировке
                decoded_header = decode_header(filename)[0]  #
                if isinstance(decoded_header[0], bytes):
                    decoded_filename = decoded_header[0].decode(decoded_header[1] or "utf-8", errors="ignore")
            except Exception as e:
                logging.warning(f"Не удалось декодировать имя файла {filename}: {e}")

            attachments_text += f"\n\n--- СОДЕРЖИМОЕ ВЛОЖЕНИЯ: {decoded_filename} ---\n"
            attachments_text += _read_attachment(part, decoded_filename)

        # --- 2. ОБРАБОТКА ТЕЛА ПИСЬМА ---
        content_type = part.get_content_type()
        if not filename and content_type.startswith("text/"):
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)  #
            if not payload:
                continue

            try:
                body_part_text = payload.decode(charset, errors="ignore")
                if content_type == "text/plain":
                    main_text_plain += body_part_text + "\n"
                elif content_type == "text/html":
                    main_text_html += body_part_text + "\n"
            except Exception as e:
                logging.error(f"Ошибка декодирования тела письма (тип: {content_type}): {e}")

    # --- 3. ФОРМИРОВАНИЕ ИТОГОВОГО ТЕКСТА ---
    final_main_text = main_text_plain.strip()
    if not final_main_text and main_text_html:
        logging.info("Plain-text версия не найдена, используется HTML-версия.")
        soup = BeautifulSoup(main_text_html, "html.parser")  #
        final_main_text = soup.get_text(separator="\n", strip=True)

    combined_text = f"=== ТЕКСТ ПИСЬМА ===\n{final_main_text}\n\n=== ВЛОЖЕНИЯ ===\n{attachments_text.strip()}"

    # Обрезаем слишком длинный текст, чтобы не перегружать GPT
    max_length = 10000
    if len(combined_text) > max_length:
        logging.warning(f"Текст письма и вложений превышает {max_length} символов. Обрезаю...")
        combined_text = combined_text[:max_length] + "\n...[текст обрезан из-за превышения лимита]"

    return combined_text, msg