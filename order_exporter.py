import os
import logging
import html
import hashlib
import re
from datetime import datetime
import xml.etree.ElementTree as ET

# Импортируем функцию сопоставления из нашего нового парсера
from nomenclature_parser import find_best_match


def validate_xml(xml_text: str) -> bool:
    """Проверяет, является ли строка корректным XML."""
    try:
        ET.fromstring(xml_text)  #
        return True
    except ET.ParseError as e:
        logging.error(f"Ошибка валидации XML: {e}")
        return False


def save_order_xml(xml_text: str, folder: str, base_filename: str, overwrite: bool = False):
    """
    Сохраняет XML-текст в файл, избегая перезаписи, если overwrite=False.

    """
    if not os.path.exists(folder):
        try:
            os.makedirs(folder)
            logging.info(f"Создана папка {folder}")
        except Exception as e:
            logging.error(f"Не удалось создать папку {folder}: {e}")
            return

    save_path = os.path.join(folder, base_filename)

    # Если перезапись запрещена и файл существует, генерируем новое имя
    if not overwrite and os.path.exists(save_path):
        index = 1
        while True:
            # Имя_файла_1.xml, Имя_файла_2.xml и т.д.
            new_filename = f"{os.path.splitext(base_filename)[0]}_{index}{os.path.splitext(base_filename)[1]}"
            new_path = os.path.join(folder, new_filename)
            if not os.path.exists(new_path):
                save_path = new_path
                break
            index += 1

    try:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(xml_text)
        logging.info(f"XML-запрос сохранён в файл: {save_path}")
    except Exception as e:
        logging.error(f"Ошибка сохранения XML: {e}")


def is_duplicate_order(xml_text: str, archive_folder: str) -> bool:
    """
    Проверяет заказ на дублирование по хешу XML и создает маркер.

    """
    order_hash = hashlib.md5(xml_text.encode("utf-8")).hexdigest()
    marker_path = os.path.join(archive_folder, f"order_{order_hash}.marker")

    if os.path.exists(marker_path):
        logging.warning("Обнаружен дубликат заказа (найден хеш-маркер).")
        return True
    else:
        try:
            # Создаем пустой файл-маркер
            with open(marker_path, "w") as f:
                f.write(datetime.now().isoformat())  # Записываем время, чтобы знать, когда он был создан
            logging.info(f"Создан маркер заказа: {marker_path}")
        except Exception as e:
            logging.error(f"Ошибка создания маркера заказа: {e}")
        return False


def cleanup_docs_folder_if_markers_exist(folder: str):
    """
    Очищает папку 'docs' (C:\1s\docs), если в ней есть маркеры обработки .processed.

    """
    marker_found = False
    try:
        files = os.listdir(folder)
    except FileNotFoundError:
        logging.warning(f"Папка {folder} не найдена для очистки.")
        return

    for filename in files:
        if ".processed." in filename:
            marker_found = True
            break

    if marker_found:
        logging.info("Найдены маркеры .processed. Очищаю папку от XML-файлов...")
        for filename in files:
            if filename.endswith(".xml"):
                file_path = os.path.join(folder, filename)
                try:
                    os.remove(file_path)
                    logging.info(f"Удалён обработанный файл: {file_path}")
                except Exception as e:
                    logging.error(f"Ошибка удаления файла {file_path}: {e}")
        logging.info("Папка очищена от обработанных (помеченных) файлов.")


def extract_phone_number(phone_str: str) -> str:
    """Простая очистка телефонного номера."""
    matches = re.findall(r'(\+?\d[\d\-\(\)\s]{5,}\d)', phone_str)
    if matches:
        return matches[0].strip()
    return phone_str.strip()


def generate_order_xml(order_data: dict, config: dict):
    """
    Формирует XML-файл заказа, сопоставляя каждую позицию с номенклатурой.

    """
    products = order_data.get("order", {}).get("products", [])
    if not products:
        logging.error("Нет товаров для формирования XML. Заказ не будет создан.")
        return

    company = order_data.get("company", {})
    contact = order_data.get("order", {}).get("contact_person", {})
    phone = extract_phone_number(contact.get("phone", ""))

    xml = f"""<Заказ>
  <Контрагент>
    <ИНН>{company.get("INN", "")}</ИНН>
    <КПП>{company.get("KPP", "")}</КПП>
    <НаименованиеПолное>{html.escape(company.get("name", ""))}</НаименованиеПолное>
    <НаименованиеКраткое>{html.escape(company.get("name", ""))}</НаименованиеКраткое>
    <ЮрАдрес>{html.escape(company.get("legal_address", ""))}</ЮрАдрес>
    <ФактАдрес>{html.escape(company.get("actual_address", ""))}</ФактАдрес>
    <Телефон>{html.escape(phone)}</Телефон>
    <Email>{html.escape(contact.get("email", ""))}</Email>
    <КонтактноеЛицо>{html.escape(contact.get("full_name", ""))}</КонтактноеЛицо>
  </Контрагент>
  <Ответственный>{html.escape(contact.get("full_name", ""))}</Ответственный>
  <Товары>
"""
    for prod in products:
        prod_name = prod.get("full_name", prod.get("name", ""))
        prod_quantity = prod.get("quantity", 0)

        # --- БЛОК ПОИСКА ПО НОМЕНКЛАТУРЕ ---
        # Используем функцию из импортированного модуля
        matched_item = find_best_match(prod_name)

        final_code = ""
        final_name = prod_name  # По умолчанию используем оригинальное имя

        if matched_item:
            # Если совпадение найдено, используем данные из номенклатуры
            # 'Код' должен быть без '\ufeff' благодаря 'utf-8-sig' в load_nomenclature
            final_code = matched_item.get('Код', '')
            final_name = matched_item.get("Полное наименование", prod_name)
        else:
            # Если совпадение не найдено, оставляем код пустым
            pass  # final_code и так "", final_name и так оригинальное

        xml += f"""    <Товар>
          <Код>{html.escape(str(final_code))}</Код>
          <ПолноеНаименование>{html.escape(final_name)}</ПолноеНаименование>
          <Количество>{prod_quantity}</Количество>
        </Товар>
"""
    xml += "  </Товары>\n</Заказ>"
    logging.info("Сформированный XML заказа:\n%s", xml)

    if not validate_xml(xml):
        logging.error("XML невалиден – остановка обработки заказа.")
        return

    order_folder = config.get("ORDER_XML_FOLDER", r"C:\1s\docs")
    if not os.path.exists(order_folder):
        try:
            os.makedirs(order_folder)
            logging.info(f"Создана папка {order_folder}")
        except Exception as e:
            logging.error(f"Не удалось создать папку {order_folder}: {e}")
            return

    # Проверяем, есть ли маркеры .processed.
    cleanup_docs_folder_if_markers_exist(order_folder)

    # Определяем, нужно ли перезаписывать zakaz.xml
    # Логика из main.py: если папка пуста, то перезаписываем.
    target_overwrite = True if not os.listdir(order_folder) else False
    save_order_xml(xml, folder=order_folder, base_filename="zakaz.xml", overwrite=target_overwrite)

    # --- Архивирование ---
    archive_folder = config["ARCHIVE_FOLDER"]
    if not os.path.exists(archive_folder):
        try:
            os.makedirs(archive_folder)
        except Exception as e:
            logging.error(f"Ошибка создания архива: {e}")

    if is_duplicate_order(xml, archive_folder):
        logging.warning("Дублирование заказа обнаружено – заказ не сохраняется повторно в архив.")
    else:
        # Сохраняем в архив с датой
        today_date = datetime.now().strftime("%d.%m.%Y_%H%M%S")
        date_filename = f"заказ_{today_date}.xml"
        # В архиве никогда не перезаписываем, а создаем _1, _2 и т.д.
        save_order_xml(xml, folder=archive_folder, base_filename=date_filename, overwrite=False)