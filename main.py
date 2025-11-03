#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import io
import json
import base64
import logging
import imaplib
import email
import html
import re
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
import time
import pytesseract
from PIL import Image
from email.header import decode_header
import subprocess
import zipfile
import tempfile
import shutil
import csv
import hashlib
from bs4 import BeautifulSoup
from PyPDF2 import PdfReader
import docx2txt
import openpyxl


#---------------------
#GLOBALS
#---------------------
last_email_hash = None
synonyms_type = None
logs = 1 #1 - логи полноценные, 0 - без

# -----------------------------
# Настройка логирования и OCR
# -----------------------------
import logging
import os
from datetime import datetime

# --- Настройка конфига ---
def load_config(path="config.json") -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)

config = load_config()

#Загрузка путей
regex = config["REGEX_PATH"]
parameters = config["PARAMETERS_PATH"]
synonyms_data = config["SYNONYMS_PATH"]

# Уровень логирования
# для детальной отладки парсинга лучше поставить DEBUG
log_level = logging.DEBUG

# Формат сообщений (добавил имя файла и строку для удобства отладки)
log_format = "%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s"
formatter = logging.Formatter(log_format)

# ---- НАСТРОЙКА ВЫВОДА В КОНСОЛЬ ----
logger = logging.getLogger()
logger.setLevel(log_level)



# Обработчик для вывода в консоль (терминал)
console_handler = logging.StreamHandler()
console_handler.setLevel(log_level)  # Уровень для консоли
console_handler.setFormatter(formatter)  # Формат для консоли
if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):  # Добавляем, если еще нет
    logger.addHandler(console_handler)

# ---- НАСТРОЙКА ВЫВОДА В ФАЙЛ ----
log_directory = config["LOGS_FOLDER"]  # Путь к папке логов (raw string для Windows)
log_file_name = f"order_processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"  # Имя файла с датой и временем
log_file_path = os.path.join(log_directory, log_file_name)

try:
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
        print(
            f"INFO: Создана папка для логов: {log_directory}")  # Используем print, т.к. логгер может быть еще не готов
except OSError as e:
    print(f"ERROR: Не удалось создать папку для логов {log_directory}: {e}")

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    logging.warning("Библиотека pywin32 не найдена. Обработка .doc файлов будет недоступна.")

try:
    with open(regex, 'r', encoding='utf-8') as f:
        REGEX_PATTERNS = json.load(f)
    logging.info("Словарь регулярных выражений успешно загружен.")
except Exception as e:
    logging.error(f"Не удалось загрузить regex_patterns.json: {e}")
    REGEX_PATTERNS = {} # Создаем пустой словарь, чтобы скрипт не упал

try:
    file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    if not any(isinstance(h, logging.FileHandler) and h.baseFilename == file_handler.baseFilename for h in
               logger.handlers):
        logger.addHandler(file_handler)

    logging.info(f"Логирование в файл настроено: {log_file_path}")

except Exception as e:
    logging.error(f"Ошибка настройки логирования в файл {log_file_path}: {e}", exc_info=True)

# --- Остальные настройки OCR и т.д. ---
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"




try:
    with open(synonyms_data, 'r', encoding='utf-8') as f:
        synonyms_type = json.load(f)
    if synonyms_type:
        print("Словарь синонимов успешно загружен:")
        if "тройник" in synonyms_type:
            print(f"Синонимы для 'тройник': {synonyms_type['тройник']}")
    else:
        print("Не удалось загрузить данные, переменная synonyms_type пуста.")

except FileNotFoundError:
    print(f"Ошибка: Файл '{synonyms_data}' не найден. Убедитесь, что файл существует и путь указан верно.")
except json.JSONDecodeError:
    print(f"Ошибка: Не удалось декодировать JSON из файла '{synonyms_data}'. Проверьте корректность формата JSON в файле.")
except Exception as e:
    print(f"Произошла непредвиденная ошибка: {e}")


# Вставьте этот блок после блока "GLOBALS" или перед функцией load_config

with open(parameters, 'r', encoding='utf-8') as f:
    PART_SPECIFICATIONS = json.load(f)

material_aliases = {
    "ст3": ["ст3", "ст.3", "сталь 3", "ст3сп", "ст3пс", "ст3кп", "s235jr", "s235", "st37-2", "q235", "a36"],
    "ст10": ["ст10", "сталь 10", "ст.10", "10кп", "10пс", "c10e", "ck10", "1010", "s10c"],
    "ст20": ["ст20", "сталь 20", "ст.20", "20кп", "20пс", "c22e", "ck22", "1020", "s20c"],
    "ст35": ["ст35", "сталь 35", "ст.35", "c35e", "ck35", "1035", "s35c"],
    "ст45": ["ст45", "сталь 45", "ст.45", "c45e", "ck45", "1045", "s45c"],
    "09г2с": ["09г2с", "09г2c", "09g2s", "s355j2", "s355", "st52-3", "q345", "16mn"],
    "17г1с": ["17г1с", "17гс", "17г1су", "s355j0", "s355", "st52-3"],
    "10г2": ["10г2"],
    "13хфа": ["13хфа", "13хф"],
    "40х": ["40х", "40cr", "5140", "scr440"],
    "30хгса": ["30хгса", "хромансиль", "30chgsa", "30hgsa"],
    "12х18н10т": ["12х18н10т", "12x18h10t", "08х18н10т", "aisi 321", "aisi321"],
    "aisi 304": ["aisi 304", "aisi304", "08х18н10"],
    "aisi 316": ["aisi 316", "aisi316", "08х17н13м2", "10х17н13м2"],
}
default_material = "ст20"

def normalize_string(s):
    """Приводит строку к единому, чистому виду для сравнения."""
    if not isinstance(s, str): return ""
    s = s.lower()
    s = s.replace('гр.', 'градусов').replace('гр ', 'градусов ')
    s = s.replace(',', '.')
    s = s.replace('х', 'x')
    s = re.sub(r'[\s\-_/]+', ' ', s).strip()
    return s


def normalize_text(s: str) -> str:
    s = s.lower().strip()
    s = s.replace('х', 'x').replace('×', 'x')
    s = s.replace(',', '.')
    s = re.sub(r'\bгр\.?\b', ' градусов ', s)
    s = re.sub(r'\bвып\.?\b', ' выпуск ', s)
    s = re.sub(r'\bсер\.?\b', ' серия ', s)
    s = re.sub(r'\bд[уy]\b', ' dn ', s)
    s = re.sub(r'\bр[уy]\b', ' pn ', s)
    s = re.sub(r'\bмат\.?\b', ' материал ', s)
    s = re.sub(r'\bшт\.?\b', ' штук ', s)
    s = re.sub(r'\bкомпл\.?\b', ' комплект ', s)


    s = re.sub(r'\b(гост|gost|ост|ost|ту|тс|asme|din|en|iso)\s*[\d\.\-]+(?:[\s\-]+[\d\.\-]+)*', ' ', s)
    s = re.sub(r'\b(сталь|ст\.?|steel|aisi|mat)\s*[\w\d\.\-]+', ' ', s)
    s = re.sub(r'[^0-9a-zа-яё.\-\s/]', ' ', s)

    # 7) Убираем лишние пробелы
    return re.sub(r'\s+', ' ', s).strip()


# --- 1) Изменяем normalize_text, чтобы он не вырезал точку (.) из размеров:

def normalize_for_parsing(text: str) -> str:
    """
    Специальная нормализация для парсинга: замена кириллицы на латиницу (транслит),
    удаление лишних символов, чтобы улучшить работу regex.
    """
    if not text:
        return ""
    text = text.lower()
    # Таблица транслитерации (дополните по необходимости)
    translit_table = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'zh',
        'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
        'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts',
        'ч': 'ch', 'ш': 'sh', 'щ': 'shch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
        # Дополнительные символы и частые замены, если нужно
    }
    # Применяем транслитерацию и замены
    normalized_text = ""
    for char in text:
        normalized_text += translit_table.get(char, char)
    normalized_text = re.sub(r'[^\w\s\.\-xх,]', '', normalized_text)
    normalized_text = re.sub(r'\s+', ' ', normalized_text).strip()
    return normalized_text


def extract_numbers(s: str) -> list[str]:
    return re.findall(r'\d+\.?\d*', normalize_text(s))


def load_nomenclature(file_path: str) -> list[dict]:
    """
    Загружает номенклатуру из текстового файла с разделителями-табуляторами.
    Первая строка считается заголовком.
    """
    if not os.path.exists(file_path):
        logging.error(f"Файл номенклатуры не найден по пути: {file_path}")
        return []

    nomenclature_data = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter='\t')

            try:
                header = next(reader)
                logging.info(f"Заголовки номенклатуры: {header}")
            except StopIteration:
                logging.error("Файл номенклатуры пуст.")
                return []

            for row in reader:
                if not row:
                    continue
                row_data = dict(zip(header, row))
                nomenclature_data.append(row_data)

        logging.info(f"Успешно загружено {len(nomenclature_data)} позиций из номенклатуры.")
        return nomenclature_data

    except Exception as e:
        logging.error(f"Ошибка при чтении файла номенклатуры {file_path}: {e}", exc_info=True)
        return []


def extract_key_features(name: str) -> dict:
    """
    ФИНАЛЬНАЯ ВЕРСИЯ. Максимально точное извлечение всех параметров.
    """
    features = {}
    processed_name = name.replace(',', '.').lower()

    dims_dxs = re.findall(r'(\d+\.?\d*)\s*[xх*]\s*(\d+\.?\d*)', processed_name)
    if dims_dxs:
        features['dims'] = sorted([f"{float(d[0]):g}x{float(d[1]):g}" for d in dims_dxs])

    dims_dnpn = re.findall(r'\b(\d+)\s*-\s*(\d+\.?\d*)\b', processed_name)
    if dims_dnpn:
        features['dims_dn_pn'] = sorted([f"{d[0]}-{d[1]}" for d in dims_dnpn])

    angle_match = re.search(r'\b(15|30|45|60|90)\b', processed_name)
    if angle_match and not re.search(r'\d' + re.escape(angle_match.group(0)), processed_name):
        features['angle'] = angle_match.group(1)

    mat_match = re.search(r'\b(ст\.?[\s\d\w\-]+|09г2с|12х18н10т|13хфа|aisi[\s\d]+)\b', processed_name, re.IGNORECASE)
    if mat_match:
        material_str = mat_match.group(0)
        potential_angle = re.sub(r'ст\.?', '', material_str, flags=re.IGNORECASE).strip()
        if not (potential_angle in ['90', '45', '60'] and f"{potential_angle}-" in processed_name):
            features['material'] = re.sub(r'[\s.]', '', material_str)

    std_match = re.search(r'\b(гост\s*[\d\-]+|ту\s*[\d\-.\s]+|атк\s*[\d.\-]+)\b', processed_name, re.IGNORECASE)
    if std_match:
        features['standard'] = " ".join(std_match.group(0).split())

    exec_match = re.search(r'-\b([a-f])\b-', processed_name, re.IGNORECASE)
    if exec_match:
        features['execution_seal'] = exec_match.group(1)

    return features


def find_best_match(order_product_name: str, nomenclature_data: list[dict]) -> dict | None:
    """
    ФИНАЛЬНАЯ ГИБРИДНАЯ ВЕРСИЯ.
    Использует строгую фильтрацию по размерам для максимальной точности.
    """
    parsed_order_info = parse_order_name(order_product_name)

    order_type = parsed_order_info.get("type")
    order_params = parsed_order_info.get("params", {})
    order_dimensions = order_params.get("dimensions")

    if not order_type:
        logging.warning(f"Не удалось определить тип для '{order_product_name}', поиск невозможен.")
        return None

    logging.debug(f"ИЩУ '{order_product_name}' по параметрам: {order_params}")


    filter_keyword = synonyms_type.get(order_type, [order_type])[0].lower()
    candidates = [
        item for item in nomenclature_data
        if filter_keyword in item.get('Полное наименование', '').lower()
    ]

    if not candidates:
        return None

    strict_candidates = []
    if order_dimensions:
        for item in candidates:
            item_info = parse_order_name(item.get('Полное наименование', ''))
            item_dimensions = item_info.get("params", {}).get("dimensions")

            if order_dimensions == item_dimensions:
                strict_candidates.append(item)

        if not strict_candidates:
            logging.warning(
                f"Для '{order_product_name}' не найдено ни одного товара с точным совпадением размеров '{order_dimensions}'.")
            return None  # Если нет совпадения по размерам - совпадения нет вообще

        candidates = strict_candidates
        logging.debug(f"После фильтра по размерам '{order_dimensions}' осталось кандидатов: {len(candidates)}")

    best_score = -1
    best_match = None

    for item in candidates:
        current_score = 0
        item_info = parse_order_name(item.get('Полное наименование', ''))
        item_params = item_info.get("params", {})

        for key, order_value in order_params.items():
            if key != 'dimensions' and item_params.get(key) == order_value:
                current_score += 1

        if current_score > best_score:
            best_score = current_score
            best_match = item

    if best_match:
        logging.debug(
            f"ИТОГ: Для '{order_product_name}' выбран товар '{best_match.get('Полное наименование')}' с финальным счетом {best_score}")

    return best_match

def extract_key_features(name: str) -> dict:
    """
    Улучшенная версия. Более гибко извлекает размеры (в т.ч. с плавающей точкой),
    сталь, стандарт и другие ключевые параметры.
    """
    features = {}
    processed_name = name.lower().replace(',', '.')

    dims_dxs = re.findall(r'(\d+\.?\d*)\s*[xх]\s*(\d+\.?\d*)', processed_name)
    if dims_dxs:
        features['dims'] = sorted([f"{float(d[0]):g}x{float(d[1]):g}" for d in dims_dxs])

    if not dims_dxs:
        dims_dnpn = re.findall(r'\b(\d+)\s*-\s*(\d+\.?\d*)\b', processed_name)
        if dims_dnpn:
            features['dims_dn_pn'] = sorted([f"{d[0]}-{d[1]}" for d in dims_dnpn if len(d[0]) < 5])

    angle_match = re.search(r'\b(15|30|45|60|90)\b', processed_name)
    if angle_match:
        features['angle'] = angle_match.group(1)

    normalized_material = None
    for canonical_name, aliases in material_aliases.items():
        for alias in aliases:
            if re.search(rf'\b{re.escape(alias)}\b', processed_name):
                normalized_material = canonical_name
                break
        if normalized_material:
            break
    if normalized_material:
        features['material'] = normalized_material
    else:
        if re.search(r'\b(20)\b', processed_name) and not re.search(r'гост', processed_name):
             features['material'] = 'ст20'

    std_match = re.search(r'\b(гост\s*[\d.\-]+|ту\s*[\d.\s\-]+|атк\s*[\d.\-]+)\b', processed_name, re.IGNORECASE)
    if std_match:
        features['standard'] = " ".join(std_match.group(0).split())

    exec_match = re.search(r'-\b([a-f\d])\b-', processed_name, re.IGNORECASE)
    if exec_match:
        features['execution_seal'] = exec_match.group(1)

    return features


import logging
import re
def clean_yandex_gpt_json_response(raw_response_text: str) -> str:
    """
    Очищает сырой текстовый ответ от YandexGPT, удаляя Markdown-обрамление
    для JSON (```json ... ``` или ``` ... ```).
    Возвращает строку, готовую для json.loads(), или исходную строку, если очистка не требуется.
    """
    cleaned_text = raw_response_text.strip()

    if cleaned_text.startswith("```json"):
        cleaned_text = cleaned_text[len("```json"):].strip()
        logging.debug("Удалено начальное '```json' из ответа YandexGPT.")
    elif cleaned_text.startswith("```"):
        cleaned_text = cleaned_text[len("```"):].strip()
        logging.debug("Удалено начальное '```' из ответа YandexGPT.")

    if cleaned_text.endswith("```"):
        cleaned_text = cleaned_text[:-len("```")].strip()
        logging.debug("Удалено конечное '```' из ответа YandexGPT.")

    return cleaned_text



def parse_order_name(order_name: str) -> dict:
    """
    ФИНАЛЬНАЯ ГИБРИДНАЯ ФУНКЦИЯ ПАРСИНГА.
    Сочетает точный поиск по regex-правилам и гибкий поиск по словарю параметров.
    """
    if logs:
        logging.debug(f"Начало гибридного парсинга строки: '{order_name}'")

    # --- Шаг 1: Определение типа детали (без изменений) ---
    item_type = None
    work_string = f" {order_name.lower().replace(',', '.')} "

    global synonyms_type, REGEX_PATTERNS, PART_SPECIFICATIONS

    product_types = {k: v for k, v in synonyms_type.items() if k.lower() != "комментарий"}
    sorted_keys = sorted(product_types.keys(), key=len, reverse=True)

    for t in sorted_keys:
        for synonym in product_types[t]:
            if f" {synonym.lower()} " in work_string:
                item_type = t
                work_string = work_string.replace(f" {synonym.lower()} ", " ", 1)
                break
        if item_type:
            break

    if not item_type:
        if logs:
            logging.warning(f"Не удалось определить тип для: '{order_name}'")
        return {"original_name": order_name, "type": None, "params": {}}

    # --- Инициализация словарей для сбора параметров ---
    parsed_data = {"original_name": order_name, "type": item_type, "params": {}}
    found_params = {}  # <--- Вот словарь, которого не хватало

    # --- Шаг 2: Точный поиск по regex-правилам из 'regex_patterns.json' ---
    patterns_for_type = REGEX_PATTERNS.get(item_type, {})
    for param_name, patterns_list in patterns_for_type.items():
        for pattern in patterns_list:
            match = re.search(pattern, work_string)
            if match:
                found_groups = list(filter(None, match.groups()))
                separator = 'x' if 'x' in match.group(0).lower() else '-'
                value = separator.join(found_groups) if len(found_groups) > 1 else (
                    found_groups[0] if found_groups else match.group(0))

                # Сразу кладем в основной словарь params
                parsed_data["params"][param_name] = value.strip()
                work_string = re.sub(re.escape(match.group(0)), ' ', work_string, 1)
                break

    specs = PART_SPECIFICATIONS.get(item_type, {})
    if specs:
        props_to_find = [key for key in specs.keys() if key not in parsed_data["params"]]

        for prop_key in sorted(props_to_find, key=len, reverse=True):
            prop_config = specs[prop_key]
            possible_values = prop_config.get("values", [])

            for value in sorted(possible_values, key=len, reverse=True):
                if re.search(rf'\b{re.escape(str(value).lower())}\b', work_string):
                    found_params[prop_key] = str(value)
                    work_string = re.sub(rf'\b{re.escape(str(value).lower())}\b', ' ', work_string, count=1)
                    break

    if specs:
        for prop_key in specs.keys():
            if prop_key not in parsed_data["params"] and prop_key not in found_params:
                default_value = specs[prop_key].get("default")
                if default_value is not None:
                    found_params[prop_key] = str(default_value)

    parsed_data["params"].update(found_params)

    if logs:
        logging.info(f"Результат гибридного парсинга: {json.dumps(parsed_data, ensure_ascii=False, indent=2)}")

    return parsed_data


# =============================
# Функции для работы с файлами и сохранения XML
# =============================
def save_order_xml(xml_text: str, folder, base_filename, overwrite=False):
    if not os.path.exists(folder):
        try:
            os.makedirs(folder)
            logging.info(f"Создана папка {folder}")
        except Exception as e:
            logging.error(f"Не удалось создать папку {folder}: {e}")
            return
    save_path = os.path.join(folder, base_filename)
    if not overwrite and os.path.exists(save_path):
        index = 1
        while True:
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

def validate_xml(xml_text: str) -> bool:
    try:
        ET.fromstring(xml_text)
        return True
    except ET.ParseError as e:
        logging.error(f"Ошибка валидации XML: {e}")
        return False

def is_duplicate_order(xml_text: str, archive_folder: str) -> bool:
    order_hash = hashlib.md5(xml_text.encode("utf-8")).hexdigest()
    marker_path = os.path.join(archive_folder, f"order_{order_hash}.marker")
    if os.path.exists(marker_path):
        return True
    else:
        try:
            with open(marker_path, "w") as f:
                f.write("")
        except Exception as e:
            logging.error(f"Ошибка создания маркера заказа: {e}")
        return False


def generate_order_xml(order_data: dict, config: dict, nomenclature_data: list) -> str:
    """
    Формирует XML-файл заказа, предварительно находя каждую позицию в номенклатуре.
    """
    products = order_data.get("order", {}).get("products", [])
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
        matched_item = find_best_match(prod_name, nomenclature_data)

        final_code = ""
        final_name = prod_name  # По умолчанию используем оригинальное имя

        if matched_item:
            # Если совпадение найдено, используем данные из номенклатуры
            # Примечание: '\ufeffКод' - это исправление для возможной проблемы с кодировкой файла (BOM)
            final_code = matched_item.get('\ufeffКод') or matched_item.get('Код', '')
            final_name = matched_item.get("Полное наименование", prod_name)
            logging.info(f"✅ НАЙДЕНО: Для '{prod_name}' -> '{final_name}' (Код: {final_code})")
        else:
            # Если совпадение не найдено, оставляем код пустым и используем исходное имя
            logging.warning(f"⚠️ НЕ НАЙДЕНО: Для '{prod_name}'. Позиция будет добавлена с оригинальным наименованием.")

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
        return xml


    order_folder = config.get("ORDER_XML_FOLDER", r"C:\1s\docs")
    if not os.path.exists(order_folder):
        try:
            os.makedirs(order_folder)
            logging.info(f"Создана папка {order_folder}")
        except Exception as e:
            logging.error(f"Не удалось создать папку {order_folder}: {e}")
            return xml

    if os.listdir(order_folder):
        cleanup_docs_folder_if_markers_exist(order_folder)

    target_overwrite = True if not os.listdir(order_folder) else False
    save_order_xml(xml, folder=order_folder, base_filename="zakaz.xml", overwrite=target_overwrite)

    archive_folder = config["ARCHIVE_FOLDER"]
    if not os.path.exists(archive_folder):
        try:
            os.makedirs(archive_folder)
        except Exception as e:
            logging.error(f"Ошибка создания архива: {e}")
    if is_duplicate_order(xml, archive_folder):
        logging.warning("Дублирование заказа обнаружено – заказ не сохраняется повторно в архив.")
    else:
        today_date = datetime.now().strftime("%d.%m")
        date_filename = f"заказ{today_date}.xml"
        save_order_xml(xml, folder=archive_folder, base_filename=date_filename, overwrite=True)
    return xml


def cleanup_docs_folder_if_markers_exist(folder):
    marker_found = False
    for filename in os.listdir(folder):
        if ".processed." in filename:
            marker_found = True
            break
    if marker_found:
        for filename in os.listdir(folder):
            if filename.endswith(".xml"):
                file_path = os.path.join(folder, filename)
                try:
                    os.remove(file_path)
                    logging.info(f"Удалён файл: {file_path}")
                except Exception as e:
                    logging.error(f"Ошибка удаления файла {file_path}: {e}")
        logging.info("Папка очищена от обработанных (помеченных) файлов.")

def extract_phone_number(phone_str: str) -> str:
    matches = re.findall(r'(\+?\d[\d\-\(\)\s]{5,}\d)', phone_str)
    if matches:
        return matches[0].strip()
    return phone_str.strip()

# =============================
# Функции для обработки писем (IMAP)
# =============================
def get_email_text_with_attachments(config: dict) -> tuple:
    """
    Финальная версия: корректно читает и тело письма, и вложения.
    """
    logging.info("Подключаюсь к IMAP для получения последнего письма...")
    try:
        mail = imaplib.IMAP4_SSL(config["IMAP_SERVER"])
        mail.login(config["MAIL_USER"], config["MAIL_PASSWORD"])
        mail.select("INBOX")
        _, data = mail.search(None, "ALL")
        mail_ids = data[0].split()

        if not mail_ids:
            mail.logout()
            logging.info("Почтовый ящик пуст.")
            return "", None

        latest_id = mail_ids[-1]
        _, msg_data = mail.fetch(latest_id, '(RFC822)')
        mail.logout()
    except Exception as e:
        logging.error(f"Ошибка подключения или чтения почты: {e}", exc_info=True)
        return "", None

    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    main_text_plain = ""
    main_text_html = ""
    attachments_text = ""

    for part in msg.walk():
        # Пропускаем только контейнеры, остальное обрабатываем по логике ниже
        if part.get_content_type() == "multipart/alternative":
            continue

        filename = part.get_filename()

        # --- 1. ОБРАБОТКА ВЛОЖЕНИЙ ---
        if filename:
            try:
                decoded_filename, enc = decode_header(filename)[0]
                if isinstance(decoded_filename, bytes):
                    decoded_filename = decoded_filename.decode(enc or "utf-8", errors="ignore")

                payload = part.get_payload(decode=True)
                if not payload:
                    logging.warning(f"Вложение {decoded_filename} не имеет данных (пустое).")
                    continue

                attachments_text += f"\n\n--- СОДЕРЖИМОЕ ВЛОЖЕНИЯ: {decoded_filename} ---\n"
                lower_filename = decoded_filename.lower()

                if lower_filename.endswith((".txt", ".csv")):
                    attachments_text += payload.decode("utf-8-sig", errors="ignore")

                elif lower_filename.endswith(".pdf"):
                    pdf_text = "".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(payload)).pages)
                    attachments_text += pdf_text

                elif lower_filename.endswith(".docx"):
                    attachments_text += docx2txt.process(io.BytesIO(payload))

                elif lower_filename.endswith(".xlsx"):
                    workbook = openpyxl.load_workbook(io.BytesIO(payload), data_only=True)
                    for sheet in workbook.worksheets:
                        attachments_text += f"\nЛист: {sheet.title}\n"
                        for row in sheet.iter_rows(values_only=True):
                            attachments_text += "\t".join(
                                [str(cell) if cell is not None else "" for cell in row]) + "\n"

                elif lower_filename.endswith((".png", ".jpg", "jpeg")):
                    ocr_text = pytesseract.image_to_string(Image.open(io.BytesIO(payload)), lang='rus+eng')
                    attachments_text += f"[Распознанный текст с изображения]:\n{ocr_text}"

                else:
                    attachments_text += f"[Формат файла '{decoded_filename}' не поддерживается для чтения]"
                    logging.warning(f"Вложение '{decoded_filename}' имеет неподдерживаемый тип.")

            except Exception as e:
                logging.error(f"Не удалось прочитать вложение {decoded_filename}: {e}", exc_info=True)
                attachments_text += f"[ОШИБКА ЧТЕНИЯ ФАЙЛА {decoded_filename}]"

        # --- 2. ОБРАБОТКА ТЕЛА ПИСЬМА ---
        content_type = part.get_content_type()
        if not filename and content_type.startswith("text/"):
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
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

    # Выбираем, какой текст использовать: plain-text в приоритете
    final_main_text = main_text_plain.strip()
    if not final_main_text and main_text_html:
        logging.info("Plain-text версия не найдена, используется HTML-версия.")
        soup = BeautifulSoup(main_text_html, "html.parser")
        final_main_text = soup.get_text(separator="\n", strip=True)

    # --- 3. ФОРМИРОВАНИЕ ИТОГОВОГО ТЕКСТА ---
    combined_text = f"=== ТЕКСТ ПИСЬМА ===\n{final_main_text}\n\n=== ВЛОЖЕНИЯ ===\n{attachments_text.strip()}"

    max_length = 10000  # Можно настроить
    if len(combined_text) > max_length:
        combined_text = combined_text[:max_length] + "\n...[текст обрезан из-за превышения лимита]"

    return combined_text, msg


# =============================
# Функции для извлечения товаров из письма (fallback, regex, GPT)
# =============================
def fallback_extract_products_new(email_text: str) -> list:
    products = []
    start = email_text.find("Номенклатура")
    if start == -1:
        return products
    block = email_text[start:]
    signature_markers = ["С уважением", "С уважением,"]
    for marker in signature_markers:
        idx = block.find(marker)
        if idx != -1:
            block = block[:idx]
            break
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    headers = {"Номенклатура", "Ед.изм.", "Кол-во по спец."}
    while lines and lines[0] in headers:
        lines.pop(0)
    if not lines:
        return products
    if len(lines) % 3 != 0:
        logging.warning("Количество строк в блоке номенклатуры не кратно 3, обработка может быть неполной.")
    i = 0
    while i < len(lines) - 2:
        name = lines[i]
        quantity_str = lines[i + 2]
        try:
            quantity = int(quantity_str)
        except ValueError:
            try:
                quantity = float(quantity_str)
            except ValueError:
                quantity = 1
        products.append({
            "name": name,
            "code": "",
            "quantity": quantity,
            "sum": 0.0
        })
        i += 3
    return products

def regex_extract_products(email_text: str) -> list:
    products = []
    pattern = re.compile(r'^(?P<name>.+?)\s+(?P<quantity>\d+)\s*$', re.MULTILINE)
    for m in pattern.finditer(email_text):
        name = m.group("name").strip()
        try:
            quantity = int(m.group("quantity"))
        except ValueError:
            quantity = 1
        products.append({
            "name": name,
            "code": "",
            "quantity": quantity,
            "sum": 0.0
        })
    return products


def gpt_extract_products(email_text: str) -> list:
    headers = {
        "Authorization": f"Api-Key {config['YANDEX_SA_API_KEY']}",
        "Content-Type": "application/json"
    }

    prompt_text = (  # Переименовал prompt в prompt_text для ясности
        "Извлеки из текста письма список товаров. Каждый товар должен быть представлен объектом JSON с полями: "
        "\"name\" (название товара), \"code\" (оставь пустым), \"quantity\" (число) и \"sum\" (0.0). "
        "Если какая-либо информация не найдена, оставь соответствующее поле пустым. "
        "Выдай корректный JSON-массив (список объектов), например: "
        "[{\"name\": \"Штуцер-елочка НР G1\\\" x шл. 25мм AISI 316\", \"code\": \"\", \"quantity\": 1, \"sum\": 0.0}, ...]. "
        "Важно: ответ должен быть строго в формате JSON-массива, без обрамления в markdown ```json ... ```.\n"  # Добавлено уточнение
        f"Текст письма:\n{email_text}"
    )

    payload = {
        "modelUri": config["YANDEX_GPT_MODEL_URI_PATTERN"],
        "completionOptions": {"stream": False, "temperature": 0.0, "maxTokens": "1500"},
        "messages": [{"role": "user", "text": prompt_text}]
    }

    try:
        logging.debug(
            f"Запрос на извлечение товаров к Yandex GPT API: Payload={json.dumps(payload, ensure_ascii=False)[:300]}...")
        response = requests.post(config["YANDEX_GPT_API_ENDPOINT"], headers=headers, json=payload)
        response.raise_for_status()
        response_data = response.json()
        result_text_raw = ""

        if 'result' in response_data and 'alternatives' in response_data['result'] and \
                len(response_data['result']['alternatives']) > 0 and \
                'message' in response_data['result']['alternatives'][0] and \
                'text' in response_data['result']['alternatives'][0]['message']:
            result_text_raw = response_data['result']['alternatives'][0]['message']['text'].strip()
        else:
            logging.error(f"Неожиданная структура ответа от Yandex GPT API при извлечении товаров: {response_data}")
            return []

        logging.info(f"Ответ от Yandex GPT (raw) для извлечения товаров: {result_text_raw[:300]}...")

        # ВЫЗЫВАЕМ НОВУЮ ФУНКЦИЮ ОЧИСТКИ
        result_text_cleaned = clean_yandex_gpt_json_response(result_text_raw)
        logging.info(f"Ответ от Yandex GPT (cleaned) для извлечения товаров: {result_text_cleaned[:300]}...")

        products = json.loads(result_text_cleaned)
        if not isinstance(products, list):
            # ... (обработка ошибки как раньше) ...
            logging.error(f"Yandex GPT вернул корректный JSON, но это не список (массив) товаров: {type(products)}")
            return []
        return products

    except requests.exceptions.RequestException as e:
        # ... (обработка ошибок как раньше) ...
        logging.error(f"Ошибка вызова Yandex GPT API для извлечения товаров: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Содержимое ответа Yandex GPT API: {e.response.text}")
        return []
    except json.JSONDecodeError as e:
        logging.error(
            f"Yandex GPT вернул некорректный JSON для извлечения товаров (после очистки): '{result_text_cleaned if 'result_text_cleaned' in locals() else 'Нет очищенного ответа'}'. Ошибка: {e}")
        return []
    except Exception as e:
        logging.error(f"Неизвестная ошибка при извлечении товаров через Yandex GPT: {e}", exc_info=True)
        return []

def extract_products_multifallback(email_text: str) -> list:
    products = fallback_extract_products_new(email_text)
    if products:
        logging.info("Товары извлечены через fallback_extract_products_new")
        return products
    products = regex_extract_products(email_text)
    if products:
        logging.info("Товары извлечены через regex_extract_products")
        return products
    products = gpt_extract_products(email_text)
    if products:
        logging.info("Товары извлечены через gpt_extract_products")
        return products
    logging.warning("Не удалось извлечь товары ни одним методом.")
    return []

# =============================
# Функция анализа письма через GPT для извлечения данных заказа
# =============================
def analyze_email_with_gpt(email_text: str, config: dict) -> dict:
    # Используем API-ключ сервисного аккаунта для аутентификации
    headers = {
        "Authorization": f"Api-Key {config['YANDEX_SA_API_KEY']}", # ИЗМЕНЕНО: Используем YANDEX_SA_API_KEY
        "Content-Type": "application/json"
        # "x-folder-id": config["YANDEX_FOLDER_ID"], # Обычно не нужен при аутентификации по Api-Key, т.к. папка указывается в modelUri
    }

    system_prompt = (
        "Ты AI-помощник для обработки заказов. Твоя задача – анализировать текст письма и извлекать из него все необходимые реквизиты заказа. "
        "Письмо может содержать данные о компании, контактном лице, дату заказа, номер договора, а также подробный список товаров. "
        "Обрати внимание, что после команды \"прошу выставить счет/кп:\" в письме обычно следует раздел с заголовком \"Номенклатура\", "
        "за которым идут данные о товарах: название, единица измерения и количество. \n\n"
        "Извлеки следующие данные:\n"
        "- Данные компании: name, INN, KPP, legal_address, actual_address, checking_account.\n"
        "- Данные контактного лица: full_name, email, phone.\n"
        "- Дату заказа\n"
        "- Список товаров. Каждый товар должен быть представлен объектом с полями: \"name\" (название товара), \"code\" (если не указан, оставь пустым), "
        "\"quantity\" (количество, число) и \"sum\" (стоимость, число).\n\n"
        "Если какая-либо информация не найдена, оставь соответствующее поле пустым. "
        "Выдай корректный JSON по следующей структуре:\n"
        "{\n"
        "  \"company\": {\"name\": \"\", \"INN\": \"\", \"KPP\": \"\", \"legal_address\": \"\", \"actual_address\": \"\", \"checking_account\": \"\"},\n"
        "  \"order\": {\"contact_person\": {\"full_name\": \"\", \"email\": \"\", \"phone\": \"\"}, \"products\": [ {\"name\": \"\", \"code\": \"\", \"quantity\": 0, \"sum\": 0.0} ], \"datetime\": \"\"},\n"
        "}\n"
        "Если какая-либо информация не найдена, оставь поле пустым."
        "Важно: ответ должен быть строго в формате JSON-массива, без обрамления в markdown ```json ... ```.\n"  # Добавлено уточнение
    )
    user_prompt = f"Текст письма:\n{email_text}"
    payload = {
        "modelUri": config["YANDEX_GPT_MODEL_URI_PATTERN"],
        "completionOptions": {"stream": False, "temperature": 0.0, "maxTokens": "2000"},
        "messages": [{"role": "system", "text": system_prompt}, {"role": "user", "text": user_prompt}]
    }

    try:
        logging.debug(
            f"Запрос к Yandex GPT API: URL={config['YANDEX_GPT_API_ENDPOINT']}, Headers Auth Key={headers['Authorization'][:15]}..., Payload={json.dumps(payload, ensure_ascii=False)[:500]}...")
        response = requests.post(config["YANDEX_GPT_API_ENDPOINT"], headers=headers, json=payload)
        response.raise_for_status()
        response_data = response.json()
        gpt_text_raw = ""

        if 'result' in response_data and 'alternatives' in response_data['result'] and \
                len(response_data['result']['alternatives']) > 0 and \
                'message' in response_data['result']['alternatives'][0] and \
                'text' in response_data['result']['alternatives'][0]['message']:
            gpt_text_raw = response_data['result']['alternatives'][0]['message']['text']
        else:
            logging.error(f"Неожиданная структура ответа от Yandex GPT API: {response_data}")
            return {}

        logging.info(f"Ответ от Yandex GPT (raw): {gpt_text_raw[:500]}...")

        # ВЫЗЫВАЕМ НОВУЮ ФУНКЦИЮ ОЧИСТКИ
        gpt_text_cleaned = clean_yandex_gpt_json_response(gpt_text_raw)
        logging.info(f"Ответ от Yandex GPT (cleaned): {gpt_text_cleaned[:500]}...")

        try:
            data = json.loads(gpt_text_cleaned)
            return data
        except json.JSONDecodeError as e:
            logging.error(f"Yandex GPT вернул некорректный JSON после очистки: {gpt_text_cleaned}")
            logging.debug(f"Ошибка декодирования: {e}")
            return {}

    except requests.exceptions.RequestException as e:
        # ... (обработка ошибок как раньше) ...
        logging.error(f"Ошибка вызова Yandex GPT API: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Содержимое ответа Yandex GPT API: {e.response.text}")
        return {}
    except Exception as e:
        logging.error(f"Неизвестная ошибка при работе с Yandex GPT: {e}", exc_info=True)
        return {}

# =============================
# Основная функция обработки заказов
# =============================
def main():
    global last_email_hash

    # --- ЗАГРУЗКА НОМЕНКЛАТУРЫ ПРИ СТАРТЕ ---
    logging.info("Загрузка номенклатуры...")
    nomenclature_file_path = config.get("NOMENCLATURE_PATH", r"C:\1s\refs\nomenclature.txt")
    nomenclature_data = load_nomenclature(nomenclature_file_path)
    if not nomenclature_data:
        logging.error("Номенклатура не загружена или пуста. Сопоставление будет невозможно.")
    # --- КОНЕЦ БЛОКА ---

    logging.info("Запуск обработки заказов...")
    logging.info(f"Версия {config["VERSION"]}")
    while True:
        try:
            email_text, msg = get_email_text_with_attachments(config)
            if not email_text:
                logging.info("Новых писем не найдено. Ожидание...")
                time.sleep(30)
                continue

            current_hash = hashlib.md5(email_text.encode("utf-8")).hexdigest()
            if last_email_hash is not None and current_hash == last_email_hash:
                logging.info("Письмо уже обработано, пропускаем его. Ожидание...")
                time.sleep(30)
                continue
            last_email_hash = current_hash

            logging.info("Первые 5000 символов письма:\n%s", email_text[:5000])
            logging.info("Письмо получено. Анализирую письмо через GPT...")
            order_data = analyze_email_with_gpt(email_text, config)
            if not order_data:
                logging.error("Не удалось извлечь данные заказа из письма.")
                time.sleep(30)
                continue
            logging.info("Извлеченные данные заказа:\n%s", json.dumps(order_data, ensure_ascii=False, indent=2))

            if not order_data.get("order", {}).get("products"):
                logging.info("Список товаров пуст, пробую извлечь товары с многоступенчатым методом.")
                fallback_products = extract_products_multifallback(email_text)
                logging.info("Многоступенчатое извлечение товаров вернуло: %s", fallback_products)
                if fallback_products:
                    order_data.setdefault("order", {})["products"] = fallback_products
                else:
                    logging.warning("Не удалось извлечь товары ни одним методом.")

            order_data["email_msg"] = msg
            order_data["email_text"] = email_text

            # --- ВЫЗОВ ФУНКЦИИ С ПЕРЕДАЧЕЙ НОМЕНКЛАТУРЫ ---
            generate_order_xml(order_data, config, nomenclature_data)
            # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            logging.info("Обработка заказа завершена.")
        except Exception as e:
            logging.error(f"Общая ошибка в обработке: {e}", exc_info=True)
        logging.info("Ожидание новых писем...")
        time.sleep(30)

if __name__ == "__main__":
    main()