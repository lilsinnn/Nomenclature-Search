import logging
import re
import json
import csv
import os
from gpt_client import gpt_extract_products  # Импортируем fallback-функцию из gpt_client

# --- Глобальные переменные уровня модуля ---
# Они будут загружены один раз при старте, в main.py
SYNONYMS_TYPE = {}
REGEX_PATTERNS = {}
PART_SPECIFICATIONS = {}
NOMENCLATURE_DATA = []

# Словарь синонимов материалов (взят из main.py)
MATERIAL_ALIASES = {
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
DEFAULT_MATERIAL = "ст20"


def load_all_parser_data(config: dict):
    """
    Загружает все необходимые для парсера данные (синонимы, regex, параметры)
    в глобальные переменные этого модуля.
    """
    global SYNONYMS_TYPE, REGEX_PATTERNS, PART_SPECIFICATIONS, NOMENCLATURE_DATA

    # Динамически импортируем, чтобы избежать циклических зависимостей
    from config_loader import load_json_data

    SYNONYMS_TYPE = load_json_data(config["SYNONYMS_PATH"], "Синонимы типов")
    REGEX_PATTERNS = load_json_data(config["REGEX_PATH"], "Регулярные выражения")
    PART_SPECIFICATIONS = load_json_data(config["PARAMETERS_PATH"], "Параметры деталей")
    NOMENCLATURE_DATA = load_nomenclature(config.get("NOMENCLATURE_PATH", r"C:\1s\refs\nomenclature.txt"))


def load_nomenclature(file_path: str) -> list[dict]:
    """
    Загружает номенклатуру из текстового файла с разделителями-табуляторами.

    """
    if not os.path.exists(file_path):
        logging.error(f"Файл номенклатуры не найден по пути: {file_path}")
        return []

    nomenclature_data = []
    try:
        # Используем 'utf-8-sig' для правильной обработки BOM (byte order mark)
        # Это частая проблема, когда '\ufeffКод' появляется вместо 'Код'
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter='\t')

            try:
                header = next(reader)
                # Очистим заголовки от возможных BOM и пробелов
                header = [h.strip() for h in header]
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


def parse_order_name(order_name: str, logs: bool = True) -> dict:
    """
    Гибридная функция парсинга.
    Сочетает точный поиск по regex-правилам и гибкий поиск по словарю параметров.

    """
    if logs:
        logging.debug(f"Начало гибридного парсинга строки: '{order_name}'")

    # --- Шаг 1: Определение типа детали ---
    item_type = None
    # Добавляем пробелы по краям для точного поиска (чтобы "кран" не нашелся в "экране")
    work_string = f" {order_name.lower().replace(',', '.')} "

    global SYNONYMS_TYPE, REGEX_PATTERNS, PART_SPECIFICATIONS

    # Исключаем "комментарий", если он есть в synonyms_data.json
    product_types = {k: v for k, v in SYNONYMS_TYPE.items() if k.lower() != "комментарий"}

    # Сортируем ключи, чтобы сначала искались самые длинные (например, "кольцо-заглушка", а не "кольцо")
    sorted_keys = sorted(product_types.keys(), key=len, reverse=True)

    for t in sorted_keys:
        for synonym in product_types[t]:
            if f" {synonym.lower()} " in work_string:
                item_type = t
                # Удаляем найденный синоним, чтобы он не мешал дальнейшему парсингу
                work_string = work_string.replace(f" {synonym.lower()} ", " ", 1)
                break
        if item_type:
            break

    if not item_type:
        if logs:
            logging.warning(f"Не удалось определить тип для: '{order_name}'")
        return {"original_name": order_name, "type": None, "params": {}}

    if logs:
        logging.debug(f"Определен тип: '{item_type}'. Оставшаяся строка: '{work_string.strip()}'")

    # --- Инициализация словарей для сбора параметров ---
    parsed_data = {"original_name": order_name, "type": item_type, "params": {}}
    found_params = {}

    # --- Шаг 2: Точный поиск по regex-правилам (regex.json) ---
    # Это для параметров, которые вы не хотите искать "выборочно"
    patterns_for_type = REGEX_PATTERNS.get(item_type, {})
    for param_name, patterns_list in patterns_for_type.items():
        for pattern in patterns_list:
            match = re.search(pattern, work_string, re.IGNORECASE)  # Добавим IGNORECASE для надежности
            if match:
                # Собираем группы, которые не None
                found_groups = list(filter(None, match.groups()))

                # Определяем разделитель
                separator = 'x' if 'x' in match.group(0).lower() else '-'

                value = separator.join(found_groups) if len(found_groups) > 1 else (
                    found_groups[0] if found_groups else match.group(0))

                # Сразу кладем в основной словарь params
                parsed_data["params"][param_name] = value.strip()
                # Удаляем найденный фрагмент из строки, чтобы он не мешал шагу 3
                work_string = re.sub(re.escape(match.group(0)), ' ', work_string, 1, re.IGNORECASE)
                logging.debug(f"REGEX-ПОИСК: Найден параметр '{param_name}': '{value}'.")
                break  # Переходим к следующему имени параметра (например, к 'angle')

    if logs:
        logging.debug(f"Строка после REGEX-поиска: '{work_string.strip()}'")

    # --- Шаг 3: "Строгий" выборочный поиск (parameters.json) ---
    # Это как раз тот метод, о котором вы говорили,
    # он ищет известные значения в любом порядке.
    specs = PART_SPECIFICATIONS.get(item_type, {})
    if specs:
        # Ищем только те параметры, которые еще не были найдены на Шаге 2
        props_to_find = [key for key in specs.keys() if key not in parsed_data["params"]]

        # Сортируем, чтобы сначала искались более "важные" или "длинные" ключи
        # (хотя в данном случае важнее сортировка значений)
        for prop_key in sorted(props_to_find, key=len, reverse=True):
            prop_config = specs[prop_key]
            possible_values = prop_config.get("values", [])

            # Сортируем значения от самого длинного к самому короткому
            # чтобы "100" не нашлось раньше, чем "1000"
            for value in sorted(possible_values, key=lambda x: len(str(x)), reverse=True):
                # Ищем значение как отдельное слово (\b)
                if re.search(rf'\b{re.escape(str(value).lower())}\b', work_string):
                    found_params[prop_key] = str(value)
                    logging.debug(f"ВЫБОРОЧНЫЙ ПОИСК: Найден параметр '{prop_key}': '{value}'.")
                    # Удаляем найденное значение из строки
                    work_string = re.sub(rf'\b{re.escape(str(value).lower())}\b', ' ', work_string, count=1)
                    break  # Нашли значение для этого prop_key, идем к следующему prop_key

    # --- Шаг 4: Применение значений по умолчанию ---
    if specs:
        for prop_key in specs.keys():
            # Если параметр не найден ни на шаге 2, ни на шаге 3
            if prop_key not in parsed_data["params"] and prop_key not in found_params:
                default_value = specs[prop_key].get("default")
                if default_value is not None:
                    found_params[prop_key] = str(default_value)
                    logging.debug(f"DEFAULT: Применен параметр '{prop_key}': '{default_value}'.")

    # Объединяем параметры, найденные по Regex (Шаг 2) и выборочно (Шаг 3 + 4)
    parsed_data["params"].update(found_params)

    if logs:
        logging.info(f"Результат гибридного парсинга: {json.dumps(parsed_data, ensure_ascii=False, indent=2)}")

    return parsed_data


def find_best_match(order_product_name: str) -> dict | None:
    """
    Ищет "сырое" название позиции в загруженной номенклатуре.

    """
    global NOMENCLATURE_DATA

    # Сначала парсим имя из заказа, чтобы понять, что мы ищем
    parsed_order_info = parse_order_name(order_product_name)

    order_type = parsed_order_info.get("type")
    order_params = parsed_order_info.get("params", {})

    # --- Ключевая логика для точного сопоставления ---
    # Пытаемся извлечь "ключевые" размеры (например, "57x3.5" или "100-16")
    order_dimensions = order_params.get("dimensions")  # Из regex.json

    # (Можно добавить и другие ключевые параметры, если "dimensions" не нашлось)
    # ...

    if not order_type:
        logging.warning(f"Не удалось определить тип для '{order_product_name}', поиск в номенклатуре невозможен.")
        return None

    logging.debug(f"ИЩУ В НОМЕНКЛАТУРЕ: '{order_product_name}' (Тип: {order_type}, Размеры: {order_dimensions})")

    # 1. Фильтруем номенклатуру по типу
    # Берем первый синоним (обычно основное название)
    filter_keyword = SYNONYMS_TYPE.get(order_type, [order_type])[0].lower()
    candidates = [
        item for item in NOMENCLATURE_DATA
        if filter_keyword in item.get('Полное наименование', '').lower()
    ]

    if not candidates:
        logging.warning(f"Для типа '{order_type}' не найдено ни одного товара в номенклатуре.")
        return None

    logging.debug(f"Найдено {len(candidates)} кандидатов по типу '{order_type}'.")

    # 2. Строгая фильтрация по размерам (если они были найдены)
    # Это самый важный шаг для точности
    strict_candidates = []
    if order_dimensions:
        for item in candidates:
            # Парсим полное наименование из номенклатуры
            item_info = parse_order_name(item.get('Полное наименование', ''),
                                         logs=False)  # Выключаем логи парсинга номенклатуры
            item_dimensions = item_info.get("params", {}).get("dimensions")

            # Сравниваем "ключевые" размеры
            if order_dimensions == item_dimensions:
                strict_candidates.append(item)
                # Если после строгой фильтрации остались кандидаты
                if strict_candidates:
                    if len(strict_candidates) == 1:
                        logging.debug(f"После фильтра по размерам '{order_dimensions}' остался 1 точный кандидат.")
                        candidates = strict_candidates
                    else:
                        # Если >1 кандидата, используем их
                        logging.debug(
                            f"После фильтра по размерам '{order_dimensions}' осталось {len(strict_candidates)} кандидатов.")
                        candidates = strict_candidates
                elif not strict_candidates and order_dimensions:
                    # Если строгая фильтрация не дала_результатов, но размеры БЫЛИ
                    logging.warning(
                        f"Для '{order_product_name}' не найдено ни одного товара с точным совпадением размеров '{order_dimensions}'. Поиск прерван.")
                    return None  # Если нет совпадения по размерам - совпадения нет вообще
                # Если order_dimensions не было, то strict_candidates будет пуст, и мы просто используем
                # всех кандидатов (candidates) по типу, что корректно.

            # 3. Оценка оставшихся кандидатов
            # Мы сравниваем "второстепенные" параметры (сталь, угол, исп.)
            best_score = -1
            best_match = None

            for item in candidates:
                current_score = 0
                # Снова парсим, но на этот раз берем все параметры
                item_info = parse_order_name(item.get('Полное наименование', ''), logs=False)
                item_params = item_info.get("params", {})

                # Сравниваем все параметры из заказа со всеми параметрами из номенклатуры
                for key, order_value in order_params.items():
                    # Размеры уже проверили, их можно пропустить
                    if key != 'dimensions' and item_params.get(key) == order_value:
                        current_score += 1

                if current_score > best_score:
                    best_score = current_score
                    best_match = item
                elif current_score == best_score:
                    # (Опционально) Здесь можно добавить логику, что делать при одинаковом счете
                    # Например, выбирать более короткое/длинное имя или первое попавшееся
                    pass

            if best_match:
                logging.info(
                    f"✅ НАЙДЕНО: Для '{order_product_name}' -> '{best_match.get('Полное наименование')}' (Код: {best_match.get('Код')})")
            else:
                logging.warning(f"⚠️ НЕ НАЙДЕНО: Для '{order_product_name}'. Позиция не сопоставлена.")

            return best_match

        def extract_products_multifallback(email_text: str, config: dict) -> list:
            """
            Пытается извлечь список товаров из текста, если GPT их не нашел.

            """

            # 1. Сначала пробуем простой regex
            products = regex_extract_products(email_text)
            if products:
                logging.info(f"Fallback: Товары извлечены через regex_extract_products ({len(products)} шт.)")
                return products

            # 2. Пробуем более сложный fallback из main.py
            products = fallback_extract_products_new(email_text)
            if products:
                logging.info(f"Fallback: Товары извлечены через fallback_extract_products_new ({len(products)} шт.)")
                return products

            # 3. Если ничего не помогло - зовем GPT специально для поиска товаров
            logging.warning(
                "Ни один из локальных fallback-методов не сработал. Вызываю GPT для прицельного извлечения товаров...")
            products = gpt_extract_products(email_text, config)  #
            if products:
                logging.info(f"Fallback: Товары извлечены через gpt_extract_products ({len(products)} шт.)")
                return products

            logging.error("Не удалось извлечь товары ни одним из методов.")
            return []

        def fallback_extract_products_new(email_text: str) -> list:
            """Один из fallback-методов, перенесенный из main.py"""
            products = []
            # Ищем начало блока
            start = email_text.find("Номенклатура")
            if start == -1:
                return products

            block = email_text[start:]

            # Ищем конец блока (подпись)
            signature_markers = ["С уважением", "С уважением,", "Best regards"]
            for marker in signature_markers:
                idx = block.find(marker)
                if idx != -1:
                    block = block[:idx]
                    break

            lines = [line.strip() for line in block.splitlines() if line.strip()]

            # Убираем строки заголовков
            headers = {"Номенклатура", "Ед.изм.", "Кол-во по спец."}
            while lines and lines[0] in headers:
                lines.pop(0)

            if not lines:
                return products

            # Логика парсинга "каждая третья строка"
            if len(lines) % 3 != 0:
                logging.warning(
                    "Fallback: Количество строк в блоке номенклатуры не кратно 3, обработка может быть неполной.")

            i = 0
            while i < len(lines) - 2:
                name = lines[i]
                unit = lines[i + 1]  # Ед. изм.
                quantity_str = lines[i + 2]

                # Проверяем, похоже ли name на имя, а unit на ед.изм. (например, "шт")
                if len(unit) > 5:
                    # Вероятно, сбилась структура "каждый третий"
                    logging.warning(f"Fallback: Пропускаем строку, не похожую на Ед.Изм.: {unit}")
                    i += 1
                    continue

                try:
                    quantity = float(quantity_str.replace(',', '.').replace(' ', ''))
                except ValueError:
                    logging.warning(f"Fallback: Не удалось распознать количество '{quantity_str}', ставлю 1.")
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
            """Более простой regex-fallback."""
            products = []
            # Ищет строки, которые заканчиваются на число
            pattern = re.compile(r'^(?P<name>.+?)\s+(?P<quantity>\d+([.,]\d+)?)\s*$', re.MULTILINE)
            for m in pattern.finditer(email_text):
                name = m.group("name").strip()

                # Отсеиваем явный мусор
                if name.lower() in ["инн", "кпп", "телефон", "email"]:
                    continue

                try:
                    quantity = float(m.group("quantity").replace(',', '.'))
                except ValueError:
                    quantity = 1

                products.append({
                    "name": name,
                    "code": "",
                    "quantity": quantity,
                    "sum": 0.0
                })
            return products