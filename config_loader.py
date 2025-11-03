import json
import logging
import os
from datetime import datetime


def load_config(path="config.json") -> dict:
    """Загружает основной файл конфигурации."""
    if not os.path.exists(path):
        logging.critical(f"Критическая ошибка: файл конфигурации {path} не найден.")
        # В реальной системе лучше падать, если нет конфига
        raise FileNotFoundError(f"Файл конфигурации {path} не найден.")

    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def load_json_data(path: str, description: str) -> dict:
    """Универсальная функция для загрузки JSON-данных (regex, synonyms, params)."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        logging.info(f"Словарь '{description}' успешно загружен из {path}.")
        return data
    except FileNotFoundError:
        logging.error(f"Файл '{description}' не найден по пути: {path}. Будет использован пустой словарь.")
        return {}
    except json.JSONDecodeError:
        logging.error(
            f"Ошибка декодирования JSON в файле '{description}' по пути: {path}. Будет использован пустой словарь.")
        return {}
    except Exception as e:
        logging.error(f"Не удалось загрузить '{description}' из {path}: {e}", exc_info=True)
        return {}


def setup_logging(config: dict) -> None:
    """Настраивает систему логирования (в консоль и в файл)."""

    # Уровень по умолчанию DEBUG, если не указан в config.json
    log_level_str = config.get("LOG_LEVEL", "DEBUG").upper()
    log_level = getattr(logging, log_level_str, logging.DEBUG)

    log_format = "%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s"
    formatter = logging.Formatter(log_format)

    # Получаем корневой логгер
    logger = logging.getLogger()
    logger.setLevel(log_level)  # Устанавливаем уровень для самого логгера

    # --- Обработчик для консоли ---
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)  # Уровень для консоли
    console_handler.setFormatter(formatter)

    # Добавляем, только если еще нет обработчика StreamHandler
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        logger.addHandler(console_handler)

    # --- Обработчик для файла ---
    log_directory = config.get("LOGS_FOLDER", "logs")
    try:
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)
            # Используем print, так как логгер может еще не писать в консоль
            print(f"INFO: Создана папка для логов: {log_directory}")
    except OSError as e:
        print(f"ERROR: Не удалось создать папку для логов {log_directory}: {e}")
        return  # Если не можем создать папку, не падаем, но в файл писать не будем

    log_file_name = f"order_processing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    log_file_path = os.path.join(log_directory, log_file_name)

    try:
        file_handler = logging.FileHandler(log_file_path, mode='a', encoding='utf-8')
        file_handler.setLevel(log_level)  # Уровень для файла
        file_handler.setFormatter(formatter)

        # Добавляем, только если еще нет файлового обработчика
        if not any(isinstance(h, logging.FileHandler) for h in logger.handlers):
            logger.addHandler(file_handler)

        logging.info(f"Логирование в файл настроено: {log_file_path}")
    except Exception as e:
        logging.error(f"Ошибка настройки логирования в файл {log_file_path}: {e}", exc_info=True)