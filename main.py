import logging
import time
import hashlib
import json

# Импортируем наши новые, разделенные модули
import config_loader
import email_processor
import gpt_client
import nomenclature_parser
import order_exporter

# Глобальная переменная для предотвращения повторной обработки
# (такая же логика, как в вашем старом main.py)
last_email_hash = None


def main_loop():
    """Главный цикл обработки."""
    global last_email_hash

    # --- 1. ЗАГРУЗКА КОНФИГУРАЦИИ И НАСТРОЙКА ЛОГОВ ---
    try:
        config = config_loader.load_config("config.json")
        config_loader.setup_logging(config)
    except FileNotFoundError:
        # Критическая ошибка, если config.json не найден
        print("Ошибка: config.json не найден! Работа программы невозможна.")
        return  # Завершаем работу, если нет конфига
    except Exception as e:
        print(f"Критическая ошибка при инициализации конфига или логов: {e}")
        return

    logging.info(f"Запуск Nomenclature-Search v{config.get('VERSION', 'N/A')}")

    # --- 2. ЗАГРУЗКА ДАННЫХ ДЛЯ ПАРСЕРА ---
    # Загружаем синонимы, regex, параметры и номенклатуру
    try:
        nomenclature_parser.load_all_parser_data(config)
    except Exception as e:
        logging.critical(f"Критическая ошибка при загрузке данных парсера (номенклатура, regex): {e}", exc_info=True)
        # Если номенклатура не загрузилась, продолжать нет смысла
        return

    logging.info("Инициализация завершена. Вхожу в главный цикл обработки...")

    # --- 3. ГЛАВНЫЙ ЦИКЛ ---
    while True:
        try:
            # --- Шаг А. Получение письма и вложений ---
            email_text, msg_object = email_processor.get_email_text_with_attachments(config)

            if not email_text:
                logging.info("Новых писем не найдено. Ожидание 30 сек...")
                time.sleep(30)
                continue

            # --- Шаг Б. Проверка на дубликаты ---
            # (Логика из вашего старого main.py)
            current_hash = hashlib.md5(email_text.encode("utf-8")).hexdigest()
            if last_email_hash is not None and current_hash == last_email_hash:
                logging.info("Письмо уже было обработано (хеш совпадает). Ожидание 30 сек...")
                time.sleep(30)
                continue

            # Обновляем хеш *перед* обработкой
            logging.info(f"Получено новое письмо. Хеш: {current_hash}. Начинаю обработку...")
            logging.debug(f"Текст письма для анализа (первые 5000 символов):\n{email_text[:5000]}")

            # --- Шаг В. Анализ через GPT ---
            order_data = gpt_client.analyze_email_with_gpt(email_text, config)  #

            if not order_data:
                logging.error("Не удалось извлечь данные заказа из письма (GPT вернул пустой результат).")
                # Все равно обновляем хеш, чтобы не зациклиться на этом письме
                last_email_hash = current_hash
                time.sleep(30)
                continue

            logging.info("Данные из GPT получены. Извлечено ключей: %s", list(order_data.keys()))

            # --- Шаг Г. Fallback-логика (если GPT не нашел товары) ---
            if not order_data.get("order", {}).get("products"):
                logging.warning("GPT не вернул список товаров. Запускаю fallback-логику...")
                fallback_products = nomenclature_parser.extract_products_multifallback(email_text, config)  #

                if fallback_products:
                    order_data.setdefault("order", {})["products"] = fallback_products
                    logging.info(f"Fallback-логика успешно извлекла {len(fallback_products)} позиций.")
                else:
                    logging.error("Fallback-логика также не смогла извлечь товары. Заказ не будет создан.")

            # Добавляем в данные сырой текст и объект письма (на всякий случай)
            order_data["email_msg"] = msg_object
            order_data["email_text"] = email_text

            # --- Шаг Д. Генерация XML ---
            order_exporter.generate_order_xml(order_data, config)  #

            # --- Шаг Е. Успешное завершение ---
            logging.info("Обработка заказа завершена. Обновляю хеш.")
            last_email_hash = current_hash  # Фиксируем хеш только после ПОЛНОЙ обработки

        except Exception as e:
            logging.critical(f"Необработанная ошибка в главном цикле: {e}", exc_info=True)
            # При любой ошибке ждем, чтобы не попасть в цикл быстрых падений
            time.sleep(60)

        logging.info("Цикл завершен. Ожидание 30 сек...")
        time.sleep(30)


if __name__ == "__main__":
    main_loop()