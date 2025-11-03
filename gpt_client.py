import requests
import json
import logging


def clean_yandex_gpt_json_response(raw_response_text: str) -> str:
    """
    Очищает сырой текстовый ответ от YandexGPT, удаляя Markdown-обрамление
    для JSON (```json ... ``` или ``` ... ```).

    """
    cleaned_text = raw_response_text.strip()

    # Удаляем начальное обрамление
    if cleaned_text.startswith("```json"):
        cleaned_text = cleaned_text[len("```json"):].strip()
    elif cleaned_text.startswith("```"):
        cleaned_text = cleaned_text[len("```"):].strip()

    # Удаляем конечное обрамление
    if cleaned_text.endswith("```"):
        cleaned_text = cleaned_text[:-len("```")].strip()

    return cleaned_text


def analyze_email_with_gpt(email_text: str, config: dict) -> dict:
    """
    Анализирует полный текст письма для извлечения данных о компании и заказе.

    """
    headers = {
        "Authorization": f"Api-Key {config['YANDEX_SA_API_KEY']}",  #
        "Content-Type": "application/json"
    }

    # Этот системный промпт взят из вашего main.py
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
        "Важно: ответ должен быть строго в формате JSON-массива, без обрамления в markdown ```json ... ```.\n"
    )  #

    user_prompt = f"Текст письма:\n{email_text}"

    payload = {
        "modelUri": config["YANDEX_GPT_MODEL_URI_PATTERN"],  #
        "completionOptions": {"stream": False, "temperature": 0.0, "maxTokens": "2000"},
        "messages": [{"role": "system", "text": system_prompt}, {"role": "user", "text": user_prompt}]
    }

    try:
        logging.debug(f"Запрос к Yandex GPT API (analyze_email): URL={config['YANDEX_GPT_API_ENDPOINT']}")
        response = requests.post(config["YANDEX_GPT_API_ENDPOINT"], headers=headers, json=payload)  #
        response.raise_for_status()  # Проверка на HTTP ошибки (4xx, 5xx)
        response_data = response.json()
        gpt_text_raw = ""

        # Безопасное извлечение текста из ответа
        if 'result' in response_data and 'alternatives' in response_data['result'] and \
                len(response_data['result']['alternatives']) > 0 and \
                'message' in response_data['result']['alternatives'][0] and \
                'text' in response_data['result']['alternatives'][0]['message']:
            gpt_text_raw = response_data['result']['alternatives'][0]['message']['text']
        else:
            logging.error(f"Неожиданная структура ответа от Yandex GPT API: {response_data}")
            return {}

        logging.info(f"Ответ от Yandex GPT (raw): {gpt_text_raw[:500]}...")

        gpt_text_cleaned = clean_yandex_gpt_json_response(gpt_text_raw)  #
        logging.info(f"Ответ от Yandex GPT (cleaned): {gpt_text_cleaned[:500]}...")

        try:
            data = json.loads(gpt_text_cleaned)
            return data
        except json.JSONDecodeError as e:
            logging.error(f"Yandex GPT вернул некорректный JSON после очистки: {gpt_text_cleaned}")
            logging.debug(f"Ошибка декодирования: {e}")
            return {}

    except requests.exceptions.RequestException as e:
        logging.error(f"Ошибка вызова Yandex GPT API: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Содержимое ответа Yandex GPT API: {e.response.text}")
        return {}
    except Exception as e:
        logging.error(f"Неизвестная ошибка при работе с Yandex GPT: {e}", exc_info=True)
        return {}


def gpt_extract_products(email_text: str, config: dict) -> list:
    """Извлекает список товаров из текста (используется как fallback)."""
    headers = {
        "Authorization": f"Api-Key {config['YANDEX_SA_API_KEY']}",
        "Content-Type": "application/json"
    }

    # Промпт для извлечения только товаров
    prompt_text = (
        "Извлеки из текста письма список товаров. Каждый товар должен быть представлен объектом JSON с полями: "
        "\"name\" (название товара), \"code\" (оставь пустым), \"quantity\" (число) и \"sum\" (0.0). "
        "Если какая-либо информация не найдена, оставь соответствующее поле пустым. "
        "Выдай корректный JSON-массив (список объектов), например: "
        "[{\"name\": \"Штуцер-елочка НР G1\\\" x шл. 25мм AISI 316\", \"code\": \"\", \"quantity\": 1, \"sum\": 0.0}, ...]. "
        "Важно: ответ должен быть строго в формате JSON-массива, без обрамления в markdown ```json ... ```.\n"
        f"Текст письма:\n{email_text}"
    )  #

    payload = {
        "modelUri": config["YANDEX_GPT_MODEL_URI_PATTERN"],
        "completionOptions": {"stream": False, "temperature": 0.0, "maxTokens": "1500"},
        "messages": [{"role": "user", "text": prompt_text}]
    }

    try:
        logging.debug(f"Запрос на извлечение товаров к Yandex GPT API (gpt_extract_products)")
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

        result_text_cleaned = clean_yandex_gpt_json_response(result_text_raw)  #
        logging.info(f"Ответ от Yandex GPT (cleaned) для извлечения товаров: {result_text_cleaned[:300]}...")

        products = json.loads(result_text_cleaned)
        if not isinstance(products, list):
            logging.error(f"Yandex GPT вернул корректный JSON, но это не список (массив) товаров: {type(products)}")
            return []
        return products

    except requests.exceptions.RequestException as e:
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