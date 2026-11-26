"""
ОПИСАНИЕ СКРИПТА:
Этот универсальный код предназначен для автоматической генерации XML-фидов (выгрузок товаров) 
брендов Prompower и Unimat сразу для нескольких интернет-магазинов ("Чип и Дип" и "Stoking"),
а также для ведения умного лога изменений цен и остатков (Delta-лог).

Где используется: В репозитории feed-for-onlinestore на GitHub Actions.

Как работает скрипт (пошагово):
1. Подключается к серверу Prompower по API и запрашивает список всех категорий и товаров обоих брендов.
2. Загружает локальный кэш PDF-ссылок (чтобы не перегружать сайт поставщика ежедневным сканированием).
3. 1-го числа каждого месяца (или в режиме отладки) обходит страницы товаров на сайте Prompower и собирает ссылки на PDF-документы.
4. Для каждого клиента (сначала Чип и Дип, затем Stoking) применяет индивидуальные правила расчета цен (costPrice и rPrice),
   форматирования комментариев и набора XML-тегов, добавляя в шапку дату и время создания файла по МСК.
5. Сохраняет готовые XML-файлы на GitHub Pages, а также создает для каждого клиента аварийный файл ZEROwarehouse с нулевыми остатками.
6. Если режим отладки ВЫКЛЮЧЕН — сравнивает сегодняшние цены (price с НДС 1.22, costPrice, rPrice) и остатки со вчерашними 
   (из logs/last_state.json) и дописывает все произошедшие изменения в помесячный файл logs/history_ГГГГ-ММ.csv для открытия в Excel.
"""

import requests
import json
import os
import re
import csv
import datetime
import time
import urllib.parse
from bs4 import BeautifulSoup
from xml.sax.saxutils import escape

# ==============================================================================
# 1. НАСТРОЙКИ ДОСТУПА, ПУТЕЙ И РЕЖИМА ОТЛАДКИ
# ==============================================================================

# Получаем логин и ключ API из безопасного хранилища GitHub Secrets
EMAIL = os.getenv("API_EMAIL")
KEY = os.getenv("API_KEY")

# Определяем, включен ли режим отладки (если запускали вручную с галочкой)
DEBUG_MODE = os.getenv("DEBUG_MODE", "false").lower() == "true"
try:
    DEBUG_LIMIT = int(os.getenv("DEBUG_LIMIT", "3"))
except (ValueError, TypeError):
    DEBUG_LIMIT = 3

# Если секретные ключи не переданы, останавливаем программу с понятной ошибкой
if not EMAIL or not KEY:
    print("КРИТИЧЕСКАЯ ОШИБКА: Не заданы секреты API_EMAIL или API_KEY в GitHub Secrets!")
    exit(1)

# Базовые адреса сайта-источника и вашего репозитория
REPO_NAME = "feed-for-onlinestore"
API_URL = "https://prompower.ru/api/prod/"
SITE_URL = "https://www.prompower.ru"
PAGES_BASE_URL = f"https://brilka.github.io/{REPO_NAME}/"

# Названия файлов кэша и папка для логов цен/остатков
CACHE_FILENAME = "pdf_cache.json"
OLD_CACHE_FILENAME = "chipidip_pdf_cache.json"
LOGS_DIR = "logs"
STATE_FILENAME = os.path.join(LOGS_DIR, "last_state.json")

# Список клиентов и имен их итоговых XML-файлов
CLIENTS_CONFIG = [
    {
        "key": "chipdip",
        "xml_file": "feedChipdip.xml",
        "zero_xml_file": "feedChipdip-ZEROwarehouse.xml"
    },
    {
        "key": "stoking",
        "xml_file": "feedStoking.xml",
        "zero_xml_file": "feedStoking-ZEROwarehouse.xml"
    }
]


# ==============================================================================
# 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Экранирование XML, форматирование чисел, кэш и API)
# ==============================================================================

def xml_escape(text):
    """
    Защищает XML от поломки из-за спецсимволов.
    Преобразует символы <, >, &, двойные и одинарные кавычки в безопасные XML-коды.
    """
    if text is None:
        return ""
    return escape(str(text), {'"': '&quot;', "'": '&apos;'})


def format_excel_num(val):
    """
    Форматирует число для русского Excel (заменяет точку на запятую, убирает лишний '.0').
    Например: 2020.0 -> '2020', 1250.55 -> '1250,55'.
    """
    if val is None or val == "":
        return ""
    try:
        num = float(val)
        if num.is_integer():
            return str(int(num))
        return str(round(num, 2)).replace(".", ",")
    except (ValueError, TypeError):
        return str(val).replace(".", ",")


# Словарь сопоставления категорий поставщика (Prompower/Unimat) и кодов групп для Чип и Дип
GROUP_MAP = {
    "19“ комплектующие": "3073",
    "Аксессуары": "3073",
    "Двери": "3059",
    "Коллаборативные роботы": "2584",
    "Модификационные комплекты для моторов": "2178",
    "Модули расширения ПЛК PMP301": "2273",
    "Монтажные панели": "3054",
    "Моторные дроссели": "2926",
    "Моторы": "2178",
    "Серводвигатели": "2585",
    "Сетевые дроссели": "2926",
    "Соединительные комплекты": "3073",
    "Шасси": "3073",
    "Шкафы электротехнические": "3073",
    "Опции для двигателей": "2178",
    "Аксессуары для ПЛК": "2273",
    "Аксессуары для реле": "2459",
    "Дополнительные контактные приставки": "2947",
    "Заземление": "3073",
    "Опции для преобразователей частоты": "2954",
    "Дополнительные контактные приставки PULSE": "2947",
    "Колодки для реле": "3010",
    "Прокладка кабеля": "3065",
    "MCB (Miniature Circuit Breaker)": "3109",
    "MCB": "3109",
    "Реле общего назначения": "1559",
    "Контакторы PULSE": "2947",
    "Панели основания": "3067",
    "Аксессуары для сервосистем": "2585",
    "Контакторы": "2947",
    "Миниатюрные силовые реле": "1559",
    "Сувенирная продукция": "2598",
    "Реле тонкие": "3624",
    "Кабели для датчиков": "1266",
    "Миниконтакторы": "2947",
    "Миниконтакторы PULSE": "2947",
    "Цоколи": "3062",
    "Климат + Свет": "3184",
    "Блок питания HDR в пластиковом корпусе": "2939",
    "Пластроны": "3124",
    "Блок питания MDR в пластиковом корпусе": "2939",
    "Боковые панели": "3069",
    "Секционирование": "3062",
    "Индуктивные датчики": "1266",
    "Дополнительные контактные приставки для MCB": "3115",
    "Опции для устройств плавного пуска": "2968",
    "Модули расширения ПЛК PMP20/PMP30": "2273",
    "Блок питания NDR в металлическом корпусе": "2939",
    "Автоматы защиты двигателя PULSE": "2930",
    "Фотоэлектрические датчики": "2744",
    "Полки": "3071",
    "Потолочные панели": "3067",
    "Преобразователи частоты PD100": "2954",
    "Преобразователи частоты PD101": "2954",
    "Тормозные резисторы": "1667",
    "Преобразователи частоты PD150": "2954",
    "Панели оператора PH1": "2622",
    "Задние панели": "3069",
    "Промышленные коммутаторы": "3256",
    "Панели оператора PH": "2273",
    "ЭМС фильтры": "2926",
    "Сейсмостойкость": "3062",
    "Дроссели dU/dt": "2926",
    "Устройства плавного пуска P2S 050": "2968",
    "Дроссели для цепей постоянного тока": "2926",
    "Преобразователи частоты PD210": "2954",
    "Преобразователи частоты PD110": "2954",
    "Устройства плавного пуска P2S 100": "2968",
    "Сервоприводы": "2585",
    "Регуляторы мощности": "1556",
    "Программируемые логические контроллеры PMP20": "2273",
    "Преобразователи частоты PD310": "2954",
    "Электродвигатели класс энергоэфф. IE1": "2178",
    "Электродвигатели класс энергоэффективности IE1": "2178",
    "ПЛК PMP301": "2273",
    "Программируемые логические контроллеры PMP301": "2273",
    "Каркасы": "3072",
    "Синус-фильтры": "2926",
    "Внешние тормозные модули для ПЧ": "2954",
    "Преобразователи частоты PD310 IP54": "2954",
    "Промышленный монитор": "2273",
    "Устройства плавного пуска P2S 300": "2968",
    "Промышленный ПК": "2273",
    "ПЛК PMP30": "2273",
    "Программируемые логические контроллеры PMP30": "2273",
    "Панельный ПК": "2273",
    "Кабели и аксессуары": "2273",
    "Модули для ПЛК": "2273",
    "Панели оператора UniMAT": "2273",
    "ПЛК UniMAT": "2273",
    "Программируемые логические контроллеры UniMAT": "2273",
    "Серво": "2585"
}

# Создаем копию словаря в нижнем регистре, чтобы поиск не зависел от заглавных букв
NORMALIZED_GROUP_MAP = {k.strip().lower(): v for k, v in GROUP_MAP.items()}

# Статичный список 5 фотографий для товаров бренда Unimat
UNIMAT_PICTURES = [
    "https://unimat-russia.ru/uploads/product-1654003344077-0.4960815358606392.png",
    "https://unimat-russia.ru/uploads/product-1654005665354-0.5424921694625866.jpg",
    "https://unimat-russia.ru/uploads/product-1703188936539-0.01815614639060681.jpg",
    "https://unimat-russia.ru/uploads/product-1654002861798-0.35138493486299605.jpg",
    "https://unimat-russia.ru/uploads/product-33.png"
]

# Статичный список 9 PDF-документов для товаров бренда Unimat
UNIMAT_PDFS = [
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/HMI%20Catalogue%206-18.pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20120%20Series%20PLC%20(1-2).pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20120%20Series%20PLC%20(61-82(%E6%9B%B4%E6%96%B0).pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%201200%20series%20PLC.pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20200%20Series%20PLC%20(29-60(%E6%9B%B4%E6%96%B0%EF%BC%89).pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20300%20%D0%BC%D0%BE%D0%B4%D1%83%D0%BB%D0%B8%20(3-28).pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UniMAT%20%D0%9B%D0%B8%D1%81%D1%82%D0%BE%D0%B2%D0%BA%D0%B0.pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/Unimat%20-%202025%20%D0%B1%D1%80%D0%BE%D1%88%D1%8E%D1%80%D0%B0.pdf",
    f"https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/%D0%9C%D0%BE%D0%B4%D1%83%D0%BB%D0%B8%20UniMAT%20UN%20300%20%3D%20Siemens%20S7-300%20(100%25%20%D0%B0%D0%BD%D0%B0%D0%BB%D0%BE%D0%B3).pdf"
]

# Ссылка на общий файл сертификатов ТР ТС (используется в каждом товаре обоих брендов)
COMMON_CERT_RAW_URL = f"https://github.com/brilka/{REPO_NAME}/raw/cc58148495a4db33e65ae255e720914c295e4fb0/%D0%A1%D1%81%D1%8B%D0%BB%D0%BA%D0%B0%20%D0%BD%D0%B0%20%D1%81%D0%B5%D1%80%D1%82%D0%B8%D1%84%D0%B8%D0%BA%D0%B0%D1%82%D1%8B%20%D0%A2%D0%A0%20%D0%A2%D0%A1.pdf"
COMMON_CERT_COMMENT_URL = f"https://github.com/brilka/{REPO_NAME}/blob/cc58148495a4db33e65ae255e720914c295e4fb0/%D0%A1%D1%81%D1%8B%D0%BB%D0%BA%D0%B0%20%D0%BD%D0%B0%20%D1%81%D0%B5%D1%80%D1%82%D0%B8%D1%84%D0%B8%D0%BA%D0%B0%D1%82%D1%8B%20%D0%A2%D0%A0%20%D0%A2%D0%A1.pdf"

# Генерируем 10 ссылок на запасные картинки (1.png - 10.png) в репозитории
DEFAULT_PROMPOWER_PICTURES = [
    f"{PAGES_BASE_URL}{i}.png" for i in range(1, 11)
]


def load_pdf_cache():
    """
    Загружает сохраненный ранее кэш ссылок на PDF.
    Если новый файл pdf_cache.json еще не создан, подхватывает данные из старого файла chipidip_pdf_cache.json.
    """
    target_file = CACHE_FILENAME if os.path.exists(CACHE_FILENAME) else (OLD_CACHE_FILENAME if os.path.exists(OLD_CACHE_FILENAME) else None)
    if target_file:
        try:
            with open(target_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось прочитать кэш ({e}), создаем новый.")
    return {"last_update_month": -1, "urls": {}}


def save_pdf_cache(cache_data):
    """
    Сохраняет обновленный кэш PDF в файл pdf_cache.json.
    В режиме отладки сохранение отключено, чтобы тестовый запуск не перезаписал полный кэш.
    """
    if DEBUG_MODE:
        return
    try:
        with open(CACHE_FILENAME, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[ОШИБКА] Не удалось сохранить кэш PDF: {e}")


def make_api_request(endpoint):
    """Отправляет POST-запрос с ключами авторизации к API Prompower и возвращает список товаров."""
    url = f"{API_URL}{endpoint}"
    payload = {"email": EMAIL, "key": KEY, "format": "json"}
    headers = {"Content-type": "application/json"}
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"[ОШИБКА API] при запросе к {url}: {e}")
        return []


def get_categories_dict():
    """Скачивает дерево категорий Prompower (и при наличии Unimat) и превращает в словарь {ID: данные}."""
    categories = {}
    try:
        resp = requests.get("https://prompower.ru/api/categories", timeout=30)
        if resp.status_code == 200:
            for cat in resp.json():
                categories[int(cat['id'])] = {
                    'title': cat.get('title', 'Без названия'),
                    'parentId': cat.get('parentId')
                }
    except Exception as e:
        print(f"[ОШИБКА] Не удалось загрузить категории Prompower: {e}")

    endpoints_to_try = ["https://prompower.ru/api/unimatCategories", "https://prompower.ru/api/unimat-categories"]
    for ep in endpoints_to_try:
        try:
            resp = requests.get(ep, timeout=10)
            if resp.status_code == 200 and isinstance(resp.json(), list):
                for cat in resp.json():
                    categories[int(cat['id'])] = {
                        'title': cat.get('title', 'Без названия'),
                        'parentId': cat.get('parentId')
                    }
        except Exception:
            continue
    return categories


def scrape_docs(url):
    """Заходит на веб-страницу товара (или раздела), ищет там ссылки на PDF-файлы и их названия."""
    docs = []
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            for a_tag in soup.find_all('a', href=True):
                href = a_tag['href']
                if '.pdf' in href.lower():
                    div_tag = a_tag.find('div', class_=lambda x: x and "text-caption" in x)
                    doc_name = div_tag.text.strip() if div_tag else "Документация"
                    full_link = SITE_URL + href if href.startswith('/') else href
                    if not any(d['url'] == full_link for d in docs):
                        docs.append({"url": full_link, "name": doc_name})
    except Exception:
        pass
    return docs


# ==============================================================================
# 3. ФИНАНСОВЫЕ МОДЕЛИ РАСЧЕТА ЦЕН (ДЛЯ ЧИП И ДИП И ДЛЯ STOKING)
# ==============================================================================

# СЛУЖЕБНЫЙ КОММЕНТАРИЙ ПО ЦЕНООБРАЗОВАНИЮ ДЛЯ ВНУТРЕННЕГО ИСПОЛЬЗОВАНИЯ:
"""
<!--  Для Чип и Дип: 
costPrice - цена продажи (за единицу измерения) поставщика для Чип и Дип. Для Prompower и Unimat costPrice определяется так:
Вариант1. если MRPPercent в API для данной позиции отсутствует или равен 0, то нужно price (из API Prompower) умножить на НДС (1.22) и полученное значение разделить на 0.85. 
Вариант2. если MRPPercent в API для данной позиции присутствует и не равен 0, то логика следующая: 
costPrice=A в том случае, если B/A меньше или равно 0.85, где 
A считается как price (из API Prompower) умножить на НДС (1.22) и разделить на коэффициент K (изначально при расчёте для каждой позиции K=1.3); 
B считается как price (из API Prompower) умножить на НДС (1.22) и умножить на (100-discountPercent)/100. discountPercent взять из API Prompower. 
Если B/A выше 0.85 - то нужно пересчитать A, уменьшая коэффициент K с шагом 0.01 (т.е. 1.3, 1.29, 1.28 и т.д.) до тех пор, пока соотношение B/A не станет меньше или равно 0.85.

rPrice - рекомендуемая цена продажи на сайте Чип и Дип. Не обязательное поле (если не указано - будет использовано заданное ценообразование для поставщика, т.е. costPrice*1.25). Для всех позиций Prompower и Unimat rPrice равен price (из API Prompower) умножить на НДС (1.22). 


Для Stoking:
costPrice - цена продажи (за единицу измерения) поставщика для Stoking. Т.е. цена, по которой Поставщик продаёт товар компании Stoking. Для Prompower и Unimat costPrice определяется так: 
Вариант1. если MRPPercent в API для данной позиции отсутствует или равен 0, то нужно price (из API Prompower) умножить на НДС (1.22) и полученное значение разделить на 0.8. 
Вариант2. если MRPPercent в API для данной позиции присутствует и не равен 0, то логика следующая: 
costPrice=A в том случае, если B/A меньше или равно 0.8, где 
A считается как price (из API Prompower) умножить на НДС (1.22) и разделить на коэффициент K (изначально при расчёте для каждой позиции K=1.2); 
B считается как price (из API Prompower) умножить на НДС (1.22) и умножить на (100-discountPercent)/100. discountPercent взять из API Prompower. 
Если B/A выше 0.8 - то нужно пересчитать A, уменьшая коэффициент K с шагом 0.01 (т.е. 1.2, 1.19, 1.18 и т.д.) до тех пор, пока соотношение B/A не станет меньше или равно 0.8.

rPrice - рекомендуемая (прайсовая) цена продажи на сайте Stoking. В интернет-магазине нельзя указывать цену ниже рекомендованной, т.е. ниже прайсовой. Поэтому в интернет-магазине Stoking должна быть указана цена не ниже rPrice. Для всех позиций Prompower и Unimat rPrice равен price (из API Prompower) умножить на НДС (1.22). Если costPrice определяется по Варианту 1, то rPrice равен costPrice/0.9   -->
"""

def calculate_prices(price_val, mrp_percent, discount_percent, client_key):
    """
    Рассчитывает закупочную цену (costPrice) и рекомендованную цену (rPrice) 
    в зависимости от выбранного клиента ('chipdip' или 'stoking').
    Возвращает два числа, округленных до 2 знаков после запятой: (cost_price, r_price).
    """
    # Умножаем базовую цену из API на НДС 22% (1.22) — это база для всех расчетов
    price_with_vat = price_val * 1.22

    # Выбираем параметры финансовой модели под конкретного клиента
    if client_key == "chipdip":
        base_divisor_var1 = 0.85   # Делитель для Варианта 1 (без РРЦ)
        initial_k = 1.3            # Начальный коэффициент K для Варианта 2 (обновлено с 1.25 на 1.3)
        target_ratio = 0.85        # Максимально допустимое отношение B/A
    elif client_key == "stoking":
        base_divisor_var1 = 0.8    # Делитель для Варианта 1 (без РРЦ)
        initial_k = 1.2            # Начальный коэффициент K для Варианта 2
        target_ratio = 0.8         # Максимально допустимое отношение B/A
    else:
        raise ValueError(f"Неизвестный клиент: {client_key}")

    # Вычисляем costPrice и rPrice по Варианту 1 или Варианту 2
    if mrp_percent == 0:
        # ВАРИАНТ 1: MRPPercent отсутствует или равен 0
        cost_price = price_with_vat / base_divisor_var1

        if client_key == "stoking":
            # Для Stoking в Варианте 1: rPrice = costPrice / 0.9
            r_price = cost_price / 0.9
        else:
            # Для Чип и Дип: rPrice = price * 1.22
            r_price = price_with_vat
    else:
        # ВАРИАНТ 2: MRPPercent присутствует и не равен 0
        K = initial_k
        # Вычисляем B = (price * 1.22) * (100 - discountPercent) / 100
        B = price_with_vat * ((100.0 - discount_percent) / 100.0)

        # Постепенно снижаем K с шагом 0.01, пока B/A не станет меньше или равно target_ratio
        while True:
            # Защита от бесконечного цикла, если K упрется в ноль
            if K <= 0.01:
                A = B / target_ratio
                break

            A = price_with_vat / K
            if A == 0:
                break

            # Проверяем соотношение B/A
            if (B / A) > target_ratio:
                K = round(K - 0.01, 2)  # Округляем до сотых, чтобы избежать погрешностей Python
            else:
                break

        cost_price = A
        # В Варианте 2 для обоих клиентов rPrice = price * 1.22
        r_price = price_with_vat

    return round(cost_price, 2), round(r_price, 2)


# ==============================================================================
# 4. ПОИСК КАТЕГОРИИ, ДОКУМЕНТОВ И ФОРМИРОВАНИЕ XML ДЛЯ ОДНОГО КЛИЕНТА
# ==============================================================================

def find_chipdip_item_group(prod, categories_dict):
    """
    Определяет код группы товара (itemGroupId) для каталога Чип и Дип.
    Сначала ищет по текстовому названию 'category' в товаре, а при неудаче поднимается по дереву ID.
    """
    item_group_id = ""
    cat_id = prod.get('categoryId', '')

    cat_raw = None
    for k, v in prod.items():
        if k.lower() == 'category':
            cat_raw = v
            break

    direct_category_name = ""
    if isinstance(cat_raw, str):
        direct_category_name = cat_raw
    elif isinstance(cat_raw, dict):
        direct_category_name = cat_raw.get('title', '') or cat_raw.get('name', '')
    elif isinstance(cat_raw, list) and len(cat_raw) > 0:
        direct_category_name = str(cat_raw[0])

    direct_category_name = direct_category_name.strip()

    if direct_category_name:
        item_group_id = NORMALIZED_GROUP_MAP.get(direct_category_name.lower(), "")

    if not item_group_id and cat_id and int(cat_id) in categories_dict:
        current_id = int(cat_id)
        for _ in range(5):
            if current_id not in categories_dict:
                break
            cat_data = categories_dict[current_id]
            cat_name = cat_data['title'].strip()
            item_group_id = NORMALIZED_GROUP_MAP.get(cat_name.lower(), "")
            if item_group_id:
                break
            if cat_data.get('parentId'):
                current_id = int(cat_data['parentId'])
            else:
                break

    return item_group_id


def get_prompower_docs(url, pdf_cache, need_global_pdf_update, scraped_this_run):
    """
    Собирает список PDF-документов для товара Prompower (со страницы товара + со страницы на уровень выше + референс-лист).
    Использует множество scraped_this_run, чтобы не сканировать одну и ту же страницу дважды за один запуск скрипта.
    """
    all_product_docs = []

    if (need_global_pdf_update and url not in scraped_this_run) or (url not in pdf_cache["urls"]):
        product_docs = scrape_docs(url)
        if product_docs or url not in pdf_cache["urls"]:
            pdf_cache["urls"][url] = product_docs
        scraped_this_run.add(url)
    else:
        product_docs = pdf_cache["urls"][url]
    all_product_docs.extend(product_docs)

    parent_url = "/".join(url.split("/")[:-1])
    if (need_global_pdf_update and parent_url not in scraped_this_run) or (parent_url not in pdf_cache["urls"]):
        parent_docs = scrape_docs(parent_url)
        if parent_docs or parent_url not in pdf_cache["urls"]:
            pdf_cache["urls"][parent_url] = parent_docs
        scraped_this_run.add(parent_url)
    else:
        parent_docs = pdf_cache["urls"][parent_url]

    existing_urls = [d['url'] for d in all_product_docs]
    for p_doc in parent_docs:
        if p_doc['url'] not in existing_urls:
            all_product_docs.append(p_doc)

    has_ref = any(d['name'] == "Краткий референс-лист проектов по автоматизации" for d in all_product_docs)
    if not has_ref:
        all_product_docs.append({
            "url": "https://prompower.ru/docs/referenceList.pdf",
            "name": "Краткий референс-лист проектов по автоматизации"
        })

    return all_product_docs


def process_products_for_client(products, brand, categories_dict, pdf_cache, need_global_pdf_update, scraped_this_run, is_first_offer, client_key, client_snapshot):
    """
    Преобразует список товаров одного бренда в готовые XML-блоки <offer> 
    с учетом правил, комментариев и тегов конкретного клиента ('chipdip' или 'stoking'),
    а также наполняет словарь client_snapshot данными для последующего логирования изменений.
    """
    items_xml = []
    param_regex = re.compile(r"^(.*?)(?:\s*\((.*?)\))?$")

    for prod in products:
        # Шаг 1: Проверяем обязательные поля (наличие артикула и ненулевой цены)
        article = str(prod.get('article', '')).strip()
        if not article:
            continue

        raw_price = prod.get('price')
        if not raw_price or float(raw_price) <= 0:
            continue

        # Шаг 2: Извлекаем и очищаем исходные данные товара (используем "or 0" для защиты от null)
        price_val = float(raw_price)
        mrp_percent = float(prod.get('MRPPercent') or 0)
        discount_percent = float(prod.get('discountPercent') or 0)
        instock = str(prod.get('instock') if prod.get('instock') is not None else 0).strip()
        description = str(prod.get('description') or '')
        title = str(prod.get('title') or '')
        tnved = str(prod.get('tnved') or '').strip()
        cat_id = prod.get('categoryId', '')

        # Шаг 3: Формируем ссылку на страницу товара Prompower (для Unimat ссылки нет)
        url = ""
        if brand == "Prompower" and prod.get('path'):
            path_str = prod.get('path')
            if not path_str.startswith('/'):
                path_str = '/' + path_str
            url = f"{SITE_URL}/catalog{path_str}"

        # Шаг 4: Рассчитываем цены под финансовую модель текущего клиента
        cost_price, r_price = calculate_prices(price_val, mrp_percent, discount_percent, client_key)

        # Записываем текущие цены (price из API умножаем на НДС 1.22!) и остаток этого товара в «снимок дня» (для лога изменений)
        client_snapshot[article] = {
            "price": round(price_val * 1.22, 2),
            "costPrice": cost_price,
            "rPrice": r_price,
            "qty": instock
        }

        # Шаг 5: Начинаем собирать XML-блок <offer>
        offer_xml = ["<offer>"]

        # --- ID ---
        if is_first_offer:
            offer_xml.append('<!--  уникальный идентификатор товара поставщика. может быть буквенно-цифровой. используется для дальнейшей трансляции заказов поставщику. У Prompower и Unimat это article в API.  -->')
        offer_xml.append(f"<id>{xml_escape(article)}</id>")

        # --- QTY (Остаток) ---
        if is_first_offer:
            offer_xml.append('<!--  кол-во товара, доступное для продажи. В API Prompower и Unimat это instock -->')
        offer_xml.append(f"<qty>{instock}</qty>")

        # --- URL ---
        if url:
            if is_first_offer:
                if client_key == "stoking":
                    offer_xml.append('<!--  ссылка на карточку товара на сайте поставщика. -->')
                else:
                    offer_xml.append('<!--  ссылка на карточку товара на сайте поставщика. Используется для просмотра информации о товаре складом или отделом закупок Чип и Дип. У Prompower это значение в path в API (но в path в API указан неполный путь, например, /mcb/ESM163C12 - поэтому в начале нужно дописать www.prompower.ru/catalog ). Для Unimat url недоступен.  -->')
            offer_xml.append(f"<url>{xml_escape(url)}</url>")

        # --- PRICE (costPrice и rPrice) ---
        if is_first_offer:
            if client_key == "stoking":
                price_comment = """<!--  costPrice – цена продажи (за единицу измерения) поставщика для Stoking. Т.е. цена, по которой Поставщик продаёт товар компании Stoking.
rPrice – рекомендуемая (прайсовая) цена продажи на сайте Stoking. В интернет-магазине нельзя указывать цену ниже рекомендованной, т.е. ниже прайсовой. Поэтому в интернет-магазине Stoking должна быть указана цена не ниже rPrice -->"""
            else:
                price_comment = """<!--  costPrice - цена продажи (за единицу измерения) поставщика для Чип и Дип. Для Prompower и Unimat costPrice определяется так: 
Вариант1. если MRPPercent в API для данной позиции отсутствует или равен 0, то нужно price (из API Prompower) умножить на НДС (1.22) и полученное значение разделить на 0.85. 
Вариант2. если MRPPercent в API для данной позиции присутствует и не равен 0, то логика следующая: 
costPrice=A в том случае, если B/A меньше или равно 0.85, где 
A считается как price (из API Prompower) умножить на НДС (1.22) и разделить на коэффициент K (изначально при расчёте для каждой позиции K=1.3); 
B считается как price (из API Prompower) умножить на НДС (1.22) и умножить на (100-discountPercent)/100. discountPercent взять из API Prompower. 
Если B/A выше 0.85 - то нужно пересчитать A, уменьшая коэффициент K с шагом 0.01 (т.е. 1.3, 1.29, 1.28 и т.д.) до тех пор, пока соотношение B/A не станет меньше или равно 0.85.

rPrice - рекомендуемая цена продажи на сайте Чип и Дип. Не обязательное поле (если не указано - будет использовано заданное ценообразование для поставщика, т.е. costPrice*1.25). Для всех позиций Prompower и Unimat rPrice равен price (из API Prompower) умножить на НДС (1.22).  -->"""
            offer_xml.append(price_comment)
        offer_xml.append(f'<price qty="1" costPrice="{cost_price}" rPrice="{r_price}"/>')

        # --- CATEGORY ID ---
        if cat_id:
            if is_first_offer:
                if client_key == "stoking":
                    offer_xml.append('<!--  принадлежность товара к категории поставщика. Код категории должен быть указан в списке categories. -->')
                else:
                    offer_xml.append('<!--  принадлежность товара к категории поставщика. Код категории должен быть указан в списке categories. Не обязателльное поле  -->')
            offer_xml.append(f"<categoryId>{cat_id}</categoryId>")

        # --- PICTURE (Фотографии товара) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  Список ссылок на фото данного товара. -->')
            else:
                offer_xml.append(f'<!--  Список ссылок на фото данного товара. Максимум 10 фото. Фото должны быть без водяных знаков. Не обязательное поле. Для Prompower фото загружаются по API - у разных товаров может быть разное количество фото: нужно предусмотреть, чтобы код правильно обработал подгрузку всех доступных фото. Если в каком-то img API Prompower нет значения, т.е. отсутствует ссылка на фото, то нужно подгрузить все фото, которые лежат здесь: {PAGES_BASE_URL} (в коде нужно указать все ссылки в списке DEFAULT_PROMPOWER_PICTURES ). Для Unimat у всех позиций нужно указать 5 фото со следующими адресами: https://unimat-russia.ru/uploads/product-1654003344077-0.4960815358606392.png ; https://unimat-russia.ru/uploads/product-1654005665354-0.5424921694625866.jpg ; https://unimat-russia.ru/uploads/product-1703188936539-0.01815614639060681.jpg ; https://unimat-russia.ru/uploads/product-1654002861798-0.35138493486299605.jpg ; https://unimat-russia.ru/uploads/product-33.png  -->')

        final_images = []
        if brand == "Unimat":
            final_images = UNIMAT_PICTURES
        else:
            api_images = prod.get('img') or []
            if isinstance(api_images, str):
                api_images = [api_images]
            if not api_images and prod.get('image'):
                api_images = [prod.get('image')]

            for img in api_images:
                if img and str(img).strip():
                    img_url = img if img.startswith('http') else SITE_URL + img
                    final_images.append(img_url)

            if len(final_images) == 0:
                final_images = DEFAULT_PROMPOWER_PICTURES

        for pic in final_images[:10]:
            offer_xml.append(f"<picture>{xml_escape(pic)}</picture>")

        # --- NAME (Наименование до 250 символов) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  Наименование товара. Для Prompower и Unimat это description в API  -->')
            else:
                offer_xml.append('<!--  Наименование товара. Макс. 250 символов. Обязателльное поле. Для Prompower и Unimat это description в API  -->')
        safe_name = (description if description else (title if title else "Товар без названия"))[:250]
        offer_xml.append(f"<name>{xml_escape(safe_name)}</name>")

        # --- PARTNUMBER (Оригинальный артикул производителя) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  Артикул (оригинальный парт номер) по каталогу производителя данного товара. Для Prompower и Unimat это title в API  -->')
            else:
                offer_xml.append('<!--  Артикул (оригинальный парт номер) по каталогу производителя данного товара. Не оябязательное поле. Для Prompower и Unimat это title в API  -->')
        if title:
            offer_xml.append(f"<partNumber>{xml_escape(title)}</partNumber>")

        # --- VENDOR (Производитель) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  Название производителя (бренда) товара. Для Prompower нужно указывать Prompower. Для Unimat нужно указывать Unimat  -->')
            else:
                offer_xml.append('<!--  Название производеителя (бренда) товара. Может быть полное или сокращенное название. Не оябязательное поле. Для Prompower нужно указывать Prompower. Для Unimat нужно указывать Unimat  -->')
        offer_xml.append(f"<vendor>{brand}</vendor>")

        # --- TNVEDCODE (Код ТН ВЭД) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  Код ТН ВЭД товара. У Prompower в API это tnved. -->')
            else:
                offer_xml.append('<!--  Код ТН ВЭД товара. Не обязательное поле. У Prompower в API это tnved. -->')
        if tnved:
            offer_xml.append(f"<tnvedcode>{xml_escape(tnved)}</tnvedcode>")

        # --- DESCRIPTION (Полное описание + доп. текст совместимости для Unimat) ---
        if is_first_offer:
            if client_key == "stoking":
                desc_comment = """<!--  Описание товара. Для Prompower и Unimat это description в API.
Для Unimat дополнительно добавляется следующий текст:
Модули Unimat на 100% совместимы (взаимозаменяемы) с соответствующими модулями Siemens S7-200, S7-300, S7-1200. Чтобы получить артикул Unimat, в большинстве случаев нужно у артикула Siemens заменить 6ES7- на UN-, а остальную часть артикула оставить без изменений.

Данный модуль Unimat [тут нужно вставить обозначение Unimat из тега partNumber] соответствует модулю Siemens [тут нужно указать то же обозначение из тега partNumber, только нужно заменить "UN " на "6ES7"].

Модули Unimat поддерживаются на складе в России. 
Кабель Profibus и ProfiNET от Unimat соответствует кабелю Siemens.

Модули UniMAT монтируются в оригинальный ПЛК Siemens, определяются средой разработки как оригинальные модули Siemens без каких-либо дополнительных манипуляций со стороны программиста.

Сценарии использования модулей Unimat: 

+ в новых проектах (с оригинальным CPU Siemens) 
+ в проектах модернизации производства, где установлены ПЛК Siemens S7-200, S7-300 и S7-1200 и требуется их расширение (или замена модулей)
+ станции удалённого (распределённого) ввода-вывода от Unimat (Profibus; Profinet)
  -->"""
            else:
                desc_comment = """<!--  Описание товара. Может быть в формате html. Не обязательное поле. Для Prompower и Unimat это description в API. 
Для Unimat дополнительно добавляется следующий текст:

Модули Unimat на 100% совместимы (взаимозаменяемы) с соответствующими модулями Siemens S7-200, S7-300, S7-1200. Чтобы получить артикул Unimat, нужно у артикула Siemens заменить 6ES7- на UN-, а остальную часть артикула оставить без изменений.

Данный модуль Unimat [тут нужно вставить обозначение Unimat из тега partNumber] соответствует модулю Siemens [тут нужно указать то же обозначение из тега partNumber, только нужно заменить "UN " на "6ES7"].

Модули Unimat поддерживаются на складе в России. 
Кабель Profibus и ProfiNET от Unimat соответствует кабелю Siemens.

Модули UniMAT монтируются в оригинальный ПЛК Siemens, определяются средой разработки как оригинальные модули Siemens без каких-либо дополнительных манипуляций со стороны программиста.

Сценарии использования модулей Unimat: 
+ в новых проектах (с оригинальным CPU Siemens) 
+ в проектах модернизации производства, где установлены ПЛК Siemens S7-200, S7-300 и S7-1200 и требуется их расширение (или замена модулей)
+ станции удалённого (распределённого) ввода-вывода от Unimat (Profibus; Profinet)
  -->"""
            offer_xml.append(desc_comment)

        final_desc = description
        if brand == "Unimat" and title:
            siemens_title = title.replace("UN ", "6ES7")
            if client_key == "stoking":
                unimat_addition = f"""<br><br>Модули Unimat на 100% совместимы (взаимозаменяемы) с соответствующими модулями Siemens S7-200, S7-300, S7-1200. Чтобы получить артикул Unimat, в большинстве случаев нужно у артикула Siemens заменить 6ES7- на UN-, а остальную часть артикула оставить без изменений.<br><br>Данный модуль Unimat {title} соответствует модулю Siemens {siemens_title}.<br><br>Модули Unimat поддерживаются на складе в России.<br>Кабель Profibus и ProfiNET от Unimat соответствует кабелю Siemens.<br><br>Модули UniMAT монтируются в оригинальный ПЛК Siemens, определяются средой разработки как оригинальные модули Siemens без каких-либо дополнительных манипуляций со стороны программиста.<br><br>Сценарии использования модулей Unimat:<br><br>+ в новых проектах (с оригинальным CPU Siemens)<br>+ в проектах модернизации производства, где установлены ПЛК Siemens S7-200, S7-300 и S7-1200 и требуется их расширение (или замена модулей)<br>+ станции удалённого (распределённого) ввода-вывода от Unimat (Profibus; Profinet)"""
            else:
                unimat_addition = f"""<br><br>Модули Unimat на 100% совместимы (взаимозаменяемы) с соответствующими модулями Siemens S7-200, S7-300, S7-1200. Чтобы получить артикул Unimat, нужно у артикула Siemens заменить 6ES7- на UN-, а остальную часть артикула оставить без изменений.<br><br>Данный модуль Unimat {title} соответствует модулю Siemens {siemens_title}.<br><br>Модули Unimat поддерживаются на складе в России.<br>Кабель Profibus и ProfiNET от Unimat соответствует кабелю Siemens.<br><br>Модули UniMAT монтируются в оригинальный ПЛК Siemens, определяются средой разработки как оригинальные модули Siemens без каких-либо дополнительных манипуляций со стороны программиста.<br><br>Сценарии использования модулей Unimat:<br>+ в новых проектах (с оригинальным CPU Siemens)<br>+ в проектах модернизации производства, где установлены ПЛК Siemens S7-200, S7-300 и S7-1200 и требуется их расширение (или замена модулей)<br>+ станции удалённого (распределённого) ввода-вывода от Unimat (Profibus; Profinet)"""
            final_desc += unimat_addition

        if final_desc:
            offer_xml.append(f"<description><![CDATA[{final_desc}]]></description>")

        # --- PARAM (Технические характеристики, макс. 20 шт.) ---
        if is_first_offer:
            if client_key == "stoking":
                offer_xml.append('<!--  список параметров товара. -->')
            else:
                offer_xml.append('<!--  список параметров товара. Максимум 20 параметров для одного товара. -->')
        props_list = (prod.get('props') or [])[:20]
        for prop in props_list:
            p_name, p_val = prop.get('name', ''), prop.get('value', '')
            match = param_regex.match(p_name)
            clean_name = match.group(1).strip() if match else p_name
            unit = match.group(2) if match and match.group(2) else ""
            offer_xml.append(f'<param name="{xml_escape(clean_name)}" unit="{xml_escape(unit)}">{xml_escape(str(p_val))}</param>')

        # --- DOCFILE (Документация и сертификаты) ---
        if is_first_offer:
            if client_key == "stoking":
                doc_comment = '<!--  Дополнительные файлы для скачивания. Например, чертеж товара, файл документации (datasheet), инструкция по использованию. -->'
            else:
                doc_comment = f"""<!--  Дополнительные файлы для скачивания. Например чертеж товара, файл документации (datasheet), инструкция по использованию. Файлы не должны содержать ссылок и логотипов на сайт поставщика. Допустимый формат файла - pdf, docx.  Не обязательное поле. 
Для Prompower доступные файлы pdf определяются следующим образом: нужно спарсить их названия и адрес со страницы, которая ранее получилась в тегах url. На этих страницах у некоторых товаров есть доступные файлы pdf. В коде страницы они расположены в разделе после названия раздела "Техническая документация и материалы для скачивания" и в следующем разделе после названия раздела "Чертежи, 3D-модели" (в class="v-list-group__items"). Сначала встречается ссылка на файл, например, href="/docs/pulse-nka/LVC_PULSE_Catalog.pdf" (ссылка неполная, поэтому в начале ещё нужно добавить www.prompower.ru), затем, после тега <div class="text-caption col col-10"> встречается название файла, например, "Низковольтная коммутационная аппаратура PROMPOWER. Спецификация продукта". 
Дополнительно нужно добавить файлы pdf, которые отсутствуют на странице конкретного товара, но есть на странице сайта на уровень выше - парсить их нужно по такому же принципу, который описан выше.
Если для какой-либо позиции Prompower после парсинга нет файла с названием "Краткий референс-лист проектов по автоматизации", то нужно добавить такой файл, его адрес https://prompower.ru/docs/referenceList.pdf 
Парсинг с кэшированием (обновляется 1-го числа каждого месяца).

Для позиций Unimat нужно добавлять следующие файлы:
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/HMI%20Catalogue%206-18.pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20120%20Series%20PLC%20(1-2).pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20120%20Series%20PLC%20(61-82(%E6%9B%B4%E6%96%B0).pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%201200%20series%20PLC.pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20200%20Series%20PLC%20(29-60(%E6%9B%B4%E6%96%B0%EF%BC%89).pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UN%20300%20%D0%BC%D0%BE%D0%B4%D1%83%D0%BB%D0%B8%20(3-28).pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/UniMAT%20%D0%9B%D0%B8%D1%81%D1%82%D0%BE%D0%B2%D0%BA%D0%B0.pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/Unimat%20-%202025%20%D0%B1%D1%80%D0%BE%D1%88%D1%8E%D1%80%D0%B0.pdf
https://github.com/brilka/{REPO_NAME}/blob/main/Unimat/%D0%9C%D0%BE%D0%B4%D1%83%D0%BB%D0%B8%20UniMAT%20UN%20300%20%3D%20Siemens%20S7-300%20(100%25%20%D0%B0%D0%BD%D0%B0%D0%BB%D0%BE%D0%B3).pdf
название файлов указано после последнего слэш до ".pdf"

Также для всех товаров Prompower и Unimat добавляется файл (документ со ссылкой на Яндекс Диск с сертификатами ТР ТС):
{COMMON_CERT_COMMENT_URL}
  -->"""
            offer_xml.append(doc_comment)

        if brand == "Prompower" and url:
            all_product_docs = get_prompower_docs(url, pdf_cache, need_global_pdf_update, scraped_this_run)
            for doc in all_product_docs:
                offer_xml.append(f'<docFile url="{xml_escape(doc["url"])}" name="{xml_escape(doc["name"])}"/>')
        elif brand == "Unimat":
            for pdf_url in UNIMAT_PDFS:
                file_name_encoded = pdf_url.split('/')[-1].replace('.pdf', '')
                file_name = urllib.parse.unquote(file_name_encoded)
                download_url = pdf_url.replace("/blob/main/", "/raw/main/")
                offer_xml.append(f'<docFile url="{xml_escape(download_url)}" name="{xml_escape(file_name)}"/>')

        offer_xml.append(f'<docFile url="{xml_escape(COMMON_CERT_RAW_URL)}" name="Ссылка на сертификаты ТР ТС"/>')

        # --- ITEMGROUPID (Только для Чип и Дип; для Stoking удален) ---
        if client_key != "stoking":
            item_group_id = find_chipdip_item_group(prod, categories_dict)
            if is_first_offer:
                long_comment = """<!--  Код группы товара из каталога Чип и Дип. Если указан - товар будет размещен в данный раздел товара сайта Чип и Дип. Не обязательное поле. Для Prompower и Unimat вот сопоставление кодов и категорий: 
3073;19“ комплектующие;
3073;Аксессуары;
3059;Двери;
2584;Коллаборативные роботы;
2178;Модификационные комплекты для моторов;
2273;Модули расширения ПЛК PMP301;
3054;Монтажные панели;
2926;Моторные дроссели;
2178;Моторы;
2585;Серводвигатели;
2926;Сетевые дроссели;
3073;Соединительные комплекты;
3073;Шасси;
3073;Шкафы электротехнические;
2178;Опции для двигателей;
2273;Аксессуары для ПЛК;
2459;Аксессуары для реле;
2947;Дополнительные контактные приставки;
3073;Заземление;
2954;Опции для преобразователей частоты;
2947;Дополнительные контактные приставки PULSE;
3010;Колодки для реле;
3065;Прокладка кабеля;
3109;MCB (Miniature Circuit Breaker);
1559;Реле общего назначения;
2947;Контакторы PULSE;
3067;Панели основания;
2585;Аксессуары для сервосистем;
2947;Контакторы;
1559;Миниатюрные силовые реле;
2598;Сувенирная продукция;
3624;Реле тонкие;
1266;Кабели для датчиков;
2947;Миниконтакторы;
2947;Миниконтакторы PULSE;
3062;Цоколи;
3184;Климат + Свет;
2939;Блок питания HDR в пластиковом корпусе;
3124;Пластроны;
2939;Блок питания MDR в пластиковом корпусе;
3069;Боковые панели;
3062;Секционирование;
1266;Индуктивные датчики;
3115;Дополнительные контактные приставки для MCB;
2968;Опции для устройств плавного пуска;
2273;Модули расширения ПЛК PMP20/PMP30;
2939;Блок питания NDR в металлическом корпусе;
2930;Автоматы защиты двигателя PULSE;
2744;Фотоэлектрические датчики;
3071;Полки;
3067;Потолочные панели;
2954;Преобразователи частоты PD100;
2954;Преобразователи частоты PD101;
1667;Тормозные резисторы;
2954;Преобразователи частоты PD150;
2622;Панели оператора PH1;
3069;Задние панели;
3256;Промышленные коммутаторы;
2273;Панели оператора PH;
2926;ЭМС фильтры;
3062;Сейсмостойкость;
2926;Дроссели dU/dt;
2968;Устройства плавного пуска P2S 050;
2926;Дроссели для цепей постоянного тока;
2954;Преобразователи частоты PD210;
2954;Преобразователи частоты PD110;
2968;Устройства плавного пуска P2S 100;
2585;Сервоприводы;
1556;Регуляторы мощности;
2273;Программируемые логические контроллеры PMP20;
2954;Преобразователи частоты PD310;
2178;Электродвигатели класс энергоэфф. IE1;
2273;ПЛК PMP301;
3072;Каркасы;
2926;Синус-фильтры;
2954;Внешние тормозные модули для ПЧ;
2954;Преобразователи частоты PD310 IP54;
2273;Промышленный монитор;
2968;Устройства плавного пуска P2S 300;
2273;Промышленный ПК;
2273;ПЛК PMP30;
2273;Панельный ПК;
2273;Кабели и аксессуары;
2273;Модули для ПЛК;
2273;Панели оператора UniMAT;
2273;ПЛК UniMAT;
2585;Серво;
   -->"""
                offer_xml.append(long_comment)
            if item_group_id:
                offer_xml.append(f"<itemGroupId>{item_group_id}</itemGroupId>")

        # --- UNITID (Единица измерения) ---
        if is_first_offer:
            offer_xml.append('<!--  Единица измерения товара. Допустимые значения (соглано ОКЕИ) -->')
        offer_xml.append("<unitID>шт</unitID>")

        # --- WEIGHT (Вес в граммах, переводим из кг умножением на 1000) ---
        if is_first_offer:
            offer_xml.append('<!--  Вес товара в граммах. Используется для вычисления тарифов по доставке товара. -->')
        weight = prod.get('weight')
        if brand == "Prompower" and weight:
            try:
                weight_grams = int(float(weight) * 1000)
                offer_xml.append(f"<weight>{weight_grams}</weight>")
            except (ValueError, TypeError):
                offer_xml.append(f"<weight>{xml_escape(str(weight))}</weight>")

        # --- WIDTH, HEIGHT, DEPTH (Габариты товара в миллиметрах из props) ---
        item_width = None
        item_height = None
        item_depth = None

        if brand == "Prompower":
            for prop in (prod.get('props') or []):
                p_name = prop.get('name', '').strip().lower()
                p_val = prop.get('value')

                if p_val in [0, 0.0, "0", "", None]:
                    continue

                if p_name in ['ширина (мм)', 'ширина']:
                    if item_width is None:
                        item_width = p_val
                elif p_name in ['высота (мм)', 'высота']:
                    if item_height is None:
                        item_height = p_val
                elif p_name in ['глубина (мм)', 'глубина']:
                    if item_depth is None:
                        item_depth = p_val

        if is_first_offer:
            offer_xml.append('<!--  Ширина товара, в миллиметрах. В API Prompower находится в props среди остальных записей. У Unimat отсутствуют данные. -->')
        if item_width is not None:
            offer_xml.append(f"<width>{xml_escape(str(item_width))}</width>")

        if is_first_offer:
            offer_xml.append('<!--  Высота товара, в миллиметрах. В API Prompower находится в props среди остальных записей. У Unimat отсутствуют данные. -->')
        if item_height is not None:
            offer_xml.append(f"<height>{xml_escape(str(item_height))}</height>")

        if is_first_offer:
            offer_xml.append('<!--  Глубина товара, в миллиметрах. В API Prompower находится в props среди остальных записей. У Unimat отсутствуют данные. -->')
        if item_depth is not None:
            offer_xml.append(f"<depth>{xml_escape(str(item_depth))}</depth>")

        offer_xml.append("</offer>")
        items_xml.append("\n".join(offer_xml))
        is_first_offer = False

    return items_xml, is_first_offer


def generate_client_feeds(client_cfg, prompower_products, unimat_products, categories_dict, pdf_cache, need_global_pdf_update, scraped_this_run):
    """
    Собирает полный XML-файл для одного клиента (основной + аварийный ZEROwarehouse),
    добавляет комментарий с датой и временем создания (МСК) в шапку, сохраняет файлы на диск
    и возвращает словарь актуальных цен/остатков (snapshot) для лога.
    """
    client_key = client_cfg["key"]
    xml_filename = client_cfg["xml_file"]
    zero_xml_filename = client_cfg["zero_xml_file"]

    print(f"\n--- Формирование фида для клиента: {client_key} ({xml_filename}) ---")

    all_offers_xml = []
    client_snapshot = {}
    is_first_offer = True

    if prompower_products:
        xml_data, is_first_offer = process_products_for_client(
            prompower_products, "Prompower", categories_dict, pdf_cache, need_global_pdf_update, scraped_this_run, is_first_offer, client_key, client_snapshot
        )
        all_offers_xml.extend(xml_data)

    if unimat_products:
        xml_data, is_first_offer = process_products_for_client(
            unimat_products, "Unimat", categories_dict, pdf_cache, need_global_pdf_update, scraped_this_run, is_first_offer, client_key, client_snapshot
        )
        all_offers_xml.extend(xml_data)

    # Вычисляем текущую дату и время по Москве (UTC + 3 часа) для шапки XML
    msk_now = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=3)
    feed_datetime_str = msk_now.strftime("%H:%M %d.%m.%Y")

    xml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<shop>',
        '<name>Prompower и Unimat</name>',
        '<company>Мотрум</company>',
        f'<url>{PAGES_BASE_URL}</url>',
        f'<!-- Дата и время создания XML-feed (МСК): {feed_datetime_str} -->',
        '<categories>'
    ]

    for cat_id, data in categories_dict.items():
        parent_attr = f' parentId="{data["parentId"]}"' if data['parentId'] else ''
        xml_lines.append(f'<category id="{cat_id}"{parent_attr}>{xml_escape(data["title"])}</category>')

    xml_lines.append('</categories>')
    xml_lines.append('<!--  список товаров к продаже  -->')
    xml_lines.append('<offers>')
    xml_lines.extend(all_offers_xml)
    xml_lines.append('</offers>')
    xml_lines.append('</shop>')

    try:
        with open(xml_filename, "w", encoding="utf-8") as f:
            f.write("\n".join(xml_lines))
        print(f"[УСПЕХ] Файл {xml_filename} сгенерирован! (Обработано позиций: {len(all_offers_xml)})")
    except Exception as e:
        print(f"[ОШИБКА] Не удалось сохранить файл {xml_filename}: {e}")

    zero_xml_lines = []
    for line in xml_lines:
        if "<qty>" in line:
            line = re.sub(r'<qty>.*?</qty>', '<qty>0</qty>', line)
        zero_xml_lines.append(line)

    try:
        with open(zero_xml_filename, "w", encoding="utf-8") as f:
            f.write("\n".join(zero_xml_lines))
        print(f"[УСПЕХ] Аварийный файл {zero_xml_filename} сгенерирован!")
    except Exception as e:
        print(f"[ОШИБКА] Не удалось сохранить аварийный файл {zero_xml_filename}: {e}")

    return client_snapshot


# ==============================================================================
# 5. ВЕДЕНИЕ УМНОГО ДЕЛЬТА-ЛОГА ЦЕН И ОСТАТКОВ (ДЛЯ EXCEL В ФОРМАТЕ CSV)
# ==============================================================================

def update_price_stock_log(all_clients_snapshot):
    """
    Сравнивает сегодняшние цены (price с НДС 1.22, costPrice, rPrice) и остатки со вчерашними (из logs/last_state.json).
    В файл logs/history_ГГГГ-ММ.csv дописываются ТОЛЬКО те товары, по которым произошли изменения
    (новая позиция, изменилась какая-то цена/остаток, либо товар пропал из фида).
    В режиме отладки функция НЕ запускается, чтобы не портить боевую историю.
    """
    if DEBUG_MODE:
        print("\n[ЛОГ ИЗМЕНЕНИЙ] Включен режим отладки — запись в лог цен и остатков отключена.")
        return

    # Создаем папку logs, если ее еще нет
    os.makedirs(LOGS_DIR, exist_ok=True)

    # Загружаем предыдущее состояние (вчерашний снимок)
    last_state = {}
    if os.path.exists(STATE_FILENAME):
        try:
            with open(STATE_FILENAME, "r", encoding="utf-8") as f:
                last_state = json.load(f)
        except Exception as e:
            print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось прочитать {STATE_FILENAME} ({e}). Начинаем историю с нуля.")

    # Определяем текущую дату и время по Москве (UTC + 3 часа)
    msk_now = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=3)
    date_str = msk_now.strftime("%d.%m.%Y")
    time_str = msk_now.strftime("%H:%M:%S")
    month_str = msk_now.strftime("%Y-%m")
    csv_filename = os.path.join(LOGS_DIR, f"history_{month_str}.csv")

    new_rows = []
    # Переводим названия полей в понятные русские подписи для столбца "Событие"
    field_labels = {
        "price": "price (API с НДС)",
        "costPrice": "costPrice",
        "rPrice": "rPrice",
        "qty": "qty"
    }

    # Проходим по всем клиентам из нашего списка (сейчас: chipdip и stoking)
    for client_key, current_items in all_clients_snapshot.items():
        prev_items = last_state.get(client_key, {})

        # 1. Ищем новые товары и товары с изменившимися ценами/остатками
        for article, curr in current_items.items():
            if article not in prev_items:
                # Это новый товар (или самый первый запуск логирования)
                new_rows.append([
                    date_str,
                    time_str,
                    client_key,
                    article,
                    format_excel_num(curr["price"]),
                    format_excel_num(curr["costPrice"]),
                    format_excel_num(curr["rPrice"]),
                    curr["qty"],
                    "Первая запись / Новый товар"
                ])
            else:
                # Товар уже был раньше — сравниваем каждое поле
                prev = prev_items[article]
                changes = []
                for field in ["price", "costPrice", "rPrice", "qty"]:
                    if str(prev.get(field)) != str(curr.get(field)):
                        old_val_fmt = format_excel_num(prev.get(field)) if field != "qty" else prev.get(field)
                        new_val_fmt = format_excel_num(curr.get(field)) if field != "qty" else curr.get(field)
                        changes.append(f"{field_labels[field]}: {old_val_fmt} -> {new_val_fmt}")

                # Если хотя бы одно значение изменилось — добавляем строчку в лог
                if changes:
                    new_rows.append([
                        date_str,
                        time_str,
                        client_key,
                        article,
                        format_excel_num(curr["price"]),
                        format_excel_num(curr["costPrice"]),
                        format_excel_num(curr["rPrice"]),
                        curr["qty"],
                        "Изменение (" + ", ".join(changes) + ")"
                    ])

        # 2. Проверяем товары, которые были раньше, но сегодня пропали из фида (например, цена стала 0)
        for article, prev in prev_items.items():
            if article not in current_items:
                new_rows.append([
                    date_str,
                    time_str,
                    client_key,
                    article,
                    "0",
                    "0",
                    "0",
                    "0",
                    "Позиция пропала из фида (цена 0 или удалена)"
                ])

    # Если есть новые строки — дописываем их в помесячный CSV-файл
    if new_rows:
        # Проверяем, существует ли уже файл, чтобы кодировку utf-8-sig (с меткой BOM для Excel) записать ТОЛЬКО в самом начале файла
        file_exists = os.path.exists(csv_filename) and os.path.getsize(csv_filename) > 0
        encoding_to_use = "utf-8" if file_exists else "utf-8-sig"

        try:
            with open(csv_filename, "a", encoding=encoding_to_use, newline="") as f:
                writer = csv.writer(f, delimiter=";")
                # Если файл создается впервые, сначала пишем строку с заголовками столбцов
                if not file_exists:
                    writer.writerow([
                        "Дата", "Время (МСК)", "Client", "id (Артикул)",
                        "price (API с НДС)", "costPrice", "rPrice", "qty", "Событие"
                    ])
                writer.writerows(new_rows)
            print(f"[ЛОГ ИЗМЕНЕНИЙ] В файл {csv_filename} добавлено строк: {len(new_rows)}")
        except Exception as e:
            print(f"[ОШИБКА] Не удалось записать строки в {csv_filename}: {e}")
    else:
        print(f"[ЛОГ ИЗМЕНЕНИЙ] Изменений цен и остатков с прошлого запуска не обнаружено.")

    # Сохраняем текущее состояние как "вчерашнее" для следующего запуска
    try:
        with open(STATE_FILENAME, "w", encoding="utf-8") as f:
            json.dump(all_clients_snapshot, f, ensure_ascii=False)
    except Exception as e:
        print(f"[ОШИБКА] Не удалось сохранить {STATE_FILENAME}: {e}")


# ==============================================================================
# 6. ГЛАВНАЯ ТОЧКА ВХОДА В ПРОГРАММУ
# ==============================================================================

def main():
    start_time = time.time()
    print("=========================================")
    if DEBUG_MODE:
        print("!!! РЕЖИМ ОТЛАДКИ ВКЛЮЧЕН !!!")
        print(f"ВНИМАНИЕ: Скрипт обработает по {DEBUG_LIMIT} товаров каждого бренда. Запись в лог и кэш отключена.")
    print("=========================================")

    # Шаг 1: Один раз скачиваем категории и товары из API
    categories_dict = get_categories_dict()
    prompower_products = make_api_request("getProducts")
    unimat_products = make_api_request("getUnimatProducts")

    print(f"Загружено из API -> Prompower: {len(prompower_products)} шт., Unimat: {len(unimat_products)} шт.")

    # Если включен режим отладки, оставляем только первые DEBUG_LIMIT позиций
    if DEBUG_MODE:
        prompower_products = prompower_products[:DEBUG_LIMIT]
        unimat_products = unimat_products[:DEBUG_LIMIT]

    # Шаг 2: Загружаем кэш PDF-файлов и определяем, нужно ли сегодня сканировать сайт заново
    pdf_cache = load_pdf_cache()
    today = datetime.datetime.now()
    is_first_of_month = (today.day == 1)

    need_global_pdf_update = False
    if DEBUG_MODE:
        need_global_pdf_update = True
    elif is_first_of_month and pdf_cache.get("last_update_month") != today.month:
        need_global_pdf_update = True
        pdf_cache["last_update_month"] = today.month

    scraped_this_run = set()

    # Сюда соберем актуальные цены и остатки по всем клиентам для записи в лог
    all_clients_snapshot = {}

    # Шаг 3: По очереди генерируем фиды для всех клиентов из списка CLIENTS_CONFIG
    for client_cfg in CLIENTS_CONFIG:
        client_snapshot = generate_client_feeds(
            client_cfg,
            prompower_products,
            unimat_products,
            categories_dict,
            pdf_cache,
            need_global_pdf_update,
            scraped_this_run
        )
        all_clients_snapshot[client_cfg["key"]] = client_snapshot

    # Шаг 4: Сохраняем обновленный кэш PDF на диск
    save_pdf_cache(pdf_cache)

    # Шаг 5: Обновляем дельта-лог цен и остатков (только в обычном режиме, не в отладке)
    update_price_stock_log(all_clients_snapshot)

    print("\n=========================================")
    print(f"Все задачи успешно выполнены за {time.time() - start_time:.2f} секунд.")
    print("=========================================")


if __name__ == "__main__":
    main()
