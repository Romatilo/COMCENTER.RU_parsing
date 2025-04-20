from dotenv import load_dotenv
import urllib3
import requests
import pandas as pd
import json
import os
import re
from bs4 import BeautifulSoup
import uuid
from tqdm import tqdm
import datetime

# Путь к файлу сертификата
cert_path = "C:/!Work/COMCENTER/fullchain.pem"

# Путь для сохранения данных
output_dir = "COMCENTER.ru_database"
log_file = "comcenter_parser.log"
xls_url = "https://comcenter.ru/Content/PriceList/price.xls"
xls_output_file = os.path.join(output_dir, "DATABASE_recent.json")
printers_output_file = os.path.join(output_dir, "Laser_Printers.json")
compatibility_output_file = os.path.join(output_dir, "PRINTERS_compatibility.json")
compatibility_actual_output_file = os.path.join(output_dir, "PRINTERS_compatibility_actual.json")
cartridges_parts_output_file = os.path.join(output_dir, "DATABASE_cartridges&Parts.json")
all_cartridges_parts_output_file = os.path.join(output_dir, "DATABASE_all_cartridges&Parts.json")
comcenter_products_output_file = os.path.join(output_dir, "DATABASE_comcenter_products.json")

class ConsoleOutputHandler:
    """Обработчик вывода для консоли с записью в файл"""
    def log(self, message):
        print(message)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def progress(self, current, total):
        """Обработка прогресса для консоли"""
        percentage = (current / total) * 100
        print(f"Прогресс: {current}/{total} ({percentage:.1f}%)")

class CancelFlag:
    """Флаг для отслеживания отмены операции"""
    def __init__(self):
        self.cancelled = False

    def cancel(self):
        self.cancelled = True

    def is_cancelled(self):
        return self.cancelled

def setup_session(output_handler):
    """Настройка сессии с учетом сертификата и авторизации"""
    if not os.path.exists(cert_path):
        output_handler.log(f"Файл сертификата {cert_path} не найден")
        return None

    # Загружаем переменные из .env
    load_dotenv()
    LOGIN = os.getenv('COMCENTER.RU_LOGIN')
    PASSWORD = os.getenv('COMCENTER.RU_PASSWORD')

    if not LOGIN or not PASSWORD:
        output_handler.log("Логин или пароль не заданы в файле .env")
        return None

    # Создаем сессию
    session = requests.Session()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36',
        'Referer': 'https://comcenter.ru/',
        'Content-Type': 'application/x-www-form-urlencoded',
    }

    # Проверяем доступность сайта
    try:
        response = requests.get("https://comcenter.ru", headers=headers, verify=cert_path, timeout=10)
        if response.status_code != 200:
            output_handler.log("Не удалось подключиться к сайту comcenter.ru")
            return None
    except requests.exceptions.RequestException as e:
        output_handler.log(f"Ошибка при проверке сайта: {e}")
        return None

    # Авторизация
    login_url = 'https://comcenter.ru/Account/LogOn'
    login_data = {
        'UserName': LOGIN,
        'Password': PASSWORD,
        'RememberMe': 'false',
    }

    try:
        response = session.post(login_url, data=login_data, headers=headers, timeout=10, verify=cert_path)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        error_message = soup.find('h1', class_='dark-red-color')
        if error_message and "неверное имя или пароль" in error_message.text.lower():
            output_handler.log("Ошибка входа: неверный логин или пароль")
            return None
        output_handler.log("Успешно вошли в систему!")
        return session, headers
    except requests.exceptions.RequestException as e:
        output_handler.log(f"Ошибка при авторизации: {e}")
        return None

def get_laser_printers_database(session, headers, output_handler, cancel_flag):
    """Получение базы данных лазерных принтеров"""
    url = 'https://comcenter.ru/Store/Browse/400000006580/printery-lazernye-i-mfu'

    try:
        response = session.get(url, headers=headers, timeout=10, verify=cert_path)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        product_ids = []
        for a_tag in soup.select('a.cells-wrapper'):
            if cancel_flag.is_cancelled():
                output_handler.log("Операция отменена")
                return
            href = a_tag.get('href')
            if href and '/Store/Details/' in href:
                match = re.search(r'/Store/Details/(\d{12})', href)
                if match:
                    product_ids.append(match.group(1))

        product_ids = list(set(product_ids))
        os.makedirs(output_dir, exist_ok=True)
        with open(printers_output_file, 'w', encoding='utf-8') as f:
            json.dump(product_ids, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Найдено {len(product_ids)} товаров. ID сохранены в '{printers_output_file}'.")
    except requests.exceptions.RequestException as e:
        output_handler.log(f"Ошибка при загрузке страницы: {e}")

def download_xls_file(session, headers, output_handler, cancel_flag):
    """Скачивание xls-файла с использованием сессии"""
    try:
        response = session.get(xls_url, headers=headers, verify=cert_path, timeout=10)
        response.raise_for_status()
        with open("temp_price.xls", "wb") as file:
            file.write(response.content)
        output_handler.log("Файл успешно скачан")
        return True
    except requests.exceptions.RequestException as e:
        output_handler.log(f"Ошибка при скачивании файла: {e}")
        return False

def process_xls_file(output_handler, cancel_flag):
    """Обработка xls-файла для поиска 12-значных номеров"""
    try:
        xls = pd.ExcelFile("temp_price.xls")
        twelve_digit_numbers = []
        for sheet_name in xls.sheet_names:
            if cancel_flag.is_cancelled():
                output_handler.log("Операция отменена")
                return None
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            for column in df.columns:
                for value in df[column]:
                    if isinstance(value, str) and re.match(r'^\d{12}$', value):
                        twelve_digit_numbers.append(value)
        return twelve_digit_numbers
    except Exception as e:
        output_handler.log(f"Ошибка при обработке xls файла: {e}")
        return None

def save_to_json(data, filename, output_handler):
    """Сохранение данных в JSON"""
    try:
        os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        output_handler.log(f"Данные сохранены в {filepath}")
    except Exception as e:
        output_handler.log(f"Ошибка при сохранении JSON: {e}")

def process_xls_database(session, headers, output_handler, cancel_flag):
    """Получение базы данных из xls-файла"""
    if download_xls_file(session, headers, output_handler, cancel_flag):
        numbers = process_xls_file(output_handler, cancel_flag)
        if numbers:
            save_to_json(numbers, "DATABASE_recent.json", output_handler)
            try:
                os.remove("temp_price.xls")
            except:
                pass

def parse_printer_compatibility(session, headers, output_handler, cancel_flag):
    """Парсинг совместимости для всех принтеров из Laser_Printers.json"""
    if not os.path.exists(printers_output_file):
        output_handler.log(f"Файл {printers_output_file} не найден")
        return

    try:
        with open(printers_output_file, 'r', encoding='utf-8') as f:
            printer_ids = json.load(f)
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {printers_output_file}: {e}")
        return

    if not printer_ids:
        output_handler.log("Список ID принтеров пуст")
        return

    compatibility_data = {}
    total = len(printer_ids)
    current = 0

    for printer_id in printer_ids:
        if cancel_flag.is_cancelled():
            output_handler.log("Операция отменена")
            return
        current += 1
        output_handler.progress(current, total)
        url = f'https://comcenter.ru/Store/Details/{printer_id}'
        output_handler.log(f"Обрабатывается принтер ID: {printer_id}")

        try:
            response = session.get(url, headers=headers, timeout=10, verify=cert_path)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            grid_sections = soup.select('div.grid.space-top')
            cartridge_ids = []
            part_ids = []

            found_cartridges = False
            found_parts = False

            for grid in grid_sections:
                header = grid.select_one('div.grid-header h2.title')
                if not header:
                    continue

                section_title = header.text.strip()

                if section_title == "Картриджи":
                    found_cartridges = True
                    links = grid.select('a.cells-wrapper')
                    for link in links:
                        href = link.get('href')
                        if href and '/Store/Details/' in href:
                            match = re.search(r'/Store/Details/(\d{12})', href)
                            if match:
                                cartridge_ids.append(match.group(1))

                elif section_title == "Запчасти" and found_cartridges:
                    found_parts = True
                    links = grid.select('a.cells-wrapper')
                    for link in links:
                        href = link.get('href')
                        if href and '/Store/Details/' in href:
                            match = re.search(r'/Store/Details/(\d{12})', href)
                            if match:
                                part_ids.append(match.group(1))

            cartridge_ids = list(set(cartridge_ids))
            part_ids = list(set(part_ids))

            compatibility_data[printer_id] = {
                "cartridges": cartridge_ids,
                "parts": part_ids
            }

            output_handler.log(f"Принтер {printer_id}: найдено картриджей: {len(cartridge_ids)}, запчастей: {len(part_ids)}")

        except requests.exceptions.RequestException as e:
            output_handler.log(f"Ошибка при загрузке страницы для принтера {printer_id}: {e}")
            continue

    if compatibility_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(compatibility_output_file, 'w', encoding='utf-8') as f:
            json.dump(compatibility_data, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Совместимость для {len(compatibility_data)} принтеров сохранена в '{compatibility_output_file}'.")
    else:
        output_handler.log("Не удалось собрать данные о совместимости")

def filter_compatibility_by_stock(output_handler, cancel_flag):
    """Фильтрация совместимости по товарам в наличии"""
    if not os.path.exists(compatibility_output_file):
        output_handler.log(f"Файл {compatibility_output_file} не найден")
        return
    if not os.path.exists(xls_output_file):
        output_handler.log(f"Файл {xls_output_file} не найден")
        return

    try:
        with open(compatibility_output_file, 'r', encoding='utf-8') as f:
            compatibility_data = json.load(f)
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {compatibility_output_file}: {e}")
        return

    try:
        with open(xls_output_file, 'r', encoding='utf-8') as f:
            stock_ids = set(json.load(f))
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {xls_output_file}: {e}")
        return

    if not compatibility_data:
        output_handler.log("Данные о совместимости пусты")
        return

    filtered_data = {}
    total = len(compatibility_data)
    current = 0

    for printer_id in compatibility_data.keys():
        if cancel_flag.is_cancelled():
            output_handler.log("Операция отменена")
            return
        current += 1
        output_handler.progress(current, total)
        data = compatibility_data[printer_id]
        filtered_cartridges = [cid for cid in data.get("cartridges", []) if cid in stock_ids]
        filtered_parts = [pid for pid in data.get("parts", []) if pid in stock_ids]

        if filtered_cartridges or filtered_parts:
            filtered_data[printer_id] = {
                "cartridges": filtered_cartridges,
                "parts": filtered_parts
            }
            output_handler.log(f"Принтер {printer_id}: сохранено картриджей: {len(filtered_cartridges)}, запчастей: {len(filtered_parts)}")
        else:
            output_handler.log(f"Принтер {printer_id}: удален, так как нет товаров в наличии")

    if filtered_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(compatibility_actual_output_file, 'w', encoding='utf-8') as f:
            json.dump(filtered_data, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Отфильтрованные данные для {len(filtered_data)} принтеров сохранены в '{compatibility_actual_output_file}'.")
    else:
        output_handler.log("Нет данных для сохранения после фильтрации")

def load_in_transit_data(output_handler):
    """Чтение данных о товарах в пути из temp_price.xls"""
    in_transit_data = {}
    try:
        if not os.path.exists("temp_price.xls"):
            output_handler.log("Файл temp_price.xls не найден")
            return in_transit_data

        xls = pd.ExcelFile("temp_price.xls")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            # Проверяем наличие колонок 3 ("Код") и 6 ("В пути")
            if len(df.columns) >= 6:
                for _, row in df.iterrows():
                    product_id = row.iloc[2]  # Колонка 3 (индекс 2)
                    in_transit = row.iloc[5]  # Колонка 6 (индекс 5)
                    if isinstance(product_id, str) and re.match(r'^\d{12}$', product_id):
                        # Проверяем, является ли in_transit числом
                        try:
                            in_transit_value = int(float(in_transit)) if pd.notna(in_transit) else 0
                        except (ValueError, TypeError):
                            in_transit_value = 0
                        in_transit_data[product_id] = in_transit_value
        output_handler.log(f"Загружено {len(in_transit_data)} записей о товарах в пути")
        return in_transit_data
    except Exception as e:
        output_handler.log(f"Ошибка при чтении temp_price.xls: {e}")
        return in_transit_data

def ensure_xls_file(session, headers, output_handler, cancel_flag):
    """Проверяет наличие temp_price.xls и скачивает его, если отсутствует"""
    if os.path.exists("temp_price.xls"):
        output_handler.log("Файл temp_price.xls уже существует")
        return True
    output_handler.log("Файл temp_price.xls отсутствует, выполняется скачивание")
    return download_xls_file(session, headers, output_handler, cancel_flag)

def parse_cartridges_and_parts(session, headers, output_handler, cancel_flag):
    """Парсинг данных о актуальных картриджах и запчастях из PRINTERS_compatibility_actual.json"""
    if not os.path.exists(compatibility_actual_output_file):
        output_handler.log(f"Файл {compatibility_actual_output_file} не найден")
        return

    try:
        with open(compatibility_actual_output_file, 'r', encoding='utf-8') as f:
            compatibility_data = json.load(f)
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {compatibility_actual_output_file}: {e}")
        return

    if not compatibility_data:
        output_handler.log("Данные о совместимости пусты")
        return

    # Собираем все уникальные ID картриджей и запчастей
    all_ids = set()
    for printer_id, data in compatibility_data.items():
        all_ids.update(data.get("cartridges", []))
        all_ids.update(data.get("parts", []))

    if not all_ids:
        output_handler.log("Нет ID картриджей или запчастей для парсинга")
        return

    output_handler.log(f"Найдено {len(all_ids)} уникальных ID для парсинга")

    # Проверяем и скачиваем temp_price.xls, если он отсутствует
    if not ensure_xls_file(session, headers, output_handler, cancel_flag):
        output_handler.log("Не удалось скачать temp_price.xls, данные 'in_transit' не будут загружены")
        in_transit_data = {}
    else:
        # Загружаем данные о товарах в пути
        in_transit_data = load_in_transit_data(output_handler)

    # Словарь для хранения данных
    parsed_data = {}
    total = len(all_ids)
    current = 0

    for product_id in all_ids:
        if cancel_flag.is_cancelled():
            output_handler.log("Операция отменена")
            return
        current += 1
        output_handler.progress(current, total)
        url = f'https://comcenter.ru/Store/Details/{product_id}'
        output_handler.log(f"Обрабатывается ID: {product_id}")

        try:
            response = session.get(url, headers=headers, timeout=10, verify=cert_path)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Извлечение наименования товара
            name_element = soup.select_one('div.grid-body.text-left.space-top-tiny h1')
            product_name = name_element.text.strip() if name_element else ""

            # Извлечение наличия
            availability_element = soup.select_one('span.product-count')
            availability = int(availability_element.text.strip()) if availability_element and availability_element.text.strip().isdigit() else 0

            # Извлечение данных о товарах в пути
            in_transit = in_transit_data.get(product_id, 0)

            # Извлечение цен
            price_element = soup.select_one('div.product-price-container span[data-bind*="getBrowsingPrice"]')
            retail_price = 0.0
            wholesale_price = 0.0
            if price_element:
                data_bind = price_element.get('data-bind', '')
                match = re.search(r'getBrowsingPrice\((\d+\.\d+), (\d+\.?\d*)\)', data_bind)
                if match:
                    retail_price = float(match.group(1))
                    wholesale_price = float(match.group(2))
                else:
                    output_handler.log(f"Не удалось извлечь цены для ID {product_id}: {data_bind}")

            # Извлечение характеристик
            characteristics = {}
            characteristics_table = soup.select_one('div.product-properties-container table.price-list')
            if characteristics_table:
                for row in characteristics_table.select('tr'):
                    cells = row.select('td')
                    if len(cells) == 2:
                        key = cells[0].text.strip()
                        value = cells[1].text.strip()
                        characteristics[key] = value

            # Извлечение описания товара
            description_section = soup.select_one('div.grid.space-top div.grid-body.text-left.space-top-tiny')
            description = ""
            if description_section:
                description = ' '.join(description_section.get_text(strip=True).split())
                description = re.sub(r'\s+', ' ', description).strip()

            # Формирование данных для текущего ID
            parsed_data[product_id] = {
                "name": product_name,
                "availability": availability,
                "in_transit": in_transit,
                "wholesale_price": wholesale_price,
                "retail_price": retail_price,
                "characteristics": characteristics,
                "description": description
            }

            output_handler.log(f"ID {product_id}: успешно обработан")

        except requests.exceptions.RequestException as e:
            output_handler.log(f"Ошибка при загрузке страницы для ID {product_id}: {e}")
            continue
        except Exception as e:
            output_handler.log(f"Ошибка при парсинге данных для ID {product_id}: {e}")
            continue

    # Сохранение данных в JSON
    if parsed_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(cartridges_parts_output_file, 'w', encoding='utf-8') as f:
            json.dump(parsed_data, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Данные для {len(parsed_data)} элементов сохранены в '{cartridges_parts_output_file}'.")
    else:
        output_handler.log("Не удалось собрать данные")

def parse_all_cartridges_and_parts(session, headers, output_handler, cancel_flag):
    """Парсинг данных о ВСЕХ картриджах и запчастях из PRINTERS_compatibility.json"""
    if not os.path.exists(compatibility_output_file):
        output_handler.log(f"Файл {compatibility_output_file} не найден")
        return

    try:
        with open(compatibility_output_file, 'r', encoding='utf-8') as f:
            compatibility_data = json.load(f)
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {compatibility_output_file}: {e}")
        return

    if not compatibility_data:
        output_handler.log("Данные о совместимости пусты")
        return

    # Собираем все уникальные ID картриджей и запчастей
    all_ids = set()
    for printer_id, data in compatibility_data.items():
        all_ids.update(data.get("cartridges", []))
        all_ids.update(data.get("parts", []))

    if not all_ids:
        output_handler.log("Нет ID картриджей или запчастей для парсинга")
        return

    output_handler.log(f"Найдено {len(all_ids)} уникальных ID для парсинга")

    # Проверяем и скачиваем temp_price.xls, если он отсутствует
    if not ensure_xls_file(session, headers, output_handler, cancel_flag):
        output_handler.log("Не удалось скачать temp_price.xls, данные 'in_transit' не будут загружены")
        in_transit_data = {}
    else:
        # Загружаем данные о товарах в пути
        in_transit_data = load_in_transit_data(output_handler)

    # Словарь для хранения данных
    parsed_data = {}
    total = len(all_ids)
    current = 0

    for product_id in all_ids:
        if cancel_flag.is_cancelled():
            output_handler.log("Операция отменена")
            return
        current += 1
        output_handler.progress(current, total)
        url = f'https://comcenter.ru/Store/Details/{product_id}'
        output_handler.log(f"Обрабатывается ID: {product_id}")

        try:
            response = session.get(url, headers=headers, timeout=10, verify=cert_path)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Извлечение наименования товара
            name_element = soup.select_one('div.grid-body.text-left.space-top-tiny h1')
            product_name = name_element.text.strip() if name_element else ""

            # Извлечение наличия
            availability_element = soup.select_one('span.product-count')
            availability = int(availability_element.text.strip()) if availability_element and availability_element.text.strip().isdigit() else 0

            # Извлечение данных о товарах в пути
            in_transit = in_transit_data.get(product_id, 0)

            # Извлечение цен
            price_element = soup.select_one('div.product-price-container span[data-bind*="getBrowsingPrice"]')
            retail_price = 0.0
            wholesale_price = 0.0
            if price_element:
                data_bind = price_element.get('data-bind', '')
                match = re.search(r'getBrowsingPrice\((\d+\.\d+), (\d+\.?\d*)\)', data_bind)
                if match:
                    retail_price = float(match.group(1))
                    wholesale_price = float(match.group(2))
                else:
                    output_handler.log(f"Не удалось извлечь цены для ID {product_id}: {data_bind}")

            # Извлечение характеристик
            characteristics = {}
            characteristics_table = soup.select_one('div.product-properties-container table.price-list')
            if characteristics_table:
                for row in characteristics_table.select('tr'):
                    cells = row.select('td')
                    if len(cells) == 2:
                        key = cells[0].text.strip()
                        value = cells[1].text.strip()
                        characteristics[key] = value

            # Извлечение описания товара
            description_section = soup.select_one('div.grid.space-top div.grid-body.text-left.space-top-tiny')
            description = ""
            if description_section:
                description = ' '.join(description_section.get_text(strip=True).split())
                description = re.sub(r'\s+', ' ', description).strip()

            # Формирование данных для текущего ID
            parsed_data[product_id] = {
                "name": product_name,
                "availability": availability,
                "in_transit": in_transit,
                "wholesale_price": wholesale_price,
                "retail_price": retail_price,
                "characteristics": characteristics,
                "description": description
            }

            output_handler.log(f"ID {product_id}: успешно обработан")

        except requests.exceptions.RequestException as e:
            output_handler.log(f"Ошибка при загрузке страницы для ID {product_id}: {e}")
            continue
        except Exception as e:
            output_handler.log(f"Ошибка при парсинге данных для ID {product_id}: {e}")
            continue

    # Сохранение данных в JSON
    if parsed_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(all_cartridges_parts_output_file, 'w', encoding='utf-8') as f:
            json.dump(parsed_data, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Данные для {len(parsed_data)} элементов сохранены в '{all_cartridges_parts_output_file}'.")
    else:
        output_handler.log("Не удалось собрать данные")

def parse_comcenter_products(session, headers, output_handler, cancel_flag):
    """Парсинг данных о актуальных товарах Comcenter из DATABASE_recent.json"""
    if not os.path.exists(xls_output_file):
        output_handler.log(f"Файл {xls_output_file} не найден")
        return

    try:
        with open(xls_output_file, 'r', encoding='utf-8') as f:
            product_ids = json.load(f)
    except Exception as e:
        output_handler.log(f"Ошибка при чтении файла {xls_output_file}: {e}")
        return

    if not product_ids:
        output_handler.log("Список ID товаров пуст")
        return

    output_handler.log(f"Найдено {len(product_ids)} уникальных ID для парсинга")

    # Проверяем и скачиваем temp_price.xls, если он отсутствует
    if not ensure_xls_file(session, headers, output_handler, cancel_flag):
        output_handler.log("Не удалось скачать temp_price.xls, данные 'in_transit' не будут загружены")
        in_transit_data = {}
    else:
        # Загружаем данные о товарах в пути
        in_transit_data = load_in_transit_data(output_handler)

    # Словарь для хранения данных
    parsed_data = {}
    total = len(product_ids)
    current = 0

    for product_id in product_ids:
        if cancel_flag.is_cancelled():
            output_handler.log("Операция отменена")
            return
        current += 1
        output_handler.progress(current, total)
        url = f'https://comcenter.ru/Store/Details/{product_id}'
        output_handler.log(f"Обрабатывается ID: {product_id}")

        try:
            response = session.get(url, headers=headers, timeout=10, verify=cert_path)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Извлечение наименования товара
            name_element = soup.select_one('div.grid-body.text-left.space-top-tiny h1')
            product_name = name_element.text.strip() if name_element else ""

            # Извлечение наличия
            availability_element = soup.select_one('span.product-count')
            availability = int(availability_element.text.strip()) if availability_element and availability_element.text.strip().isdigit() else 0

            # Извлечение данных о товарах в пути
            in_transit = in_transit_data.get(product_id, 0)

            # Извлечение цен
            price_element = soup.select_one('div.product-price-container span[data-bind*="getBrowsingPrice"]')
            retail_price = 0.0
            wholesale_price = 0.0
            if price_element:
                data_bind = price_element.get('data-bind', '')
                match = re.search(r'getBrowsingPrice\((\d+\.\d+), (\d+\.?\d*)\)', data_bind)
                if match:
                    retail_price = float(match.group(1))
                    wholesale_price = float(match.group(2))
                else:
                    output_handler.log(f"Не удалось извлечь цены для ID {product_id}: {data_bind}")

            # Извлечение характеристик
            characteristics = {}
            characteristics_table = soup.select_one('div.product-properties-container table.price-list')
            if characteristics_table:
                for row in characteristics_table.select('tr'):
                    cells = row.select('td')
                    if len(cells) == 2:
                        key = cells[0].text.strip()
                        value = cells[1].text.strip()
                        characteristics[key] = value

            # Извлечение описания товара
            description_section = soup.select_one('div.grid.space-top div.grid-body.text-left.space-top-tiny')
            description = ""
            if description_section:
                description = ' '.join(description_section.get_text(strip=True).split())
                description = re.sub(r'\s+', ' ', description).strip()

            # Формирование данных для текущего ID
            parsed_data[product_id] = {
                "name": product_name,
                "availability": availability,
                "in_transit": in_transit,
                "wholesale_price": wholesale_price,
                "retail_price": retail_price,
                "characteristics": characteristics,
                "description": description
            }

            output_handler.log(f"ID {product_id}: успешно обработан")

        except requests.exceptions.RequestException as e:
            output_handler.log(f"Ошибка при загрузке страницы для ID {product_id}: {e}")
            continue
        except Exception as e:
            output_handler.log(f"Ошибка при парсинге данных для ID {product_id}: {e}")
            continue

    # Сохранение данных в JSON
    if parsed_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(comcenter_products_output_file, 'w', encoding='utf-8') as f:
            json.dump(parsed_data, f, ensure_ascii=False, indent=4)
        output_handler.log(f"Данные для {len(parsed_data)} элементов сохранены в '{comcenter_products_output_file}'.")
    else:
        output_handler.log("Не удалось собрать данные")

def run_action(choice, output_handler, cancel_flag):
    """Запуск выбранного действия"""
    session_info = setup_session(output_handler)
    if not session_info:
        output_handler.log("Не удалось авторизоваться. Программа завершена.")
        return

    session, headers = session_info

    if choice == "1":
        output_handler.log("Получение базы данных лазерных принтеров...")
        get_laser_printers_database(session, headers, output_handler, cancel_flag)
    
    elif choice == "2":
        output_handler.log("Получение базы данных из xls-файла...")
        process_xls_database(session, headers, output_handler, cancel_flag)
    
    elif choice == "3":
        output_handler.log("Парсинг совместимости для всех принтеров...")
        parse_printer_compatibility(session, headers, output_handler, cancel_flag)
    
    elif choice == "4":
        output_handler.log("Фильтрация совместимости по товарам в наличии...")
        filter_compatibility_by_stock(output_handler, cancel_flag)
    
    elif choice == "5":
        output_handler.log("Парсинг актуальных картриджей и запчастей...")
        parse_cartridges_and_parts(session, headers, output_handler, cancel_flag)
    
    elif choice == "6":
        output_handler.log("Парсинг ВСЕХ картриджей и запчастей...")
        parse_all_cartridges_and_parts(session, headers, output_handler, cancel_flag)
    
    elif choice == "7":
        output_handler.log("Парсинг актуальных товаров Comcenter...")
        parse_comcenter_products(session, headers, output_handler, cancel_flag)
    
    else:
        output_handler.log("Неверный выбор. Пожалуйста, выберите 0, 1, 2, 3, 4, 5, 6 или 7")

def console_main():
    """Консольный интерфейс программы"""
    output_handler = ConsoleOutputHandler()
    
    while True:
        print("\nМеню:")
        print("1. Получить базу данных лазерных принтеров")
        print("2. Получить базу данных из xls-файла")
        print("3. Парсинг совместимости принтера")
        print("4. Совместимость только по товарам в наличии")
        print("5. Парсинг актуальных картриджей и запчастей")
        print("6. Парсинг ВСЕХ картриджей и запчастей")
        print("7. Парсинг актуальных товаров Comcenter")
        print("0. Выход")
        
        choice = input("Выберите действие (0-7): ")
        
        if choice == "0":
            output_handler.log("Программа завершена")
            break
        
        cancel_flag = CancelFlag()
        run_action(choice, output_handler, cancel_flag)

if __name__ == "__main__":
    console_main()