from dotenv import load_dotenv
import urllib3
import requests
import pandas as pd
import json
import os
import re
from bs4 import BeautifulSoup

# Путь к файлу сертификата
cert_path = "C:/!Work/COMCENTER/fullchain.pem"

# Путь для сохранения данных
output_dir = "COMCENTER.ru_database"
xls_url = "https://comcenter.ru/Content/PriceList/price.xls"
xls_output_file = os.path.join(output_dir, "DATABASE_recent.json")
printers_output_file = os.path.join(output_dir, "Laser_Printers.json")
compatibility_output_file = os.path.join(output_dir, "PRINTERS_compatibility.json")

def setup_session():
    """Настройка сессии с учетом сертификата и авторизации"""
    if not os.path.exists(cert_path):
        print(f"Файл сертификата {cert_path} не найден")
        return None

    # Загружаем переменные из .env
    load_dotenv()
    LOGIN = os.getenv('COMCENTER.RU_LOGIN')
    PASSWORD = os.getenv('COMCENTER.RU_PASSWORD')

    if not LOGIN or not PASSWORD:
        print("Логин или пароль не заданы в файле .env")
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
            print("Не удалось подключиться к сайту comcenter.ru")
            return None
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при проверке сайта: {e}")
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
            print("Ошибка входа: неверный логин или пароль")
            return None
        print("Успешно вошли в систему!")
        return session, headers
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при авторизации: {e}")
        return None

def get_laser_printers_database(session, headers):
    """Получение базы данных лазерных принтеров"""
    url = 'https://comcenter.ru/Store/Browse/400000006580/printery-lazernye-i-mfu'

    try:
        response = session.get(url, headers=headers, timeout=10, verify=cert_path)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        product_ids = []
        for a_tag in soup.select('a.cells-wrapper'):
            href = a_tag.get('href')
            if href and '/Store/Details/' in href:
                match = re.search(r'/Store/Details/(\d{12})', href)
                if match:
                    product_ids.append(match.group(1))

        product_ids = list(set(product_ids))
        os.makedirs(output_dir, exist_ok=True)
        with open(printers_output_file, 'w', encoding='utf-8') as f:
            json.dump(product_ids, f, ensure_ascii=False, indent=4)
        print(f"Найдено {len(product_ids)} товаров. ID сохранены в '{printers_output_file}'.")
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при загрузке страницы: {e}")

def download_xls_file(session, headers):
    """Скачивание xls-файла с использованием сессии"""
    try:
        response = session.get(xls_url, headers=headers, verify=cert_path, timeout=10)
        response.raise_for_status()
        with open("temp_price.xls", "wb") as file:
            file.write(response.content)
        print("Файл успешно скачан")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при скачивании файла: {e}")
        return False

def process_xls_file():
    """Обработка xls-файла для поиска 12-значных номеров"""
    try:
        xls = pd.ExcelFile("temp_price.xls")
        twelve_digit_numbers = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            for column in df.columns:
                for value in df[column]:
                    if isinstance(value, str) and re.match(r'^\d{12}$', value):
                        twelve_digit_numbers.append(value)
        return twelve_digit_numbers
    except Exception as e:
        print(f"Ошибка при обработке xls файла: {e}")
        return None

def save_to_json(data, filename):
    """Сохранение данных в JSON"""
    try:
        os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Данные сохранены в {filepath}")
    except Exception as e:
        print(f"Ошибка при сохранении JSON: {e}")

def process_xls_database(session, headers):
    """Получение базы данных из xls-файла"""
    if download_xls_file(session, headers):
        numbers = process_xls_file()
        if numbers:
            save_to_json(numbers, "DATABASE_recent.json")
            try:
                os.remove("temp_price.xls")
            except:
                pass

def parse_printer_compatibility(session, headers):
    """Парсинг совместимости для всех принтеров из Laser_Printers.json"""
    # Проверяем существование файла Laser_Printers.json
    if not os.path.exists(printers_output_file):
        print(f"Файл {printers_output_file} не найден")
        return

    # Читаем список ID принтеров
    try:
        with open(printers_output_file, 'r', encoding='utf-8') as f:
            printer_ids = json.load(f)
    except Exception as e:
        print(f"Ошибка при чтении файла {printers_output_file}: {e}")
        return

    if not printer_ids:
        print("Список ID принтеров пуст")
        return

    compatibility_data = {}

    for printer_id in printer_ids:
        # Формируем URL для текущего принтера
        url = f'https://comcenter.ru/Store/Details/{printer_id}'
        print(f"Обрабатывается принтер ID: {printer_id}")

        try:
            response = session.get(url, headers=headers, timeout=10, verify=cert_path)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Находим все секции grid
            grid_sections = soup.select('div.grid.space-top')
            cartridge_ids = []
            part_ids = []

            # Переменные для отслеживания нужных секций
            found_cartridges = False
            found_parts = False

            for grid in grid_sections:
                header = grid.select_one('div.grid-header h2.title')
                if not header:
                    continue

                section_title = header.text.strip()

                # Проверяем секцию "Картриджи"
                if section_title == "Картриджи":
                    found_cartridges = True
                    links = grid.select('a.cells-wrapper')
                    for link in links:
                        href = link.get('href')
                        if href and '/Store/Details/' in href:
                            match = re.search(r'/Store/Details/(\d{12})', href)
                            if match:
                                cartridge_ids.append(match.group(1))

                # Проверяем секцию "Запчасти"
                elif section_title == "Запчасти" and found_cartridges:
                    found_parts = True
                    links = grid.select('a.cells-wrapper')
                    for link in links:
                        href = link.get('href')
                        if href and '/Store/Details/' in href:
                            match = re.search(r'/Store/Details/(\d{12})', href)
                            if match:
                                part_ids.append(match.group(1))

            # Удаляем дубликаты
            cartridge_ids = list(set(cartridge_ids))
            part_ids = list(set(part_ids))

            # Добавляем данные в словарь
            compatibility_data[printer_id] = {
                "cartridges": cartridge_ids,
                "parts": part_ids
            }

            print(f"Принтер {printer_id}: найдено картриджей: {len(cartridge_ids)}, запчастей: {len(part_ids)}")

        except requests.exceptions.RequestException as e:
            print(f"Ошибка при загрузке страницы для принтера {printer_id}: {e}")
            continue

    # Сохраняем данные в JSON
    if compatibility_data:
        os.makedirs(output_dir, exist_ok=True)
        with open(compatibility_output_file, 'w', encoding='utf-8') as f:
            json.dump(compatibility_data, f, ensure_ascii=False, indent=4)
        print(f"Совместимость для {len(compatibility_data)} принтеров сохранена в '{compatibility_output_file}'.")
    else:
        print("Не удалось собрать данные о совместимости")

def show_menu():
    """Отображение меню"""
    print("\nМеню:")
    print("1. Получить базу данных лазерных принтеров")
    print("2. Получить базу данных из xls-файла")
    print("3. Парсинг совместимости принтера")
    print("0. Выход")

def main():
    """Основная функция с меню"""
    # Выполняем авторизацию до вывода меню
    session_info = setup_session()
    if not session_info:
        print("Не удалось авторизоваться. Программа завершена.")
        return

    session, headers = session_info

    while True:
        show_menu()
        choice = input("Выберите действие (0-3): ")
        
        if choice == "1":
            print("Получение базы данных лазерных принтеров...")
            get_laser_printers_database(session, headers)
        
        elif choice == "2":
            print("Получение базы данных из xls-файла...")
            process_xls_database(session, headers)
        
        elif choice == "3":
            print("Парсинг совместимости для всех принтеров...")
            parse_printer_compatibility(session, headers)
        
        elif choice == "0":
            print("Программа завершена")
            break
        
        else:
            print("Неверный выбор. Пожалуйста, выберите 0, 1, 2 или 3")

if __name__ == "__main__":
    main()