from dotenv import load_dotenv
import urllib3
import requests
import json
from bs4 import BeautifulSoup
import re
import os

# Путь к файлу сертификата
cert_path = "C:/!Work/COMCENTER/fullchain.pem"

# Проверка существования файла сертификата
if not os.path.exists(cert_path):
    print(f"Файл сертификата {cert_path} не найден")
    exit(1)

# Загружаем переменные из файла .env
load_dotenv()

LOGIN = os.getenv('COMCENTER.RU_LOGIN')
PASSWORD = os.getenv('COMCENTER.RU_PASSWORD')

# Проверка наличия данных
if not LOGIN or not PASSWORD:
    print("Логин или пароль не заданы в файле .env")
    exit(1)

# Создаем сессию
session = requests.Session()

# Устанавливаем заголовки
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36',
    'Referer': 'https://comcenter.ru/',
    'Content-Type': 'application/x-www-form-urlencoded',
}

# Проверяем доступность сайта
try:
    response = requests.get("https://comcenter.ru", headers=headers, verify=cert_path, timeout=10)
    print(f"Статус проверки сайта: {response.status_code}")
    if response.status_code != 200:
        print("Не удалось подключиться к сайту comcenter.ru")
        exit(1)
except requests.exceptions.SSLError as e:
    print(f"Ошибка SSL при проверке сайта: {e}")
    exit(1)
except requests.exceptions.RequestException as e:
    print(f"Ошибка при проверке сайта: {e}")
    exit(1)

# URL для авторизации
login_url = 'https://comcenter.ru/Account/LogOn'

# Данные формы для входа
login_data = {
    'UserName': LOGIN,
    'Password': PASSWORD,
    'RememberMe': 'false',  # Можно установить 'true', если нужно запомнить сессию
}

try:
    # Выполняем POST-запрос для входа с указанным сертификатом
    response = session.post(login_url, data=login_data, headers=headers, timeout=10, verify=cert_path)

    # Проверяем успешность входа
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Проверяем, есть ли сообщение об ошибке на странице
    error_message = soup.find('h1', class_='dark-red-color')
    if error_message and "неверное имя или пароль" in error_message.text.lower():
        print("Ошибка входа: неверный логин или пароль")
        exit(1)

    print("Успешно вошли в систему!")

except requests.exceptions.SSLError as e:
    print(f"Ошибка SSL при авторизации: {e}")
    exit(1)
except requests.exceptions.RequestException as e:
    print(f"Ошибка при выполнении запроса: {e}")
    exit(1)

# URL страницы для парсинга
url = 'https://comcenter.ru/Store/Browse/400000000796/kartridzhi-dlya-lazernykh-printerov-originalnye'

try:
    # Выполняем GET-запрос с использованием сессии и указанным сертификатом
    response = session.get(url, headers=headers, timeout=10, verify=cert_path)
    response.raise_for_status()
    html_content = response.text

except requests.exceptions.SSLError as e:
    print(f"Ошибка SSL при загрузке страницы: {e}")
    exit(1)
except requests.exceptions.RequestException as e:
    print(f"Ошибка при загрузке страницы: {e}")
    exit(1)


# Парсим HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Массив для ID товаров
product_ids = []

# Ищем все <a> с классом 'cells-wrapper'
for a_tag in soup.select('a.cells-wrapper'):
    href = a_tag.get('href')
    if href and '/Store/Details/' in href:
        # Ищем 12-значный ID внутри URL с помощью регулярного выражения
        match = re.search(r'/Store/Details/(\d{12})', href)
        if match:
            product_ids.append(match.group(1))

# Удаляем дубли
product_ids = list(set(product_ids))

# Создаем папку и сохраняем файл в JSON
os.makedirs('COMCENTER.ru_database', exist_ok=True)
with open('COMCENTER.ru_database/Laser_Printers.json', 'w', encoding='utf-8') as f:
    json.dump(product_ids, f, ensure_ascii=False, indent=4)

print(f"Найдено {len(product_ids)} товаров. ID сохранены в 'Laser_Printers.json'.")
