import requests
import json
from bs4 import BeautifulSoup

# URL страницы для парсинга
url = 'https://comcenter.ru/Store/Browse/400000000796/kartridzhi-dlya-lazernykh-printerov-originalnye'

# Заголовки (имитируем браузер)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                  ' Chrome/90.0.4430.85 Safari/537.36'
}

try:
    # Получение HTML страницы с отключенной проверкой SSL
    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()
    html_content = response.text
except requests.exceptions.SSLError as e:
    print("SSL ошибка:", e)
    exit(1)

# Парсинг HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Собираем все ссылки товаров
product_links = []

# Ищем все <a> с классом 'cells-wrapper'
for a_tag in soup.select('a.cells-wrapper'):
    href = a_tag.get('href')
    if href and '/Store/Details/' in href:
        # Полный URL
        full_url = 'https://comcenter.ru' + href
        product_links.append(full_url)

# Удаляем дубликаты
product_links = list(set(product_links))

# Сохраняем в JSON файл
with open('product_links.json', 'w', encoding='utf-8') as f:
    json.dump(product_links, f, ensure_ascii=False, indent=4)

print(f"Найдено {len(product_links)} товаров. Ссылки сохранены в 'product_links.json'.")
