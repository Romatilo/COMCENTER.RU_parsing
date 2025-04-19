import requests

cert_path = "C:/!Work/COMCENTER/fullchain.pem"
response = requests.get("https://comcenter.ru", verify=cert_path)
print(response.status_code)  # Должно вернуть 200