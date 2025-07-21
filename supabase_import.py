import pandas as pd
from supabase import create_client, Client

# Настройка подключения к Supabase
url = "https://musltekbzezvxheqafks.supabase.co"
key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im11c2x0ZWtiemV6dnhoZXFhZmtzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDY0NTg3MTQsImV4cCI6MjA2MjAzNDcxNH0.3YQInSZ-X-3vwIqUm4uQMbm5OldSOylHg2y2AZUHTrQ"
supabase: Client = create_client(url, key)

# Чтение CSV
df = pd.read_csv("D:/!tilo/COMCENTER/parsing_COMCENTER.RU/COMCENTER.ru_database/DATABASE_comcenter_products.csv")

# Очистка старых данных
supabase.table("products").delete().neq("product_id", "").execute()

# Загрузка новых данных
data = df.to_dict(orient="records")
supabase.table("products").insert(data).execute()

print("Данные успешно обновлены!")