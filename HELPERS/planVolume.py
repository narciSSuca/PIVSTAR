import pandas as pd
import mysql.connector

# Параметры подключения к MySQL
DB_HOST = "78.40.219.237"
DB_PORT = 3306
DB_USER = "gen_user"
DB_PASSWORD = "nS5zN{|,Pw*)zC"
DB_NAME = "default_db"  # Имя базы данных
TABLE_NAME = "volume_plan_category"
FILE_PATH = "volume_plan.xlsx"  # Путь к файлу

# Читаем Excel-файл
try:
    df = pd.read_excel(FILE_PATH, engine="openpyxl")
except Exception as e:
    print(f"Ошибка при чтении файла {FILE_PATH}: {e}")
    exit()

# Проверяем наличие обязательных колонок
base_columns = ['DSM', 'TSM', 'ESR']
categories = [col for col in df.columns if col not in base_columns]  # Все остальные колонки - категории

if not all(col in df.columns for col in base_columns):
    print(f"Ошибка: недостающие колонки в файле. Ожидались: {base_columns}")
    exit()

# Преобразуем данные в "длинный" формат
melted_df = df.melt(id_vars=base_columns, var_name="category_plan", value_name="Plan")

# Удаляем строки, где значение 'Plan' пустое (NaN)
melted_df = melted_df.dropna(subset=["Plan"])

# Подключение к MySQL
conn = mysql.connector.connect(
    host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASSWORD
)
cursor = conn.cursor()

# Создаем базу данных (если нет)
cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
cursor.execute(f"USE {DB_NAME}")

# Создаем таблицу (если нет)
sql_create = f"""
    CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
        id INT AUTO_INCREMENT PRIMARY KEY,
        DSM VARCHAR(255),
        TSM VARCHAR(255),
        ESR VARCHAR(255),
        category_plan VARCHAR(255),
        Plan DECIMAL(15,2)
    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
"""
cursor.execute(sql_create)

# Очищаем таблицу перед загрузкой новых данных
cursor.execute(f"DELETE FROM {TABLE_NAME}")

# Вставка данных
sql_insert = f"""
    INSERT INTO {TABLE_NAME} (DSM, TSM, ESR, category_plan, Plan)
    VALUES (%s, %s, %s, %s, %s)
"""

data_to_insert = [tuple(row) for row in melted_df.itertuples(index=False, name=None)]
cursor.executemany(sql_insert, data_to_insert)
conn.commit()

# Закрываем соединение
cursor.close()
conn.close()
print(f"Загрузка файла {FILE_PATH} в таблицу {TABLE_NAME} завершена.")
