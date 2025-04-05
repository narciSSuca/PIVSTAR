import pandas as pd
import mysql.connector
import re

# Параметры подключения к MySQL
DB_HOST = "78.40.219.237"
DB_PORT = 3306
DB_USER = "gen_user"
DB_PASSWORD = "nS5zN{|,Pw*)zC"
DB_NAME = "default_db"  # Замените на имя вашей базы данных
TABLE_NAME = "four_plan"
FILE_PATH = "coverage_plan.xlsx"  # Замените на путь к вашему Excel-файлу

# Читаем Excel-файл
try:
    df = pd.read_excel(FILE_PATH, engine="openpyxl")
except Exception as e:
    print(f"Ошибка при чтении файла {FILE_PATH}: {e}")
    exit()

# Проверка на наличие необходимых колонок
required_columns = ['DSM', 'TSM', 'ESR', 'Plan']
if not all(col in df.columns for col in required_columns):
    print(f"Ошибка: недостающие колонки в файле. Ожидались: {required_columns}")
    exit()

# Подключение к MySQL
conn = mysql.connector.connect(
    host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASSWORD
)
cursor = conn.cursor()

# Создаем базу данных (если нет)
cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
cursor.execute(f"USE {DB_NAME}")

# Создаем таблицу four_plan (если нет)
sql_create = f"""
    CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
        id INT AUTO_INCREMENT PRIMARY KEY,
        DSM VARCHAR(255),
        TSM VARCHAR(255),
        ESR VARCHAR(255),
        Plan VARCHAR(255)
    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
"""
cursor.execute(sql_create)

# Очищаем таблицу перед загрузкой новых данных
cursor.execute(f"DELETE FROM {TABLE_NAME}")

# Вставка данных
for i, row in df.iterrows():
    cleaned_row = [None if pd.isna(x) else x for x in row[required_columns]]  # NaN → None
    placeholders = ", ".join(["%s"] * len(required_columns))
    sql_insert = f"INSERT INTO {TABLE_NAME} ({', '.join(required_columns)}) VALUES ({placeholders})"

    try:
        cursor.execute(sql_insert, tuple(cleaned_row))
    except mysql.connector.Error as err:
        print(f"Ошибка при вставке данных: {err}")
        print("Строка с ошибкой:", cleaned_row)
        break  # Прерываем загрузку при ошибке

    conn.commit()

# Закрываем соединение
cursor.close()
conn.close()
print(f"Загрузка файла {FILE_PATH} в таблицу {TABLE_NAME} завершена.")
