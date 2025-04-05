import pandas as pd
import mysql.connector
import re
import asyncio
from aiogram.exceptions import TelegramBadRequest, TelegramNetworkError

# Параметры подключения к MySQL
DB_HOST = "78.40.219.237"
DB_PORT = 3306
DB_USER = "gen_user"
DB_PASSWORD = "nS5zN{|,Pw*)zC"

BATCH_SIZE = 500  # Размер пакета для executemany()
CHUNK_SIZE = 5000  # Количество строк в одном чанке
RETRY_ATTEMPTS = 3  # Количество попыток при разрыве соединения


def clean_column_name(name):
    """Очищает название колонок, заменяя неподходящие символы на _"""
    if not isinstance(name, str) or not name.strip():
        return None
    name = re.sub(r'[^a-zA-Zа-яА-Я0-9_]', '_', name).strip('_')
    return name


async def execute_with_retry(query, data, db_name):
    """Выполняет SQL-запрос с повторными попытками при сбоях соединения"""
    for attempt in range(RETRY_ATTEMPTS):
        try:
            with mysql.connector.connect(
                host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASSWORD, database=db_name,
                connection_timeout=600
            ) as conn:
                with conn.cursor() as cursor:
                    for batch_start in range(0, len(data), BATCH_SIZE):
                        batch_end = batch_start + BATCH_SIZE
                        cursor.executemany(query, data[batch_start:batch_end])
                        conn.commit()
            return
        except mysql.connector.Error as e:
            print(f"Ошибка MySQL: {e}")
            if attempt < RETRY_ATTEMPTS - 1:
                print("Попытка переподключения...")
                await asyncio.sleep(2)
            else:
                print("Не удалось выполнить запрос после нескольких попыток.")


async def load_excel_to_mysql(db_name, table_name, file_paths, column_types, progress_message, clear_table=True):
    """Загружает данные из нескольких Excel-файлов в MySQL"""
    dataframes = []
    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            df.columns = [clean_column_name(col) for col in df.columns]
            df = df.loc[:, df.columns.notnull()]
            dataframes.append(df)
        except Exception as e:
            print(f"Ошибка при чтении файла {file_path}: {e}")
            return

    combined_df = pd.concat(dataframes, ignore_index=True)

    # Обработка дубликатов колонок
    seen = {}
    final_columns = []
    for col in combined_df.columns:
        if col in seen:
            seen[col] += 1
            new_col = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
            new_col = col
        final_columns.append(new_col)
    combined_df.columns = final_columns

    # Подключение к MySQL для создания БД и таблицы
    with mysql.connector.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASSWORD,
        connection_timeout=600
    ) as conn:
        with conn.cursor() as cursor:
            conn.autocommit = False
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS {db_name} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            cursor.execute(f"USE {db_name}")

            # Создание таблицы
            sql_columns = []
            for col in final_columns:
                col_type = "TEXT"
                if column_types:
                    for sql_type, cols in column_types.items():
                        if col in cols:
                            col_type = sql_type
                            break
                sql_columns.append(f"{col} {col_type}")

            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {table_name} (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    {', '.join(sql_columns)}
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
            """)

            # Очистка таблицы, если нужно
            if clear_table:
                cursor.execute(f"DELETE FROM {table_name}")
                conn.commit()

    total_rows = len(combined_df)
    if total_rows == 0:
        if progress_message:
            await safe_edit_text(progress_message, "Нет данных для загрузки 🚫")
        return

    update_step = max(1, total_rows // 20)  # Обновление раз в 5% загрузки
    progress = 0

    # Вставка данных чанками с новыми соединениями
    placeholders = ", ".join(["%s"] * len(final_columns))
    sql_insert = f"INSERT INTO {table_name} ({', '.join(final_columns)}) VALUES ({placeholders})"

    for start in range(0, total_rows, CHUNK_SIZE):
        end = min(start + CHUNK_SIZE, total_rows)
        chunk = combined_df.iloc[start:end]

        batch_data = [tuple(row.where(pd.notna(row), None)) for _, row in chunk.iterrows()]
        await execute_with_retry(sql_insert, batch_data, db_name)

        new_progress = round((end) / total_rows * 100)
        if new_progress % 5 == 0 and new_progress > progress:
            progress = new_progress
            if progress_message:
                await safe_edit_text(progress_message, f"{progress}%")

    if progress_message:
        await safe_edit_text(progress_message, "100% — Загрузка завершена! ✅")

    print(f"Данные загружены в {table_name} из {len(file_paths)} файлов.")


async def safe_edit_text(message, text):
    """Безопасное редактирование сообщения, чтобы избежать ошибок TelegramBadRequest"""
    for _ in range(3):  # До 3 попыток
        try:
            await message.edit_text(text)
            break
        except TelegramBadRequest:
            break  # Ошибка означает, что текст не изменился — пропускаем
        except TelegramNetworkError:
            await asyncio.sleep(1)  # Ждём 1 секунду и пробуем снова
