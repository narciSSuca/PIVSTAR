import pandas as pd
from openpyxl import load_workbook

def find_column_name(df, possible_names):
    """ Ищет правильное название столбца в DataFrame """
    for col in df.columns:
        if col.strip() in possible_names:  # Убираем лишние пробелы
            return col
    raise ValueError(f"Не найден ни один из столбцов: {possible_names}")

def process_reportFour(report_file, reference_file, sheet_to_keep, output_file):
    # Загружаем отчёт
    with pd.ExcelFile(report_file) as xls:
        if sheet_to_keep not in xls.sheet_names:
            raise ValueError(f"Лист '{sheet_to_keep}' не найден в файле {report_file}")
        report_df = xls.parse(sheet_to_keep)

    # Загружаем справочник
    reference_df = pd.read_excel(reference_file)

    # Определяем корректные названия столбцов
    report_code_col = find_column_name(report_df, ["код ТТ КИС"])
    reference_code_col = find_column_name(reference_df, ["Код"])

    # Проверяем, что в справочнике есть нужные колонки
    if not {"TSM", "ESR"}.issubset(reference_df.columns):
        raise ValueError("В справочнике отсутствуют колонки 'TSM' и 'ESR'")

    # Приводим коды к единому формату (убираем пробелы, приводим к нижнему регистру)
    report_df[report_code_col] = report_df[report_code_col].astype(str).str.strip().str.lower()
    reference_df[reference_code_col] = reference_df[reference_code_col].astype(str).str.strip().str.lower()

    # Создаём словарь соответствий
    mapping = reference_df.set_index(reference_code_col)[["TSM", "ESR"]].to_dict("index")

    for index, row in report_df.iterrows():
        code = row[report_code_col]
        old_tsm, old_esr = row.get("TSM", None), row.get("ESR", None)
        new_values = mapping.get(code, {"TSM": None, "ESR": None})

        if new_values["TSM"] is not None or new_values["ESR"] is not None:
            report_df.at[index, "TSM"] = new_values["TSM"]
            if new_values["TSM"] in {"GC_Калуга_TSM", "GC_Тула_TSM", "GC_Смоленск_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Тула_Калуга_DSM"
            elif new_values["TSM"] in {"GC_Тамбов_TSM", "GC_Воронеж_TSM", "GC_Липецк_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Воронеж_Тамбов_DSM"
            elif new_values["TSM"] in {"GC_Курск_TSM", "GC_Белгород_TSM", "GC_Орел_TSM", "GC_Брянск_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Белгород_Курск_DSM"

            report_df.at[index, "ESR"] = new_values["ESR"]
            report_df.at[index, "виртуальность"] = "Реальный"  # Обновляем "Виртуальность"

    # Сохраняем новый файл
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name=sheet_to_keep, index=False)

    print(f"Файл '{output_file}' сохранён.")

def process_reportFive(report_file, reference_file, category_file, sheet_to_keep, output_file):
    # Загружаем отчёт
    with pd.ExcelFile(report_file) as xls:
        if sheet_to_keep not in xls.sheet_names:
            raise ValueError(f"Лист '{sheet_to_keep}' не найден в файле {report_file}")
        report_df = xls.parse(sheet_to_keep)

    # Загружаем справочник с TSM и ESR
    reference_df = pd.read_excel(reference_file)

    # Загружаем справочник категорий
    category_df = pd.read_excel(category_file)

    # Ищем название столбца с кодом в каждом файле
    report_code_col = find_column_name(report_df, ["Код", "XCRM GUID TT", "код ТТ КИС"])
    reference_code_col = find_column_name(reference_df, ["Код", "XCRM GUID TT", "код ТТ КИС"])

    # Проверяем наличие TSM и ESR в справочнике
    required_cols = {"TSM", "ESR"}
    if not required_cols.issubset(reference_df.columns):
        raise ValueError(f"В справке отсутствуют столбцы: {required_cols}")

    # Фильтруем строки, где TSM и ESR равны "агент"
    mask = (report_df["TSM"] == "агент") & (report_df["ESR"] == "агент")

    # Обновляем значения по коду
    mapping = reference_df.set_index(reference_code_col)[["TSM", "ESR"]].to_dict("index")
    report_df.loc[mask, ["TSM", "ESR"]] = report_df.loc[mask, report_code_col].map(mapping).apply(pd.Series)

    # Обновляем категории
    category_mapping = category_df.set_index("category_name")["category"].to_dict()
    report_df["category_correct"] = report_df["категория"].map(category_mapping).fillna("без категории")

    # Сохраняем в новый файл
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name=sheet_to_keep, index=False)

    print(f"Файл '{output_file}' сохранён. Оставлен только лист '{sheet_to_keep}'.")

def process_reportSeven(report_file, reference_file, category_file, new_reference_file, sheet_to_keep, output_file):
    # Загружаем отчёт
    with pd.ExcelFile(report_file) as xls:
        if sheet_to_keep not in xls.sheet_names:
            raise ValueError(f"Лист '{sheet_to_keep}' не найден в файле {report_file}")
        report_df = xls.parse(sheet_to_keep)

    # Загружаем справочники
    reference_df = pd.read_excel(reference_file)
    category_df = pd.read_excel(category_file)
    new_reference_df = pd.read_excel(new_reference_file)  # Новый справочник

    report_code_col = find_column_name(report_df, ["код ТТ КИС"])
    reference_code_col = find_column_name(reference_df, ["Код"])

    # Проверяем наличие TSM и ESR в справочнике
    if not {"TSM", "ESR"}.issubset(reference_df.columns):
        raise ValueError("В справочнике отсутствуют колонки 'TSM' и 'ESR'")

    # Приводим коды к единому формату (убираем пробелы, приводим к нижнему регистру)
    report_df[report_code_col] = report_df[report_code_col].astype(str).str.strip().str.lower()
    reference_df[reference_code_col] = reference_df[reference_code_col].astype(str).str.strip().str.lower()

    # Создаём словарь соответствий
    mapping = reference_df.set_index(reference_code_col)[["TSM", "ESR"]].to_dict("index")

    for index, row in report_df.iterrows():
        code = row[report_code_col]
        old_tsm, old_esr = row.get("TSM", None), row.get("ESR", None)
        new_values = mapping.get(code, {"TSM": None, "ESR": None})

        if new_values["TSM"] is not None or new_values["ESR"] is not None:
            report_df.at[index, "TSM"] = new_values["TSM"]
            if new_values["TSM"] in {"GC_Калуга_TSM", "GC_Тула_TSM", "GC_Смоленск_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Тула_Калуга_DSM"
            elif new_values["TSM"] in {"GC_Тамбов_TSM", "GC_Воронеж_TSM", "GC_Липецк_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Воронеж_Тамбов_DSM"
            elif new_values["TSM"] in {"GC_Курск_TSM", "GC_Белгород_TSM", "GC_Орел_TSM", "GC_Брянск_TSM"}:
                 report_df.at[index, "DSM"] = "GC_PI_Белгород_Курск_DSM"

            report_df.at[index, "ESR"] = new_values["ESR"]
            report_df.at[index, "виртуальность"] = "Реальный"  # Обновляем "Виртуальность"

    # Обновляем категории
    category_mapping = category_df.set_index("category_name")["category"].to_dict()
    report_df["category_correct"] = report_df["категория"].map(category_mapping).fillna("без категории")

    # Обновляем "Источник заказа" по новому справочнику
    order_source_mapping = {
        10: "B2B",
        20: "B2B",
        30: "SFA",
        40: "Telesales",
        50: "заказ ч/з 1С"
    }

    new_reference_dict = new_reference_df.set_index("НомерДок")["КодпротоколЦен"].to_dict()
    report_df["Источник заказа"] = report_df["номер документа"].map(new_reference_dict).map(order_source_mapping)

    # Сохраняем в новый файл
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name=sheet_to_keep, index=False)

    print(f"Файл '{output_file}' сохранён. Оставлен только лист '{sheet_to_keep}'.")
# Пример вызова функции05315d67-1a4d-e711-80fa-005056011415
