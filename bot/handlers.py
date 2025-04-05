from aiogram import types, Router
from aiogram.filters import Command
from bot.keyboard import main_keyboard  # Импорт кнопок
from parsers.parser import selenium_parserSixtyOne, selenium_parserFourTula, selenium_parserFourVoronej, selenium_parserFiveVoronej, selenium_parserFiveTula, selenium_parserSevenVoronej, selenium_parserSevenTula, selenium_parserCumulativeCoverageVoronej, selenium_parserCumulativeCoverageTula, selenium_parserCumulativeCoverageVoronejDec,  selenium_parserCumulativeCoverageTulaDec  # Импорт парсера
from processors.processor import process_reportFour, process_reportFive, process_reportSeven  # Импорт обработки Excel
from pathlib import Path
from database.save_to_db import load_excel_to_mysql  # Функция сохранения в БД


router = Router()

@router.message(Command("start"))  #  Теперь это работает
async def start_handler(message: types.Message):
    await message.answer("Выберите выгрузку:", reply_markup=main_keyboard)


@router.message(lambda message: message.text == "Обновить 4 отчёт")
async def fetch_data(message: types.Message):
    await message.answer("Начинаю парсинг...")
    selenium_parserFourTula("http://95.181.206.11/WebES250/Account/Login")
    selenium_parserFourVoronej("http://95.181.206.11/WebES250/Account/Login")

    await message.answer("Обрабатываю данные...")
    process_reportFour("reports/Report04.xlsx", "reports/mp.xlsx", "Отчёт №04", "reports/Report04.xlsx")
    process_reportFour("reports/Report04 (1).xlsx", "reports/mp.xlsx", "Отчёт №04", "reports/Report04 (1).xlsx")
    column_types = {
        "INT": ["код_Региона", "Код_площадки", "код_ESРа", "код_ТТ", "Количество_шт"],
        "FLOAT": ["Сумма_продаж_руб", "Сумма_продаж_БЕЗ_НДС_руб", "Вес_кг"],
        "TEXT": ["Регион", "Площадка", "RSM", "TSM", "ESR", "сеть", "город", "ГОД_МЕСЯЦ"]
    }

    # Отправляем начальное сообщение о прогрессе
    progress_message = await message.answer("0%")

    # Загружаем данные в БД с обновлением прогресса
    await load_excel_to_mysql('default_db', 'report_four', ["reports/Report04.xlsx", "reports/Report04 (1).xlsx"], column_types,  progress_message)

    file_path_one = Path("reports/Report04.xlsx")
    file_path_two = Path("reports/Report04 (1).xlsx")

    if file_path_one.exists() and file_path_two.exists():
        file_path_two.unlink()
        file_path_one.unlink()
    await message.answer("База обновлена, хеш-данные удалены.")


@router.message(lambda message: message.text == "Обновить 5 отчёт")
async def fetch_data(message: types.Message):
    await message.answer("Начинаю парсинг...")
    selenium_parserFiveTula("http://95.181.206.11/WebES250/Account/Login")
    selenium_parserFiveVoronej("http://95.181.206.11/WebES250/Account/Login")

    await message.answer("Обрабатываю данные...")
    process_reportFive("reports/Report05.xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "Отчёт №05", "reports/Report05.xlsx")
    process_reportFive("reports/Report05 (1).xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "Отчёт №05", "reports/Report05 (1).xlsx")

    column_types = {
        "INT": ["код Региона", "Код площадки", "код ESRа", "код ТТ", "код товара", "Год", "Количество (шт.)"],
        "FLOAT": [
            "Табельный номер TSMа", "атрибут команды Пурины", "Паспорт ESR", "Дата приема на работу", "Дата увольнения",
            "ассортимент на КПК", "Бизнес ТТ", "тип ТТ Мороженое", "тип импульсного проекта Мороженое",
            "кол-во конечных торговых точек Мороженое", "Код Ключевой розницы", "Код Ключевого опта",
            "Пурина код клуба", "Пурина код заводчика", "Пурина код владельца", "Код ТУ традиция", "ТУ спецканал",
            "Код ТУ спецканал", "Код ТУ сети", "Сумма продаж (руб.)", "Сумма продаж БЕЗ НДС (руб.)",
            "Количество (уп.)", "Количество (кор.)", "Вес (кг.)"
        ],
        "VARCHAR(255)": [
            "Регион", "Площадка", "Бизнес RSM", "RSM", "Табельный номер RSMа", "DSM", "Табельный номер DSMа", "TSM",
            "ФИО TSM", "ESR", "ФИО ESR", "Tier", "принадлежность к стриму", "Стрим ТК", "Бизнес ESR", "виртуальность",
            "канал продаж ТП", "эксклюзивность", "XCRM GUID TT", "Дата заведения ТТ", "код ТТ КИС", "наименование ТТ",
            "краткое наименование ТТ", "адрес ТТ", "Специализация ТТ", "канал ТТ", "канал ТТ Пурина", "Ключевая розница",
            "Ключевой опт", "Передано IN", "ТУ традиция", "ТУ сети", "сеть", "протокол", "город", "MML", "Источник заказа",
            "DBC", "наименование DBC", "товар", "категория", "группа", "подгруппа", "Месяц","category_correct"
        ]
    }
    # Отправляем начальное сообщение о прогрессе
    progress_message = await message.answer("0%")

    # Загружаем данные в БД с обновлением прогресса
    await load_excel_to_mysql('default_db', 'report_five', ["reports/Report05.xlsx", "reports/Report05 (1).xlsx"], column_types,  progress_message)

    file_path_one = Path("reports/Report05.xlsx")
    file_path_two = Path("reports/Report05 (1).xlsx")

    if file_path_one.exists() and file_path_two.exists():
        file_path_two.unlink()
        file_path_one.unlink()
    await message.answer("База обновлена, хеш-данные удалены.")
    await message.answer("Готово! Данные сохранены.")


@router.message(lambda message: message.text == "Обновить 7 отчёт")
async def fetch_data(message: types.Message):
    await message.answer("Начинаю парсинг...")
    selenium_parserSevenTula("http://95.181.206.11/WebES250/Account/Login")
    selenium_parserSevenVoronej("http://95.181.206.11/WebES250/Account/Login")

    await message.answer("Обрабатываю данные...")
    process_reportSeven("reports/Report07.xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_tula.xlsx", "Отчёт №07", "reports/Report07.xlsx")
    process_reportSeven("reports/Report07 (1).xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_voronej.xlsx", "Отчёт №07", "reports/Report07 (1).xlsx")

    column_types = {
        "INT": ["код Региона", "Код площадки", "код ESRа", "код ТТ", "код товара", "Год", "Количество (шт.)"],
        "FLOAT": [
            "Табельный номер TSMа", "атрибут команды Пурины", "Паспорт ESR", "Дата приема на работу", "Дата увольнения",
            "ассортимент на КПК", "Бизнес ТТ", "тип ТТ Мороженое", "тип импульсного проекта Мороженое",
            "кол-во конечных торговых точек Мороженое", "Код Ключевой розницы", "Код Ключевого опта",
            "Пурина код клуба", "Пурина код заводчика", "Пурина код владельца", "Код ТУ традиция", "ТУ спецканал",
            "Код ТУ спецканал", "Код ТУ сети", "Сумма продаж (руб.)", "Сумма продаж БЕЗ НДС (руб.)",
            "Количество (уп.)", "Количество (кор.)", "Вес (кг.)"
        ],
        "VARCHAR(255)": [
            "Регион", "Площадка", "Бизнес RSM", "RSM", "Табельный номер RSMа", "DSM", "Табельный номер DSMа", "TSM",
            "ФИО TSM", "ESR", "ФИО ESR", "Tier", "принадлежность к стриму", "Стрим ТК", "Бизнес ESR", "виртуальность",
            "канал продаж ТП", "эксклюзивность", "XCRM GUID TT", "Дата заведения ТТ", "код ТТ КИС", "наименование ТТ",
            "краткое наименование ТТ", "адрес ТТ", "Специализация ТТ", "канал ТТ", "канал ТТ Пурина", "Ключевая розница",
            "Ключевой опт", "Передано IN", "ТУ традиция", "ТУ сети", "сеть", "протокол", "город", "MML", "Источник заказа",
            "DBC", "наименование DBC", "товар", "категория", "группа", "подгруппа", "Месяц","category_correct"
        ]
    }
    # Отправляем начальное сообщение о прогрессе
    progress_message = await message.answer("0%")

    # Загружаем данные в БД с обновлением прогресса
    await load_excel_to_mysql('default_db', 'report_seven', ["reports/Report07.xlsx", "reports/Report07 (1).xlsx"], column_types,  progress_message)

    file_path_one = Path("reports/Report07.xlsx")
    file_path_two = Path("reports/Report07 (1).xlsx")

    if file_path_one.exists() and file_path_two.exists():
        file_path_two.unlink()
        file_path_one.unlink()
    await message.answer("База обновлена, хеш-данные удалены.")



@router.message(lambda message: message.text == "Обновить накопительное Январь - Март")
async def fetch_data(message: types.Message):
    await message.answer("Начинаю парсинг...")
    selenium_parserCumulativeCoverageTula("http://95.181.206.11/WebES250/Account/Login")
    selenium_parserCumulativeCoverageVoronej("http://95.181.206.11/WebES250/Account/Login")

    selenium_parserCumulativeCoverageTulaDec("http://95.181.206.11/WebES250/Account/Login")
    selenium_parserCumulativeCoverageVoronejDec("http://95.181.206.11/WebES250/Account/Login")


    await message.answer("Обрабатываю данные...")
    process_reportSeven("reports/Report07.xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_tula.xlsx", "Отчёт №07", "reports/Report07.xlsx")
    process_reportSeven("reports/Report07 (1).xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_voronej.xlsx", "Отчёт №07", "reports/Report07 (1).xlsx")

    process_reportSeven("reports/Report07 (2).xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_tula.xlsx", "Отчёт №07", "reports/Report07 (2).xlsx")
    process_reportSeven("reports/Report07 (3).xlsx", "reports/mp.xlsx", "reports/volume_category.xlsx", "reports/dinamika_voronej.xlsx", "Отчёт №07", "reports/Report07 (3).xlsx")


    column_types = {
        "INT": ["код Региона", "Код площадки", "код ESRа", "код ТТ", "код товара", "Год", "Количество (шт.)"],
        "FLOAT": [
            "Табельный номер TSMа", "атрибут команды Пурины", "Паспорт ESR", "Дата приема на работу", "Дата увольнения",
            "ассортимент на КПК", "Бизнес ТТ", "тип ТТ Мороженое", "тип импульсного проекта Мороженое",
            "кол-во конечных торговых точек Мороженое", "Код Ключевой розницы", "Код Ключевого опта",
            "Пурина код клуба", "Пурина код заводчика", "Пурина код владельца", "Код ТУ традиция", "ТУ спецканал",
            "Код ТУ спецканал", "Код ТУ сети", "Сумма продаж (руб.)", "Сумма продаж БЕЗ НДС (руб.)",
            "Количество (уп.)", "Количество (кор.)", "Вес (кг.)"
        ],
        "VARCHAR(255)": [
            "Регион", "Площадка", "Бизнес RSM", "RSM", "Табельный номер RSMа", "DSM", "Табельный номер DSMа", "TSM",
            "ФИО TSM", "ESR", "ФИО ESR", "Tier", "принадлежность к стриму", "Стрим ТК", "Бизнес ESR", "виртуальность",
            "канал продаж ТП", "эксклюзивность", "XCRM GUID TT", "Дата заведения ТТ", "код ТТ КИС", "наименование ТТ",
            "краткое наименование ТТ", "адрес ТТ", "Специализация ТТ", "канал ТТ", "канал ТТ Пурина", "Ключевая розница",
            "Ключевой опт", "Передано IN", "ТУ традиция", "ТУ сети", "сеть", "протокол", "город", "MML", "Источник заказа",
            "DBC", "наименование DBC", "товар", "категория", "группа", "подгруппа", "Месяц","category_correct"
        ]
    }
    # Отправляем начальное сообщение о прогрессе
    progress_message = await message.answer("0%")

    # Загружаем данные в БД с обновлением прогресса
    await load_excel_to_mysql('default_db', 'cumulative_coverage', ["reports/Report07.xlsx", "reports/Report07 (1).xlsx", "reports/Report07 (2).xlsx", "reports/Report07 (3).xlsx"], column_types,  progress_message)

    file_path_one = Path("reports/Report07.xlsx")
    file_path_two = Path("reports/Report07 (1).xlsx")
    file_path_two = Path("reports/Report07 (2).xlsx")
    file_path_two = Path("reports/Report07 (3).xlsx")

    if file_path_one.exists() and file_path_two.exists():
        file_path_two.unlink()
        file_path_one.unlink()
    await message.answer("База обновлена, хеш-данные удалены.")
