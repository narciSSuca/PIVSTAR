import os
import time
import calendar
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium.webdriver.support import expected_conditions as EC

# Загрузка переменных окружения (логин и пароль из .env)
# load_dotenv()
USERNAME = 'VRG_PI_NM_X005'
PASSWORD = 'Pa$$295249X005'


def get_chrome_driver():
    """Создает и настраивает экземпляр WebDriver для Chrome."""
    try:
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # Убрал для теста загрузок
        options.add_argument("--start-maximized")
        options.add_argument("window-size=1920x1080")
        options.add_argument('--disable-gpu')
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--ignore-certificate-errors")

        download_directory = os.path.abspath(os.path.join(os.getcwd(), "reports"))
        if not os.path.exists(download_directory):
            os.makedirs(download_directory)

        prefs = {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            "profile.default_content_settings.popups": 0
        }

        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(service=service, options=options)
        print("Chrome Driver успешно инициализирован.")
        return driver
    except Exception as e:
        print(f"Ошибка при создании драйвера Chrome: {e}")
        raise

def wait_for_download(download_directory, timeout=60):
    """Ожидает завершения загрузки файла"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = os.listdir(download_directory)
        for file in files:
            if file.endswith(".xlsx"):  # Проверяем, что Excel-файл скачался
                print(f"Файл загружен: {file}")
                return os.path.join(download_directory, file)
        time.sleep(1)
    raise Exception("Файл не был загружен за отведенное время")

def wait_and_click(driver, by, value, timeout=30):
    """Ожидает, пока элемент станет кликабельным, затем кликает по нему."""
    try:
        print(f"Ожидание клика по элементу: {value}")
        wait = WebDriverWait(driver, timeout)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "animationload")))
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
        print(f"Элемент найден: {value}")
        element.click()
    except Exception as e:
        print(f"Ошибка при клике по элементу {value}: {e}")
        raise


def wait_and_type(driver, by, value, text, timeout=30):
    """Ожидает, пока поле ввода станет доступным, затем вводит текст."""
    try:
        print(f"Ожидание ввода в поле: {value}")
        wait = WebDriverWait(driver, timeout)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "animationload")))
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
        element.clear()
        element.send_keys(text)
        print(f"Введен текст в поле: {value}")
    except Exception as e:
        print(f"Ошибка при вводе текста в поле {value}: {e}")
        raise


def select_dropdown(driver, element_id, value):
    """Выбирает значение в выпадающем списке."""
    try:
        print(f"Ожидание выбора из выпадающего списка {element_id}")
        wait = WebDriverWait(driver, 30)
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "animationload")))
        select_element = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, element_id)))
        Select(select_element).select_by_value(value)
        print(f"Выбрано значение {value} из списка {element_id}")
    except Exception as e:
        print(f"Ошибка при выборе значения из выпадающего списка {element_id}: {e}")
        raise


def get_first_day_of_month():
    try:
        return datetime.today().replace(day=1).strftime("%d.%m.%Y")
    except Exception as e:
        print(f"Ошибка при получении первого дня месяца: {e}")
        raise


def get_current_date():
    try:
        return datetime.today().strftime("%d.%m.%Y")
    except Exception as e:
        print(f"Ошибка при получении текущей даты: {e}")
        raise


def get_last_day_of_month():
    """Получает последний день текущего месяца."""
    try:
        today = datetime.today()
        first_day_weekday, last_day = calendar.monthrange(today.year, today.month)
        last_day_date = today.replace(day=last_day)
        return last_day_date.strftime("%d.%m.%Y")
    except Exception as e:
        print(f"Ошибка при получении последнего дня месяца: {e}")
        raise

def uncheck_checkbox(driver, by, value):
    checkbox = driver.find_element(by, value)
    if checkbox.is_selected():
        checkbox.click()

def login(driver):
    """Авторизация на сайте."""
    try:
        wait_and_type(driver, By.ID, 'MainContent_UserName', USERNAME)
        wait_and_type(driver, By.ID, 'MainContent_Password', PASSWORD)
        wait_and_click(driver, By.NAME, 'ctl00$MainContent$ctl05')
    except Exception as e:
        print(f"Ошибка при авторизации: {e}")
        raise


def select_bd_anp_space(driver, space):
    """Выбор базы данных/пространства."""
    try:
        driver.save_screenshot("debug_screenshot.png")
        with open("debug_page.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        wait_and_click(driver, By.ID, 'MainContent_cmdGetDB')
        select_dropdown(driver, "MainContent_DBComboBox", space)
    except Exception as e:
        print(f"Ошибка при выборе базы данных: {e}")
        raise


def select_org(driver):
    """Выбор организации."""
    try:
        wait_and_click(driver, By.ID, "MainContent_TreeSalesn15CheckBox")
    except Exception as e:
        print(f"Ошибка при выборе организации: {e}")
        raise

def select_org_Voronej(driver):
    """Выбор организации."""
    try:
        wait_and_click(driver, By.ID, "MainContent_TreeSalesn22CheckBox")
    except Exception as e:
        print(f"Ошибка при выборе организации: {e}")
        raise


def configure_report_four(driver):
    """Настройка параметров отчета 04."""
    try:
        print("Настройка отчета 04...")
        select_dropdown(driver, "MainContent_ReportComboBox", "04 Продажи (Руб/Шт/Уп/Кор/Кг) по ТТ/каналам/сетям, ТК")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', get_first_day_of_month())
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', get_current_date())
        wait_and_click(driver, By.CSS_SELECTOR, 'label[for="attr_vkl2"]', 50)
        time.sleep(2)
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn126')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn130CheckBox')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn132CheckBox')
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn130CheckBox')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn132CheckBox')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn126')
        time.sleep(10)
        # Возвращаю селекторы Атрибутов к исходному состоянию
    except Exception as e:
        print(f"Ошибка при настройке отчета 04: {e}")
        raise

def configure_report_sixty_one(driver):
    """Настройка параметров отчета 61."""
    try:
        print("Настройка отчета 61...")
        select_dropdown(driver, "MainContent_ReportComboBox", "61 Call Rate Strike Rate по ТК до ESR")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', get_first_day_of_month())
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', get_last_day_of_month())
        wait_and_click(driver, By.CSS_SELECTOR, 'label[for="attr_vkl2"]', 50)
        time.sleep(2)
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn126')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn130CheckBox')
        wait_and_click(driver, By.ID, 'MainContent_TreeAttributesn132CheckBox')
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        time.sleep(10)
        # Возвращаю селекторы Атрибутов к исходному состоянию
    except Exception as e:
        print(f"Ошибка при настройке отчета 61: {e}")
        raise

def configure_report_five(driver):
    """Настройка параметров отчета 5."""
    try:
        print("Настройка отчета 5...")
        select_dropdown(driver, "MainContent_ReportComboBox", "05 Продажи (Руб/Шт/Уп/Кор/Кг) по ТТ/каналам/сетям, ПрК/MML/DBC/SKU, ТК")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', get_first_day_of_month())
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', get_current_date())
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        time.sleep(10)

    except Exception as e:
        print(f"Ошибка при настройке отчета 5: {e}")
        raise

def configure_report_seven(driver):
    """Настройка параметров отчета 7."""
    try:
        print("Настройка отчета 7...")
        select_dropdown(driver, "MainContent_ReportComboBox", "07 Продажи (Руб/Шт/Уп/Кор/Кг) по документам, ТТ/каналам/сетям, ПрК/MML/DBC/SKU, ТК")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', '01.03.2025')
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', get_current_date())
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        time.sleep(10)

    except Exception as e:
        print(f"Ошибка при настройке отчета 5: {e}")
        raise

def configure_cumulative_coverage_jun_mar(driver):
    """Настройка параметров отчета 7 для накопительного покрытия."""
    try:
        print("Настройка отчета 7 для накопительного покрытия")
        select_dropdown(driver, "MainContent_ReportComboBox", "07 Продажи (Руб/Шт/Уп/Кор/Кг) по документам, ТТ/каналам/сетям, ПрК/MML/DBC/SKU, ТК")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', '01.02.2025')
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', get_current_date())
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        time.sleep(10)

    except Exception as e:
        print(f"Ошибка при настройке отчета 5: {e}")
        raise

def configure_cumulative_coverage_dek(driver):
    """Настройка параметров отчета 7 для накопительного покрытия."""
    try:
        print("Настройка отчета 7 для накопительного покрытия")
        select_dropdown(driver, "MainContent_ReportComboBox", "07 Продажи (Руб/Шт/Уп/Кор/Кг) по документам, ТТ/каналам/сетям, ПрК/MML/DBC/SKU, ТК")
        select_dropdown(driver, "MainContent_LimitComboBox", "0")
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker1', '01.01.2025')
        wait_and_type(driver, By.ID, 'MainContent_DateTimePicker2', '31.01.2025')
        wait_and_click(driver, By.ID, "MainContent_cmdRunReport")
        time.sleep(10)

    except Exception as e:
        print(f"Ошибка при настройке отчета 5: {e}")
        raise




def selenium_parserFourTula(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_report_four(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserFourVoronej(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_296 - (296) VRG-Пивстар (Воронеж)")
            time.sleep(3)
            select_org_Voronej(driver)
            time.sleep(3)
            configure_report_four(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserFiveVoronej(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_296 - (296) VRG-Пивстар (Воронеж)")
            time.sleep(3)
            select_org_Voronej(driver)
            time.sleep(3)
            configure_report_five(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserFiveTula(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_report_five(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserSevenVoronej(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_296 - (296) VRG-Пивстар (Воронеж)")
            time.sleep(3)
            select_org_Voronej(driver)
            time.sleep(3)
            configure_report_seven(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserSevenTula(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_report_seven(driver)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserCumulativeCoverageVoronej(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_296 - (296) VRG-Пивстар (Воронеж)")
            time.sleep(3)
            select_org_Voronej(driver)
            time.sleep(3)
            configure_cumulative_coverage_jun_mar(driver)
            time.sleep(10)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserCumulativeCoverageTula(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_cumulative_coverage_jun_mar(driver)
            time.sleep(10)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise




def selenium_parserCumulativeCoverageVoronejDec(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_296 - (296) VRG-Пивстар (Воронеж)")
            time.sleep(3)
            select_org_Voronej(driver)
            time.sleep(3)
            configure_cumulative_coverage_dek(driver)
            time.sleep(10)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise

def selenium_parserCumulativeCoverageTulaDec(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_cumulative_coverage_dek(driver)
            time.sleep(10)
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise




def selenium_parserSixtyOne(url):
    """Основная функция: открывает сайт, авторизуется и выгружает отчет."""
    try:
        with get_chrome_driver() as driver:
            driver.get(url)
            print("Страница загружена.")
            time.sleep(3)  # Ждем загрузку страницы
            login(driver)
            time.sleep(5)  # Ожидание загрузки после логина
            select_bd_anp_space(driver, "Mobile_Nestle_295 - (295) VRG-Пивстар (Тула)")
            time.sleep(3)
            select_org(driver)
            time.sleep(3)
            configure_report_sixty_one(driver)
            input('Press ENTER to exit')
    except Exception as e:
        print(f"Произошла ошибка в selenium_parser: {e}")
        raise


# if __name__ == "__main__":
#     try:
#         URL = "http://95.181.206.11/WebES250/Account/Login"
#         selenium_parser(URL)
#     except Exception as e:
#         print(f"Произошла ошибка при выполнении основного кода: {e}")
