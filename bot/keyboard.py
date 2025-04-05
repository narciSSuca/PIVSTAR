from aiogram.types import ReplyKeyboardMarkup, KeyboardButton

# Создаём кнопки
button_downloadRepFour = KeyboardButton(text="Обновить 4 отчёт")
button_downloadRepFive = KeyboardButton(text="Обновить 5 отчёт")
button_downloadRepSeven = KeyboardButton(text="Обновить 7 отчёт")
button_downloadCumulativeCoverage = KeyboardButton(text="Обновить накопительное Январь - Март")
button_test = KeyboardButton(text="Скачать данные")

# Создаём клавиатуру с кнопкой "Скачать данные"
main_keyboard = ReplyKeyboardMarkup(
    keyboard=[[button_downloadRepFour, button_downloadRepFive, button_downloadRepSeven, button_downloadCumulativeCoverage]],  # Здесь должен быть список списков кнопок
    resize_keyboard=True
)
