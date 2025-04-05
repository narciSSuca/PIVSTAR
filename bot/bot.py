import os
import logging
from aiogram import Bot, Dispatcher
from dotenv import load_dotenv
from bot.handlers import router  # Теперь импортируем router  # Импортируем обработчики команд

# Загружаем переменные из .env
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

# Создаём объект бота и диспетчер для обработки команд
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

async def run_bot():
    dp.include_router(router)
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(run_bot())
