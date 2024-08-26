from aiogram import Bot, Dispatcher, types
from TOKEN import API_TOKEN
import logging
from aiogram import Bot, Dispatcher, F, Router, html
from aiogram.enums import ParseMode
from main import form_router

async def main():
    bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
    dp = Dispatcher()
    dp.include_router(form_router)

    await dp.start_polling(bot)


bot = Bot(token=API_TOKEN, parse_mode="HTML")
dp = Dispatcher()

logging.basicConfig(level=logging.INFO)