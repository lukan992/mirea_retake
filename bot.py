import asyncio
from aiogram import Bot, Dispatcher
from aiogram.types import Message
from aiogram.filters import CommandStart, Command


bot = Bot(token='8039359957:AAGxtqGUr27gl278dMpzJ-t8QqbMwzdWsYM')
dp = Dispatcher()


@dp.message(CommandStart())
async def cmd_start(message: Message):
    await message.answer('Введи номер студенческого билета')


async def main():
    await dp.start_polling(bot)


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print('бот выключен!')
