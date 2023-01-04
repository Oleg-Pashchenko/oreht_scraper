import dataclasses
import logging

import asyncio
import math

from aiogram import Bot, Dispatcher, types
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.reader.excel import load_workbook

API_TOKEN = "5398420004:AAF1PxwUSTnPPYZCC2hnVpOovMM2YcOTdEc"
logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)


@dataclasses.dataclass
class Data:
    article: str
    commission: str
    logistic: str
    price: float


async def scraper(code: str) -> float:
    response = requests.get(
        f"https://www.oreht.ru/modules.php?name=orehtPriceLS&op=ShowInfo&code={code}"
    )
    soup = BeautifulSoup(response.text, features="html.parser")
    try:
        return float(
            soup.find("span", {"class": "mg-price-n"}).text
            + "."
            + soup.find_all("span", {"class": "mg-price-n"})[-1].text
        )
    except:
        return 0


async def read_input_file(filename: str) -> list[Data]:
    dataframe = openpyxl.load_workbook(filename).active
    return [Data(article=str(row[0].value).split('-')[-1], commission=str(row[3].value), logistic=str(row[5].value)) for
            row in dataframe]


async def calculate_value(data: Data):
    b4 = data.commission
    c4 = data.logistic
    d4 = 35
    m4 = data.price
    o4 = m4
    p4 = 37
    q4 = o4 * p4
    r4 = b4 / 100
    v2 = 7
    t4 = c4 * d4 * 20
    t1 = 100
    u4 = (t4 + 33) / (t1 - 33)
    w4 = (o4 + q4 + u4) / (1 - r4) / (1 - v2)
    if w4 * 0.05 >= 250:
        a4 = 250
    else:
        if w4 * 0.05 >= 20:
            if w4 * 0.05 < 20:
                a4 = 20
            else:
                a4 = w4 * 0.05
        else:
            a4 = 20
    s4 = w4 * r4
    v4 = w4 * v2
    x2 = 30
    y2 = 0
    z4 = w4 / (1 - x2) / (1 - y2)
    x4 = z4 * x2
    y4 = (z4 - x4) * y2
    aa4 = math.ceil(z4) - 10
    ab4 = w4 / o4
    ac4 = q4 / w4
    ad4 = w4 - v4 - u4 - s4 - o4
    ae4 = ad4 / o4


async def write_to_output_file():
    pass


@dp.message_handler(content_types=["document"])
async def bot_handler(message: types.Message):
    filename = message.document.file_name
    await message.document.download(destination_file=filename)  # download document
    info = await read_input_file(filename)
    for i in range(len(info)):
        info[i].price = await scraper(info[i].article)

    await message.reply_document(open(filename, "rb"))  # send document


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(calculate_value())
    # asyncio.run(main())
