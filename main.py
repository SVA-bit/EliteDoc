import requests
import sqlite3
import asyncio
import logging
import sys
import os
import locale
import pymorphy3

from docxtpl import DocxTemplate
from datetime import datetime

from typing import Any, Dict

from TOKEN import API_TOKEN, API_KEY

from aiogram import Bot, Dispatcher, F, Router, html
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InputFile,
    FSInputFile,
)

morph = pymorphy3.MorphAnalyzer()
router = Router()

from aiogram.utils.chat_action import ChatActionSender
from aiogram.types import Message, FSInputFile, InputMediaPhoto, InputMediaVideo

#Машина состояний

class Form(StatesGroup):    #Формирование договора
    ogrn = State()          #1. Получение ОГРН Юр.(13) ИП(15) API - запрос          // Вносим в БД
    address = State()       #2. Прописки ИП нет в открытом доступе запрос           // Вносим в БД
    bik = State()           #3. Банковские реквизиты API - запрос по БИКу           // Вносим в БД
    rs = State()            #4. Получаем расчётный счёт                             // Вносим в БД
    mail = State()          #5. Получаем E-mail                                     // Вносим в БД
    tel = State()           #6. Получаем Телефон, Достаём все ранее полученные данные из БД
                            ## Подставляем данные в шаблоны из папки /doc отправляем ползователю 


class User(StatesGroup):    # Данные о пользователе для формирования служебки /copy
    user_name = State()     #1. Ф.И.О. 
    user_post = State()     #2. Должность


#Хендлер /start 
#1. В случае отсутствия БД создаёт её 
#2. Добавляет пользователя по Telegram ID

@router.message(CommandStart())
async def command_start(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    connection = sqlite3.connect('database.db')
    cursor = connection.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS Buyers (
    id INTEGER PRIMARY KEY,
    user_id TEXT,
    user_name TEXT,
    user_post TEXT,
    user_tel TEXT,
    user_sity TEXT, 
    ogrn TEXT,
    inn TEXT,
    kpp TEXT,
    datareg TEXT,
    fullname TEXT,
    name TEXT,
    dir TEXT,
    sex TEXT,
    address TEXT,
    bik TEXT,
    bank TEXT,
    ks TEXT,
    rs TEXT,
    mail TEXT,
    tel TEXT,
    addressg TEXT
    )
    ''')
    connection.commit()
    connection.close()
    try:
        sqlite_connection = sqlite3.connect('database.db')
        cursor = sqlite_connection.cursor()
        print("Подключен к SQLite")

        sqlite_insert_with_param = """INSERT INTO Buyers
                              (id, user_id)
                              VALUES (?,?);"""
        data_tuple = (user_id, user_id)
        cursor.execute(sqlite_insert_with_param, data_tuple)
        sqlite_connection.commit()
        print("Запись успешно вставлена ​​в таблицу sqlitedb_developers ")
        cursor.close()

    except sqlite3.Error as error:
        print("Ошибка при работе с SQLite", error)
    finally:
        if sqlite_connection:
            sqlite_connection.close()
            print("Соединение с SQLite закрыто")
    await state.set_state(Form.ogrn) 
    await message.answer("🗄 Введите ОГРН контрагента")

#Хендлер /limit
#Достаёт данные из БД
#По len(ogrn) выбирает шаблон доп. соглашения
#Формирует из них доп. соглашение по шаблону 

@router.message(Command("limit"))
@router.message(F.text.casefold() == "limit")
async def limit_handler(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    await message.answer(f"💵")
    await message.answer("📎Формирую Лимит📎")

    sqlite_connection = sqlite3.connect('database.db')
    cursor = sqlite_connection.cursor()
    sql_select_query = """SELECT * from Buyers WHERE user_id = ?"""
    cursor.execute(sql_select_query, (user_id,))
    records = cursor.fetchall()
    for row in records:
        ogrn = row[6]
        inn = row[7]
        kpp = row[8]
        datareg = row[9]
        fullname = row[10]
        name = row[11]
        dir = row[12]
        sex = row[13]
        address = row[14]
        bik = row[15]
        bank = row[16]
        ks = row[17]
        rs = row[18]
        mail = row[19]
        tel = row[20]
    cursor.close()
    print (records)

    if len(ogrn) == 13:
        words = dir.split()
        new_words = []
        for word in words:
           p = morph.parse(word)[0]
           new_word = p.inflect({'gent'}).word
           new_words.append(new_word)
           word = (" ".join(new_words))
        word = [int(x) if x.isdigit() else x for x in word.split()]
        word = [item.capitalize() for item in word]
        dir = (" ".join(word))
        sexs = sex.split()
        new_sexs = []
        for sex in sexs:
            p = morph.parse(sex)[0]
            new_sex = p.inflect({'gent'}).word
            new_sexs.append(new_sex)
            sex = (" ".join(new_sexs))
        sex = [int(x) if x.isdigit() else x for x in sex.split()]
        sex = [item.capitalize() for item in sex]
        sex = (" ".join(sex))

        doc = DocxTemplate("doc/Company_limit.docx")
        context = {'fullname': fullname, 'datareg': datareg, 'name': name, 'ogrn': ogrn, 'sex': sex, 'inn': inn, 'dir': dir, 
             'address': address, 'kpp': kpp, 'bik': bik, 'bank': bank, 'ks': ks, 'rs': rs, 'mail': mail, 'tel': tel}
        doc.render(context)
        name = str(name).replace('"', '')
        doc.save(f'Лимит {name}.docx')
        print(name)
        try:
            new_buyer = FSInputFile(f'Лимит {name}.docx')
            result = await message.answer_document(
                new_buyer,
            )
        finally:
            os.remove(f'Лимит {name}.docx')
        await state.set_state(Form.ogrn)
    elif len(ogrn) == 15:
        doc = DocxTemplate("doc/Entrepreneur_limit.docx")
        print("Взяли документ")
        context = {'fullname': fullname, 'datareg': datareg, 'name': name, 'ogrn': ogrn, 'inn': inn, 'dir': dir, 
             'address': address, 'kpp': kpp, 'bik': bik, 'bank': bank, 'ks': ks, 'rs': rs, 'mail': mail, 'tel': tel}
        doc.render(context)
        name = str(name).replace('"', '')
        doc.save(f'Лимит {name}.docx')
        try:
            new_buyer = FSInputFile(f'Лимит {name}.docx')
            result = await message.answer_document(
                new_buyer,
            )
        finally:
            os.remove(f'Лимит {name}.docx') 

    await state.set_state(Form.ogrn)
    await message.answer("🗄 Введите ОГРН")

#Хендлер /alert
#Достаёт данные из БД
#Формирует из них Уведомление о работе без печати по шаблону 

@router.message(Command("alert"))
@router.message(F.text.casefold() == "alert")
async def alert_handler(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    await message.answer(f"Уведомление о работе без печати")
    await message.answer("📎Ваш файлик📎")

    sqlite_connection = sqlite3.connect('database.db')
    cursor = sqlite_connection.cursor()
    print("Подключен к SQLite")
    sql_select_query = """SELECT * from Buyers WHERE user_id = ?"""
    cursor.execute(sql_select_query, (user_id,))
    records = cursor.fetchall()
    for row in records:
        ogrn = row[6]
        inn = row[7]
        name = row[11]
        address = row[14]
    cursor.close()
    print (records)
    if len(ogrn) == 13:
        await message.answer("Уведомление только для ИП")
    elif len(ogrn) == 15:
        doc = DocxTemplate("doc/Alert.docx")
        print("Взяли документ")
        context = {'name': name, 'ogrn': ogrn, 'inn': inn, 'address': address}
        doc.render(context)
        name = str(name).replace('"', '')
        doc.save(f'Уведомление {name}.docx')
        try:
            new_buyer = FSInputFile(f'Уведомление {name}.docx')
            result = await message.answer_document(
                new_buyer,
            )
        finally:
            os.remove(f'Уведомление {name}.docx')

    await state.set_state(Form.ogrn)
    await message.answer("🗄 Введите ОГРН")

#Хендлер /copy
#Достаёт данные из БД
#Проверяет на наличие Ф.И.О. и должности, 
#Формирует из них доп. соглашение по шаблону 
# в случае отсутствия данных, показывает команду на хендлер /hello  

@router.message(Command("copy"))
@router.message(F.text.casefold() == "copy")
async def copy_handler(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    await message.answer(f"Отсутствие оригиналов документов")
    

    sqlite_connection = sqlite3.connect('database.db')
    cursor = sqlite_connection.cursor()
    print("Подключен к SQLite")
    sql_select_query = """SELECT * from Buyers WHERE user_id = ?"""
    cursor.execute(sql_select_query, (user_id,))
    records = cursor.fetchall()
    for row in records:
        user_name = row[2]
        user_post = row[3]
        ogrn = row[6]
        name = row[11]
    cursor.close()
    print (records)
    if user_name == None:
        await message.answer("Необходимо познакомится /hello")
    else:
        await message.answer("📎Ваш файлик📎")
        doc = DocxTemplate("doc/Copy.docx")
        context = {'name': name, 'ogrn': ogrn, 'user_name': user_name, 'user_post': user_post}
        doc.render(context)
        name = str(name).replace('"', '')
        doc.save(f'Копии {name}.docx')
        try:
            new_buyer = FSInputFile(f'Копии {name}.docx')
            result = await message.answer_document(
                new_buyer
            )
            await state.set_state(Form.ogrn)
            await message.answer("🗄 Введите ОГРН")   
        finally:
            os.remove(f'Копии {name}.docx')

#Хендлер /hello
#Вносит Ф.И.О. и должность в БД

@router.message(Command("hello"))
@router.message(F.text.casefold() == "hello")
async def cancel_handler(message: Message, state: FSMContext) -> None:
    await state.set_state(User.user_name)
    await message.answer("✅(1/2) Введите ваше ФИО")
@router.message(User.user_name) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if message.text:
        user_name = message.text
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            sql_update_query = """Update Buyers set user_name = ? WHERE user_id = ?"""
            data4 = (user_name, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()

        await state.set_state(User.user_post)
        await message.answer("✅(2/2) Введите вашу должность")

@router.message(User.user_post) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if message.text:
        user_post = message.text
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            sql_update_query = """Update Buyers set user_post = ? WHERE user_id = ?"""
            data4 = (user_post, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()
                print("Соединение с SQLite закрыто")
        await state.clear()
        await message.answer("/copy")

# Основное Состояние пользователя 
# Получает сообщение ОГРН
# Если длина сообщения равна 13 знакам, то делаем API запрос по базе ЕГРЮЛ 
# Если длина сообщения равна 15 знакам, то делаем API запрос по базе ЕГРИП 
# Полученный json ответ разбираем на переменные и сохраняем в БД

# Если ИП переводим в состояние address
# Если ЮЛ переводим в состояние bik


@router.message(Form.ogrn) 
async def process_name(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id

    if len(message.text) == 13:

        url = ("https://api.checko.ru/v2/company")
        params = {"key": API_KEY, "ogrn": message.text, "source": "true"}
        
        r = requests.get(url, params)
        data = r.json()
        print(data)
        if len(data["data"]) > 5:

            ogrn = data["data"]["ОГРН"]
            inn = data["data"]["ИНН"]
            kpp = data["data"]["КПП"]
            fullname = data["data"]["НаимПолн"]
            fullname = str(fullname).replace('|', '') 
            name = data["data"]["НаимСокр"]
            name = str(name).replace('|', '')
            address = data["data"]["ЮрАдрес"]["АдресРФ"]
            address = str(address).replace('|', '')
            dir = data["data"]["Руковод"][0]["ФИО"]
            sex = data["data"]["Руковод"][0]["НаимДолжн"]
            datareg = ('NULL')

            try:
                sqlite_connection = sqlite3.connect('database.db')
                cursor = sqlite_connection.cursor()
                sql_update_query = """Update Buyers set ogrn = ?, inn = ?, kpp = ?, fullname = ?, name = ?, address = ?, dir = ?, datareg = ?, sex = ? WHERE user_id = ?"""
                data4 = (ogrn, inn, kpp, fullname, name, address, dir, datareg, sex, user_id)
                cursor.execute(sql_update_query, data4)
                sqlite_connection.commit()
                cursor.close()
            except sqlite3.Error as error:
                print("Ошибка при работе с SQLite", error)
            finally:
                if sqlite_connection:
                    sqlite_connection.close()
                

            await message.answer(f" ✅(1/5) = {name}")
            await state.set_state(Form.bik)
            await message.answer("🏦 Введите БИК \n \n<i>(В случае отсутствия банковских реквизитов введите '-')</i>")
        else:
            await message.answer("Не найдено ни одной организации с указанными реквизитами")

    elif len(message.text) == 15:
        url = ("https://api.checko.ru/v2/entrepreneur")
        params = {"key": API_KEY, "ogrn": message.text, "source": "true"}
        r = requests.get(url, params)
        data = r.json()
        print(data)

        if len(data["data"]) > 5:

            ogrn = data["data"]["ОГРНИП"]
            inn = data["data"]["ИНН"]
            kpp = data["data"]["ОКПО"]
            datareg = data["data"]["ДатаОГРНИП"]
            locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
            date_object = datetime.strptime(datareg, '%Y-%m-%d')
            datareg = date_object.strftime('%d.%m.%Y')
            fullname = data["data"]["Тип"]
            fullname = str(fullname).replace('|', '')
            name = data["data"]["ФИО"]
            name = str(name).replace('|', '')
            dir = ('NULL')
            sex = ('NULL')
            try:
                sqlite_connection = sqlite3.connect('database.db')
                cursor = sqlite_connection.cursor()
                sql_update_query = """Update Buyers set ogrn = ?, inn = ?, kpp = ?, datareg = ?, fullname = ?, name = ?, dir = ?, sex = ? WHERE user_id = ?"""
                data4 = (ogrn, inn, kpp, datareg, fullname, name, dir, sex, user_id)
                cursor.execute(sql_update_query, data4)
                sqlite_connection.commit()
                cursor.close()
            except sqlite3.Error as error:
                print("Ошибка при работе с SQLite", error)
            finally:
                if sqlite_connection:
                    sqlite_connection.close()


            await state.set_state(Form.address)
            await message.answer(f" ✅(1/6) = {name}")
            await message.answer("🏡 Введите прописку")
        else:
            await message.answer("ОГРНИП не найден")
    else:
        await message.answer("⛔️ ОГРН указан неверно")

# Состояние address
# Вносим данные о Юридическом адресе ИП

@router.message(Form.address) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if message.text:
        address = message.text
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            sql_update_query = """Update Buyers set address = ? WHERE user_id = ?"""
            data4 = (address, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()

        await state.set_state(Form.bik)
        await message.answer("🏦 Введите БИК \n <i>В случае отсутствия реквизитов введите '-'</i>")

# Состояние bik
# Проверка кол-во символов (9)
# Делаем API запорос 
# Полученный json ответ разбивам на переменные и сохраняем в БД

@router.message(Form.bik) 
async def process_bik(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id

    if len(message.text) == 9:
#        ogr = ("https://api.checko.ru/v2/bank") #Экономим запросы, лимит 100 зап/сут.
#        params = {"key": API_KEY, "bic": message.text} 
        ogr = ("https://bik-info.ru/api.html?type=json&bik=")
        params = {"type": "json", "bik": message.text}
        r = requests.get(ogr, params)
        data = r.json()
        print(data)
        if len(data["bik"]) > 5:
#            bik = data["data"]["БИК"]
#            bank = data["data"]["Наим"]
#            ks = data["data"]["КорСчет"]["Номер"]
            bik = data["bik"]
            ks = data["ks"]
            bank = data["namemini"]
            city = data["city"]
            bank = (f"{bank} г. {city}")
            print(bank)
            try:
                sqlite_connection = sqlite3.connect('database.db')
                cursor = sqlite_connection.cursor()
                sql_update_query = """Update Buyers set bik = ?, bank = ?, ks = ? WHERE user_id = ?"""
                data4 = (bik, bank, ks, user_id)
                cursor.execute(sql_update_query, data4)
                sqlite_connection.commit()
                cursor.close()
            except sqlite3.Error as error:
                print("Ошибка при работе с SQLite", error)
            finally:
                if sqlite_connection:
                    sqlite_connection.close()

            await message.answer(f" ✅(2/6) {bank}")
            await state.set_state(Form.rs)
            await message.answer("💰 Введите расчётный счёт")
        else:
            await message.answer("БИК указан неверно")

    elif message.text == "-":
        bik = bank = ks = rs = "-"
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            print("Подключен к SQLite")
            sql_update_query = """Update Buyers set bik = ?, bank = ?, ks = ?, rs = ? WHERE user_id = ?"""
            data4 = (bik, bank, ks, rs, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()
        
        await message.answer("Реквизиты отсутствуют")
        await message.answer("✉️ Введите E-mail")
        await state.set_state(Form.mail)
    else:
        le = len(message.text)
        print(le)
        await message.answer(f"⛔️ БИК указан неверно {le}/9")

# Состояние rs
# Проверка кол-во символов (20)
# Проверка первый символ (4)
# Вносим в БД

@router.message(Form.rs) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if len(message.text) == 20:
        rs = message.text
        if rs[0] == "4":   
            try:
                sqlite_connection = sqlite3.connect('database.db')
                cursor = sqlite_connection.cursor()
                sql_update_query = """Update Buyers set rs = ? WHERE user_id = ?"""
                data4 = (rs, user_id)
                cursor.execute(sql_update_query, data4)
                sqlite_connection.commit()
                cursor.close()
            except sqlite3.Error as error:
                print("Ошибка при работе с SQLite", error)
            finally:
                if sqlite_connection:
                    sqlite_connection.close()

            await message.answer(f"✅(3/5)") #= {rs}
            await state.set_state(Form.mail)
            await message.answer("✉️ Введите E-mail")
        else:
            await message.answer(f"Указан к/с, расчётный счёт начинается с 4... ")
    else:
        le = len(message.text)
        await message.answer(f"⛔️ Р/С указан неверно - {le}/20")

# Состояние mail
# Вносим в БД

@router.message(Form.mail) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if message.text:
        mail = message.text
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            sql_update_query = """Update Buyers set mail = ? WHERE user_id = ?"""
            data4 = (mail, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()

        await message.answer(f" ✅(4/5)") #= {mail}
        await state.set_state(Form.tel)
        await message.answer("📞 Введите номер телефона")


# Состояние tel
# Вносим номер телефона в БД

# По Telegram ID Достаём из БД все ранее полученные данные
# Разбираем их на переменные 
# По len(ogrn) выбираем шаблон договора
# Подставляем данные в шаблон 
# Отправляем файл пользователю

@router.message(Form.tel) 
async def process_rs(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id
    if message.text:
        tel = message.text
        try:
            sqlite_connection = sqlite3.connect('database.db')
            cursor = sqlite_connection.cursor()
            sql_update_query = """Update Buyers set tel = ? WHERE user_id = ?"""
            data4 = (tel, user_id)
            cursor.execute(sql_update_query, data4)
            sqlite_connection.commit()
            cursor.close()
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if sqlite_connection:
                sqlite_connection.close()

        await message.answer(f" ✅(5/5)") #= {tel}
        await message.answer("📎Формирование договора📎")

        sqlite_connection = sqlite3.connect('database.db')
        cursor = sqlite_connection.cursor()
        sql_select_query = """SELECT * from Buyers WHERE user_id = ?"""
        cursor.execute(sql_select_query, (user_id,))
        records = cursor.fetchall()
        for row in records:
            ogrn = row[6]
            inn = row[7]
            kpp = row[8]
            datareg = row[9]
            fullname = row[10]
            name = row[11]
            dir = row[12]
            sex = row[13]
            address = row[14]
            bik = row[15]
            bank = row[16]
            ks = row[17]
            rs = row[18]
            mail = row[19]
            tel = row[20]
        cursor.close()
        print (records)


        if len(ogrn) == 13:
            words = dir.split()
            new_words = []
            for word in words:
               p = morph.parse(word)[0]
               new_word = p.inflect({'gent'}).word
               new_words.append(new_word)
               word = (" ".join(new_words))
            word = [int(x) if x.isdigit() else x for x in word.split()]
            word = [item.capitalize() for item in word]
            dir = (" ".join(word))
            sexs = sex.split()
            new_sexs = []
            for sex in sexs:
                p = morph.parse(sex)[0]
                new_sex = p.inflect({'gent'}).word
                new_sexs.append(new_sex)
                sex = (" ".join(new_sexs))
            sex = [int(x) if x.isdigit() else x for x in sex.split()]
            sex = [item.capitalize() for item in sex]
            sex = (" ".join(sex))


            doc = DocxTemplate("doc/Company.docx")
            context = {'fullname': fullname, 'datareg': datareg, 'name': name, 'ogrn': ogrn, 'sex': sex, 'inn': inn, 'dir': dir, 
                 'address': address, 'kpp': kpp, 'bik': bik, 'bank': bank, 'ks': ks, 'rs': rs, 'mail': mail, 'tel': tel}
            doc.render(context)
            name = str(name).replace('"', '')
            doc.save(f'{name}.docx')
            print(name)
            try:
                new_buyer = FSInputFile(f'{name}.docx')
                result = await message.answer_document(
                    new_buyer,
                    caption="Список документов:"
                )
                await message.answer("Проверте ФИО в шапке, я только учусь 🙄")
            finally:
                os.remove(f'{name}.docx')  # Удаляет новый файл
            await state.set_state(Form.ogrn)
            await message.answer("/limit - Доп. соглашение на отсрочку \n/copy - Согласование по сканам")
            await message.answer("Введите ОГРН")


        elif len(ogrn) == 15:
            doc = DocxTemplate("doc/Entrepreneur.docx")
            print("Взяли документ")
            context = {'fullname': fullname, 'datareg': datareg, 'name': name, 'ogrn': ogrn, 'inn': inn, 'dir': dir, 
                 'address': address, 'kpp': kpp, 'bik': bik, 'bank': bank, 'ks': ks, 'rs': rs, 'mail': mail, 'tel': tel}
            doc.render(context)
            name = str(name).replace('"', '')
            doc.save(f'{name}.docx')

            try:
                new_buyer = FSInputFile(f'{name}.docx')
                result = await message.answer_document(
                    new_buyer,
                    caption="Список документов:"
                )
            finally:
                os.remove(f'{name}.docx')
            await state.set_state(Form.ogrn)
            await message.answer('/limit - доп. соглашение на отсрочку \n/alert - Уведомление о работе без печати \n/copy - Согласование по сканам')
            await message.answer("🗄 Введите ОГРН")


####################################################################################

async def main():
    bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
    dp = Dispatcher()
    dp.include_router(router)

    await dp.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, stream=sys.stdout)
    asyncio.run(main())
