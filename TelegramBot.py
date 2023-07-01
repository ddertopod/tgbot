import logging
import datetime
import telebot
import telebot_calendar
from telebot_calendar import CallbackData, Calendar, RUSSIAN_LANGUAGE
from telebot.types import ReplyKeyboardRemove, CallbackQuery
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from telebot import types
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv()
API_TOKEN = os.getenv('SUPERIMPORTANT')
logger = telebot.logger
telebot.logger.setLevel(logging.DEBUG)
bot = telebot.TeleBot(API_TOKEN)
calendar = Calendar(language=RUSSIAN_LANGUAGE)
calendar_1_callback = CallbackData("calendar_1", "action", "year", "month", "day")
calendar_2_callback = CallbackData("calendar_2", "action", "year", "month", "day")
calendar_3_callback = CallbackData("calendar_3", "action", "year", "month", "day")
calendar_4_callback = CallbackData("calendar_4", "action", "year", "month", "day")
@bot.message_handler(commands=['start'])
def start_message(message):
  markup=types.ReplyKeyboardMarkup(resize_keyboard=True)
  item1=types.KeyboardButton("Сырьевые материалы")
  item2=types.KeyboardButton("Стальная продукция")
  item3=types.KeyboardButton("Ферросплавы (Кремний и марганец)")
  item4=types.KeyboardButton("Ферросплавы (Хром)")
  markup.add(item1, item2, item3, item4)
  bot.send_message(message.chat.id,"Привет! Для получения отчета выбери тип отчета из списка ниже: ", reply_markup=markup)
@bot.message_handler(content_types='text')
def message_reply(message):
    if message.text=="Сырьевые материалы":
        now = datetime.now()
        bot.send_message(
            message.chat.id,
            "Выбор даты отчета",
            reply_markup=calendar.create_calendar(
                name=calendar_1_callback.prefix,
                year=now.year,
                month=now.month,
            ),
        )    
    elif message.text=="Стальная продукция":
        now = datetime.now()
        bot.send_message(
            message.chat.id,
            "Выбор даты отчета",
            reply_markup=calendar.create_calendar(
                name=calendar_2_callback.prefix,
                year=now.year,
                month=now.month,
            ),
        ) 
    elif message.text=="Ферросплавы (Кремний и марганец)":
        now = datetime.now()
        bot.send_message(
            message.chat.id,
            "Выбор даты отчета",
            reply_markup=calendar.create_calendar(
                name=calendar_3_callback.prefix,
                year=now.year,
                month=now.month,
            ),
        )  
    elif message.text=="Ферросплавы (Хром)":
        now = datetime.now()
        bot.send_message(
            message.chat.id,
            "Выбор даты отчета",
            reply_markup=calendar.create_calendar(
                name=calendar_4_callback.prefix,
                year=now.year,
                month=now.month,
            ),
        )   
global_date = None
@bot.callback_query_handler(
    func=lambda call: call.data.startswith(calendar_1_callback.prefix)
)
def callback_inline1(call: CallbackQuery):
    global global_date
    name, action, year, month, day = call.data.split(calendar_1_callback.sep)
    date1 = calendar.calendar_query_handler(
        bot=bot, call=call, name=name, action=action, year=year, month=month, day=day
    )
    if action == "DAY":
        bot.send_message(
            chat_id=call.from_user.id,
            text=f"You have chosen1 {date1.strftime('%d.%m.%Y')}",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_1_callback}: Day: {date1.strftime('%d.%m.%Y')}")
    elif action == "CANCEL":
        bot.send_message(
            chat_id=call.from_user.id,
            text="Cancellation",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_2_callback}: Cancellation")
    global_date = date1
    print(f"date1: {global_date}")
    with open("datetime.txt", "w") as file:
        file.write(global_date.strftime("%d.%m.%Y"))
    import Poiskdata
    import ConvertSur
    document = open('Sur.png', 'rb')
    bot.send_document(chat_id = call.from_user.id, document = document)
DATESur = global_date
global_date = None
@bot.callback_query_handler(
    func=lambda call: call.data.startswith(calendar_2_callback.prefix)
)
def callback_inline2(call: CallbackQuery):
    global global_date
    name, action, year, month, day = call.data.split(calendar_2_callback.sep)
    date2 = calendar.calendar_query_handler(
        bot=bot, call=call, name=name, action=action, year=year, month=month, day=day
    )
    if action == "DAY":
        bot.send_message(
            chat_id=call.from_user.id,
            text=f"You have chosen2 {date2.strftime('%d.%m.%Y')}",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_2_callback}: Day: {date2.strftime('%d.%m.%Y')}")
    elif action == "CANCEL":
        bot.send_message(
            chat_id=call.from_user.id,
            text="Cancellation",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_2_callback}: Cancellation")
    global_date = date2
    print(f"date2: {global_date}")
    with open("datetime.txt", "w") as file:
        file.write(global_date.strftime("%d.%m.%Y"))
    import Poiskdata
    import ConvertStal
    document = open('Stal.png', 'rb')
    bot.send_document(chat_id = call.from_user.id, document = document)
DATEStal = global_date
global_date = None
@bot.callback_query_handler(
    func=lambda call: call.data.startswith(calendar_3_callback.prefix)
)
def callback_inline3(call: CallbackQuery):
    global global_date
    name, action, year, month, day = call.data.split(calendar_3_callback.sep)
    date3 = calendar.calendar_query_handler(
        bot=bot, call=call, name=name, action=action, year=year, month=month, day=day
    )
    if action == "DAY":
        bot.send_message(
            chat_id=call.from_user.id,
            text=f"You have chosen3 {date3.strftime('%d.%m.%Y')}",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_3_callback}: Day: {date3.strftime('%d.%m.%Y')}")
    elif action == "CANCEL":
        bot.send_message(
            chat_id=call.from_user.id,
            text="Cancellation",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_3_callback}: Cancellation")
    global_date = date3
    print(f"date3: {global_date}")
    with open("datetime.txt", "w") as file:
        file.write(global_date.strftime("%d.%m.%Y"))
    import Poiskdata
    import ConvertFer1
    document = open('Fer1.png', 'rb')
    bot.send_document(chat_id = call.from_user.id, document = document)
DATEFER1 = global_date
global_date = None
@bot.callback_query_handler(
    func=lambda call: call.data.startswith(calendar_4_callback.prefix)
)
def callback_inline4(call: CallbackQuery):
    global global_date
    name, action, year, month, day = call.data.split(calendar_4_callback.sep)
    date4 = calendar.calendar_query_handler(
        bot=bot, call=call, name=name, action=action, year=year, month=month, day=day
    )
    if action == "DAY":
        bot.send_message(
            chat_id=call.from_user.id,
            text=f"You have chosen4 {date4.strftime('%d.%m.%Y')}",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_4_callback}: Day: {date4.strftime('%d.%m.%Y')}")
    elif action == "CANCEL":
        bot.send_message(
            chat_id=call.from_user.id,
            text="Cancellation",
            reply_markup=ReplyKeyboardRemove(),
        )
        print(f"{calendar_4_callback}: Cancellation")
    global_date = date4
    print(f"date4: {global_date}")
    with open("datetime.txt", "w") as file:
        file.write(global_date.strftime("%d.%m.%Y"))
    import Poiskdata
    import ConvertFer2
    document = open('Fer2.png', 'rb')
    bot.send_document(chat_id = call.from_user.id, document = document)
DATEFER2 = global_date
@bot.message_handler(commands=['data'])
def data_message(message):
    bot.send_message(message.chat.id, ' Дата: {}'.format(DATESur))
bot.infinity_polling()
