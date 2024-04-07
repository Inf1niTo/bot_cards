from datetime import datetime

import openpyxl
import telebot
from telebot import types

book = openpyxl.open('/Users/antonplotnikov/Desktop/cards_for_bot.xlsx', read_only=True)
sheet = book.worksheets[0]

count = 3

iForStolbik = []

arrTypeToWrire = []
arrTypeToRead = []
arrToWrite = []

max_length = 30

while count != 100:
    if sheet[count][3].value not in arrTypeToRead:
        arrTypeToRead.append(sheet[count][3].value)
        if len(str(sheet[count][3].value)) > 30:
            item = sheet[count][3].value[:max_length] + "..."
            arrTypeToWrire.append(item)
            count += 1
        else:
            arrTypeToWrire.append(sheet[count][3].value)
            count += 1
    else:
        count += 1
print(arrTypeToWrire)

bot = telebot.TeleBot('6528949920:AAH-owQ6O0xHlbuy5A1KSg0xKbhRS3yKNvY')


@bot.message_handler(commands=['start', 'main'])
def main(message):
    markup = types.InlineKeyboardMarkup()

    markup.add(types.InlineKeyboardButton('🎢 Развлечения', callback_data='fun'))
    markup.add(types.InlineKeyboardButton('🥗 Еда', callback_data='food'))

    bot.send_message(message.chat.id, "Здравствуйте, вас приветствует бот путеводитель по Москве."
                                      " Выберите категорию которая вам интересна:", reply_markup=markup)


@bot.callback_query_handler(func=lambda callback: True)
def callback_message(callback):
    if callback.data == 'delete':
        bot.delete_message(callback.message.chat.id, callback.message.message_id)

    elif callback.data == 'edit':
        bot.edit_message_text('edit message', callback.message.chat.id, callback.message.message_id)

    elif callback.data == 'fun':

        keyboard = []

        for element in arrTypeToWrire:
            button = types.InlineKeyboardButton(text=element, callback_data=element)
            keyboard.append([button])
        reply_markup = types.InlineKeyboardMarkup(keyboard)

        bot.send_message(callback.message.chat.id, "Выберите интересующую вас категорию развлечений:",
                         reply_markup=reply_markup)

    elif callback.data in arrTypeToWrire:

        index = arrTypeToWrire.index(callback.data)

        print(index)

        arrToWrite.clear()
        iForStolbik.clear()

        for i in range(3, 500):

            if sheet[i][3].value == callback.data:
                if len(sheet[i][1].value) > 30:
                    item = sheet[i][1].value[:max_length] + "..."
                    arrToWrite.append(item)
                    iForStolbik.append(i)
                else:
                    arrToWrite.append(sheet[i][1].value)
                    iForStolbik.append(i)
            else:
                print("vremya categoriy   " + str(datetime.now().strftime("%H:%M:%S")))

        keyboard_2 = []

        for element in arrToWrite:
            button = types.InlineKeyboardButton(text=element, callback_data=element)
            keyboard_2.append([button])

        reply_markup = types.InlineKeyboardMarkup(keyboard_2)

        bot.send_message(callback.message.chat.id, "Выберите интересующее вас место:", reply_markup=reply_markup)

    elif callback.data in arrToWrite:

        index = arrToWrite.index(callback.data)

        print("\n\n\n")

        iForStolbik[index]

        bot.send_message(callback.message.chat.id, f'🍻 <strong>{sheet[iForStolbik[index]][1].value}</strong> 🍻\n '
                                                   f'\n<b>Ссылка на источник:</b>{sheet[iForStolbik[index]][2].value}\n'
                                                   f'\n<b>Описание: </b>{sheet[iForStolbik[index]][4].value}\n'
                                                   f'\n{sheet[iForStolbik[index]][5].value}\n'
                                                   f'\n<b>Время работы: </b>{sheet[iForStolbik[index]][7].value}',
                         parse_mode="html")

        print("vremya poiska karti  " + str(datetime.now().strftime("%H:%M:%S")))


bot.polling(none_stop=True)
