from telebot import *
from telebot.types import *

file_name = "None"

markupboshmenyu = InlineKeyboardMarkup(row_width=2)
item1 = InlineKeyboardButton("Excel",callback_data="Excel")

markupboshmenyu.add(item1)

channel_name = "@test_pythonbotchannel"
bot = TeleBot("BOT_TOKEN")

def Excel_File_Name(message):
    data = message.text.split("$")
    return data[1]

def Excel(excel_data,file_name):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    data = excel_data.split("$")
    print(data)
    
    data.remove("/excel")
    print(data)
   
    for i in range(0,len(data),2):
        print(data[i],data[i+1]) 
        ws["{0}".format(data[i])] = data[i+1]
        


    wb.save("{0}.xlsx".format(file_name))




@bot.message_handler(commands=['start',])
def start(message):
    human = message.from_user.id
    humans_in_channel = bot.get_chat_member(channel_name,human)
    
    if humans_in_channel.status != "left":
        bot.send_message(message.chat.id , "<b>Assalomu Alaykum {0} ! \nBotimizga Xush Kelibsiz ! </b>\n\nTanlang :",reply_markup=markupboshmenyu,parse_mode="html")
    else:
        bot.send_message(message.chat.id,"Kanalimizga ulaning : @test_pythonbotchannel")
    


@bot.message_handler(content_types=['text'])
def text(message):
    try:
        global file_name
        if "/excel" in message.text:
            bot.send_message(message.chat.id,"Tayyorlanmoqda...")
            Excel(message.text,file_name)
            bot.send_message(message.chat.id,"Tayyor!!!")
            doc = open('./{0}.xlsx'.format(file_name), 'rb')
            bot.send_document(message.chat.id,data=doc)
        
        if "/file_name" in message.text : 
            file_name = Excel_File_Name(message) 
            bot.send_message(message.chat.id,"Fayl Nomini Kiritildi!")
    except Exception as e:
        bot.send_message(message.chat.id,"Xatolik yuz berdi ! Iltimos /start ni qaytadan bosing.")
    

@bot.callback_query_handler(func=lambda call: True)
def call(call):
    if call.data == "Excel":
        #photo = open('./Excel.png', 'rb')
        photo = "https://cs5-1.4pda.to/17501056.png"
        bot.send_photo(call.message.chat.id, photo = photo , caption="Excel Botga Xush Kelibsiz !\n\nFoydalanish tartibi :\n/file_name$[fayl_nomi] buyrug'i orqali fayilingiz nomini kiritasiz.\n/excel$[katak_nomi]$[qiymat] orqali ma'lumot kiritasiz.")

    

bot.polling(none_stop=True)
