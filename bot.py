import telebot
from telebot import apihelper
import openpyxl
from datetime import datetime
import os
from fuzzywuzzy import fuzz
from flask import Flask, request

# Telegram Bot Token
TOKEN = "8518508714:AAH3b_A2UlaGiI8MaBRqkDBGZKwj_r5tTHM"
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)

# Har bir foydalanuvchining ma'lumotlarini vaqtinchalik xotirada saqlash uchun
user_data = {}

def ismlarni_standartlash(ism):
    """O'zbek ismlaridagi eng ko'p imlo xatolarini tozalash va standartlash"""
    if not ism:
        return ""
    ism = str(ism).strip().lower()
    # Har xil tutuq belgilari va o'zbekcha o'ziga xos harflarni standartlaymiz
    ism = ism.replace("`", "").replace("ʻ", "").replace("‘", "").replace("’", "").replace("'", "")
    ism = ism.replace("о‘", "o").replace("o‘", "o").replace("o'", "o").replace("о'", "o")
    ism = ism.replace("g‘", "g").replace("g'", "g").replace("г‘", "g")
    ism = ism.replace("ch", "c").replace("sh", "s")
    ism = ism.replace("x", "h").replace("ya", "a").replace("yu", "u")
    ism = ism.replace("i", "e").replace("a", "e")
    return "".join(ism.split())

@app.route('/' + TOKEN, methods=['POST'])
def getMessage():
    """Render tarmog'idan kelgan xabarlarni tutib olish ko'prigi"""
    json_string = request.get_data().decode('utf-8')
    update = telebot.types.Update.de_json(json_string)
    bot.process_new_updates([update])
    return "!", 200

@app.route("/")
def webhook():
    """Server holatini tekshirish uchun bosh sahifa"""
    return "Bot Render bulutida 24/7 rejimida xavfsiz faol!", 200

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    chat_id = message.chat.id
    user_data[chat_id] = {"holat": "sana_kutish"}
    
    bot.reply_to(message, "Salom! Kontraktlarni avtomatik yangilovchi tizim ishga tushdi. 🚀\n\n"
                          "To'lovlarni **qaysi sanadan boshlab** hisoblayiy?\n"
                          "Format namunasi: `27.06.2026` shaklida yozib yuboring:")

@bot.message_handler(func=lambda message: message.chat.id in user_data and user_data[message.chat.id].get("holat") == "sana_kutish")
def qabul_qilish_sanasi(message):
    chat_id = message.chat.id
    sana_matni = message.text.strip()
    
    try:
        cheklov_sanasi = datetime.strptime(sana_matni, "%d.%m.%Y")
        user_data[chat_id]["sana"] = cheklov_sanasi
        user_data[chat_id]["holat"] = "fayl_kutish"
        
        bot.send_message(chat_id, f"✅ Sana tasdiqlandi: **{sana_matni}**.\n\n"
                                  f"1. Endi **Asosiy bazangizni** (.xlsx) faylini yuboring.")
    except ValueError:
        bot.reply_to(message, "❌ Noto'g'ri format kiritdingiz. Iltimos nuqtalar bilan `DD.MM.YYYY` ko'rinishida yozing (Masalan: 27.06.2026):")

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    if chat_id not in user_data or "sana" not in user_data[chat_id]:
        bot.reply_to(message, "❌ Iltimos, avval /start buyrug'ini bosing va filtr sanasini kiriting!")
        return

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        baza_nomi = f"baza_{chat_id}.xlsx"
        deb_nomi = f"deb_{chat_id}.xlsx"
        natija_nomi = f"Tayyor_Yangilangan_{chat_id}.xlsx"

        # 1-Fayl: Asosiy baza yuklanganda
        if "baza_yuklandi" not in user_data[chat_id]:
            with open(baza_nomi, 'wb') as f:
                f.write(downloaded_file)
            user_data[chat_id]["baza_yuklandi"] = True
            user_data[chat_id]["baza_path"] = baza_nomi
            bot.reply_to(message, "✅ Asosiy baza muvaffaqiyatli qabul qilindi!\n\n"
                                  "2. Endi bankdan kelgan yangi **Debitorka** faylini yuboring.")
            return
        
        # 2-Fayl: Debitorka yuklanganda hisob-kitobni boshlash
        elif "deb_yuklandi" not in user_data[chat_id]:
            with open(deb_nomi, 'wb') as f:
                f.write(downloaded_file)
            user_data[chat_id]["deb_yuklandi"] = True
            
            baza_path = user_data[chat_id]["baza_path"]
            cheklov_sanasi = user_data[chat_id]["sana"]
            
            bot.reply_to(message, "🔄 Xavfsiz tahlil va taqqoslash algoritmi boshlandi, iltimos kuting...")

            # SAFE OPENING: Formulalarni buzmaslik uchun parallel ikki xil rejimda ochamiz
            wb_baza_write = openpyxl.load_workbook(baza_path, data_only=False) # Formulalarni tirik saqlaydi
            wb_baza_read = openpyxl.load_workbook(baza_path, data_only=True)   # Sonlarni o'qiydi
            wb_deb = openpyxl.load_workbook(deb_nomi, data_only=True)

            varoq_nomi = 'KONTRAKTLAR' if 'KONTRAKTLAR' in wb_baza_write.sheetnames else wb_baza_write.sheetnames[0]
            sheet_write = wb_baza_write[varoq_nomi]
            sheet_read = wb_baza_read[varoq_nomi]
            sheet_deb = wb_deb['bank']

            # 1. Asosiy bazadan talabalarni yig'ish (Jadval 25-qatordan pastda joylashgan)
            baza_talabalari = []
            for row in range(20, sheet_read.max_row + 1):
                fio = sheet_read.cell(row=row, column=3).value  # C ustun (F.I.SH)
                if fio and str(fio).strip() and not str(fio).lower().startswith(('familiya', 'f.i.sh', 'итого', 'jami')):
                    baza_talabalari.append({
                        "row": row,
                        "original_name": str(fio).strip(),
                        "clean_name": ismlarni_standartlash(fio)
                    })

            yangilanish_tarixi = []
            topilmaganlar = []
            
            # 2. Debitorka 'bank' varog'ini satrma-satr tahlil qilish
            for row in range(2, sheet_deb.max_row + 1):
                sana_val = sheet_deb.cell(row=row, column=1).value  # A ustun (Sana)
                if not sana_val: 
                    continue

                try:
                    if isinstance(sana_val, str):
                        if '.' in sana_val:
                            to_lov_sanasi = datetime.strptime(sana_val.strip(), "%d.%m.%y")
                        else:
                            to_lov_sanasi = datetime.strptime(sana_val.strip(), "%Y-%m-%d %H:%M:%S")
                    elif isinstance(sana_val, datetime):
                        to_lov_sanasi = sana_val
                    else:
                        continue
                except ValueError:
                    continue

                # Faqat foydalanuvchi belgilagan sanadan keyingi to'lovlarni olamiz
                if to_lov_sanasi >= cheklov_sanasi:
                    summa_val = sheet_deb.cell(row=row, column=7).value    # G ustun (Summa)
                    deb_fio = sheet_deb.cell(row=row, column=9).value      # I ustun (Ism)

                    if not deb_fio or not summa_val: 
                        continue

                    yangi_summa = float(summa_val)
                    deb_fio_clean = ismlarni_standartlash(deb_fio)

                    eng_yaxshi_moslik = None
                    eng_yuqori_ball = 0

                    # Fuzzy logic yordamida imlo xatolarini tekshirish
                    for talaba in baza_talabalari:
                        ball = fuzz.token_sort_ratio(deb_fio_clean, talaba["clean_name"])
                        if ball > eng_yuqori_ball:
                            eng_yuqori_ball = ball
                            eng_yaxshi_moslik = talaba

                    # 75% o'xshashlik bo'lsa, chala yozilgan ismlarni ham muvaffaqiyatli bog'laydi
                    if eng_yaxshi_moslik and eng_yuqori_ball >= 75:
                        target_row = eng_yaxshi_moslik["row"]
                        
                        # E ustundagi (5-ustun) eski to'lov summasini toza son holatda o'qiymiz
                        eski_summa_val = sheet_read.cell(row=target_row, column=5).value
                        try:
                            eski_summa = float(eski_summa_val) if eski_summa_val else 0.0
                        except (ValueError, TypeError):
                            eski_summa = 0.0

                        # Formulalarni o'chirmasdan, faqat qiymat ustiga yozamiz
                        jami_yangi = eski_summa + yangi_summa
                        sheet_write.cell(row=target_row, column=5).value = jami_yangi
                        
                        # Tarix hisobotiga kiritish
                        yangilanish_tarixi.append(
                            f"👤 **{eng_yaxshi_moslik['original_name']}**\n"
                            f"├ 🏦 Bank kvitansiyasida: `{deb_fio}`\n"
                            f"├ ➕ Tushgan summa: `{yangi_summa:,.0f} so'm`\n"
                            f"└ 📊 Eski holat: `{eski_summa:,.0f}` ➔ Yangi Jami: `{jami_yangi:,.0f} so'm`"
                        )
                    else:
                        topilmaganlar.append(f"❓ `{deb_fio}` — {yangi_summa:,.0f} so'm")

            # 3. Faylni yangi ko'rinishda xavfsiz saqlash
            wb_baza_write.save(natija_nomi)
            wb_baza_write.close()
            wb_baza_read.close()
            wb_deb.close()

            # 4. Batafsil matnli hisobot matnini tayyorlash
            hisobot_matni = f"📊 **KONTRAKT YANGILANISH HISOBOTI**\n"
            hisobot_matni += f"📅 Filtr sanasi: {cheklov_sanasi.strftime('%d.%m.%Y')} dan boshlab\n"
            hisobot_matni += f"━━━━━━━━━━━━━━━━━━━━\n\n"
            
            if yangilanish_tarixi:
                hisobot_matni += "✅ **Muvaffaqiyatli yangilangan talabalar:**\n\n"
                for t in yangilanish_tarixi[:12]: # Telegram xabari chegarasi sababli 12 ta ko'rsatiladi
                    hisobot_matni += t + "\n\n"
                if len(yangilanish_tarixi) > 12:
                    hisobot_matni += f"i... va yana {len(yangilanish_tarixi)-12} ta talaba yangilandi.\n\n"
            else:
                hisobot_matni += "❌ Kiritilgan sanadan keyin yangi to'lovlar topilmadi.\n\n"

            if topilmaganlar:
                hisobot_matni += "⚠️ **Bazadan ismi mos kelmagan bank to'lovlari:**\n"
                for top in topilmaganlar[:8]:
                    hisobot_matni += top + "\n"
                if len(topilmaganlar) > 8:
                    hisobot_matni += f"i... va yana {len(topilmaganlar)-8} ta qator topilmadi."

            # 5. Telegramga tayyor Excel fayl va matnli hisobotni yuborish
            with open(natija_nomi, 'rb') as f_send:
                bot.send_document(chat_id, f_send, caption="📄 Hamma varaqlari va formulalari toliq saqlangan Excel faylingiz.")
            
            bot.send_message(chat_id, hisobot_matni, parse_mode="Markdown")

            # Vaqtinchalik fayllarni xavfsiz tozalash
            if os.path.exists(baza_nomi): os.remove(baza_nomi)
            if os.path.exists(deb_nomi): os.remove(deb_nomi)
            if os.path.exists(natija_nomi): os.remove(natija_nomi)
            user_data[chat_id] = {}

    except Exception as e:
        bot.send_message(chat_id, f"❌ Xatolik yuz berdi: {str(e)}")
        user_data[chat_id] = {}

# Render server ishga tushganda Webhook ulanish qismi
if __name__ == "__main__":
    bot.remove_webhook()
    # DIQQAT: O'zingizning Render ko'k havolangizni quyidagi joyga to'g'ri yozing!
    bot.set_webhook(url="https://shahrisabz-kontrakt-bot.onrender.com/" + TOKEN) 
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 5000)))
