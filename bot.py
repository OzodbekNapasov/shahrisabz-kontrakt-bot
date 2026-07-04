import telebot
import openpyxl
from datetime import datetime
import os
from fuzzywuzzy import fuzz
from flask import Flask, request

TOKEN = "8518508714:AAH3b_A2UlaGiI8MaBRqkDBGZKwj_r5tTHM"
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)
user_files = {}

def ismlarni_standartlash(ism):
    if not ism: return ""
    ism = str(ism).strip().lower()
    ism = ism.replace("`", "").replace("ʻ", "").replace("‘", "").replace("’", "").replace("'", "")
    ism = ism.replace("о‘", "o").replace("o‘", "o").replace("o'", "o").replace("о'", "o")
    ism = ism.replace("g‘", "g").replace("g'", "g").replace("г‘", "g")
    ism = ism.replace("ch", "c").replace("sh", "s")
    ism = ism.replace("x", "h").replace("ya", "a").replace("yu", "u")
    ism = ism.replace("i", "e").replace("a", "e")
    return "".join(ism.split())

@app.route('/' + TOKEN, methods=['POST'])
def getMessage():
    try:
        json_string = request.get_data().decode('utf-8')
        update = telebot.types.Update.de_json(json_string)
        bot.process_new_updates([update])
    except Exception:
        pass
    return "!", 200

@app.route("/")
def webhook():
    return "Bot Render bulutida 24/7 rejimida xavfsiz faol!", 200

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    chat_id = message.chat.id
    user_files[chat_id] = {} 
    bot.reply_to(message, "Salom Ozodbek aka! Tizim aloqasi muvaffaqiyatli o'rnatildi. 🚀\n\n"
                          "1. Birinchi bo'lib **Asosiy bazangizni** (.xlsx) faylini yuboring.")

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    if chat_id not in user_files:
        user_files[chat_id] = {}

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        baza_nomi = f"baza_{chat_id}.xlsx"
        deb_nomi = f"deb_{chat_id}.xlsx"
        natija_nomi = f"Tayyor_Yangilangan_{chat_id}.xlsx"

        # 1-Fayl: Asosiy baza
        if "baza_path" not in user_files[chat_id]:
            with open(baza_nomi, 'wb') as f:
                f.write(downloaded_file)
            user_files[chat_id]["baza_path"] = baza_nomi
            bot.reply_to(message, "✅ Asosiy baza qabul qilindi!\n\n"
                                  "2. Endi bankdan kelgan yangi **Debitorka** faylini yuboring.")
            return
        
        # 2-Fayl: Debitorka
        else:
            with open(deb_nomi, 'wb') as f:
                f.write(downloaded_file)
            
            baza_path = user_files[chat_id]["baza_path"]
            bot.reply_to(message, "🔄 Formulalarni himoyalash va aqlli taqqoslash boshlandi, biroz kuting...")

            wb_baza_write = openpyxl.load_workbook(baza_path, data_only=False)
            wb_baza_read = openpyxl.load_workbook(baza_path, data_only=True)
            wb_deb = openpyxl.load_workbook(deb_nomi, data_only=True)

            # Birinchi faol varaqlarni avtomatik tanlash
            sheet_write = wb_baza_write.active
            sheet_read = wb_baza_read.active
            sheet_deb = wb_deb.active

            # AQLLI TIZIM: Talabalar ro'yxati nechanchi qatordan boshlanishini avtomatik aniqlash
            start_row = 4
            for r in range(1, 40):
                val = sheet_read.cell(row=r, column=3).value
                if val and any(keyword in str(val).lower() for keyword in ['familiya', 'f.i.sh', 'ism', 'sharfi']):
                    start_row = r + 1
                    break

            baza_talabalari = []
            for row in range(start_row, sheet_read.max_row + 1):
                fio = sheet_read.cell(row=row, column=3).value # C ustun
                if fio and str(fio).strip() and not str(fio).lower().startswith(('итого', 'jami', 'yangilangan')):
                    baza_talabalari.append({
                        "row": row,
                        "original_name": str(fio).strip(),
                        "clean_name": ismlarni_standartlash(fio)
                    })

            yangilanish_tarixi = []
            topilmaganlar = []
            
            # Debitorka tahlili (G-summa, I-ism)
            for row in range(2, sheet_deb.max_row + 1):
                summa_val = sheet_deb.cell(row=row, column=7).value
                deb_fio = sheet_deb.cell(row=row, column=9).value

                if not deb_fio or not summa_val: continue

                try:
                    yangi_summa = float(summa_val)
                except ValueError:
                    continue

                deb_fio_clean = ismlarni_standartlash(deb_fio)
                eng_yaxshi_moslik = None
                eng_yuqori_ball = 0

                for talaba in baza_talabalari:
                    ball = fuzz.token_sort_ratio(deb_fio_clean, talaba["clean_name"])
                    if ball > eng_yuqori_ball:
                        eng_yuqori_ball = ball
                        eng_yaxshi_moslik = talaba

                if eng_yaxshi_moslik and eng_yuqori_ball >= 76:
                    target_row = eng_yaxshi_moslik["row"]
                    eski_summa_val = sheet_read.cell(row=target_row, column=5).value # E ustun
                    try:
                        eski_summa = float(eski_summa_val) if eski_summa_val else 0.0
                    except (ValueError, TypeError):
                        eski_summa = 0.0

                    sheet_write.cell(row=target_row, column=5).value = eski_summa + yangi_summa
                    
                    yangilanish_tarixi.append(
                        f"👤 **{eng_yaxshi_moslik['original_name']}**\n"
                        f"├ 🏦 Bankda: `{deb_fio}`\n"
                        f"├ ➕ Tushgan pul: `{yangi_summa:,.0f} so'm`\n"
                        f"└ 📊 Yangi Jami: `{(eski_summa + yangi_summa):,.0f} so'm`"
                    )
                else:
                    topilmaganlar.append(f"❓ `{deb_fio}` — {yangi_summa:,.0f} so'm")

            wb_baza_write.save(natija_nomi)
            wb_baza_write.close()
            wb_baza_read.close()
            wb_deb.close()

            hisobot_matni = f"📊 **KONTRAKT YANGILANISH HISOBOTI**\n━━━━━━━━━━━━━━━━━━━━\n\n"
            if yangilanish_tarixi:
                hisobot_matni += "✅ **Yangilangan talabalar:**\n\n"
                for t in yangilanish_tarixi[:12]:
                    hisobot_matni += t + "\n\n"
                if len(yangilanish_tarixi) > 12:
                    hisobot_matni += f"i... va yana {len(yangilanish_tarixi)-12} ta talaba yangilandi.\n\n"
            else:
                hisobot_matni += "❌ Yangi to'lovlar topilmadi.\n\n"

            if topilmaganlar:
                hisobot_matni += "⚠️ **Ismi topilmagan bank to'lovlari:**\n"
                for top in topilmaganlar[:8]:
                    hisobot_matni += top + "\n"

            with open(natija_nomi, 'rb') as f_send:
                bot.send_document(chat_id, f_send, caption="📄 Formulalari buzilmagan tayyor Excel faylingiz.")
            
            bot.send_message(chat_id, hisobot_matni, parse_mode="Markdown")

            if os.path.exists(baza_nomi): os.remove(baza_nomi)
            if os.path.exists(deb_nomi): os.remove(deb_nomi)
            if os.path.exists(natija_nomi): os.remove(natija_nomi)
            user_files[chat_id] = {}

    except Exception as e:
        bot.send_message(chat_id, f"❌ Xatolik yuz berdi: {str(e)}")
        user_files[chat_id] = {}

if __name__ == "__main__":
    bot.remove_webhook()
    bot.set_webhook(url="https://shahrisabz-kontrakt-bot.onrender.com/" + TOKEN) 
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 5000)))
