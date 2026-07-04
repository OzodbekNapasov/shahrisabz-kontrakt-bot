import telebot
import openpyxl
import os
import threading
import time
from fuzzywuzzy import fuzz
from flask import Flask

TOKEN = "8518508714:AAH3b_A2UlaGiI8MaBRqkDBGZKwj_r5tTHM"
bot = telebot.TeleBot(TOKEN)

app = Flask(__name__)
baza_ombori = {}

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

@app.route("/")
def home():
    return "Tizim Render serverida 24/7 rejimida muvaffaqiyatli ishlamoqda!", 200

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    chat_id = message.chat.id
    baza_ombori[chat_id] = None
    bot.send_message(chat_id, "Salom Ozodbek aka! Tizim bulutda to'liq barqaror holatga keltirildi. 🚀\n\n"
                              "1. Birinchi bo'lib **Asosiy bazangizni** (.xlsx) faylini shu yerga yuboring.")

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    if chat_id not in baza_ombori:
        baza_ombori[chat_id] = None

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        baza_nomi = f"baza_{chat_id}.xlsx"
        deb_nomi = f"deb_{chat_id}.xlsx"
        natija_nomi = f"Tayyor_Yangilangan_{chat_id}.xlsx"

        if baza_ombori[chat_id] is None:
            with open(baza_nomi, 'wb') as f:
                f.write(downloaded_file)
            baza_ombori[chat_id] = baza_nomi
            bot.send_message(chat_id, "✅ Asosiy baza qabul qilindi!\n\n"
                                      "2. Endi bankdan kelgan yangi **Debitorka** faylini yuboring.")
            return
        else:
            with open(deb_nomi, 'wb') as f:
                f.write(downloaded_file)
            
            bot.send_message(chat_id, "🔄 Excel formulalarini himoyalash va aqlli taqqoslash boshlandi, kuting...")

            wb_baza_write = openpyxl.load_workbook(baza_nomi, data_only=False)
            wb_baza_read = openpyxl.load_workbook(baza_nomi, data_only=True)
            wb_deb = openpyxl.load_workbook(deb_nomi, data_only=True)

            sheet_write = wb_baza_write.active
            sheet_read = wb_baza_read.active
            sheet_deb = wb_deb.active

            ism_ustun = 3  
            tolov_ustun = 5 
            
            for r in range(1, 30):
                for c in range(1, 15):
                    val = str(sheet_read.cell(row=r, column=c).value or "").lower()
                    if any(x in val for x in ['familiya', 'f.i.sh', 'ism', 'sharfi']):
                        ism_ustun = c
                    if any(x in val for x in ['jami', 'to\'lagan summasi', 'to\'lov']):
                        tolov_ustun = c

            baza_talabalari = []
            for row in range(4, sheet_read.max_row + 1):
                fio = sheet_read.cell(row=row, column=ism_ustun).value
                if fio and str(fio).strip() and not str(fio).lower().startswith(('итого', 'jami', 'yangilangan', 'guruh')):
                    baza_talabalari.append({
                        "row": row,
                        "original_name": str(fio).strip(),
                        "clean_name": ismlarni_standartlash(fio)
                    })

            yangilanish_tarixi = []
            topilmaganlar = []
            
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
                    eski_summa_val = sheet_read.cell(row=target_row, column=tolov_ustun).value
                    try:
                        eski_summa = float(eski_summa_val) if eski_summa_val else 0.0
                    except Exception:
                        eski_summa = 0.0

                    sheet_write.cell(row=target_row, column=tolov_ustun).value = eski_summa + yangi_summa
                    
                    yangilanish_tarixi.append(
                        f"👤 **{eng_yaxshi_moslik['original_name']}**\n"
                        f"├ 🏦 Bankda: `{deb_fio}`\n"
                        f"├ ➕ Pul: `+{yangi_summa:,.0f} so'm`\n"
                        f"└ 📊 Jami: `{(eski_summa + yangi_summa):,.0f} so'm`"
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
                for t in yangilanish_tarixi[:10]:
                    hisobot_matni += t + "\n\n"
                if len(yangilanish_tarixi) > 10:
                    hisobot_matni += f"i... va yana {len(yangilanish_tarixi)-10} ta talaba yangilandi.\n\n"
            else:
                hisobot_matni += "❌ Yangi mos keluvchi to'lovlar topilmadi.\n\n"

            if topilmaganlar:
                hisobot_matni += "⚠️ **Ismi topilmagan bank to'lovlari:**\n"
                for top in topilmaganlar[:6]:
                    hisobot_matni += top + "\n"

            with open(natija_nomi, 'rb') as f_send:
                bot.send_document(chat_id, f_send, caption="📄 Formulalari buzilmagan tayyor Excel faylingiz.")
            
            bot.send_message(chat_id, hisobot_matni, parse_mode="Markdown")

            if os.path.exists(baza_nomi): os.remove(baza_nomi)
            if os.path.exists(deb_nomi): os.remove(deb_nomi)
            if os.path.exists(natija_nomi): os.remove(natija_nomi)
            baza_ombori[chat_id] = None

    except Exception as e:
        bot.send_message(chat_id, f"❌ Ichki xatolik yuz berdi:\n`{str(e)}`")

# Botni alohida fonda uzluksiz yurgizish funksiyasi
def run_bot():
    bot.remove_webhook()
    while True:
        try:
            bot.polling(none_stop=True, interval=2, timeout=20)
        except Exception:
            time.sleep(5)

if __name__ == "__main__":
    # Botni orqa fonda (Thread) xavfsiz boshlaymiz
    threading.Thread(target=run_bot, daemon=True).start()
    # Flask portini Render uchun ochiq tutamiz
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 5000)))
