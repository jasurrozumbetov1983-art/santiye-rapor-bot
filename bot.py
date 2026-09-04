
import os, re
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")
TZ = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))
REPORT_HOUR = int(os.getenv("REPORT_HOUR", "15"))
REPORT_MINUTE = int(os.getenv("REPORT_MINUTE", "0"))
DATA_FILE = "raporlar.xlsx"

SITES = ["DATA CENTR","DMC","LOT13","LOT71","SKP","STADYUM","BWC","KÖKSARAY","MMP","MOS","PİRAMİT","RMC","TYM","YHP"]
ALIASES = {
    "DATA CENTER":"DATA CENTR", "DATA CENTR":"DATA CENTR",
    "STADIUM":"STADYUM", "STADYUM":"STADYUM",
    "PIRAMIT":"PİRAMİT", "PİRAMİT":"PİRAMİT",
    "PIRAMIT TOWER":"PİRAMİT", "PİRAMİT TOWER":"PİRAMİT",
    "BWC":"BWC", "SKP":"SKP", "LOT71":"LOT71", "LOT13":"LOT13",
    "DMC":"DMC", "MMP":"MMP", "MOS":"MOS", "RMC":"RMC", "TYM":"TYM", "YHP":"YHP",
    "KOKSARAY":"KÖKSARAY", "KÖKSARAY":"KÖKSARAY",
    "ELLIPSE GARDEN":"ELLIPSE GARDEN"
}
# Extra sites can be added here without changing the parser.
if "ELLIPSE GARDEN" not in SITES:
    SITES.append("ELLIPSE GARDEN")

MENU = ReplyKeyboardMarkup([["📊 Rapor Durumu"],["📥 Excel"]], resize_keyboard=True)

def ensure_excel():
    if os.path.exists(DATA_FILE): return
    wb = Workbook()
    ws = wb.active; ws.title = "Raporlar"
    ws.append(["Tarih","Saat","Şantiye","Gönderen","İşçi Sayısı","Rapor Metni","Mesaj ID"])
    wb.save(DATA_FILE)

def normalize(s):
    return s.upper().replace("İ","I").replace("Ş","S").replace("Ğ","G").replace("Ü","U").replace("Ö","O").replace("Ç","C")

def detect_site(text):
    n = normalize(text)
    # longest aliases first
    for alias, site in sorted(ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        if normalize(alias) in n:
            return site
    return None

def detect_total_workers(text):
    n = normalize(text)
    patterns = [
        r"(?:GENEL\s+TOPLAM|TOPLAM)\s*[:\-]?\s*(\d+)\s*(?:KISI|KİŞİ)?",
        r"(\d+)\s*(?:KISI|KİŞİ)\s*$"
    ]
    for p in patterns:
        hits = re.findall(p, n, flags=re.MULTILINE)
        if hits:
            try: return int(hits[-1])
            except: pass
    return None

def save_report(site, sender, workers, text, msg_id):
    ensure_excel()
    wb = load_workbook(DATA_FILE)
    ws = wb["Raporlar"]
    now = datetime.now(TZ)
    ws.append([now.strftime("%d.%m.%Y"), now.strftime("%H:%M"), site, sender, workers or "", text, msg_id])
    wb.save(DATA_FILE)

def today_reports():
    ensure_excel()
    wb = load_workbook(DATA_FILE, read_only=True)
    ws = wb["Raporlar"]
    today = datetime.now(TZ).strftime("%d.%m.%Y")
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == today and row[2]:
            result[row[2]] = row
    return result

def status_text():
    reports = today_reports()
    sent = [s for s in SITES if s in reports]
    missing = [s for s in SITES if s not in reports]
    t = f"🕐 {datetime.now(TZ).strftime('%H:%M')} Şantiye Rapor Durumu\n\n"
    t += f"✅ Rapor ileten şantiyeler ({len(sent)}):\n"
    t += "".join(f"• {s}\n" for s in sent) or "• Henüz yok\n"
    t += f"\n❌ Rapor iletmeyen şantiyeler ({len(missing)}):\n"
    t += "".join(f"• {s}\n" for s in missing) or "• Eksik rapor yok\n"
    t += "\n📝 Not: Yapılan işin raporunu vermek, işi yapmak kadar önemlidir.\n⚠️ Eksik olan raporları lütfen iletiniz."
    return t

async def start(update, context):
    await update.message.reply_text("🏗️ Şantiye Rapor Kontrol Botu\n\nRaporlarınızı gruba normal mesaj olarak gönderebilirsiniz.", reply_markup=MENU)

async def status(update, context):
    await update.message.reply_text(status_text(), reply_markup=MENU)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return
    text = update.message.text.strip()

    if text == "📊 Rapor Durumu":
        await status(update, context); return

    if text == "📥 Excel":
        if not ADMIN_CHAT_ID or str(update.effective_chat.id) != str(ADMIN_CHAT_ID):
            await update.message.reply_text("⛔ Bu bölüm sadece yönetici içindir."); return
        ensure_excel()
        with open(DATA_FILE, "rb") as f:
            await update.message.reply_document(f, filename="santiye_raporlari.xlsx")
        return

    site = detect_site(text)
    if not site:
        return

    workers = detect_total_workers(text)
    sender = update.effective_user.full_name if update.effective_user else "Bilinmiyor"
    save_report(site, sender, workers, text, update.message.message_id)

    # Confirm to sender only when the message looks like a real report.
    if workers is not None:
        await update.message.reply_text(f"✅ {site} raporu kaydedildi.\n👷 Toplam: {workers} kişi")
    else:
        await update.message.reply_text(f"✅ {site} raporu kaydedildi.\n⚠️ Toplam işçi sayısı raporda bulunamadı.")

async def daily_report(context):
    if ADMIN_CHAT_ID:
        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=status_text())

def main():
    if not BOT_TOKEN: raise RuntimeError("BOT_TOKEN .env içine girilmelidir.")
    ensure_excel()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("durum", status))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.job_queue.run_daily(daily_report, time=datetime.now(TZ).replace(hour=REPORT_HOUR, minute=REPORT_MINUTE, second=0, microsecond=0).timetz())
    print("Bot çalışıyor...")
    app.run_polling()

if __name__ == "__main__":
    main()
