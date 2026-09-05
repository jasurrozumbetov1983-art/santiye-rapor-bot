import os
import re
from datetime import datetime, date
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")
TZ = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))

HISTORY_FILE = "raporlar_tarih_2026_2.xlsx"
LIVE_FILE = "raporlar.xlsx"

REPORT_HOUR = int(os.getenv("REPORT_HOUR", "15"))
REPORT_MINUTE = int(os.getenv("REPORT_MINUTE", "0"))

SITES = [
    "DATA CENTER", "DMC", "LOT13", "LOT71", "SKP", "STADYUM",
    "BWC", "KÖKSARAY", "MPP", "MOS", "PİRAMİT", "RMC", "TYM",
    "YHP", "ELLIPSE GARDEN"
]

ALIASES = {
    "DATA CENTER": "DATA CENTER",
    "DMC": "DMC",
    "LOT13": "LOT13", "LOT 13": "LOT13",
    "LOT71": "LOT71", "LOT 71": "LOT71",
    "SKP": "SKP",
    "STADIUM": "STADYUM", "STADYUM": "STADYUM",
    "PIRAMIT": "PİRAMİT", "PİRAMİT": "PİRAMİT",
    "PIRAMIT TOWER": "PİRAMİT", "PİRAMİT TOWER": "PİRAMİT",
    "BWC": "BWC",
    "KOKSARAY": "KÖKSARAY", "KÖKSARAY": "KÖKSARAY",
    "MPP": "MPP", "MOS": "MOS", "RMC": "RMC", "TYM": "TYM",
    "YHP": "YHP", "ELLIPSE GARDEN": "ELLIPSE GARDEN",
}

MENU = ReplyKeyboardMarkup(
    [["🔎 Rapor Ara"], ["📊 Rapor Durumu", "📥 Excel"]],
    resize_keyboard=True
)

def normalize(text):
    if text is None:
        return ""
    text = str(text).upper()
    for old, new in {
        "İ": "I", "Ş": "S", "Ğ": "G", "Ü": "U",
        "Ö": "O", "Ç": "C", "Ə": "E"
    }.items():
        text = text.replace(old, new)
    return text.strip()

def detect_site(text):
    n = normalize(text)
    for alias, site in sorted(
        ALIASES.items(),
        key=lambda x: len(normalize(x[0])),
        reverse=True
    ):
        if normalize(alias) in n:
            return site
    return None

def parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in [
        "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y",
        "%Y-%m-%d", "%Y.%m.%d", "%d.%m.%y"
    ]:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None

def find_column(headers, names):
    normalized_headers = {}
    for index, header in enumerate(headers):
        if header is not None:
            normalized_headers[normalize(header)] = index
    for name in names:
        if normalize(name) in normalized_headers:
            return normalized_headers[normalize(name)]
    return None

def ensure_excel():
    if os.path.exists(HISTORY_FILE):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Raporlar"
    ws.append([
        "Tarih", "Saat", "Şantiye", "Gönderen",
        "İşçi Sayısı", "Rapor Metni", "Mesaj ID"
    ])
    wb.save(HISTORY_FILE)

def search_excel(start_date=None, end_date=None, site=None, keyword=None):
    if not os.path.exists(HISTORY_FILE):
        return []

    try:
        wb = load_workbook(HISTORY_FILE, read_only=True, data_only=True)
    except Exception as e:
        print("Excel ochishda xato:", e)
        return []

    results = []

    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        try:
            headers = next(rows)
        except StopIteration:
            continue

        date_col = find_column(headers, ["Tarih", "Sana", "Date", "Дата"])
        time_col = find_column(headers, ["Saat", "Vaqt", "Time", "Время"])
        site_col = find_column(headers, ["Şantiye", "Shantiye", "Шантиё"])
        sender_col = find_column(headers, ["Gönderen", "Yuboruvchi", "Sender", "Отправитель"])
        worker_col = find_column(headers, ["İşçi Sayısı", "Ishchi Sonı", "Рабочие", "Количество рабочих"])
        text_col = find_column(headers, ["Rapor Metni", "Rapor", "Hisobot", "Report", "Сообщение"])
        message_id_col = find_column(headers, ["Mesaj ID", "Message ID", "ID"])

        for row in rows:
            if not row:
                continue

            row_date = None
            if date_col is not None and date_col < len(row):
                row_date = parse_date(row[date_col])

            if (start_date or end_date) and row_date is None:
                continue
            if start_date and row_date < start_date:
                continue
            if end_date and row_date > end_date:
                continue

            row_site = ""
            if site_col is not None and site_col < len(row):
                row_site = str(row[site_col] or "").strip()

            if site and normalize(site) != normalize(row_site) and normalize(site) not in normalize(row_site):
                continue

            all_text = " ".join(str(x or "") for x in row)
            if keyword and normalize(keyword) not in normalize(all_text):
                continue

            results.append({
                "date": row_date,
                "time": row[time_col] if time_col is not None and time_col < len(row) else "",
                "site": row_site,
                "sender": row[sender_col] if sender_col is not None and sender_col < len(row) else "",
                "workers": row[worker_col] if worker_col is not None and worker_col < len(row) else "",
                "text": row[text_col] if text_col is not None and text_col < len(row) else "",
                "message_id": row[message_id_col] if message_id_col is not None and message_id_col < len(row) else "",
            })

    wb.close()
    return results

def parse_date_range(text):
    text = text.strip()

    match = re.search(
        r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s*[-–—]\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})",
        text
    )
    if match:
        d1 = parse_date(match.group(1))
        d2 = parse_date(match.group(2))
        if d1 and d2:
            return (d2, d1) if d1 > d2 else (d1, d2)

    match = re.search(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", text)
    if match:
        d = parse_date(match.group(0))
        if d:
            return d, d

    return None, None

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👷 Шантийе Рапорт Контроль Боти\n\n"
        "Рапорт қидириш учун 🔎 Rapor Ara тугмасини босинг.\n\n"
        "Мисол:\n01.01.2026 - 05.09.2026\n\n"
        "Ёки:\nPIRAMIT\n\nЁки:\n05.09.2026 PIRAMIT",
        reply_markup=MENU
    )

async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now(TZ).date()
    reports = search_excel(start_date=today, end_date=today)

    found_sites = sorted(set(r["site"] for r in reports if r["site"]))
    missing_sites = [s for s in SITES if s not in found_sites]

    text = (
        "📊 Шантийе Рапорт Ҳолати\n\n"
        f"📅 Сана: {today.strftime('%d.%m.%Y')}\n\n"
        f"✅ Рапорт берилган: {len(found_sites)} та\n"
    )

    text += "\n".join(f"• {s}" for s in found_sites) if found_sites else "• Ҳозирча йўқ"
    text += "\n\n❌ Рапорт берилмаган:\n"
    text += "\n".join(f"• {s}" for s in missing_sites) if missing_sites else "• Ҳаммаси берилган"

    await update.message.reply_text(text, reply_markup=MENU)

async def search_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🔎 Рапорт қидириш.\n\n"
        "Сана:\n05.09.2026\n\n"
        "Сана оралиғи:\n01.01.2026 - 05.09.2026\n\n"
        "Шантиё:\nPIRAMIT\n\n"
        "Ёки исталган сўзни ёзинг.",
        reply_markup=MENU
    )
    context.user_data["search_mode"] = True

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message or not update.message.text:
        return

    text = update.message.text.strip()

    if text == "🔎 Rapor Ara":
        await search_start(update, context)
        return

    if text == "📊 Rapor Durumu":
        await status(update, context)
        return

    if text == "📥 Excel":
        if not ADMIN_CHAT_ID:
            await update.message.reply_text("⚠️ ADMIN_CHAT_ID sozlanmagan.")
            return
        if str(update.effective_chat.id) != str(ADMIN_CHAT_ID):
            await update.message.reply_text("⛔ Бу бўлим фақат администратор учун.")
            return
        if not os.path.exists(HISTORY_FILE):
            await update.message.reply_text("⚠️ Excel файл топилмади.")
            return
        with open(HISTORY_FILE, "rb") as f:
            await update.message.reply_document(f, filename=HISTORY_FILE)
        return

    if not context.user_data.get("search_mode"):
        return

    start_date, end_date = parse_date_range(text)
    site = detect_site(text)
    keyword = None if start_date or site else text

    results = search_excel(
        start_date=start_date,
        end_date=end_date,
        site=site,
        keyword=keyword
    )

    if not results:
        await update.message.reply_text(
            "🔎 Рапорт топилмади.\n\n"
            "Мисол:\n01.01.2026 - 05.09.2026\n"
            "ёки\nPIRAMIT\n"
            "ёки\n05.09.2026",
            reply_markup=MENU
        )
        context.user_data["search_mode"] = False
        return

    results = results[:100]
    answer = f"🔎 Топилди: {len(results)} та рапорт\n\n"

    for i, r in enumerate(results, start=1):
        d = r["date"]
        d_text = d.strftime("%d.%m.%Y") if isinstance(d, date) else str(d or "")
        report_text = str(r["text"] or "").strip()
        if len(report_text) > 500:
            report_text = report_text[:500] + "..."

        answer += (
            f"━━━━━━━━━━━━━━\n"
            f"#{i}\n"
            f"📅 Сана: {d_text}\n"
            f"⏰ Вақт: {r['time'] or ''}\n"
            f"🏗 Шантиё: {r['site'] or '—'}\n"
            f"👤 Юборувчи: {r['sender'] or '—'}\n"
            f"👷 Ишчи: {r['workers'] or '—'}\n"
        )
        if report_text:
            answer += f"📝 {report_text}\n"

    await update.message.reply_text(answer, reply_markup=MENU)
    context.user_data["search_mode"] = False

async def daily_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    today = datetime.now(TZ).date()
    reports = search_excel(start_date=today, end_date=today)
    found_sites = sorted(set(r["site"] for r in reports if r["site"]))
    missing_sites = [s for s in SITES if s not in found_sites]

    text = (
        "📊 Кунлик шантиё рапорти\n\n"
        f"📅 Сана: {today.strftime('%d.%m.%Y')}\n\n"
        f"✅ Рапорт берган шантиёлар: {len(found_sites)} та\n"
    )

    if found_sites:
        text += "\n".join(f"• {s}" for s in found_sites)

    text += "\n\n❌ Рапорт бермаганлар:\n"
    text += "\n".join(f"• {s}" for s in missing_sites) if missing_sites else "• Ҳаммаси рапорт берган."

    try:
        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=text)
    except Exception as e:
        print("Daily report xato:", e)

def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN Render Environment Variables ichiga kiritilishi kerak.")

    ensure_excel()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.job_queue.run_daily(
        daily_report,
        time=datetime(2000, 1, 1, REPORT_HOUR, REPORT_MINUTE, tzinfo=TZ).timetz()
    )

    print("🤖 Bot ishlayapti...")
    app.run_polling()

if __name__ == "__main__":
    main()
