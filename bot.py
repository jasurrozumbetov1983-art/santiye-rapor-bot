import os
import re
from calendar import monthrange
from datetime import date, datetime
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

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_CHAT_ID_RAW = os.getenv("ADMIN_CHAT_ID", "").strip()
TIMEZONE = os.getenv("TIMEZONE", "Asia/Tashkent").strip()
TZ = ZoneInfo(TIMEZONE)

try:
    ADMIN_CHAT_ID = int(ADMIN_CHAT_ID_RAW) if ADMIN_CHAT_ID_RAW else None
except ValueError:
    ADMIN_CHAT_ID = None

# Yangi toza Excel fayl.
HISTORY_FILE = "raporlar_yangi.xlsx"

# Jadval vaqtlari
REMINDER_HOUR = int(os.getenv("REMINDER_HOUR", "12"))
REMINDER_MINUTE = int(os.getenv("REMINDER_MINUTE", "0"))
REPORT_HOUR = int(os.getenv("REPORT_HOUR", "17"))
REPORT_MINUTE = int(os.getenv("REPORT_MINUTE", "30"))
WEEKLY_HOUR = int(os.getenv("WEEKLY_HOUR", "18"))
WEEKLY_MINUTE = int(os.getenv("WEEKLY_MINUTE", "0"))
MONTHLY_HOUR = int(os.getenv("MONTHLY_HOUR", "18"))
MONTHLY_MINUTE = int(os.getenv("MONTHLY_MINUTE", "30"))

# DMC + ELLIPSE GARDEN bitta obyekt hisoblanadi.
SITES = [
    "DATA CENTER",
    "DMC",
    "LOT13",
    "LOT71",
    "SKP",
    "STADYUM",
    "BWC",
    "KÖKSARAY",
    "MMP",
    "MOS",
    "PİRAMİT",
    "RMC",
    "TYM",
    "YHP",
]

ALIASES = {
    "DATA CENTER": "DATA CENTER",
    "DMC": "DMC",
    "ELLIPSE GARDEN": "DMC",
    "LOT13": "LOT13",
    "LOT 13": "LOT13",
    "LOT71": "LOT71",
    "LOT 71": "LOT71",
    "SKP": "SKP",
    "STADIUM": "STADYUM",
    "STADYUM": "STADYUM",
    "PIRAMIT": "PİRAMİT",
    "PİRAMİT": "PİRAMİT",
    "PIRAMIT TOWER": "PİRAMİT",
    "PİRAMİT TOWER": "PİRAMİT",
    "BWC": "BWC",
    "KOKSARAY": "KÖKSARAY",
    "KÖKSARAY": "KÖKSARAY",
    "MMP": "MMP",
    "MPP": "MMP",
    "MOS": "MOS",
    "RMC": "RMC",
    "TYM": "TYM",
    "YHP": "YHP",
}

RESPONSIBLES = {
    "BWC": "@YSF1434",
    "DMC": "@uzyusufmutlu",
    "MMP": "@Orhan_Ceylan",
    "MOS": "@Orhan_Ceylan",
    "RMC": "@Orhan_Ceylan",
    "TYM": "@Orhan_Ceylan",
    "YHP": "@Orhan_Ceylan",
}

MENU = ReplyKeyboardMarkup(
    [
        ["🔎 Rapor Ara"],
        ["📊 Rapor Durumu", "📥 Excel"],
    ],
    resize_keyboard=True,
)

REPORT_NOTE = (
    "📢 Not: Şantiyenin dili verdiği rapordur; raporu olmayan iş tamamlanmış sayılmaz. "
    "Lütfen günlük raporlarınızı zamanında iletiniz."
)

REPORT_NOTE_2 = (
    "📢 Not: Yapılan işin raporunu vermek, saha yönetiminin en kritik adımıdır. "
    "Bunca çabaya rağmen rapor iletmeyen şantiyeler, lütfen rapor düzenine özen göstersin. 🙏\n"
    "Unutmayın: İşi yapmak cesarettir, raporlamak ise disiplindir. ⚠️"
)


def normalize(text):
    if text is None:
        return ""
    text = str(text).upper()
    for old, new in {
        "İ": "I",
        "Ş": "S",
        "Ğ": "G",
        "Ü": "U",
        "Ö": "O",
        "Ç": "C",
        "Ə": "E",
    }.items():
        text = text.replace(old, new)
    return text.strip()


def detect_site(text):
    normalized = normalize(text)
    for alias, site in sorted(
        ALIASES.items(),
        key=lambda item: len(normalize(item[0])),
        reverse=True,
    ):
        if normalize(alias) in normalized:
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
    for fmt in (
        "%d.%m.%Y",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%d.%m.%y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_date_range(text):
    text = text.strip()

    match = re.search(
        r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s*[-–—]\s*"
        r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})",
        text,
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


def extract_worker_count(text):
    patterns = [
        r"(?:işçi|isci|işçiler|isciler|работник|рабочих)\s*[:=-]?\s*(\d+)",
        r"(\d+)\s*(?:işçi|isci|işçiler|isciler)",
    ]
    normalized = normalize(text)
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if match:
            return int(match.group(1))
    return ""


def ensure_excel():
    if os.path.exists(HISTORY_FILE):
        try:
            wb = load_workbook(HISTORY_FILE)
            ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
            headers = [cell.value for cell in ws[1]]
            needed = [
                "Tarih",
                "Saat",
                "Şantiye",
                "Gönderen",
                "İşçi Sayısı",
                "Rapor Metni",
                "Mesaj ID",
                "Telegram Chat ID",
                "Telegram User ID",
                "Username",
            ]
            changed = False
            for header in needed:
                if header not in headers:
                    ws.cell(row=1, column=ws.max_column + 1, value=header)
                    changed = True
            if changed:
                wb.save(HISTORY_FILE)
            wb.close()
            return
        except Exception as exc:
            print(f"⚠️ Excel tekshirishda xato: {exc}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Raporlar"
    ws.append(
        [
            "Tarih",
            "Saat",
            "Şantiye",
            "Gönderen",
            "İşçi Sayısı",
            "Rapor Metni",
            "Mesaj ID",
            "Telegram Chat ID",
            "Telegram User ID",
            "Username",
        ]
    )
    wb.save(HISTORY_FILE)
    wb.close()


def header_indices(headers):
    return {
        str(value).strip(): index
        for index, value in enumerate(headers)
        if value is not None
    }


def append_report_to_excel(update, site, report_text):
    ensure_excel()

    try:
        wb = load_workbook(HISTORY_FILE)
        ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active

        headers = [cell.value for cell in ws[1]]
        needed = [
            "Tarih",
            "Saat",
            "Şantiye",
            "Gönderen",
            "İşçi Sayısı",
            "Rapor Metni",
            "Mesaj ID",
            "Telegram Chat ID",
            "Telegram User ID",
            "Username",
        ]

        for header in needed:
            if header not in headers:
                ws.cell(row=1, column=ws.max_column + 1, value=header)
                headers.append(header)

        indices = header_indices(headers)
        now = datetime.now(TZ)
        user = update.effective_user
        chat = update.effective_chat

        values = {
            "Tarih": now.date(),
            "Saat": now.strftime("%H:%M:%S"),
            "Şantiye": site,
            "Gönderen": user.full_name if user else "Noma'lum",
            "İşçi Sayısı": extract_worker_count(report_text),
            "Rapor Metni": report_text,
            "Mesaj ID": update.message.message_id if update.message else "",
            "Telegram Chat ID": chat.id if chat else "",
            "Telegram User ID": user.id if user else "",
            "Username": f"@{user.username}" if user and user.username else "",
        }

        row = ws.max_row + 1
        for key, value in values.items():
            ws.cell(row=row, column=indices[key] + 1, value=value)

        wb.save(HISTORY_FILE)
        wb.close()
        return True

    except Exception as exc:
        print(f"❌ Rapor Excel'e kaydedilirken xato: {exc}")
        return False


def search_excel(start_date=None, end_date=None, site=None, keyword=None):
    ensure_excel()

    try:
        wb = load_workbook(HISTORY_FILE, data_only=True)
        ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            wb.close()
            return []

        headers = list(rows[0])
        indices = header_indices(headers)

        def value(row, name):
            index = indices.get(name)
            if index is None or index >= len(row):
                return ""
            return row[index]

        results = []

        for row in rows[1:]:
            row_date = parse_date(value(row, "Tarih"))

            if start_date and (row_date is None or row_date < start_date):
                continue
            if end_date and (row_date is None or row_date > end_date):
                continue

            row_site = str(value(row, "Şantiye") or "").strip()
            detected = detect_site(row_site)
            if detected:
                row_site = detected

            if site and normalize(site) != normalize(row_site):
                continue

            all_text = " ".join(str(item or "") for item in row)
            if keyword and normalize(keyword) not in normalize(all_text):
                continue

            results.append(
                {
                    "date": row_date,
                    "time": value(row, "Saat"),
                    "site": row_site,
                    "sender": value(row, "Gönderen"),
                    "workers": value(row, "İşçi Sayısı"),
                    "text": value(row, "Rapor Metni"),
                }
            )

        wb.close()
        results.sort(
            key=lambda item: (
                item["date"] or date.min,
                str(item["time"] or ""),
            ),
            reverse=True,
        )
        return results

    except Exception as exc:
        print(f"❌ Excel qidirishda xato: {exc}")
        return []


def get_missing_sites(day=None):
    day = day or datetime.now(TZ).date()
    reports = search_excel(start_date=day, end_date=day)
    found = {normalize(report["site"]) for report in reports if report["site"]}
    return [
        site for site in SITES
        if normalize(site) not in found
    ]


def build_status_text(day=None):
    day = day or datetime.now(TZ).date()
    reports = search_excel(start_date=day, end_date=day)

    found_sites = []
    for report in reports:
        site = report["site"]
        if site and normalize(site) not in {normalize(x) for x in found_sites}:
            found_sites.append(site)

    missing_sites = get_missing_sites(day)

    text = (
        "📊 Şantiye Rapor Durumu\n\n"
        f"📅 Tarih: {day.strftime('%d.%m.%Y')}\n\n"
        f"✅ Rapor gönderilen: {len(found_sites)}\n"
    )

    text += (
        "\n".join(f"• {site}" for site in found_sites)
        if found_sites
        else "• Henüz yok"
    )

    text += "\n\n❌ Rapor gönderilmeyen:\n"
    text += (
        "\n".join(f"• {site}" for site in missing_sites)
        if missing_sites
        else "• Hepsi gönderildi"
    )

    return text


def build_missing_text(day=None):
    day = day or datetime.now(TZ).date()
    missing = get_missing_sites(day)

    if not missing:
        return (
            f"✅ {day.strftime('%d.%m.%Y')} kuni barcha "
            "şantiyeler rapor yubordi."
        )

    lines = [
        f"⚠️ {day.strftime('%d.%m.%Y')} — raporu olmayan şantiyeler:",
        "",
    ]

    for site in missing:
        responsible = RESPONSIBLES.get(site, "")
        suffix = f" — {responsible}" if responsible else ""
        lines.append(f"• {site}{suffix}")

    lines.extend(["", REPORT_NOTE, "", REPORT_NOTE_2])
    return "\n".join(lines)


def build_daily_report(day=None):
    day = day or datetime.now(TZ).date()
    reports = search_excel(start_date=day, end_date=day)

    by_site = {}
    for report in reports:
        site = report["site"]
        if site:
            by_site[normalize(site)] = report

    found_count = sum(
        1 for site in SITES if normalize(site) in by_site
    )
    missing_count = len(SITES) - found_count

    lines = [
        "📋 KUNLIK RAPOR NAZORATI",
        "",
        f"📅 Sana: {day.strftime('%d.%m.%Y')}",
        f"📊 Jami obyektlar: {len(SITES)}",
        f"✅ Rapor bor: {found_count}",
        f"❌ Rapor yo'q: {missing_count}",
        "",
    ]

    for site in SITES:
        report = by_site.get(normalize(site))
        if report:
            workers = report["workers"] if report["workers"] != "" else "—"
            sender = report["sender"] or "—"
            lines.append(
                f"✅ {site} — {workers} ishchi — {sender}"
            )
        else:
            responsible = RESPONSIBLES.get(site, "")
            suffix = f" ({responsible})" if responsible else ""
            lines.append(f"❌ {site}{suffix}")

    return "\n".join(lines)


def build_excel_bytes():
    ensure_excel()
    with open(HISTORY_FILE, "rb") as file:
        return file.read()


async def is_group_admin(update, context):
    user = update.effective_user
    chat = update.effective_chat

    if not user or not chat:
        return False

    # Excel tugmasi faqat guruh administratorlari uchun.
    if chat.type not in ("group", "supergroup"):
        return False

    try:
        member = await context.bot.get_chat_member(
            chat_id=chat.id,
            user_id=user.id,
        )

        return str(member.status).lower() in (
            "administrator",
            "creator",
            "owner",
        )

    except Exception as exc:
        print(f"❌ Admin tekshirish xatosi: {exc}")
        return False


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["search_mode"] = False

    await update.message.reply_text(
        "👷 Şantiye Rapor Kontrol Botu\n\n"
        "Rapor aramak için 🔎 Rapor Ara butonuna basın.\n\n"
        "Örnek:\n"
        "01.01.2026 - 05.09.2026\n\n"
        "Veya:\n"
        "PIRAMIT\n\n"
        "Veya:\n"
        "05.09.2026 PIRAMIT",
        reply_markup=MENU,
    )


async def chatid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        await update.message.reply_text(
            f"Chat ID: {update.effective_chat.id}"
        )


async def search_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["search_mode"] = True

    await update.message.reply_text(
        "🔎 Rapor arama.\n\n"
        "Tarih:\n05.09.2026\n\n"
        "Tarih aralığı:\n01.01.2026 - 05.09.2026\n\n"
        "Şantiye:\nPIRAMIT\n\n"
        "Veya istediğiniz kelimeyi yazın.",
        reply_markup=MENU,
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        build_status_text(),
        reply_markup=MENU,
    )


async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await is_group_admin(update, context):
        await update.message.reply_text(
            "⛔ Bu bölüm sadece grup yöneticileri için kullanılabilir.",
            reply_markup=MENU,
        )
        return

    try:
        data = build_excel_bytes()

        await update.message.reply_document(
            document=data,
            filename=HISTORY_FILE,
            caption="📥 Güncel rapor Excel dosyası",
        )

    except Exception as exc:
        print(f"❌ Excel yuborishda xato: {exc}")
        await update.message.reply_text(
            "⚠️ Excel dosyası gönderilemedi.",
            reply_markup=MENU,
        )


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
        await send_excel(update, context)
        return

    # Qidiruv rejimi
    if context.user_data.get("search_mode"):
        start_date, end_date = parse_date_range(text)
        site = detect_site(text)
        keyword = None if start_date or site else text

        results = search_excel(
            start_date=start_date,
            end_date=end_date,
            site=site,
            keyword=keyword,
        )

        context.user_data["search_mode"] = False

        if not results:
            await update.message.reply_text(
                "🔎 Rapor bulunamadı.\n\n"
                "Misol:\n"
                "01.01.2026 - 05.09.2026\n"
                "veya\n"
                "PIRAMIT\n"
                "veya\n"
                "05.09.2026",
                reply_markup=MENU,
            )
            return

        results = results[:100]
        answer = f"🔎 Bulunan rapor sayısı: {len(results)}\n\n"

        for i, report in enumerate(results, start=1):
            d = report["date"]
            d_text = (
                d.strftime("%d.%m.%Y")
                if isinstance(d, date)
                else str(d or "")
            )

            report_text = str(report["text"] or "").strip()
            if len(report_text) > 500:
                report_text = report_text[:500] + "..."

            answer += (
                "━━━━━━━━━━━━━━\n"
                f"#{i}\n"
                f"📅 Tarih: {d_text}\n"
                f"⏰ Saat: {report['time'] or ''}\n"
                f"🏗 Şantiye: {report['site'] or '—'}\n"
                f"👤 Gönderen: {report['sender'] or '—'}\n"
                f"👷 İşçi: {report['workers'] or '—'}\n"
            )

            if report_text:
                answer += f"📝 {report_text}\n"

        await update.message.reply_text(
            answer,
            reply_markup=MENU,
        )
        return

    # Oddiy xabar: obyekt nomi bo'lsa rapor sifatida saqlanadi.
    site = detect_site(text)

    if site:
        saved = append_report_to_excel(
            update,
            site,
            text,
        )

        if saved:
            sender = (
                update.effective_user.full_name
                if update.effective_user
                else "Noma'lum"
            )

            await update.message.reply_text(
                f"✅ Rapor alındı.\n"
                f"🏗 {site}\n"
                f"👤 {sender}",
                reply_markup=MENU,
            )
        else:
            await update.message.reply_text(
                "⚠️ Rapor alındı fakat Excel'e kaydedilemedi.",
                reply_markup=MENU,
            )


async def daily_reminder(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        print("⚠️ ADMIN_CHAT_ID yo'q; 12:00 reminder yuborilmadi.")
        return

    try:
        text = build_missing_text()

        await context.bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=f"⏰ 12:00 RAPOR ESLATMASI\n\n{text}",
        )

        print("✅ 12:00 reminder yuborildi.")

    except Exception as exc:
        print(f"❌ 12:00 reminder xatosi: {exc}")


async def daily_report_1730(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    try:
        await context.bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=build_daily_report(),
        )

        print("✅ 17:30 daily report yuborildi.")

    except Exception as exc:
        print(f"❌ 17:30 report xatosi: {exc}")


async def weekly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    try:
        data = build_excel_bytes()

        await context.bot.send_document(
            chat_id=ADMIN_CHAT_ID,
            document=data,
            filename=HISTORY_FILE,
            caption="📊 Haftalik Excel — yakshanba 18:00",
        )

        print("✅ Weekly Excel yuborildi.")

    except Exception as exc:
        print(f"❌ Weekly Excel xatosi: {exc}")


async def monthly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    now = datetime.now(TZ)
    last_day = monthrange(now.year, now.month)[1]

    if now.day != last_day:
        return

    try:
        data = build_excel_bytes()

        await context.bot.send_document(
            chat_id=ADMIN_CHAT_ID,
            document=data,
            filename=HISTORY_FILE,
            caption="📊 Oylik Excel — oyning oxirgi kuni 18:30",
        )

        print("✅ Monthly Excel yuborildi.")

    except Exception as exc:
        print(f"❌ Monthly Excel xatosi: {exc}")


async def error_handler(update, context):
    print(f"❌ Telegram xatosi: {context.error}")


def main():
    if not BOT_TOKEN:
        raise RuntimeError(
            "BOT_TOKEN environment variable topilmadi."
        )

    ensure_excel()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("chatid", chatid))

    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            handle_message,
        )
    )

    app.add_error_handler(error_handler)

    if app.job_queue is None:
        raise RuntimeError(
            "JobQueue mavjud emas. "
            "requirements.txt ichida python-telegram-bot[job-queue] "
            "o'rnatilgan bo'lishi kerak."
        )

    app.job_queue.run_daily(
        daily_reminder,
        time=datetime(
            2000,
            1,
            1,
            REMINDER_HOUR,
            REMINDER_MINUTE,
            tzinfo=TZ,
        ).timetz(),
    )

    app.job_queue.run_daily(
        daily_report_1730,
        time=datetime(
            2000,
            1,
            1,
            REPORT_HOUR,
            REPORT_MINUTE,
            tzinfo=TZ,
        ).timetz(),
    )

    # APScheduler: Monday=0 ... Sunday=6
    app.job_queue.run_daily(
        weekly_excel_report,
        time=datetime(
            2000,
            1,
            1,
            WEEKLY_HOUR,
            WEEKLY_MINUTE,
            tzinfo=TZ,
        ).timetz(),
        days=(6,),
    )

    # Funksiya har kuni tekshiradi, faqat oyning oxirgi kunida yuboradi.
    app.job_queue.run_daily(
        monthly_excel_report,
        time=datetime(
            2000,
            1,
            1,
            MONTHLY_HOUR,
            MONTHLY_MINUTE,
            tzinfo=TZ,
        ).timetz(),
    )

    print("🤖 Bot ishlayapti!")
    print(
        f"⏰ Reminder: "
        f"{REMINDER_HOUR:02d}:{REMINDER_MINUTE:02d}"
    )
    print(
        f"📋 Daily report: "
        f"{REPORT_HOUR:02d}:{REPORT_MINUTE:02d}"
    )
    print(
        f"📥 Weekly Excel: "
        f"Sunday {WEEKLY_HOUR:02d}:{WEEKLY_MINUTE:02d}"
    )
    print(
        f"📥 Monthly Excel: "
        f"last day {MONTHLY_HOUR:02d}:{MONTHLY_MINUTE:02d}"
    )
    print(f"👑 ADMIN_CHAT_ID: {ADMIN_CHAT_ID}")

    # Faqat bitta Render instance ishlashi kerak.
    app.run_polling(
        drop_pending_updates=True,
        allowed_updates=Update.ALL_TYPES,
    )


if __name__ == "__main__":
    main()
