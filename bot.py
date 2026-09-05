import os
import re
from io import BytesIO
from calendar import monthrange
from datetime import datetime, date, timedelta
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
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")  # Группа, куда 17:30 да отправляется общий отчёт
TZ = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))

HISTORY_FILE = "raporlar_tarih_2026_2.xlsx"
LIVE_FILE = "raporlar.xlsx"

HISTORY_FILES = [
    "raporlar_tarih_2026_2.xlsx",
    "raporlar_tarih_2026_2(1).xlsx",
    "raporlar_2026_2.xlsx",
    "raporlar.xlsx",
]

# 12:00 — напоминание, 17:30 — контрольный отчёт
REMINDER_HOUR = int(os.getenv("REMINDER_HOUR", "12"))
REMINDER_MINUTE = int(os.getenv("REMINDER_MINUTE", "0"))
REPORT_HOUR = int(os.getenv("REPORT_HOUR", "17"))
REPORT_MINUTE = int(os.getenv("REPORT_MINUTE", "30"))

# Ҳафталик Excel: якшанба 18:00
WEEKLY_HOUR = int(os.getenv("WEEKLY_HOUR", "18"))
WEEKLY_MINUTE = int(os.getenv("WEEKLY_MINUTE", "0"))

# Ойлик Excel: ҳар ойнинг охирги куни 18:30
MONTHLY_HOUR = int(os.getenv("MONTHLY_HOUR", "18"))
MONTHLY_MINUTE = int(os.getenv("MONTHLY_MINUTE", "30"))

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
    filename = HISTORY_FILE
    if os.path.exists(filename):
        # Старый Excel мог быть создан без Telegram ID.
        # Добавляем недостающие колонки, не удаляя старые данные.
        try:
            wb = load_workbook(filename)
            ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
            headers = [c.value for c in ws[1]]
            needed = ["Telegram Chat ID", "Telegram User ID", "Username"]
            changed = False
            for h in needed:
                if h not in headers:
                    ws.cell(row=1, column=ws.max_column + 1, value=h)
                    changed = True
            if changed:
                wb.save(filename)
            wb.close()
        except Exception as e:
            print("Excel ustunlarini tekshirishda xato:", e)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Raporlar"
    ws.append([
        "Tarih", "Saat", "Şantiye", "Gönderen",
        "İşçi Sayısı", "Rapor Metni", "Mesaj ID",
        "Telegram Chat ID", "Telegram User ID", "Username"
    ])
    wb.save(filename)
    wb.close()


def get_history_file():
    for filename in HISTORY_FILES:
        if os.path.exists(filename):
            return filename
    return None


def search_excel(start_date=None, end_date=None, site=None, keyword=None):
    filename = get_history_file()
    if not filename:
        print("Excel topilmadi. Qidirilgan fayllar:", HISTORY_FILES)
        return []

    try:
        wb = load_workbook(filename, read_only=True, data_only=True)
    except Exception as e:
        print("Excel ochishda xato:", e)
        return []

    results = []
    sheets = [wb["Raporlar"]] if "Raporlar" in wb.sheetnames else wb.worksheets

    for ws in sheets:
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
        text_col = find_column(headers, [
            "Rapor Metni", "Rapor", "Hisobot", "Report", "Сообщение",
            "Toliq_raport", "Toliq raport"
        ])
        message_id_col = find_column(headers, ["Mesaj ID", "Message ID", "ID"])
        chat_id_col = find_column(headers, ["Telegram Chat ID", "Chat ID"])
        user_id_col = find_column(headers, ["Telegram User ID", "User ID"])
        username_col = find_column(headers, ["Username", "Telegram Username"])

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

            if site and normalize(site) != normalize(row_site):
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
                "chat_id": row[chat_id_col] if chat_id_col is not None and chat_id_col < len(row) else "",
                "user_id": row[user_id_col] if user_id_col is not None and user_id_col < len(row) else "",
                "username": row[username_col] if username_col is not None and username_col < len(row) else "",
            })

    wb.close()
    return results


def append_report_to_excel(update: Update, site: str, report_text: str):
    """
    Сохраняет новый рапорт и Telegram ID отправителя.
    Это нужно для персонального напоминания в 12:00.
    """
    filename = HISTORY_FILE
    ensure_excel()

    try:
        wb = load_workbook(filename)
        ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
        headers = [c.value for c in ws[1]]

        # Добавляем недостающие колонки
        for h in ["Telegram Chat ID", "Telegram User ID", "Username"]:
            if h not in headers:
                ws.cell(row=1, column=ws.max_column + 1, value=h)
                headers = [c.value for c in ws[1]]

        def col(name):
            return headers.index(name) + 1

        now = datetime.now(TZ)
        user = update.effective_user
        chat = update.effective_chat

        row = ws.max_row + 1
        values = {
            "Tarih": now.date(),
            "Saat": now.strftime("%H:%M:%S"),
            "Şantiye": site,
            "Gönderen": (
                user.full_name if user else "Noma'lum"
            ),
            "Rapor Metni": report_text,
            "Mesaj ID": update.message.message_id if update.message else "",
            "Telegram Chat ID": chat.id if chat else "",
            "Telegram User ID": user.id if user else "",
            "Username": (
                f"@{user.username}" if user and user.username else ""
            ),
        }

        for key, value in values.items():
            if key in headers:
                ws.cell(row=row, column=col(key), value=value)

        wb.save(filename)
        wb.close()
        return True
    except Exception as e:
        print("Rapor Excel'e kaydedilirken hata:", e)
        return False


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
        "👷 Şantiye Rapor Kontrol Botu\n\n"
        "Rapor aramak için 🔎 Rapor Ara butonuna basın.\n\n"
        "Örnek:\n01.01.2026 - 05.09.2026\n\n"
        "Veya:\nPIRAMIT\n\n"
        "Veya:\n05.09.2026 PIRAMIT",
        reply_markup=MENU
    )


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.now(TZ).date()
    reports = search_excel(start_date=today, end_date=today)

    found_sites = sorted(set(r["site"] for r in reports if r["site"]))
    missing_sites = [s for s in SITES if normalize(s) not in {normalize(x) for x in found_sites}]

    text = (
        "📊 Şantiye Rapor Durumu\n\n"
        f"📅 Tarih: {today.strftime('%d.%m.%Y')}\n\n"
        f"✅ Rapor gönderilen: {len(found_sites)}\n"
    )

    text += "\n".join(f"• {s}" for s in found_sites) if found_sites else "• Henüz yok"
    text += "\n\n❌ Rapor gönderilmeyen:\n"
    text += "\n".join(f"• {s}" for s in missing_sites) if missing_sites else "• Hepsi gönderildi"

    await update.message.reply_text(text, reply_markup=MENU)


async def search_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🔎 Rapor arama.\n\n"
        "Tarih:\n05.09.2026\n\n"
        "Tarih aralığı:\n01.01.2026 - 05.09.2026\n\n"
        "Şantiye:\nPIRAMIT\n\n"
        "Veya istediğiniz kelimeyi yazın.",
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
            await update.message.reply_text("⚠️ ADMIN_CHAT_ID ayarlanmamış.")
            return
        if str(update.effective_chat.id) != str(ADMIN_CHAT_ID):
            await update.message.reply_text("⛔ Bu bölüm sadece yönetici içindir.")
            return
        filename = get_history_file()
        if not filename:
            await update.message.reply_text("⚠️ Excel dosyası bulunamadı.")
            return
        with open(filename, "rb") as f:
            await update.message.reply_document(f, filename=filename)
        return

    # Қидирув режимидаги матнни рапорт сифатида сақламаймиз.
    if context.user_data.get("search_mode"):
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
                "🔎 Rapor bulunamadı.\n\n"
                "Мисол:\n01.01.2026 - 05.09.2026\n"
                "ёки\nPIRAMIT\n"
                "ёки\n05.09.2026",
                reply_markup=MENU
            )
            context.user_data["search_mode"] = False
            return

        results = results[:100]
        answer = f"🔎 Bulunan rapor sayısı: {len(results)}\n\n"

        for i, r in enumerate(results, start=1):
            d = r["date"]
            d_text = d.strftime("%d.%m.%Y") if isinstance(d, date) else str(d or "")
            report_text = str(r["text"] or "").strip()
            if len(report_text) > 500:
                report_text = report_text[:500] + "..."

            answer += (
                f"━━━━━━━━━━━━━━\n"
                f"#{i}\n"
                f"📅 Tarih: {d_text}\n"
                f"⏰ Saat: {r['time'] or ''}\n"
                f"🏗 Şantiye: {r['site'] or '—'}\n"
                f"👤 Gönderen: {r['sender'] or '—'}\n"
                f"👷 İşçi: {r['workers'] or '—'}\n"
            )
            if report_text:
                answer += f"📝 {report_text}\n"

        await update.message.reply_text(answer, reply_markup=MENU)
        context.user_data["search_mode"] = False
        return

    # Агар оддий хабарда объект номи бўлса, янги рапорт деб қабул қиламиз.
    # Шу билан юборувчининг Telegram Chat ID сақланади ва 12:00 да шахсий
    # эслатма юбориш мумкин бўлади.
    site = detect_site(text)
    if site:
        saved = append_report_to_excel(update, site, text)
        if saved:
            await update.message.reply_text(
                f"✅ Rapor alındı.\n🏗 {site}\n"
                f"👤 {update.effective_user.full_name if update.effective_user else '—'}"
            )
        return


def build_site_sender_map():
    """
    Охирги маълумотлардан объект -> юборувчи Telegram Chat ID харитасини тузади.
    Эски Excelда Chat ID бўлмаса, у юборувчига шахсий эслатма юбориб бўлмайди.
    Янги рапорт юборилганидан кейин ID автоматик сақланади.
    """
    today = datetime.now(TZ).date()
    reports = search_excel()

    mapping = {}
    for r in reports:
        site = str(r.get("site") or "").strip()
        chat_id = r.get("chat_id")
        sender = str(r.get("sender") or "").strip()
        username = str(r.get("username") or "").strip()

        if not site or not chat_id:
            continue

        # Энг охирги рапорт юборган шахс объектга масъул сифатида олинади.
        mapping[site] = {
            "chat_id": str(chat_id),
            "sender": sender or username or str(chat_id),
            "username": username,
            "last_date": r.get("date"),
        }

    return mapping


async def reminder_1200(context: ContextTypes.DEFAULT_TYPE):
    """
    Ҳар куни 12:00 да: объектни одатда ким юборса, ўша одамга шахсий эслатма.
    Бир одам бир нечта объект юборса, битта хабарда ҳаммасини кўрсатади.
    """
    mapping = build_site_sender_map()

    by_chat = {}
    for site in SITES:
        info = mapping.get(site)
        if not info:
            continue
        chat_id = info["chat_id"]
        by_chat.setdefault(chat_id, []).append(site)

    for chat_id, sites in by_chat.items():
        sites_text = "\n".join(f"• {s}" for s in sites)
        text = (
            "🔔 Hatırlatma!\n\n"
            "Bugünkü raporu saat 17:30'a kadar göndermeyi unutmayın.\n\n"
            "Size bağlı şantiyeler:\n"
            f"{sites_text}"
        )
        try:
            await context.bot.send_message(chat_id=int(chat_id), text=text)
        except Exception as e:
            print(f"12:00 hatırlatması gönderilirken hata ({chat_id}):", e)


async def daily_report_1730(context: ContextTypes.DEFAULT_TYPE):
    """
    Ҳар куни 17:30 да ADMIN_CHAT_ID группага:
    ҳар бир объект — ким юборди ёки ким юбормагани.
    """
    if not ADMIN_CHAT_ID:
        print("ADMIN_CHAT_ID ayarlanmamış.")
        return

    today = datetime.now(TZ).date()
    reports = search_excel(start_date=today, end_date=today)

    # Объект бўйича бугунги рапортлар
    site_reports = {}
    for r in reports:
        site = str(r.get("site") or "").strip()
        if site:
            site_reports.setdefault(normalize(site), []).append(r)

    mapping = build_site_sender_map()

    lines = [
        "📊 GÜNLÜK RAPOR KONTROLÜ",
        "",
        f"📅 Tarih: {today.strftime('%d.%m.%Y')}",
        "⏰ Saat: 17:30",
        "",
    ]

    sent_count = 0
    missing_count = 0

    for site in SITES:
        rows = site_reports.get(normalize(site), [])

        if rows:
            sent_count += 1
            # Бир объектга бир нечта рапорт келса, барчасини кўрсатамиз.
            senders = []
            for r in rows:
                name = str(r.get("sender") or "").strip()
                username = str(r.get("username") or "").strip()
                display = name or username or "Bilinmiyor"
                if username and username not in display:
                    display += f" ({username})"
                if display not in senders:
                    senders.append(display)

            lines.append(f"✅ {site} — gönderildi")
            for sender in senders:
                lines.append(f"   👤 {sender}")
        else:
            missing_count += 1
            responsible = mapping.get(site)
            if responsible:
                name = responsible["sender"]
                lines.append(f"❌ {site} — gönderilmedi")
                lines.append(f"   👤 Sorumlu: {name}")
            else:
                lines.append(f"❌ {site} — gönderilmedi")
                lines.append("   👤 Sorumlunun Telegram ID'si henüz kayıtlı değil")

    lines.extend([
        "",
        f"📌 Toplam şantiye: {len(SITES)}",
        f"✅ Gönderildi: {sent_count}",
        f"❌ Gönderilmedi: {missing_count}",
    ])

    try:
        await context.bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text="\n".join(lines)
        )
    except Exception as e:
        print("17:30 daily report xato:", e)


def make_weekly_excel(target_date=None):
    """Ўтган душанба-якшанба даври учун: ҳар бир объект ва юборилмаган кунлар."""
    if target_date is None:
        target_date = datetime.now(TZ).date()

    # Ҳафта душанба кунидан бошланади.
    monday = target_date - timedelta(days=target_date.weekday())
    sunday = monday + timedelta(days=6)

    reports = search_excel(start_date=monday, end_date=sunday)

    # Ҳар объект ва сана бўйича ким юборганини сақлаймиз.
    sent = {}
    for r in reports:
        site = str(r.get("site") or "").strip()
        d = r.get("date")
        if site and isinstance(d, date):
            key = (normalize(site), d)
            sent.setdefault(key, []).append(str(r.get("sender") or "Bilinmiyor"))

    wb = Workbook()
    summary = wb.active
    summary.title = "Haftalik xulosa"
    summary.append(["Obyekt", "Masul", "Yuborilgan kunlar", "Yuborilmagan kunlar", "Holat"])

    details = {}
    mapping = build_site_sender_map()

    for site in SITES:
        missing = []
        sent_days = []
        for i in range(7):
            d = monday + timedelta(days=i)
            rows = sent.get((normalize(site), d), [])
            if rows:
                sent_days.append(d.strftime("%d.%m"))
            else:
                missing.append(d.strftime("%d.%m"))

        responsible = mapping.get(site, {}).get("sender", "Telegram ID hali saqlanmagan")
        status_text = "To'liq" if not missing else "Kamchilik bor"

        summary.append([
            site,
            responsible,
            ", ".join(sent_days) if sent_days else "—",
            ", ".join(missing) if missing else "—",
            status_text
        ])

    ws = wb.create_sheet("Kunlik tafsilot")
    ws.append(["Sana", "Obyekt", "Holat", "Yuboruvchi"])
    for i in range(7):
        d = monday + timedelta(days=i)
        for site in SITES:
            rows = sent.get((normalize(site), d), [])
            if rows:
                for sender in rows:
                    ws.append([d, site, "Yuborildi", sender])
            else:
                responsible = mapping.get(site, {}).get("sender", "—")
                ws.append([d, site, "Yuborilmadi", responsible])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream, f"Haftalik_Raport_{monday.strftime('%d.%m.%Y')}_{sunday.strftime('%d.%m.%Y')}.xlsx"


def make_monthly_excel(target_date=None):
    """Ойлик Excel: ҳар бир объект, ҳар бир кун, юборилган/юборилмаган ва масъул."""
    if target_date is None:
        target_date = datetime.now(TZ).date()

    first_day = target_date.replace(day=1)
    last_day = target_date.replace(day=monthrange(target_date.year, target_date.month)[1])

    reports = search_excel(start_date=first_day, end_date=last_day)

    sent = {}
    for r in reports:
        site = str(r.get("site") or "").strip()
        d = r.get("date")
        if site and isinstance(d, date):
            sent.setdefault((normalize(site), d), []).append(r)

    mapping = build_site_sender_map()

    wb = Workbook()
    summary = wb.active
    summary.title = "Oylik xulosa"
    summary.append([
        "Obyekt", "Masul", "Jami kun", "Yuborilgan kun",
        "Yuborilmagan kun", "Foiz", "Holat"
    ])

    for site in SITES:
        total = (last_day - first_day).days + 1
        sent_count = 0
        missing_days = []

        for i in range(total):
            d = first_day + timedelta(days=i)
            if sent.get((normalize(site), d)):
                sent_count += 1
            else:
                missing_days.append(d.strftime("%d.%m"))

        percent = round(sent_count * 100 / total, 1) if total else 0
        responsible = mapping.get(site, {}).get("sender", "Telegram ID hali saqlanmagan")
        status_text = "To'liq" if not missing_days else "Kamchilik bor"

        summary.append([
            site,
            responsible,
            total,
            sent_count,
            ", ".join(missing_days) if missing_days else "—",
            f"{percent}%",
            status_text
        ])

    # Ҳар бир объект учун алоҳида лист
    for site in SITES:
        safe_title = re.sub(r'[:\\/?*\[\]]', '_', site)[:31]
        ws = wb.create_sheet(safe_title)
        ws.append(["Sana", "Obyekt", "Holat", "Yuboruvchi", "Rapor matni"])

        total = (last_day - first_day).days + 1
        for i in range(total):
            d = first_day + timedelta(days=i)
            rows = sent.get((normalize(site), d), [])

            if rows:
                for r in rows:
                    ws.append([
                        d,
                        site,
                        "Yuborildi",
                        r.get("sender") or "Noma'lum",
                        r.get("text") or ""
                    ])
            else:
                responsible = mapping.get(site, {}).get("sender", "—")
                ws.append([d, site, "Yuborilmadi", responsible, ""])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream, f"Oylik_Raport_{first_day.strftime('%m.%Y')}.xlsx"


async def weekly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    try:
        stream, filename = make_weekly_excel(datetime.now(TZ).date())
        await context.bot.send_document(
            chat_id=ADMIN_CHAT_ID,
            document=stream,
            filename=filename,
            caption="📥 Haftalık rapor Excel\nKim hangi gün gönderdi/göndermedi."
        )
    except Exception as e:
        print("Haftalık Excel gönderilirken hata:", e)


async def monthly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

    # Job ҳар куни ишлайди; фақат ойнинг охирги кунида юборилади.
    today = datetime.now(TZ).date()
    if today.day != monthrange(today.year, today.month)[1]:
        return

    try:
        stream, filename = make_monthly_excel(today)
        await context.bot.send_document(
            chat_id=ADMIN_CHAT_ID,
            document=stream,
            filename=filename,
            caption="📥 Aylık rapor Excel\nHer şantiye için ayın tamamına ait rapor."
        )
    except Exception as e:
        print("Aylık Excel gönderilirken hata:", e)



def main():
    if not BOT_TOKEN:
        raise RuntimeError(
            "BOT_TOKEN Render Environment Variables ichiga kiritilishi kerak."
        )

    ensure_excel()

    try:
        app = Application.builder().token(BOT_TOKEN).build()
    except Exception as e:
        raise RuntimeError(
            "Bot ishga tushmadi. python-telegram-bot[job-queue] o'rnatilganini tekshiring."
        ) from e

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # 12:00 — шахсий эслатма
    app.job_queue.run_daily(
        reminder_1200,
        time=datetime(
            2000, 1, 1,
            REMINDER_HOUR, REMINDER_MINUTE,
            tzinfo=TZ
        ).timetz()
    )

    # 17:30 — группага умумий назорат
    app.job_queue.run_daily(
        daily_report_1730,
        time=datetime(
            2000, 1, 1,
            REPORT_HOUR, REPORT_MINUTE,
            tzinfo=TZ
        ).timetz()
    )

    # Якшанба 18:00 — ҳафталик Excel
    app.job_queue.run_daily(
        weekly_excel_report,
        time=datetime(
            2000, 1, 1,
            WEEKLY_HOUR, WEEKLY_MINUTE,
            tzinfo=TZ
        ).timetz(),
        days=(6,)
    )

    # Ҳар куни 18:30 да текширади; фақат ойнинг охирги куни Excel юборади.
    app.job_queue.run_daily(
        monthly_excel_report,
        time=datetime(
            2000, 1, 1,
            MONTHLY_HOUR, MONTHLY_MINUTE,
            tzinfo=TZ
        ).timetz()
    )

    print("🤖 Bot ishlayapti...")
    print(f"🔔 Reminder: {REMINDER_HOUR:02d}:{REMINDER_MINUTE:02d}")
    print(f"📊 Daily report: {REPORT_HOUR:02d}:{REPORT_MINUTE:02d}")
    print(f"📥 Weekly Excel: Sunday {WEEKLY_HOUR:02d}:{WEEKLY_MINUTE:02d}")
    print(f"📥 Monthly Excel: last day {MONTHLY_HOUR:02d}:{MONTHLY_MINUTE:02d}")

    app.run_polling()


if __name__ == "__main__":
    main()
