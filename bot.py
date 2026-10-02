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

REPORT_NOTE = """📝 Not: Şantiyenin dili verdiği rapordur; raporu olmayan iş tamamlanmış sayılmaz. ⚠️
Lütfen günlük raporlarınızı zamanında iletiniz."""

REPORT_NOTE_2 = """📊
Bunca çabaya rağmen rapor iletmeyen şantiyeler, lütfen rapor düzenine özen göstersin. 🙏
Unutmayın: İşi yapmak cesarettir, raporlamak ise disiplindir. ⚠️"""

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")
TZ = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))

HISTORY_FILE = "raporlar_yangi.xlsx"
LIVE_FILE = "raporlar.xlsx"

# Faqat yangi davrdagi raporlar saqlanadi.
# Eski Excel fayllari ataylab o'qilmaydi.
HISTORY_FILES = [
    "raporlar_yangi.xlsx",
]

REMINDER_HOUR = int(os.getenv("REMINDER_HOUR", "12"))
REMINDER_MINUTE = int(os.getenv("REMINDER_MINUTE", "0"))
REPORT_HOUR = int(os.getenv("REPORT_HOUR", "17"))
REPORT_MINUTE = int(os.getenv("REPORT_MINUTE", "30"))
WEEKLY_HOUR = int(os.getenv("WEEKLY_HOUR", "18"))
WEEKLY_MINUTE = int(os.getenv("WEEKLY_MINUTE", "0"))
MONTHLY_HOUR = int(os.getenv("MONTHLY_HOUR", "18"))
MONTHLY_MINUTE = int(os.getenv("MONTHLY_MINUTE", "30"))

SITES = [
    "DATA CENTER", "DMC", "LOT13", "LOT71", "SKP", "STADYUM",
    "BWC", "KÖKSARAY", "MMP", "MOS", "PİRAMİT", "RMC", "TYM", "YHP"
]

ALIASES = {
    "DATA CENTER": "DATA CENTER",
    "DMC": "DMC",
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
    "MPP": "MMP",
    "MMP": "MMP",
    "MOS": "MOS",
    "RMC": "RMC",
    "TYM": "TYM",
    "YHP": "YHP",
    "ELLIPSE GARDEN": "DMC",
    "DMC ELLIPSE GARDEN": "DMC",
}

SITE_RESPONSIBLE_USERNAMES = {
    "BWC": "@YSF1434",
    "DMC": "@uzyusufmutlu",
    "MMP": "@Orhan_Ceylan",
    "MOS": "@Orhan_Ceylan",
    "RMC": "@Orhan_Ceylan",
    "TYM": "@Orhan_Ceylan",
    "YHP": "@Orhan_Ceylan",
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


def canonical_site_name(site):
    s = normalize(site or "")
    aliases = {
        "ELLIPSE GARDEN": "DMC",
        "DMC ELLIPSE GARDEN": "DMC",
        "DMCELLIPSEGARDEN": "DMC",
        "MPP": "MMP",
    }
    if s in aliases:
        return aliases[s]
    if isinstance(site, str):
        return site.strip()
    return site


def responsible_username(site, mapping=None):
    canonical = canonical_site_name(site)
    if canonical in SITE_RESPONSIBLE_USERNAMES:
        return SITE_RESPONSIBLE_USERNAMES[canonical]
    if mapping:
        info = mapping.get(canonical) or mapping.get(site)
        if info:
            username = str(info.get("username") or "").strip()
            sender = str(info.get("sender") or "").strip()
            if username:
                return username if username.startswith("@") else "@" + username
            if sender:
                return sender
    return "Mas'ul aniqlanmadi"


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


def parse_date_range(text):
    if not text:
        return None, None

    matches = re.findall(
        r"\b(?:\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}|\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})\b",
        str(text)
    )

    parsed = [parse_date(x) for x in matches]
    parsed = [x for x in parsed if x is not None]

    if len(parsed) >= 2:
        start_date, end_date = parsed[0], parsed[1]
        if start_date > end_date:
            start_date, end_date = end_date, start_date
        return start_date, end_date

    if len(parsed) == 1:
        return parsed[0], parsed[0]

    return None, None


def find_column(headers, names):
    normalized_headers = {}
    for index, header in enumerate(headers):
        if header is not None:
            normalized_headers[normalize(header)] = index
    for name in names:
        if normalize(name) in normalized_headers:
            return normalized_headers[normalize(name)]
    return None


def get_history_file():
    if os.path.exists(HISTORY_FILE):
        return HISTORY_FILE
    return None


def get_history_files():
    files = []
    for filename in [HISTORY_FILE] + HISTORY_FILES:
        if filename not in files and os.path.exists(filename):
            files.append(filename)
    return files


def ensure_excel():
    filename = get_history_file() or HISTORY_FILE

    if os.path.exists(filename):
        try:
            wb = load_workbook(filename)
            ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
            headers = [c.value for c in ws[1]]

            needed = [
                "Tarih", "Saat", "Şantiye", "Gönderen",
                "İşçi Sayısı", "Rapor Metni", "Mesaj ID",
                "Telegram Chat ID", "Telegram User ID", "Username"
            ]

            changed = False
            for header in needed:
                if header not in headers:
                    ws.cell(row=1, column=ws.max_column + 1, value=header)
                    headers.append(header)
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


def extract_worker_count(text):
    patterns = [
        r"(?:ishchi|ishchilar|рабочих|рабочие|işçi|işçiler)\s*[:\-]?\s*(\d+)",
        r"(?:işçi sayısı|ishchi soni|количество рабочих)\s*[:\-]?\s*(\d+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def append_report_to_excel(update, site, text):
    try:
        ensure_excel()
        filename = get_history_file() or HISTORY_FILE

        wb = load_workbook(filename)
        ws = wb["Raporlar"] if "Raporlar" in wb.sheetnames else wb.active
        headers = [c.value for c in ws[1]]

        required = [
            "Tarih", "Saat", "Şantiye", "Gönderen",
            "İşçi Sayısı", "Rapor Metni", "Mesaj ID",
            "Telegram Chat ID", "Telegram User ID", "Username"
        ]

        for header in required:
            if header not in headers:
                ws.cell(row=1, column=ws.max_column + 1, value=header)
                headers.append(header)

        user = update.effective_user
        chat = update.effective_chat
        message = update.message
        now = datetime.now(TZ)

        values = {
            "Tarih": now.date(),
            "Saat": now.strftime("%H:%M"),
            "Şantiye": canonical_site_name(site),
            "Gönderen": user.full_name if user else "",
            "İşçi Sayısı": extract_worker_count(text),
            "Rapor Metni": text,
            "Mesaj ID": message.message_id if message else "",
            "Telegram Chat ID": chat.id if chat else "",
            "Telegram User ID": user.id if user else "",
            "Username": f"@{user.username}" if user and user.username else "",
        }

        row_number = ws.max_row + 1
        for col_number, header in enumerate(headers, start=1):
            if header in values:
                ws.cell(row=row_number, column=col_number, value=values[header])

        wb.save(filename)
        wb.close()

        print(
            f"✅ Rapor saqlandi: {canonical_site_name(site)} | "
            f"{now.strftime('%d.%m.%Y %H:%M')} | {values['Gönderen']}"
        )
        return True

    except Exception as e:
        print("❌ Rapor Excelga saqlashda xato:", e)
        return False


def search_excel(start_date=None, end_date=None, site=None, keyword=None):
    filenames = get_history_files()
    if not filenames:
        print("Excel topilmadi:", HISTORY_FILES)
        return []

    results = []
    seen = set()

    for filename in filenames:
        try:
            wb = load_workbook(filename, read_only=True, data_only=True)
        except Exception as e:
            print(f"Excel ochishda xato ({filename}): {e}")
            continue

        try:
            sheets = [wb["Raporlar"]] if "Raporlar" in wb.sheetnames else wb.worksheets

            for ws in sheets:
                rows = ws.iter_rows(values_only=True)
                try:
                    headers = next(rows)
                except StopIteration:
                    continue

                date_col = find_column(headers, ["Tarih", "Sana", "Date", "Дата"])
                time_col = find_column(headers, ["Saat", "Vaqt", "Time", "Время"])
                site_col = find_column(headers, ["Şantiye", "Shantiye", "Шантиё", "Obyekt"])
                sender_col = find_column(headers, ["Gönderen", "Yuboruvchi", "Sender", "Отправитель"])
                worker_col = find_column(headers, ["İşçi Sayısı", "Ishchi Sonı", "Ishchi Soni", "Рабочие", "Количество рабочих"])
                text_col = find_column(headers, ["Rapor Metni", "Rapor", "Hisobot", "Report", "Сообщение", "Toliq_raport", "Toliq raport"])
                message_id_col = find_column(headers, ["Mesaj ID", "Message ID", "ID"])
                chat_id_col = find_column(headers, ["Telegram Chat ID", "Chat ID"])
                user_id_col = find_column(headers, ["Telegram User ID", "User ID"])
                username_col = find_column(headers, ["Username", "Telegram Username"])

                for row in rows:
                    if not row:
                        continue

                    row_date = parse_date(row[date_col]) if date_col is not None and date_col < len(row) else None

                    if (start_date or end_date) and row_date is None:
                        continue
                    if start_date and row_date < start_date:
                        continue
                    if end_date and row_date > end_date:
                        continue

                    row_site = str(row[site_col] or "").strip() if site_col is not None and site_col < len(row) else ""
                    row_site = canonical_site_name(row_site)

                    if site and normalize(canonical_site_name(site)) != normalize(row_site):
                        continue

                    all_text = " ".join(str(x or "") for x in row)
                    if keyword and normalize(keyword) not in normalize(all_text):
                        continue

                    message_id = row[message_id_col] if message_id_col is not None and message_id_col < len(row) else ""
                    text_value = row[text_col] if text_col is not None and text_col < len(row) else ""
                    sender_value = row[sender_col] if sender_col is not None and sender_col < len(row) else ""
                    username_value = row[username_col] if username_col is not None and username_col < len(row) else ""

                    if message_id not in (None, ""):
                        key = f"{filename}|{message_id}"
                    else:
                        key = (
                            row_date,
                            str(row[time_col] if time_col is not None and time_col < len(row) else ""),
                            normalize(row_site),
                            normalize(sender_value),
                            normalize(text_value),
                        )

                    if key in seen:
                        continue
                    seen.add(key)

                    results.append({
                        "date": row_date,
                        "time": row[time_col] if time_col is not None and time_col < len(row) else "",
                        "site": row_site,
                        "sender": sender_value,
                        "workers": row[worker_col] if worker_col is not None and worker_col < len(row) else "",
                        "text": text_value,
                        "message_id": message_id,
                        "chat_id": row[chat_id_col] if chat_id_col is not None and chat_id_col < len(row) else "",
                        "user_id": row[user_id_col] if user_id_col is not None and user_id_col < len(row) else "",
                        "username": username_value,
                    })

        finally:
            wb.close()

    results.sort(
        key=lambda r: (r.get("date") or date.min, str(r.get("time") or "")),
        reverse=True
    )
    return results


async def chatid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show the current Telegram chat ID so ADMIN_CHAT_ID can be configured."""
    if update.message and update.effective_chat:
        await update.message.reply_text(
            f"🆔 Chat ID: {update.effective_chat.id}"
        )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message:
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
    found_normalized = {normalize(canonical_site_name(x)) for x in found_sites}

    missing_sites = [
        s for s in SITES
        if normalize(canonical_site_name(s)) not in found_normalized
    ]

    text = (
        "📊 Şantiye Rapor Durumu\n\n"
        f"📅 Tarih: {today.strftime('%d.%m.%Y')}\n\n"
        f"✅ Rapor gönderilen: {len(found_sites)}\n"
    )

    text += (
        "\n".join(f"• {canonical_site_name(s)}" for s in found_sites)
        if found_sites else "• Henüz yok"
    )

    text += "\n\n❌ Rapor gönderilmeyen:\n"

    if missing_sites:
        text += "\n".join(
            f"• {canonical_site_name(s)} ({responsible_username(s)})"
            for s in missing_sites
        )
    else:
        text += "• Hepsi gönderildi"

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


def format_search_results(results):
    chunks = []
    current = ""

    for i, r in enumerate(results, start=1):
        d = r["date"]
        d_text = d.strftime("%d.%m.%Y") if isinstance(d, date) else str(d or "")
        report_text = str(r["text"] or "").strip()

        if len(report_text) > 500:
            report_text = report_text[:500] + "..."

        item = (
            f"━━━━━━━━━━━━━━\n"
            f"#{i}\n"
            f"📅 Tarih: {d_text}\n"
            f"⏰ Saat: {r['time'] or ''}\n"
            f"🏗 Şantiye: {r['site'] or '—'}\n"
            f"👤 Gönderen: {r['sender'] or '—'}\n"
            f"👷 İşçi: {r['workers'] or '—'}\n"
        )

        if report_text:
            item += f"📝 {report_text}\n"

        if len(current) + len(item) > 3800:
            if current:
                chunks.append(current)
            current = item
        else:
            current += item

    if current:
        chunks.append(current)

    return chunks


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
            await update.message.reply_text("⛔ Bu bölüm sadece yönetici grubunda kullanılabilir.")
            return

        filename = get_history_file()
        if not filename:
            await update.message.reply_text("⚠️ Excel dosyası bulunamadı.")
            return

        with open(filename, "rb") as f:
            await update.message.reply_document(f, filename=filename)
        return

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
                "Misol:\n01.01.2026 - 05.09.2026\n"
                "ёки\nPIRAMIT\n"
                "ёки\n05.09.2026",
                reply_markup=MENU
            )
            context.user_data["search_mode"] = False
            return

        results = results[:100]
        chunks = format_search_results(results)

        await update.message.reply_text(
            f"🔎 Bulunan rapor sayısı: {len(results)}"
        )

        for chunk in chunks:
            await update.message.reply_text(chunk)

        context.user_data["search_mode"] = False
        await update.message.reply_text(
            "🔎 Arama tamamlandı.",
            reply_markup=MENU
        )
        return

    site = detect_site(text)

    if site:
        saved = append_report_to_excel(update, site, text)

        if saved:
            await update.message.reply_text(
                f"✅ Rapor alındı.\n"
                f"🏗 {canonical_site_name(site)}\n"
                f"👤 {update.effective_user.full_name if update.effective_user else '—'}"
            )
        else:
            await update.message.reply_text(
                "⚠️ Rapor alınamadı. Render Logs kontrol edilmeli."
            )
        return


def build_site_sender_map():
    reports = search_excel()
    mapping = {}

    for r in reports:
        site = canonical_site_name(str(r.get("site") or "").strip())
        if not site:
            continue

        user_id = str(r.get("user_id") or "").strip()
        chat_id = str(r.get("chat_id") or "").strip()
        sender = str(r.get("sender") or "").strip()
        username = str(r.get("username") or "").strip()

        target_id = user_id or chat_id
        if not target_id:
            continue

        mapping[site] = {
            "chat_id": chat_id,
            "user_id": user_id,
            "target_id": target_id,
            "sender": sender or username or target_id,
            "username": username,
            "last_date": r.get("date"),
        }

    return mapping


async def reminder_1200(context: ContextTypes.DEFAULT_TYPE):
    mapping = build_site_sender_map()
    by_target = {}

    for raw_site in SITES:
        site = canonical_site_name(raw_site)
        info = mapping.get(site)

        if not info:
            continue

        target_id = info["target_id"]
        by_target.setdefault(target_id, []).append(site)

    for target_id, sites in by_target.items():
        sites_text = "\n".join(f"• {s}" for s in sites)

        text = (
            "🔔 Hatırlatma!\n\n"
            "Bugünkü raporu saat 17:30'a kadar göndermeyi unutmayın.\n\n"
            "Size bağlı şantiyeler:\n"
            f"{sites_text}"
        )

        try:
            await context.bot.send_message(
                chat_id=int(target_id),
                text=text
            )
        except Exception as e:
            print(f"12:00 hatırlatması gönderilirken hata ({target_id}): {e}")


async def daily_report_1730(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        print("ADMIN_CHAT_ID ayarlanmamış.")
        return

    today = datetime.now(TZ).date()
    reports = search_excel(start_date=today, end_date=today)

    site_reports = {}
    for r in reports:
        site = canonical_site_name(str(r.get("site") or "").strip())
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
        canonical = canonical_site_name(site)
        rows = site_reports.get(normalize(canonical), [])

        if rows:
            sent_count += 1
            senders = []

            for r in rows:
                name = str(r.get("sender") or "").strip()
                username = str(r.get("username") or "").strip()
                display = name or username or "Bilinmiyor"

                if username and username not in display:
                    display += f" ({username})"

                if display not in senders:
                    senders.append(display)

            lines.append(f"✅ {canonical} — gönderildi")

            for sender in senders:
                lines.append(f"   👤 {sender}")

        else:
            missing_count += 1
            lines.append(f"❌ {canonical} — gönderilmedi")
            lines.append(
                f"   👤 Sorumlu: {responsible_username(canonical, mapping)}"
            )

    lines.extend([
        "",
        f"📌 Toplam şantiye: {len(SITES)}",
        f"✅ Gönderildi: {sent_count}",
        f"❌ Gönderilmedi: {missing_count}",
        "",
        REPORT_NOTE,
        "",
        REPORT_NOTE_2,
    ])

    try:
        await context.bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text="\n".join(lines)
        )
    except Exception as e:
        print("17:30 daily report xato:", e)


def make_weekly_excel(target_date=None):
    if target_date is None:
        target_date = datetime.now(TZ).date()

    monday = target_date - timedelta(days=target_date.weekday())
    sunday = monday + timedelta(days=6)

    reports = search_excel(start_date=monday, end_date=sunday)

    sent = set()

    for r in reports:
        d = r.get("date")
        site = canonical_site_name(str(r.get("site") or "").strip())

        if site and isinstance(d, date):
            sent.add((normalize(site), d))

    mapping = build_site_sender_map()

    wb = Workbook()
    ws = wb.active
    ws.title = "Yuborilmagan raporlar"
    ws.append(["Sana", "Obyekt", "Mas'ul", "Telegram", "Holat"])

    missing_rows = []

    for i in range(7):
        d = monday + timedelta(days=i)

        for raw_site in SITES:
            site = canonical_site_name(raw_site)

            if (normalize(site), d) in sent:
                continue

            info = mapping.get(site, {})
            sender = str(info.get("sender") or "Mas'ul aniqlanmadi").strip()
            username = responsible_username(site, mapping)

            ws.append([d, site, sender, username, "❌ Yuborilmadi"])

            missing_rows.append({
                "date": d,
                "site": site,
                "sender": sender,
                "username": username,
            })

    if not missing_rows:
        ws.append(["", "", "", "", "✅ Bu hafta barcha raporlar yuborilgan"])

    for col, width in {"A": 14, "B": 20, "C": 28, "D": 24, "E": 20}.items():
        ws.column_dimensions[col].width = width

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    filename = (
        f"Yuborilmagan_Raporlar_"
        f"{monday.strftime('%d.%m.%Y')}_"
        f"{sunday.strftime('%d.%m.%Y')}.xlsx"
    )

    return stream, filename, missing_rows


async def weekly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        print("ADMIN_CHAT_ID ayarlanmamış.")
        return

    try:
        stream, filename, missing_rows = make_weekly_excel(datetime.now(TZ).date())

        await context.bot.send_document(
            chat_id=ADMIN_CHAT_ID,
            document=stream,
            filename=filename,
            caption="📥 Haftalık eksik raporlar Excel — sadece gönderilmeyen günler."
        )

        if not missing_rows:
            msg = (
                "🔔 HAFTALIK RAPOR KONTROLÜ\n\n"
                "✅ Bu hafta tüm şantiyelerin günlük raporları eksiksiz gönderildi."
            )
        else:
            grouped = {}

            for item in missing_rows:
                key = item["username"]
                grouped.setdefault(key, []).append(item)

            lines = [
                "🔔 HAFTALIK RAPOR HATIRLATMASI",
                "",
                "Aşağıdaki raporlar gönderilmemiştir.",
                "Lütfen sorumlu kişiler eksik günlerin raporlarını tamamlasın.",
                "",
            ]

            for responsible, items in grouped.items():
                lines.append(f"👤 {responsible}")
                by_site = {}

                for item in items:
                    by_site.setdefault(item["site"], []).append(
                        item["date"].strftime("%d.%m.%Y")
                    )

                for site, dates in by_site.items():
                    lines.append(f"• {site}: {', '.join(dates)}")

                lines.append("")

            lines.extend([
                REPORT_NOTE,
                "",
                REPORT_NOTE_2,
            ])
            msg = "\n".join(lines)

        await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=msg)

    except Exception as e:
        print("Haftalık Excel/hatırlatma hatası:", e)


def make_monthly_excel(target_date=None):
    if target_date is None:
        target_date = datetime.now(TZ).date()

    first_day = target_date.replace(day=1)
    last_day = target_date.replace(
        day=monthrange(target_date.year, target_date.month)[1]
    )

    reports = search_excel(start_date=first_day, end_date=last_day)

    sent = {}

    for r in reports:
        site = canonical_site_name(str(r.get("site") or "").strip())
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
        canonical = canonical_site_name(site)
        total = (last_day - first_day).days + 1
        sent_count = 0
        missing_days = []

        for i in range(total):
            d = first_day + timedelta(days=i)

            if sent.get((normalize(canonical), d)):
                sent_count += 1
            else:
                missing_days.append(d.strftime("%d.%m"))

        percent = round(sent_count * 100 / total, 1) if total else 0

        responsible = mapping.get(canonical, {}).get(
            "sender",
            responsible_username(canonical, mapping)
        )

        status_text = "To'liq" if not missing_days else "Kamchilik bor"

        summary.append([
            canonical,
            responsible,
            total,
            sent_count,
            ", ".join(missing_days) if missing_days else "—",
            f"{percent}%",
            status_text
        ])

    for site in SITES:
        canonical = canonical_site_name(site)
        safe_title = re.sub(r'[:\\/?*\[\]]', '_', canonical)[:31]

        ws = wb.create_sheet(safe_title)
        ws.append(["Sana", "Obyekt", "Holat", "Yuboruvchi", "Rapor matni"])

        total = (last_day - first_day).days + 1

        for i in range(total):
            d = first_day + timedelta(days=i)
            rows = sent.get((normalize(canonical), d), [])

            if rows:
                for r in rows:
                    ws.append([
                        d,
                        canonical,
                        "Yuborildi",
                        r.get("sender") or "Noma'lum",
                        r.get("text") or ""
                    ])
            else:
                responsible = mapping.get(canonical, {}).get(
                    "sender",
                    responsible_username(canonical, mapping)
                )
                ws.append([d, canonical, "Yuborilmadi", responsible, ""])

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    return stream, f"Oylik_Raport_{first_day.strftime('%m.%Y')}.xlsx"


async def monthly_excel_report(context: ContextTypes.DEFAULT_TYPE):
    if not ADMIN_CHAT_ID:
        return

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
    app.add_handler(CommandHandler("chatid", chatid))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.job_queue.run_daily(
        reminder_1200,
        time=datetime(
            2000, 1, 1,
            REMINDER_HOUR, REMINDER_MINUTE,
            tzinfo=TZ
        ).timetz()
    )

    app.job_queue.run_daily(
        daily_report_1730,
        time=datetime(
            2000, 1, 1,
            REPORT_HOUR, REPORT_MINUTE,
            tzinfo=TZ
        ).timetz()
    )

    app.job_queue.run_daily(
        weekly_excel_report,
        time=datetime(
            2000, 1, 1,
            WEEKLY_HOUR, WEEKLY_MINUTE,
            tzinfo=TZ
        ).timetz(),
        days=(0,)
    )

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
