import logging
import csv
import os
import json
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove, BotCommand
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters, ContextTypes,
)

# ============ НАСТРОЙКИ ============
TOKEN = "8996911876:AAExcImtJ8ebUo4GPT650ZwFO_cHoh1w3Sw"
ADMIN_CHAT_ID = 1084292156
CSV_FILE = "zayavki.csv"
EXCEL_FILE = "zayavki.xlsx"
SLOTS_FILE = "available_slots.json"
BLACKLIST_FILE = "blacklist.json"

WORKING_HOURS = (9, 21)  # с 9 до 21

# ============ УСЛУГИ ============
SERVICES = [
    ("Классический маникюр — 1200 ₽", "classic"),
    ("Гель-лак — 1800 ₽", "gel"),
    ("Дизайн / Арт — +500 ₽", "design"),
    ("Наращивание — 2500 ₽", "extension"),
    ("Другое", "other"),
]

MASTERS = ["Анна", "Мария", "Любой мастер"]

BACK_CALLBACK = "back"
FIXED_TIMES = ["09:00", "12:00", "15:00", "18:00", "21:00"]

def escape_html(text: str) -> str:
    """Экранирует спецсимволы для HTML"""
    if not text:
        return ""
    return str(text).replace("&", "&").replace("<", "<").replace(">", ">")

# ============ Вспомогательные функции ============
def is_working_hours() -> bool:
    hour = datetime.now().hour
    return WORKING_HOURS[0] <= hour < WORKING_HOURS[1]

def load_blacklist():
    if os.path.exists(BLACKLIST_FILE):
        with open(BLACKLIST_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_blacklist(bl):
    with open(BLACKLIST_FILE, "w") as f:
        json.dump(list(bl), f)

def is_blacklisted(user_id) -> bool:
    return str(user_id) in load_blacklist()

def validate_name(name: str):
    name = name.strip()
    if len(name) < 2: return False, "Имя слишком короткое."
    if len(name) > 50: return False, "Имя слишком длинное."
    if not any(c.isalpha() for c in name): return False, "Имя должно содержать буквы."
    return True, name

def validate_contact(contact: str):
    import re
    contact = contact.strip()
    cleaned = re.sub(r"[^\d+]", "", contact)
    if cleaned.startswith("+7") and len(cleaned) == 12: return True, cleaned
    if cleaned.startswith("8") and len(cleaned) == 11: return True, "+7" + cleaned[1:]
    if cleaned.startswith("7") and len(cleaned) == 11: return True, "+" + cleaned
    if contact.startswith("@") and re.match(r"^@[a-zA-Z0-9_]{5,32}$", contact): return True, contact
    if re.match(r"^[a-zA-Z0-9_]{5,32}$", contact) and not contact.isdigit(): return True, "@" + contact
    if contact.isdigit() and 6 <= len(contact) <= 12: return True, contact
    if re.search(r"vk\.com/[a-zA-Z0-9_.]+", contact, re.IGNORECASE): return True, contact
    return False, "Неверный формат. Телефон, @username или ссылка на VK."

# ============ Слоты ============
def generate_month_slots():
    import calendar
    slots = {}
    today = datetime.now().date()
    last_day = calendar.monthrange(today.year, today.month)[1]
    end = today.replace(day=last_day)
    current = today + timedelta(days=1)
    while current <= end:
        slots[current.strftime("%Y-%m-%d")] = FIXED_TIMES.copy()
        current += timedelta(days=1)
    return slots

def load_slots():
    slots = generate_month_slots()
    if os.path.exists(SLOTS_FILE):
        try:
            with open(SLOTS_FILE, "r", encoding="utf-8") as f:
                old = json.load(f)
            for d, times in old.items():
                if d in slots:
                    slots[d] = [t for t in FIXED_TIMES if t in times]
        except: pass
    save_slots(slots)
    return slots

def save_slots(slots):
    with open(SLOTS_FILE, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

def get_available_dates():
    return sorted([d for d, t in load_slots().items() if t])

def get_available_times(date_str):
    return load_slots().get(date_str, [])

def book_slot(date_str, time_str):
    slots = load_slots()
    if date_str in slots and time_str in slots[date_str]:
        slots[date_str].remove(time_str)
        if not slots[date_str]: del slots[date_str]
        save_slots(slots)
        return True
    return False

def free_slot(date_str, time_str):
    slots = load_slots()
    if date_str not in slots:
        slots[date_str] = []
    if time_str not in slots[date_str]:
        slots[date_str].append(time_str)
        slots[date_str].sort()
        save_slots(slots)

WEEKDAYS_RU = {0:"пн",1:"вт",2:"ср",3:"чт",4:"пт",5:"сб",6:"вс"}

def format_date_ru(d, with_wd=True):
    try:
        dt = datetime.strptime(d, "%Y-%m-%d")
        s = dt.strftime("%d.%m")
        return f"{s} ({WEEKDAYS_RU[dt.weekday()]})" if with_wd else s
    except: return d

# ============ Сохранение ============
def save_to_csv(data):
    exists = os.path.exists(CSV_FILE)
    with open(CSV_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["datetime","name","phone","date","time","service","master","comment","user_id"])
        if not exists: w.writeheader()
        w.writerow(data)

def save_to_excel(data):
    try:
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        headers = ["Дата заявки","Имя","Контакт","Дата","Время","Услуга","Мастер","Комментарий","ID"]
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Заявки"
            ws.append(headers)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="FF69B4")
        try: nice = datetime.strptime(data["date"],"%Y-%m-%d").strftime("%d.%m.%Y")
        except: nice = data["date"]
        ws.append([data["datetime"], data["name"], data["phone"], nice, data["time"], data["service"], data.get("master",""), data["comment"], str(data["user_id"])])
        wb.save(EXCEL_FILE)
    except Exception as e:
        logger.error(f"Excel error: {e}")

def get_user_bookings(user_id):
    if not os.path.exists(CSV_FILE): return []
    bookings = []
    with open(CSV_FILE, "r", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("user_id") == str(user_id):
                bookings.append(row)
    return bookings

def cancel_booking(user_id, date, time):
    if not os.path.exists(CSV_FILE): return False
    rows = []
    found = False
    with open(CSV_FILE, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        for row in reader:
            if row.get("user_id") == str(user_id) and row.get("date") == date and row.get("time") == time:
                found = True
                continue
            rows.append(row)
    if found:
        with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(rows)
        free_slot(date, time)
    return found

# ============ Состояния ============
NAME, PHONE, MASTER, DATE, TIME, SERVICE, COMMENT = range(7)
DATES_PER_PAGE = 10

def build_dates_keyboard(dates, page=0):
    start = page * DATES_PER_PAGE
    end = start + DATES_PER_PAGE
    page_dates = dates[start:end]
    kb = []
    row = []
    for d in page_dates:
        row.append(InlineKeyboardButton(format_date_ru(d), callback_data=f"date_{d}"))
        if len(row) == 2:
            kb.append(row)
            row = []
    if row: kb.append(row)
    nav = []
    if page > 0: nav.append(InlineKeyboardButton("◀️", callback_data=f"page_{page-1}"))
    if end < len(dates): nav.append(InlineKeyboardButton("▶️", callback_data=f"page_{page+1}"))
    if nav: kb.append(nav)
    kb.append([InlineKeyboardButton("⬅️ Назад", callback_data=BACK_CALLBACK)])
    return InlineKeyboardMarkup(kb)

# ============ Хендлеры ============
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if is_blacklisted(user_id):
        await update.message.reply_text("Вы заблокированы.")
        return ConversationHandler.END
    if not is_working_hours():
        await update.message.reply_text("Сейчас нерабочее время (9:00–21:00). Напишите утром — отвечу!")
        return ConversationHandler.END

    keyboard = [
        ["📅 Записаться"],
        ["📋 Мои записи", "❌ Отменить запись"],
    ]
    await update.message.reply_text(
        "Привет! Я бот для записи на маникюр 💅\n\nВыбери действие:",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return ConversationHandler.END

async def get_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ok, res = validate_name(update.message.text)
    if not ok:
        await update.message.reply_text(res)
        return NAME
    context.user_data["name"] = res
    await update.message.reply_text("Укажи контакт (телефон, @username или VK):")
    return PHONE

async def get_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ok, res = validate_contact(update.message.text)
    if not ok:
        await update.message.reply_text(res)
        return PHONE
    context.user_data["phone"] = res
    kb = [[InlineKeyboardButton(m, callback_data=f"master_{m}")] for m in MASTERS]
    kb.append([InlineKeyboardButton("⬅️ Назад", callback_data=BACK_CALLBACK)])
    await update.message.reply_text("Выбери мастера:", reply_markup=InlineKeyboardMarkup(kb))
    return MASTER

async def master_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == BACK_CALLBACK:
        await q.edit_message_text("Укажи контакт:")
        return PHONE
    context.user_data["master"] = q.data.replace("master_", "")
    context.user_data["date_page"] = 0
    dates = get_available_dates()
    if not dates:
        await q.edit_message_text("Нет свободных дат.")
        return ConversationHandler.END
    await q.edit_message_text("Выбери дату:", reply_markup=build_dates_keyboard(dates, 0))
    return DATE

async def date_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == BACK_CALLBACK:
        kb = [[InlineKeyboardButton(m, callback_data=f"master_{m}")] for m in MASTERS]
        await q.edit_message_text("Выбери мастера:", reply_markup=InlineKeyboardMarkup(kb))
        return MASTER
    if q.data.startswith("page_"):
        page = int(q.data.split("_")[1])
        context.user_data["date_page"] = page
        await q.edit_message_text("Выбери дату:", reply_markup=build_dates_keyboard(get_available_dates(), page))
        return DATE
    date = q.data.replace("date_", "")
    context.user_data["date"] = date
    times = get_available_times(date)
    if not times:
        await q.edit_message_text("На эту дату нет времени.")
        return DATE
    kb = []
    row = []
    for t in times:
        row.append(InlineKeyboardButton(t, callback_data=f"time_{t}"))
        if len(row) == 3: kb.append(row); row = []
    if row: kb.append(row)
    kb.append([InlineKeyboardButton("⬅️ Назад", callback_data=BACK_CALLBACK)])
    await q.edit_message_text("Выбери время:", reply_markup=InlineKeyboardMarkup(kb))
    return TIME

async def time_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == BACK_CALLBACK:
        page = context.user_data.get("date_page", 0)
        await q.edit_message_text("Выбери дату:", reply_markup=build_dates_keyboard(get_available_dates(), page))
        return DATE
    context.user_data["time"] = q.data.replace("time_", "")
    kb = [[InlineKeyboardButton(t, callback_data=f"service_{c}")] for t, c in SERVICES]
    kb.append([InlineKeyboardButton("⬅️ Назад", callback_data=BACK_CALLBACK)])
    await q.edit_message_text("Выбери услугу:", reply_markup=InlineKeyboardMarkup(kb))
    return SERVICE

async def service_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == BACK_CALLBACK:
        date = context.user_data["date"]
        times = get_available_times(date)
        kb = [[InlineKeyboardButton(t, callback_data=f"time_{t}")] for t in times]
        kb.append([InlineKeyboardButton("⬅️ Назад", callback_data=BACK_CALLBACK)])
        await q.edit_message_text("Выбери время:", reply_markup=InlineKeyboardMarkup(kb))
        return TIME
    cb = q.data.replace("service_", "")
    context.user_data["service"] = next(t for t, c in SERVICES if c == cb)
    await q.edit_message_text("Комментарий или пожелания? (или напиши «нет»)")
    return COMMENT

async def get_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    comment = update.message.text.strip()
    if comment.lower() in ["нет", "-", "no"]: comment = ""
    if len(comment) > 300:
        await update.message.reply_text("Слишком длинный комментарий.")
        return COMMENT
    context.user_data["comment"] = comment

    date = context.user_data["date"]
    time = context.user_data["time"]
    if not book_slot(date, time):
        await update.message.reply_text("Это время уже заняли. Начни заново /start")
        return ConversationHandler.END

    data = {
        "datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "name": context.user_data["name"],
        "phone": context.user_data["phone"],
        "date": date,
        "time": time,
        "service": context.user_data["service"],
        "master": context.user_data.get("master", ""),
        "comment": comment,
        "user_id": update.effective_user.id,
    }
    save_to_csv(data)
    save_to_excel(data)

    nice = datetime.strptime(date, "%Y-%m-%d").strftime("%d.%m.%Y")
    text = (f"✅ Заявка принята!\n\nИмя: {data['name']}\nКонтакт: {data['phone']}\n"
            f"Мастер: {data['master']}\nДата: {nice}\nВремя: {time}\nУслуга: {data['service']}\n"
            f"Комментарий: {comment or '—'}")
    await update.message.reply_text(text)

    admin = (f"🔔 Новая заявка\n\n"
             f"Имя: {escape_html(data['name'])}\n"
             f"Контакт: {escape_html(data['phone'])}\n"
             f"Мастер: {escape_html(data['master'])}\n"
             f"Дата: {nice} {time}\n"
             f"Услуга: {escape_html(data['service'])}\n"
             f"Комментарий: {escape_html(comment or '—')}")
    try:
        await context.bot.send_message(ADMIN_CHAT_ID, admin)
    except Exception as e:
        logger.error(f"Не удалось отправить админу: {e}")

    context.user_data.clear()
    return ConversationHandler.END

async def menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "📅 Записаться":
        await update.message.reply_text("Как тебя зовут?", reply_markup=ReplyKeyboardRemove())
        return NAME
    elif text == "📋 Мои записи":
        await my_bookings(update, context)
    elif text == "❌ Отменить запись":
        await cancel_booking_cmd(update, context)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Отменено. Нажми /start чтобы начать заново",
        reply_markup=ReplyKeyboardRemove()
    )
    context.user_data.clear()
    return ConversationHandler.END

# ============ Клиентские команды ============
async def my_bookings(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    bookings = get_user_bookings(uid)
    if not bookings:
        await update.message.reply_text("У тебя нет активных записей.")
        return
    text = "📋 Твои записи:\n\n"
    for b in bookings:
        try:
            date_str = b.get("date", "")
            if date_str:
                nice = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
            else:
                nice = "дата не указана"
            time_str = b.get("time", "—")
            service = b.get("service", "—")
            master = b.get("master", "")
            text += f"• {nice} в {time_str} — {service}"
            if master:
                text += f" ({master})"
            text += "\n"
        except Exception:
            continue
    text += "\nЧтобы отменить — нажми «❌ Отменить запись»"
    await update.message.reply_text(text)

async def cancel_booking_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    bookings = get_user_bookings(uid)
    if not bookings:
        await update.message.reply_text("Нет записей для отмены.")
        return
    kb = []
    for b in bookings:
        try:
            date_str = b.get("date", "")
            time_str = b.get("time", "")
            if not date_str or not time_str:
                continue
            nice = datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m")
            kb.append([InlineKeyboardButton(f"{nice} {time_str}", callback_data=f"cancel_{date_str}_{time_str}")])
        except Exception:
            continue
    if not kb:
        await update.message.reply_text("Нет корректных записей для отмены.")
        return
    await update.message.reply_text("Выбери запись для отмены:", reply_markup=InlineKeyboardMarkup(kb))

async def cancel_booking_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    _, date, time = q.data.split("_", 2)
    uid = q.from_user.id
    if cancel_booking(uid, date, time):
        await q.edit_message_text(f"✅ Запись на {date} {time} отменена. Слот свободен.")
        try: await context.bot.send_message(ADMIN_CHAT_ID, f"Клиент отменил запись: {date} {time}")
        except: pass
    else:
        await q.edit_message_text("Не удалось отменить.")

# ============ Админ ============
def is_admin(update): return update.effective_user.id == ADMIN_CHAT_ID

async def admin_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    text = ("🔧 Админ-команды:\n\n"
            "/slots — свободные слоты\n/addslot — добавить\n/delslot — удалить\n"
            "/stats — статистика\n/export — прислать Excel\n"
            "/block <id> — заблокировать\n/unblock <id> — разблокировать\n/help")
    await update.message.reply_text(text)

async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not os.path.exists(CSV_FILE):
        await update.message.reply_text("Заявок пока нет.")
        return
    total = today = 0
    today_str = datetime.now().strftime("%Y-%m-%d")
    with open(CSV_FILE, "r", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            total += 1
            if row.get("datetime","").startswith(today_str): today += 1
    await update.message.reply_text(f"📊 Статистика:\nВсего заявок: {total}\nСегодня: {today}")

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if os.path.exists(EXCEL_FILE):
        await update.message.reply_document(document=open(EXCEL_FILE, "rb"), filename="zayavki.xlsx")
    else:
        await update.message.reply_text("Файл ещё не создан.")

async def block_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not context.args:
        await update.message.reply_text("Использование: /block 123456789")
        return
    bl = load_blacklist()
    bl.add(context.args[0])
    save_blacklist(bl)
    await update.message.reply_text(f"Пользователь {context.args[0]} заблокирован.")

async def unblock_user(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not context.args: return
    bl = load_blacklist()
    bl.discard(context.args[0])
    save_blacklist(bl)
    await update.message.reply_text(f"Пользователь {context.args[0]} разблокирован.")

async def show_slots(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    slots = load_slots()
    if not slots:
        await update.message.reply_text("Нет слотов.")
        return
    text = "📅 Свободные слоты:\n\n"
    for d in sorted(slots.keys())[:15]:
        text += f"{format_date_ru(d)}: {', '.join(slots[d])}\n"
    await update.message.reply_text(text)

# Напоминания
async def send_reminders(context: ContextTypes.DEFAULT_TYPE):
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    if not os.path.exists(CSV_FILE): return
    with open(CSV_FILE, "r", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("date") == tomorrow and row.get("user_id"):
                try:
                    nice = datetime.strptime(row["date"], "%Y-%m-%d").strftime("%d.%m.%Y")
                    await context.bot.send_message(
                        int(row["user_id"]),
                        f"🔔 Напоминание!\nЗавтра {nice} в {row['time']}\n{row['service']}\nЖдём тебя 💅"
                    )
                except: pass

async def post_init(application: Application):
    commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("mybookings", "Посмотреть мои записи"),
        BotCommand("cancel_booking", "Отменить запись"),
        BotCommand("help", "Помощь и команды"),
    ]
    await application.bot.set_my_commands(commands)


def main():
    if not os.path.exists(SLOTS_FILE):
        save_slots(generate_month_slots())

    app = Application.builder().token(TOKEN).post_init(post_init).build()

    # Нижние кнопки
    app.add_handler(MessageHandler(
        filters.Regex("^(📅 Записаться|📋 Мои записи|❌ Отменить запись)$"),
        menu_handler
    ))

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            MASTER: [CallbackQueryHandler(master_callback)],
            DATE: [CallbackQueryHandler(date_callback)],
            TIME: [CallbackQueryHandler(time_callback)],
            SERVICE: [CallbackQueryHandler(service_callback)],
            COMMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_comment)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("mybookings", my_bookings))
    app.add_handler(CommandHandler("cancel_booking", cancel_booking_cmd))
    app.add_handler(CallbackQueryHandler(cancel_booking_callback, pattern="^cancel_"))
    app.add_handler(CommandHandler("help", admin_help))
    app.add_handler(CommandHandler("stats", stats))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(CommandHandler("block", block_user))
    app.add_handler(CommandHandler("unblock", unblock_user))
    app.add_handler(CommandHandler("slots", show_slots))

    if app.job_queue:
        app.job_queue.run_daily(send_reminders, time=datetime.strptime("10:00", "%H:%M").time())

    print("Бот запущен...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
