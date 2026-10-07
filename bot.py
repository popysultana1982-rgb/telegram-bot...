import io
import json
import os
import re
import shutil
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# =========================================================
# CONFIG
# =========================================================

BOT_TOKEN = os.getenv("BOT_TOKEN")

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN environment variable is not set.")

ADMIN_SECRET_KEY = os.getenv("ADMIN_SECRET_KEY", "mysecretadmin123")

MIN_WITHDRAW = 50
WITHDRAW_FEE_PERCENT = 4
REWARD_PER_EMAIL = 25.0  # প্রতি ভ্যালিড জিমেইল অ্যাকাউন্টে ২৫ টাকা
DAILY_FILE_LIMIT = 5
MAX_EMAILS_PER_FILE = 5

REFERRAL_BONUS = 5.0
USDT_RATE = 124.0

DB_NAME = "bot.db"
PORT = int(os.environ.get("PORT", 10000))


# =========================================================
# RENDER WEB SERVER (FIXED FOR UPTIMEROBOT HEAD REQUESTS)
# =========================================================

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Telegram bot is running.")

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()

    def log_message(self, format, *args):
        return


def run_web_server():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), HealthHandler)
    print(f"Web server running on port {PORT}")
    server.serve_forever()


# =========================================================
# DATABASE
# =========================================================

db = sqlite3.connect(DB_NAME, check_same_thread=False)
db.row_factory = sqlite3.Row


def db_execute(query, params=(), fetchone=False, fetchall=False):
    cur = db.cursor()
    cur.execute(query, params)
    db.commit()
    if fetchone:
        return cur.fetchone()
    if fetchall:
        return cur.fetchall()
    return None


def init_db():
    db_execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            balance REAL DEFAULT 0,
            files_today INTEGER DEFAULT 0,
            last_file_date TEXT,
            referred_by INTEGER DEFAULT NULL,
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS buttons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            url TEXT NOT NULL
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS withdrawals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            amount REAL,
            fee REAL,
            receive_amount REAL,
            method TEXT,
            account TEXT,
            status TEXT DEFAULT 'pending',
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            file_id TEXT,
            emails_json TEXT,
            accepted_count INTEGER DEFAULT 0,
            rejected_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'pending',
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)


# =========================================================
# SYSTEM & MESSAGE HELPERS
# =========================================================

DEFAULT_MESSAGES = {
    "rules": (
        "আসসালামু আলাইকুম, {name}! 🌸\n\n"
        "💙 আপনাকে স্বাগতম আমাদের Gmail Sell Bot-এ!\n"
        "এখানে আপনি আপনার তৈরি করা Valid Gmail Account সেল/সাবমিট করতে পারবেন।\n\n"
        "📌 গুরুত্বপূর্ণ নিয়মাবলি:\n"
        "🔹 প্রতিদিন সর্বোচ্চ ৫টি Gmail Account সাবমিট করতে পারবেন।\n"
        "🔹 প্রতিটি ভ্যালিড Gmail অ্যাকাউন্টের মূল্য ৳২৫ টাকা। 💰\n"
        "🔹 প্রতিটি Gmail অবশ্যই Valid এবং ব্যবহারযোগ্য হতে হবে।\n"
        "🔹 Gmail সাবমিট করার পর সর্বোচ্চ ২৪ থেকে ৪৮ ঘণ্টার মধ্যে আপনার Gmail রিসিভ করা হবে।\n"
        "🔹 Gmail রিসিভ হওয়ার পর অ্যাডমিন আপনাকে মেসেজের মাধ্যমে জানিয়ে দেবে যে কয়টি জিমেইল রিসিভ হয়েছে।\n"
        "🔹 অ্যাডমিন রিসিভ করার সময় থেকে পরবর্তী ২৪ থেকে ৪৮ ঘণ্টা Gmail-এর স্ট্যাটাস পর্যবেক্ষণ করা হবে।\n"
        "🔹 ২৪ থেকে ৪৮ ঘণ্টা পর যে Gmail Accountগুলো ঠিক থাকবে/নষ্ট হবে না, সেই Gmailগুলোর টাকা (প্রতিটিতে ৳২৫) আপনার Balance-এ যোগ হয়ে যাবে।\n"
        "🔹 আর যে Gmail Accountগুলো নষ্ট বা বাতিল হয়ে যাবে, সেগুলোর জন্য টাকা যোগ হবে না এবং সেই Gmailগুলো আপনাকে ফেরত দেওয়া হবে।\n\n"
        "⚠️ দয়া করে শুধু Valid Gmail Account সাবমিট করুন এবং উপরের নিয়মগুলো মেনে চলুন।\n\n"
        "💚 ধন্যবাদ আমাদের সাথে থাকার জন্য।\n"
        "সুন্দর ও নিরাপদ লেনদেনের শুভকামনা!"
    ),
    "sell": (
        "📤 আপনার ফ্রেশ জিমেইল সম্বলিত এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।\n\n"
        "💰 প্রতি ভ্যালিড জিমেইল রেট: ৳২৫ BDT\n"
        "⚠️ সতর্কতা: এক ফাইলে সর্বোচ্চ ৫টি @gmail.com থাকতে হবে। অন্য কোনো মেইল গ্রহণযোগ্য নয়।"
    ),
    "support": "যে কোনো সমস্যা বা সহযোগিতার জন্য সরাসরি সাপোর্টে যোগাযোগ করুন:",
    "maintenance": (
        "⚠️ **বট আপডেটের কাজ চলছে!** 🛠\n\n"
        "সম্মানিত ইউজার, আমাদের সিস্টেমে জরুরি আপডেটের কাজ চলছে। "
        "সাময়িকভাবে বটের কার্যক্রম স্থগিত রয়েছে। কাজ সম্পন্ন হওয়ামাত্রই বট পুনরায় চালু হবে।"
    ),
    "referral": (
        "👥 রেফার করে ইনকাম করুন!\n\n"
        "প্রতি সফল রেফারে আপনি পাবেন ৳{bonus} BDT বোনাস।\n"
        "আপনার রেফারেল লিংকটি শেয়ার করুন:\n\n"
        "🔗 `{link}`"
    )
}

MESSAGE_NAMES = {
    "rules": "📜 Rules & Welcome (24-48 Hours)",
    "sell": "📤 Sell Gmail Instruction",
    "referral": "👥 Referral Message",
    "support": "📞 Support Text",
    "maintenance": "🛠 Maintenance Notice"
}


def get_admin_id():
    res = db_execute("SELECT value FROM settings WHERE key='admin_id'", fetchone=True)
    if res:
        try:
            return int(res["value"])
        except ValueError:
            return None
    return None


def set_admin_id(user_id):
    db_execute("""
        INSERT INTO settings (key, value) VALUES ('admin_id', ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (str(user_id),))


def is_maintenance_mode():
    res = db_execute("SELECT value FROM settings WHERE key='maintenance'", fetchone=True)
    return res and res["value"] == "1"


def toggle_maintenance_mode():
    current = is_maintenance_mode()
    new_val = "0" if current else "1"
    db_execute("""
        INSERT INTO settings (key, value) VALUES ('maintenance', ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (new_val,))
    return new_val == "1"


def get_custom_msg(key):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"msg_{key}",), fetchone=True)
    if res and res["value"]:
        return res["value"]
    return DEFAULT_MESSAGES.get(key, "")


def set_custom_msg(key, text):
    db_execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (f"msg_{key}", text))


def delete_custom_msg(key):
    db_execute("DELETE FROM settings WHERE key=?", (f"msg_{key}",))


def get_button_title(btn_key, default_title):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"btn_{btn_key}",), fetchone=True)
    return res["value"] if res else default_title


def set_button_title(btn_key, new_title):
    db_execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (f"btn_{btn_key}", new_title))


def clear_junk_cache():
    deleted_files = 0
    if os.path.exists("downloads"):
        for filename in os.listdir("downloads"):
            file_path = os.path.join("downloads", filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                    deleted_files += 1
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
                    deleted_files += 1
            except Exception:
                pass
    db_execute("VACUUM")
    return deleted_files


def add_user(user, referrer_id=None):
    existing = db_execute("SELECT * FROM users WHERE user_id=?", (user.id,), fetchone=True)
    if not existing:
        db_execute("""
            INSERT INTO users (user_id, username, balance, files_today, last_file_date, referred_by, created_at)
            VALUES (?, ?, 0, 0, ?, ?, ?)
        """, (user.id, user.username or "", datetime.now().strftime("%Y-%m-%d"), referrer_id, datetime.now().isoformat()))
        return True
    else:
        db_execute("UPDATE users SET username=? WHERE user_id=?", (user.username or "", user.id))
        return False


def get_user(user_id):
    return db_execute("SELECT * FROM users WHERE user_id=?", (user_id,), fetchone=True)


def get_balance(user_id):
    user = get_user(user_id)
    return float(user["balance"]) if user else 0.0


def update_balance(user_id, amount):
    db_execute("UPDATE users SET balance = balance + ? WHERE user_id=?", (amount, user_id))


def get_referral_count(user_id):
    res = db_execute("SELECT COUNT(*) AS c FROM users WHERE referred_by=?", (user_id,), fetchone=True)
    return res["c"] if res else 0


def get_today_file_count(user_id):
    user = get_user(user_id)
    today = datetime.now().strftime("%Y-%m-%d")
    if not user:
        return 0
    if user["last_file_date"] != today:
        db_execute("UPDATE users SET files_today=0, last_file_date=? WHERE user_id=?", (today, user_id))
        return 0
    return user["files_today"]


def increase_file_count(user_id):
    today = datetime.now().strftime("%Y-%m-%d")
    db_execute("UPDATE users SET files_today = files_today + 1, last_file_date = ? WHERE user_id=?", (today, user_id))


# =========================================================
# VALIDATIONS (BD PHONE, BINANCE UID & STRICT GMAIL)
# =========================================================

BD_PHONE_REGEX = re.compile(r"^01[3-9]\d{8}$")
BINANCE_UID_REGEX = re.compile(r"^\d{7,10}$")
GMAIL_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@gmail\.com$", re.IGNORECASE)


def validate_gmail_file(file_path):
    raw_entries = []
    if file_path.endswith(".csv"):
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                for token in line.strip().split(","):
                    val = token.strip()
                    if val:
                        raw_entries.append(val)
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    raw_entries.append(str(cell).strip())

    valid_gmails = [m for m in raw_entries if GMAIL_REGEX.match(m)]
    unique_gmails = []
    for g in valid_gmails:
        if g not in unique_gmails:
            unique_gmails.append(g)

    return unique_gmails


# =========================================================
# KEYBOARD & MENUS
# =========================================================

def get_bottom_keyboard():
    btn_sell = get_button_title("sell", "📤 SELL FRESH GMAIL ACCOUNT")
    btn_bal = get_button_title("balance", "💰 BALANCE")
    btn_wd = get_button_title("withdraw", "💸 WITHDRAW")
    btn_ref = get_button_title("referral", "👥 REFERRAL")
    btn_rules = get_button_title("rules", "📜 RULES")
    btn_sup = get_button_title("support", "📞 SUPPORT")

    keyboard = [
        [KeyboardButton(btn_sell)],
        [KeyboardButton(btn_bal), KeyboardButton(btn_wd)],
        [KeyboardButton(btn_ref), KeyboardButton(btn_rules)],
        [KeyboardButton(btn_sup)]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


async def show_main_menu(update, context):
    user = update.effective_user
    add_user(user)
    balance_bdt = get_balance(user.id)
    balance_usdt = balance_bdt / USDT_RATE

    name = user.first_name or "User"
    welcome_msg = get_custom_msg("rules").replace("{name}", name)

    balance_text = (
        f"\n\n═══════════════════\n"
        f"💰 Available Balance:\n"
        f"🔹 ৳{balance_bdt:.2f} BDT\n"
        f"🔹 ${balance_usdt:.2f} USDT\n"
        f"═══════════════════\n\n"
        f"নিচের বাটনগুলো চেপে অপশন নির্বাচন করুন:"
    )

    full_text = welcome_msg + balance_text

    buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
    inline_kb = None
    if buttons:
        kb_list = [[InlineKeyboardButton(b["title"], url=b["url"])] for b in buttons]
        inline_kb = InlineKeyboardMarkup(kb_list)

    if update.callback_query:
        await update.callback_query.message.reply_text(full_text, reply_markup=get_bottom_keyboard())
        if inline_kb:
            await update.callback_query.message.reply_text("Important Links:", reply_markup=inline_kb)
    else:
        await update.message.reply_text(full_text, reply_markup=get_bottom_keyboard())
        if inline_kb:
            await update.message.reply_text("Important Links:", reply_markup=inline_kb)


# =========================================================
# COMMAND HANDLERS
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    admin_id = get_admin_id()

    if is_maintenance_mode() and user.id != admin_id:
        await update.message.reply_text(get_custom_msg("maintenance"))
        return

    context.user_data.clear()

    referrer_id = None
    if context.args and len(context.args) > 0:
        try:
            potential_ref = int(context.args[0])
            if potential_ref != user.id and get_user(potential_ref):
                referrer_id = potential_ref
        except ValueError:
            referrer_id = None

    is_new = add_user(user, referrer_id)

    if is_new and referrer_id:
        update_balance(referrer_id, REFERRAL_BONUS)
        ref_new_bdt = get_balance(referrer_id)
        ref_new_usdt = ref_new_bdt / USDT_RATE
        try:
            await context.bot.send_message(
                chat_id=referrer_id,
                text=(
                    f"🎉 অভিনন্দন! আপনার একটি রেফার সফল হয়েছে!\n\n"
                    f"👤 নতুন ইউজার: @{user.username or 'No username'} ({user.id})\n"
                    f"🎁 রেফারেল বোনাস: +৳{REFERRAL_BONUS:.2f} BDT\n"
                    f"💰 বর্তমান ব্যালেন্স: ৳{ref_new_bdt:.2f} BDT (${ref_new_usdt:.2f} USDT)"
                )
            )
        except Exception:
            pass

    await show_main_menu(update, context)


async def set_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/setadmin <secret_key>`")
        return

    if context.args[0] == ADMIN_SECRET_KEY:
        set_admin_id(user_id)
        await update.message.reply_text(f"✅ আপনার আইডি ({user_id}) অ্যাডমিন হিসেবে সেট করা হয়েছে।\n/admin লিখে প্যানেল খুলুন।")
    else:
        await update.message.reply_text("❌ পাসওয়ার্ড ভুল!")


async def admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if not admin_id or update.effective_user.id != admin_id:
        await update.message.reply_text("❌ আপনি অ্যাডমিন নন। প্রথমে `/setadmin <key>` কমান্ড পাঠান।")
        return

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👨‍💼 Open Admin Panel", callback_data="admin_panel")]
    ])
    await update.message.reply_text("🔐 Admin Access Granted.", reply_markup=keyboard)


async def add_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != get_admin_id():
        return
    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার নিয়ম: `/addbalance <user_id> <amount>`")
        return
    try:
        t_id, amt = int(context.args[0]), float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ সংখ্যায় আইডি ও টাকার পরিমাণ দিন।")
        return

    if not get_user(t_id):
        await update.message.reply_text("❌ ইউজার পাওয়া যায়নি।")
        return

    update_balance(t_id, amt)
    new_bdt = get_balance(t_id)
    await update.message.reply_text(f"✅ ব্যালেন্স যোগ হয়েছে! ইউজার: {t_id}, বর্তমান ব্যালেন্স: ৳{new_bdt:.2f}")
    try:
        await context.bot.send_message(chat_id=t_id, text=f"🎁 আপনার অ্যাকাউন্টে ৳{amt:.2f} BDT যোগ করা হয়েছে!")
    except Exception:
        pass


async def cut_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != get_admin_id():
        return
    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার নিয়ম: `/cutbalance <user_id> <amount>`")
        return
    try:
        t_id, amt = int(context.args[0]), float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ সংখ্যায় আইডি ও টাকার পরিমাণ দিন।")
        return

    if not get_user(t_id):
        await update.message.reply_text("❌ ইউজার পাওয়া যায়নি।")
        return

    update_balance(t_id, -amt)
    new_bdt = get_balance(t_id)
    await update.message.reply_text(f"✂️ ব্যালেন্স কেটে নেওয়া হয়েছে! ইউজার: {t_id}, বর্তমান ব্যালেন্স: ৳{new_bdt:.2f}")
    try:
        await context.bot.send_message(chat_id=t_id, text=f"⚠️ আপনার অ্যাকাউন্ট থেকে ৳{amt:.2f} BDT কেটে নেওয়া হয়েছে।")
    except Exception:
        pass


async def reset_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != get_admin_id():
        return
    if len(context.args) < 1:
        await update.message.reply_text("ব্যবহার নিয়ম: `/resetbalance <user_id>`")
        return
    try:
        t_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ সংখ্যায় আইডি দিন।")
        return

    if not get_user(t_id):
        await update.message.reply_text("❌ ইউজার পাওয়া যায়নি।")
        return

    db_execute("UPDATE users SET balance=0 WHERE user_id=?", (t_id,))
    await update.message.reply_text(f"🔄 ব্যালেন্স রিসেট সফল! ইউজার: {t_id}, ব্যালেন্স: ৳0.00")


# =========================================================
# ADMIN CONTROL PANEL
# =========================================================

async def show_admin_panel(query):
    m_status = "🔴 বট বন্ধ (Maintenance ON)" if is_maintenance_mode() else "🟢 বট চালু (Active)"

    keyboard = [
        [InlineKeyboardButton(f"🛠 স্ট্যাটাস: {m_status}", callback_data="toggle_maintenance")],
        [InlineKeyboardButton("💬 বাটন মেসেজ কন্ট্রোল (Add/Edit/Delete)", callback_data="admin_messages_menu")],
        [InlineKeyboardButton("✏️ মেনু বাটন নাম এডিট", callback_data="admin_edit_buttons_menu")],
        [InlineKeyboardButton("🧹 ক্লিয়ার ক্যাশ / আবর্জনা মুছুন", callback_data="admin_clear_cache")],
        [InlineKeyboardButton("➕ নতুন লিংক বাটন", callback_data="admin_add_button"),
         InlineKeyboardButton("🗑 লিংক বাটন মুছুন", callback_data="admin_remove_button")],
        [InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast"),
         InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message")],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")],
        [InlineKeyboardButton("⬅️ প্যানেল বন্ধ করুন", callback_data="admin_close")]
    ]
    await query.edit_message_text("👨‍💼 ADMIN CONTROL PANEL\n\nএকটি অপশন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))


# =========================================================
# CALLBACK HANDLER & WORKFLOW
# =========================================================

admin_sub_selections = {}


def build_submission_admin_keyboard(sub_id, emails, selected_indices):
    num_buttons = []
    for idx, _ in enumerate(emails):
        icon = "✅" if idx in selected_indices else "❌"
        num_buttons.append(InlineKeyboardButton(f"[{idx+1}] {icon}", callback_data=f"sub_toggle:{sub_id}:{idx}"))

    chunked = [num_buttons[i:i + 3] for i in range(0, len(num_buttons), 3)]

    action_rows = [
        [InlineKeyboardButton(f"🚀 কনফার্ম করুন ({len(selected_indices)} রিসিভ / {len(emails)-len(selected_indices)} রিজেক্ট)", callback_data=f"sub_commit:{sub_id}")],
        [InlineKeyboardButton("✅ সবগুলো রিসিভ করুন", callback_data=f"sub_accept_all:{sub_id}"),
         InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_reject_all:{sub_id}")]
    ]

    return InlineKeyboardMarkup(chunked + action_rows)


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    admin_id = get_admin_id()
    data = query.data

    if data == "admin_panel":
        if user_id != admin_id:
            return
        await show_admin_panel(query)

    elif data == "admin_close":
        await query.message.delete()

    elif data == "toggle_maintenance":
        if user_id != admin_id:
            return
        new_state = toggle_maintenance_mode()
        msg = "বট বন্ধ করা হয়েছে (ইউজাররা রক্ষণাবেক্ষণ নোটিশ পাবে)!" if new_state else "বট সফলভাবে চালু করা হয়েছে!"
        await query.answer(msg, show_alert=True)
        await show_admin_panel(query)

    elif data == "admin_clear_cache":
        if user_id != admin_id:
            return
        deleted = clear_junk_cache()
        await query.answer(f"ক্যাশ ক্লিয়ার সম্পন্ন! {deleted}টি ফাইল মুছে ফেলা হয়েছে।", show_alert=True)
        await show_admin_panel(query)

    elif data == "admin_messages_menu":
        if user_id != admin_id:
            return
        keyboard = [
            [InlineKeyboardButton(f"{v}", callback_data=f"msg_view:{k}")] for k, v in MESSAGE_NAMES.items()
        ]
        keyboard.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text("💬 কোন বাটনের মেসেজ কাস্টমাইজ বা এডিট/ডিলিট করতে চান? নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("msg_view:"):
        if user_id != admin_id:
            return
        msg_key = data.split(":")[1]
        current_text = get_custom_msg(msg_key)
        name = MESSAGE_NAMES.get(msg_key, msg_key)
        preview = current_text if len(current_text) <= 300 else current_text[:300] + "..."

        keyboard = [
            [InlineKeyboardButton("✍️ নতুন মেসেজ লিখুন / Edit", callback_data=f"msg_set:{msg_key}")],
            [InlineKeyboardButton("🗑 মেসেজ রিসেট / মুছুন (Default)", callback_data=f"msg_del:{msg_key}")],
            [InlineKeyboardButton("⬅️ Back to Messages", callback_data="admin_messages_menu")]
        ]
        await query.edit_message_text(
            f"📌 **বাটন:** {name}\n\n"
            f"📝 **বর্তমান মেসেজ প্রিভিউ:**\n`{preview}`\n\n"
            f"নিচের অপশন বেছে নিন:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )

    elif data.startswith("msg_set:"):
        if user_id != admin_id:
            return
        msg_key = data.split(":")[1]
        context.user_data["editing_msg_key"] = msg_key
        context.user_data["state"] = "waiting_new_msg_text"
        tips = ""
        if msg_key == "rules":
            tips = "\n💡 টিপস: লেখার মধ্যে `{name}` দিলে সেখানে ইউজারের নাম বসবে।"
        elif msg_key == "referral":
            tips = "\n💡 টিপস: লেখার মধ্যে `{link}` দিলে রেফারেল লিংক এবং `{bonus}` দিলে বোনাসের টাকা বসবে।"
        await query.edit_message_text(f"✍️ **{MESSAGE_NAMES.get(msg_key)}** বাটনের জন্য নতুন মেসেজটি লিখে পাঠিয়ে দিন:{tips}\n\n(ক্যানসেল করতে /cancel পাঠান)")

    elif data.startswith("msg_del:"):
        if user_id != admin_id:
            return
        msg_key = data.split(":")[1]
        delete_custom_msg(msg_key)
        await query.answer("মেসেজ সফলভাবে রিসেট করা হয়েছে!", show_alert=True)
        await button_handler(update, context)

    elif data == "admin_edit_buttons_menu":
        if user_id != admin_id:
            return
        keyboard = [
            [InlineKeyboardButton("✏️ Sell বাটন", callback_data="edit_btn:sell"),
             InlineKeyboardButton("✏️ Balance বাটন", callback_data="edit_btn:balance")],
            [InlineKeyboardButton("✏️ Withdraw বাটন", callback_data="edit_btn:withdraw"),
             InlineKeyboardButton("✏️ Referral বাটন", callback_data="edit_btn:referral")],
            [InlineKeyboardButton("✏️ Rules বাটন", callback_data="edit_btn:rules"),
             InlineKeyboardButton("✏️ Support বাটন", callback_data="edit_btn:support")],
            [InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]
        ]
        await query.edit_message_text("✏️ কোন বাটনটির নাম পরিবর্তন করতে চান?", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("edit_btn:"):
        if user_id != admin_id:
            return
        btn_key = data.split(":")[1]
        context.user_data["editing_btn_key"] = btn_key
        context.user_data["state"] = "waiting_new_button_title"
        await query.edit_message_text(f"বাটনটির নতুন নাম লিখে পাঠান:\n(কী: {btn_key})")

    elif data == "admin_add_button":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_add_button_title"
        await query.edit_message_text("➕ Add Link Button\n\nপ্রথমে বাটনের নাম লিখে পাঠান:")

    elif data == "admin_remove_button":
        if user_id != admin_id:
            return
        buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
        if not buttons:
            await query.edit_message_text("কোনো কাস্টম লিংক বাটন নেই।", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")]]))
            return

        keyboard = [[InlineKeyboardButton(f"❌ {b['title']}", callback_data=f"delete_button:{b['id']}")] for b in buttons]
        keyboard.append([InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")])
        await query.edit_message_text("মুছে ফেলতে বাটন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_broadcast":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 ব্রডকাস্ট করার মেসেজটি লিখে পাঠান:")

    elif data == "admin_single_message":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ যাকে মেসেজ পাঠাবেন তার টেলিগ্রাম আইডি পাঠান:")

    elif data == "admin_stats":
        if user_id != admin_id:
            return
        users = db_execute("SELECT COUNT(*) AS c FROM users", fetchone=True)["c"]
        total_balance = db_execute("SELECT COALESCE(SUM(balance),0) AS total FROM users", fetchone=True)["total"]
        total_usdt = float(total_balance) / USDT_RATE
        pending = db_execute("SELECT COUNT(*) AS c FROM withdrawals WHERE status='pending'", fetchone=True)["c"]
        pending_files = db_execute("SELECT COUNT(*) AS c FROM submissions WHERE status='pending'", fetchone=True)["c"]

        await query.edit_message_text(
            f"📊 Admin Statistics\n\n"
            f"👤 মোট ইউজার: {users}\n"
            f"💰 মোট ব্যালেন্স: ৳{float(total_balance):.2f} BDT (${total_usdt:.2f} USDT)\n"
            f"💸 পেন্ডিং উইথড্র: {pending}\n"
            f"📁 পেন্ডিং ফাইল: {pending_files}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]])
        )

    # -------------------------------------------------------------
    # SUBMISSION LOGIC: ১, ২, ৩, ৪, ৫ নির্বাচন ও প্রতিটিতে ৳২৫ হিসাব
    # -------------------------------------------------------------
    elif data.startswith("sub_open:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        emails = json.loads(sub["emails_json"])
        if sub_id not in admin_sub_selections:
            admin_sub_selections[sub_id] = set(range(len(emails)))

        kb = build_submission_admin_keyboard(sub_id, emails, admin_sub_selections[sub_id])
        await query.edit_message_caption(
            caption=f"📋 সাবমিশন (#{sub_id})\nইউজার ID: {sub['user_id']}\n\nনিচের সংখ্যা বাটনগুলো চেপে যে যে জিমেইল ঠিক আছে সেগুলো ✅ (রিসিভ) রাখুন এবং সমস্যাযুক্ত জিমেইল ❌ (বাতিল) করে কনফার্ম করুন:",
            reply_markup=kb
        )

    elif data.startswith("sub_toggle:"):
        if user_id != admin_id:
            return
        parts = data.split(":")
        sub_id = int(parts[1])
        idx = int(parts[2])

        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            return

        emails = json.loads(sub["emails_json"])
        if sub_id not in admin_sub_selections:
            admin_sub_selections[sub_id] = set(range(len(emails)))

        if idx in admin_sub_selections[sub_id]:
            admin_sub_selections[sub_id].remove(idx)
        else:
            admin_sub_selections[sub_id].add(idx)

        kb = build_submission_admin_keyboard(sub_id, emails, admin_sub_selections[sub_id])
        await query.edit_message_reply_markup(reply_markup=kb)

    elif data.startswith("sub_commit:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        emails = json.loads(sub["emails_json"])
        selected = admin_sub_selections.get(sub_id, set(range(len(emails))))

        accepted_emails = [emails[i] for i in sorted(list(selected))]
        rejected_emails = [emails[i] for i in range(len(emails)) if i not in selected]

        total_accepted = len(accepted_emails)
        total_rejected = len(rejected_emails)
        reward = total_accepted * REWARD_PER_EMAIL

        db_execute("""
            UPDATE submissions 
            SET status='completed', accepted_count=?, rejected_count=?
            WHERE id=?
        """, (total_accepted, total_rejected, sub_id))

        if total_accepted > 0:
            update_balance(sub["user_id"], reward)

        admin_sub_selections.pop(sub_id, None)

        await query.edit_message_caption(
            caption=(
                f"✅ সাবমিশন (#{sub_id}) সম্পন্ন হয়েছে!\n\n"
                f"🔹 রিসিভ করা হয়েছে: {total_accepted} টি\n"
                f"🔹 রিজেক্ট করা হয়েছে: {total_rejected} টি\n"
                f"💰 প্রতি জিমেইলে ৳২৫ হারে যোগ হয়েছে: ৳{reward:.2f} BDT"
            )
        )

        user_msg = (
            f"🔔 **আপনার সাবমিট করা জিমেইল অ্যাকাউন্টের আপডেট!**\n\n"
            f"🔹 মোট জমা দেওয়া ছিল: {len(emails)} টি\n"
            f"✅ অ্যাডমিন সফলভাবে রিসিভ করেছে: {total_accepted} টি\n"
            f"💰 প্রতি অ্যাকাউন্টে ৳২৫ হারে যোগ হয়েছে: ৳{reward:.2f} BDT\n\n"
            f"📌 **জরুরি নিয়ম:**\n"
            f"রিসিভ করা জিমেইলগুলো আগামী **২৪ থেকে ৪৮ ঘণ্টা** পর্যবেক্ষণে রাখা হবে। "
            f"২৪-৪৮ ঘণ্টা পর যে জিমেইলগুলো নষ্ট হবে না, সেগুলোর টাকা স্থায়ী থাকবে।"
        )
        try:
            await context.bot.send_message(chat_id=sub["user_id"], text=user_msg, parse_mode="Markdown")
        except Exception:
            pass

        if total_rejected > 0:
            reject_file_content = "সমস্যাযুক্ত ও বাতিল হওয়া জিমেইল তালিকা:\n" + "\n".join(rejected_emails)
            bio = io.BytesIO(reject_file_content.encode('utf-8'))
            bio.name = f"rejected_gmails_sub_{sub_id}.txt"

            reject_caption = (
                f"⚠️ **{total_rejected}টি জিমেইল অ্যাকাউন্টে সমস্যা থাকায় বাতিল করা হয়েছে!**\n\n"
                f"বাতিলকৃত জিমেইলগুলো ফেরত দেওয়া হলো (সংযুক্ত ফাইলে দেখুন)। "
                f"এগুলোর জন্য কোনো টাকা যোগ হয়নি। দয়া করে ফ্রেশ ও ভ্যালিড জিমেইল সাবমিট করুন।"
            )
            try:
                await context.bot.send_document(
                    chat_id=sub["user_id"],
                    document=bio,
                    caption=reject_caption,
                    parse_mode="Markdown"
                )
            except Exception:
                pass

    elif data.startswith("sub_accept_all:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            return

        emails = json.loads(sub["emails_json"])
        total_accepted = len(emails)
        reward = total_accepted * REWARD_PER_EMAIL

        db_execute("UPDATE submissions SET status='completed', accepted_count=?, rejected_count=0 WHERE id=?", (total_accepted, sub_id))
        update_balance(sub["user_id"], reward)
        admin_sub_selections.pop(sub_id, None)

        await query.edit_message_caption(caption=f"✅ সবকটি ({total_accepted}টি) জিমেইল রিসিভ করা হয়েছে! ইউজারকে ৳{reward:.2f} BDT যোগ করে দেওয়া হয়েছে।")
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=(
                    f"🎉 অভিনন্দন! আপনার দেওয়া সবকটি ({total_accepted}টি) জিমেইল সফলভাবে রিসিভ করা হয়েছে।\n"
                    f"💰 ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT।\n\n"
                    f"পরবর্তী ২৪ থেকে ৪৮ ঘণ্টা জিমেইলগুলো পর্যবেক্ষণ করা হবে।"
                )
            )
        except Exception:
            pass

    elif data.startswith("sub_reject_all:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            return

        emails = json.loads(sub["emails_json"])
        db_execute("UPDATE submissions SET status='rejected', accepted_count=0, rejected_count=? WHERE id=?", (len(emails), sub_id))
        admin_sub_selections.pop(sub_id, None)

        await query.edit_message_caption(caption=f"❌ পুরো ফাইল ({len(emails)}টি জিমেইল) বাতিল করা হয়েছে।")

        reject_file_content = "সম্পূর্ণ বাতিলকৃত জিমেইল তালিকা:\n" + "\n".join(emails)
        bio = io.BytesIO(reject_file_content.encode('utf-8'))
        bio.name = f"all_rejected_sub_{sub_id}.txt"

        try:
            await context.bot.send_document(
                chat_id=sub["user_id"],
                document=bio,
                caption="❌ আপনার ফাইলের সবকটি জিমেইলে সমস্যা থাকায় সম্পূর্ণ ফাইলটি বাতিল ও ফেরত দেওয়া হলো।"
            )
        except Exception:
            pass

    # Withdrawals
    elif data.startswith("approve_withdraw:"):
        if user_id != admin_id:
            return
        w_id = int(data.split(":")[1])
        w = db_execute("SELECT * FROM withdrawals WHERE id=?", (w_id,), fetchone=True)
        if not w or w["status"] != "pending":
            return
        bal = get_balance(w["user_id"])
        if bal < float(w["amount"]):
            db_execute("UPDATE withdrawals SET status='rejected' WHERE id=?", (w_id,))
            await query.edit_message_text("❌ পর্যাপ্ত ব্যালেন্স না থাকায় রিজেক্ট করা হলো।")
            return
        update_balance(w["user_id"], -float(w["amount"]))
        db_execute("UPDATE withdrawals SET status='approved' WHERE id=?", (w_id,))
        rec_bdt = float(w['receive_amount'])
        rec_usdt = rec_bdt / USDT_RATE
        await query.edit_message_text(f"✅ Withdrawal Approved! পেমেন্ট: ৳{rec_bdt:.2f} (${rec_usdt:.2f})")
        try:
            msg = f"✅ উইথড্র সফল হয়েছে!\n💸 পেয়েছেন: ৳{rec_bdt:.2f} BDT"
            if w["method"] == "Binance":
                msg += f" (${rec_usdt:.2f} USDT)"
            msg += f"\n💳 মেথড: {w['method']}\n📌 একাউন্ট/UID: {w['account']}"
            await context.bot.send_message(chat_id=w["user_id"], text=msg)
        except Exception:
            pass

    elif data.startswith("reject_withdraw:"):
        if user_id != admin_id:
            return
        w_id = int(data.split(":")[1])
        w = db_execute("SELECT * FROM withdrawals WHERE id=?", (w_id,), fetchone=True)
        if not w or w["status"] != "pending":
            return
        db_execute("UPDATE withdrawals SET status='rejected' WHERE id=?", (w_id,))
        await query.edit_message_text(f"❌ Withdrawal #{w_id} Rejected.")
        try:
            await context.bot.send_message(chat_id=w["user_id"], text=f"❌ উইথড্রয়াল রিকোয়েস্ট (৳{w['amount']:.2f}) বাতিল করা হয়েছে।")
        except Exception:
            pass

    elif data.startswith("delete_button:"):
        if user_id != admin_id:
            return
        b_id = int(data.split(":")[1])
        db_execute("DELETE FROM buttons WHERE id=?", (b_id,))
        await query.answer("বাটন মুছে ফেলা হয়েছে।")
        await show_admin_panel(query)


# =========================================================
# MESSAGE HANDLER
# =========================================================

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not user:
        return

    admin_id = get_admin_id()

    if is_maintenance_mode() and user.id != admin_id:
        await update.message.reply_text(get_custom_msg("maintenance"))
        return

    add_user(user)
    text = update.message.text.strip() if update.message.text else ""
    state = context.user_data.get("state")

    btn_sell = get_button_title("sell", "📤 SELL FRESH GMAIL ACCOUNT")
    btn_bal = get_button_title("balance", "💰 BALANCE")
    btn_wd = get_button_title("withdraw", "💸 WITHDRAW")
    btn_ref = get_button_title("referral", "👥 REFERRAL")
    btn_rules = get_button_title("rules", "📜 RULES")
    btn_sup = get_button_title("support", "📞 SUPPORT")

    if text == "/cancel":
        context.user_data.clear()
        await update.message.reply_text("বাতিল করা হয়েছে।")
        return

    if user.id == admin_id and state == "waiting_new_msg_text":
        key = context.user_data.get("editing_msg_key")
        if text:
            set_custom_msg(key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! '{MESSAGE_NAMES.get(key)}' মেসেজটি সেভ হয়েছে।")
        return

    if user.id == admin_id and state == "waiting_new_button_title":
        btn_key = context.user_data.get("editing_btn_key")
        if text:
            set_button_title(btn_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ বাটন নাম আপডেট হয়েছে:\n{text}")
        return

    if text == btn_sell:
        count = get_today_file_count(user.id)
        if count >= DAILY_FILE_LIMIT:
            await update.message.reply_text(f"❌ দৈনিক লিমিট শেষ ({DAILY_FILE_LIMIT}/{DAILY_FILE_LIMIT})।")
            return

        context.user_data["state"] = "waiting_file"
        sell_msg = get_custom_msg("sell")
        await update.message.reply_text(
            f"{sell_msg}\n\n"
            f"📊 আজকের সাবমিশন: {count}/{DAILY_FILE_LIMIT}\n"
            f"💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{REWARD_PER_EMAIL:.2f} BDT"
        )
        return

    elif text == btn_bal:
        b_bdt = get_balance(user.id)
        b_usdt = b_bdt / USDT_RATE
        refs = get_referral_count(user.id)
        await update.message.reply_text(
            f"💰 YOUR ACCOUNT BALANCE\n\n"
            f"🔹 BDT Balance: ৳{b_bdt:.2f} BDT\n"
            f"🔹 USDT Balance: ${b_usdt:.2f} USDT\n"
            f"👥 Total Referrals: {refs} users"
        )
        return

    elif text == btn_wd:
        bal = get_balance(user.id)
        bal_usdt = bal / USDT_RATE
        if bal < MIN_WITHDRAW:
            await update.message.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{MIN_WITHDRAW} BDT। আপনার আছে: ৳{bal:.2f} BDT")
            return

        context.user_data["state"] = "withdraw_amount"
        await update.message.reply_text(
            f"💸 WITHDRAWAL SYSTEM\n\n"
            f"💰 ব্যালেন্স: ৳{bal:.2f} BDT (${bal_usdt:.2f} USDT)\n"
            f"🔹 সর্বনিম্ন উইথড্র: ৳{MIN_WITHDRAW} BDT\n"
            f"🔹 ফি: {WITHDRAW_FEE_PERCENT}%\n"
            f"🔹 ডলার রেট: $1 = ৳{USDT_RATE:.2f}\n\n"
            f"কত টাকা উইথড্র করতে চান? সংখ্যায় লিখে পাঠান:"
        )
        return

    elif text == btn_ref:
        b_info = await context.bot.get_me()
        ref_link = f"https://t.me/{b_info.username}?start={user.id}"
        refs = get_referral_count(user.id)
        earned = refs * REFERRAL_BONUS

        ref_template = get_custom_msg("referral")
        custom_ref = ref_template.replace("{link}", ref_link).replace("{bonus}", str(REFERRAL_BONUS))

        msg = (
            f"{custom_ref}\n\n"
            f"📊 আপনার রেফারেল পরিসংখ্যান:\n"
            f"🔹 মোট রেফার: {refs} জন\n"
            f"🔹 মোট আয়: ৳{earned:.2f} BDT"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")
        return

    elif text == btn_rules:
        name = user.first_name or "User"
        rules = get_custom_msg("rules").replace("{name}", name)
        await update.message.reply_text(rules)
        return

    elif text == btn_sup:
        sup_text = get_custom_msg("support")
        keyboard = InlineKeyboardMarkup([[InlineKeyboardButton("💬 মেসেজ পাঠান", url="https://t.me/Talha_juba098")]])
        await update.message.reply_text(sup_text, reply_markup=keyboard)
        return

    # Admin actions
    if user.id == admin_id and state == "admin_add_button_title":
        context.user_data["button_title"] = text
        context.user_data["state"] = "admin_add_button_url"
        await update.message.reply_text("বাটনের URL দিন:")
        return

    if user.id == admin_id and state == "admin_add_button_url":
        if not text.startswith(("http://", "https://", "tg://")):
            await update.message.reply_text("❌ সঠিক URL দিন।")
            return
        t = context.user_data.get("button_title")
        db_execute("INSERT INTO buttons (title, url) VALUES (?, ?)", (t, text))
        context.user_data.clear()
        await update.message.reply_text(f"✅ বাটন যুক্ত হয়েছে: {t} -> {text}")
        return

    if user.id == admin_id and state == "admin_broadcast":
        users = db_execute("SELECT user_id FROM users", fetchall=True)
        count = 0
        for u in users:
            try:
                await context.bot.send_message(chat_id=u["user_id"], text=text)
                count += 1
            except Exception:
                pass
        context.user_data.clear()
        await update.message.reply_text(f"📢 ব্রডকাস্ট সম্পন্ন: {count}/{len(users)}")
        return

    if user.id == admin_id and state == "admin_single_user":
        try:
            context.user_data["target_user"] = int(text)
            context.user_data["state"] = "admin_single_text"
            await update.message.reply_text("মেসেজটি লিখে পাঠান:")
        except ValueError:
            await update.message.reply_text("❌ সঠিক নিউমেরিক আইডি দিন।")
        return

    if user.id == admin_id and state == "admin_single_text":
        t_id = context.user_data.get("target_user")
        try:
            await context.bot.send_message(chat_id=t_id, text=text)
            await update.message.reply_text("✅ মেসেজ পাঠানো হয়েছে।")
        except Exception as e:
            await update.message.reply_text(f"❌ ব্যর্থ হয়েছে: {e}")
        context.user_data.clear()
        return

    # Withdraw input
    if state == "withdraw_amount":
        try:
            amount = float(text)
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যা লিখুন।")
            return

        if amount < MIN_WITHDRAW:
            await update.message.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{MIN_WITHDRAW} BDT।")
            return

        if amount > get_balance(user.id):
            await update.message.reply_text("❌ ব্যালেন্স পর্যাপ্ত নেই।")
            return

        fee = amount * WITHDRAW_FEE_PERCENT / 100
        rec = amount - fee
        rec_usdt = rec / USDT_RATE

        context.user_data.update({"withdraw_amount": amount, "withdraw_fee": fee, "withdraw_receive": rec})
        context.user_data["state"] = "withdraw_method"

        keyboard = [
            [InlineKeyboardButton("💳 bKash", callback_data="method_bkash")],
            [InlineKeyboardButton("💳 Nagad", callback_data="method_nagad")],
            [InlineKeyboardButton("💰 Binance (UID)", callback_data="method_binance")]
        ]
        await update.message.reply_text(
            f"💸 উইথড্র সামারি:\nউত্তোলন: ৳{amount:.2f} BDT\nফি: ৳{fee:.2f} BDT\nপাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\nমেথড নির্বাচন করুন:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    if state == "withdraw_account":
        method = context.user_data.get("withdraw_method")

        if method in ["Bkash", "Nagad"] and not BD_PHONE_REGEX.match(text):
            await update.message.reply_text(f"❌ ভুল {method} নম্বর! ১১ ডিজিটের বাংলাদেশি নম্বর দিন।")
            return
        elif method == "Binance" and not BINANCE_UID_REGEX.match(text):
            await update.message.reply_text("❌ ভুল Binance UID! শুধুমাত্র ৭-১০ ডিজিটের সংখ্যা দিন।")
            return

        amt = context.user_data["withdraw_amount"]
        fee = context.user_data["withdraw_fee"]
        rec = context.user_data["withdraw_receive"]
        rec_usdt = rec / USDT_RATE

        db_execute("""
            INSERT INTO withdrawals (user_id, amount, fee, receive_amount, method, account, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (user.id, amt, fee, rec, method, text, datetime.now().isoformat()))

        w_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        context.user_data.clear()

        await update.message.reply_text(f"✅ উইথড্র রিকোয়েস্ট সফল হয়েছে!\nমেথড: {method}\nঅ্যাকাউন্ট: {text}\nপাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)")

        if admin_id:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")]
            ])
            await context.bot.send_message(
                chat_id=admin_id,
                text=f"🔔 নতুন উইথড্র রিকোয়েস্ট (#{w_id})\nইউজার: {user.id}\nটাকা: ৳{rec:.2f} (${rec_usdt:.2f})\nমেথড: {method}\nঅ্যাকাউন্ট: {text}",
                reply_markup=keyboard
            )
        return

    # File input
    if state == "waiting_file":
        doc = update.message.document
        if not doc:
            await update.message.reply_text("❌ এক্সেল বা সিএসভি ফাইল (.xlsx, .csv) পাঠান।")
            return

        fn = doc.file_name.lower()
        if not (fn.endswith(".xlsx") or fn.endswith(".xls") or fn.endswith(".csv")):
            await update.message.reply_text("❌ শুধুমাত্র Excel বা CSV ফাইল সাপোর্টেড।")
            return

        os.makedirs("downloads", exist_ok=True)
        local_path = os.path.join("downloads", f"{user.id}_{doc.file_name}")
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(local_path)

        try:
            valid_emails = validate_gmail_file(local_path)
        except Exception as e:
            if os.path.exists(local_path):
                os.remove(local_path)
            await update.message.reply_text(f"❌ ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        if os.path.exists(local_path):
            os.remove(local_path)

        if not valid_emails:
            await update.message.reply_text("❌ কোনো ভ্যালিড @gmail.com পাওয়া যায়নি!")
            return

        if len(valid_emails) > MAX_EMAILS_PER_FILE:
            await update.message.reply_text(
                f"❌ ফাইলে {len(valid_emails)}টি জিমেইল রয়েছে।\n"
                f"একটি ফাইলে সর্বোচ্চ {MAX_EMAILS_PER_FILE}টি জিমেইল দেওয়া অনুমোদিত।"
            )
            return

        emails_json = json.dumps(valid_emails)
        db_execute("""
            INSERT INTO submissions (user_id, file_id, emails_json, status, created_at)
            VALUES (?, ?, ?, 'pending', ?)
        """, (user.id, doc.file_id, emails_json, datetime.now().isoformat()))

        sub_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        increase_file_count(user.id)
        context.user_data.clear()

        await update.message.reply_text(
            f"✅ ফাইল সফলভাবে জমা হয়েছে!\n\n"
            f"🔍 মোট ভ্যালিড Gmail: {len(valid_emails)} টি।\n"
            f"⏳ সর্বোচ্চ ২৪ থেকে ৪৮ ঘণ্টার মধ্যে রিসিভ করা হবে। রিসিভ হওয়ার পর প্রতি ভ্যালিড অ্যাকাউন্টে ৳{REWARD_PER_EMAIL:.2f} করে যোগ হবে।"
        )

        if admin_id:
            mail_preview = "\n".join([f"{i+1}. {m}" for i, m in enumerate(valid_emails)])
            admin_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("📥 রিসিভ (Select Accounts)", callback_data=f"sub_open:{sub_id}")],
                [InlineKeyboardButton("✅ সব রিসিভ (All Accept)", callback_data=f"sub_accept_all:{sub_id}"),
                 InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_reject_all:{sub_id}")]
            ])
            await context.bot.send_document(
                chat_id=admin_id,
                document=doc.file_id,
                caption=(
                    f"📁 নতুন জিমেইল সাবমিশন (#{sub_id})\n"
                    f"👤 ইউজার: {user.id}\n"
                    f"✉️ জিমেইল সংখ্যা: {len(valid_emails)} টি\n\n"
                    f"📋 তালিকা:\n{mail_preview}"
                ),
                reply_markup=admin_kb
            )
        return

    await update.message.reply_text("Please use the buttons below.", reply_markup=get_bottom_keyboard())


async def withdrawal_method_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data.startswith("method_"):
        raw_m = query.data.replace("method_", "").lower()
        if raw_m == "bkash":
            method, hint = "Bkash", "১১ ডিজিটের বিকাশ নম্বর দিন:"
        elif raw_m == "nagad":
            method, hint = "Nagad", "১১ ডিজিটের নগদ নম্বর দিন:"
        else:
            method, hint = "Binance", "বাইনান্স ইউআইডি (Binance UID) দিন (৭-১০ ডিজিট):"

        context.user_data["withdraw_method"] = method
        context.user_data["state"] = "withdraw_account"
        await query.edit_message_text(f"💳 নির্বাচিত: {method}\n\n👉 {hint}")


# =========================================================
# MAIN
# =========================================================

def main():
    init_db()

    web_thread = threading.Thread(target=run_web_server, daemon=True)
    web_thread.start()

    application = Application.builder().token(BOT_TOKEN).build()

    # Commands
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("admin", admin))
    application.add_handler(CommandHandler("setadmin", set_admin_cmd))
    application.add_handler(CommandHandler("addbalance", add_balance_cmd))
    application.add_handler(CommandHandler("cutbalance", cut_balance_cmd))
    application.add_handler(CommandHandler("resetbalance", reset_balance_cmd))

    # Handlers
    application.add_handler(CallbackQueryHandler(withdrawal_method_handler, pattern=r"^method_"))
    application.add_handler(CallbackQueryHandler(button_handler))
    application.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, message_handler))

    print("Telegram bot is running...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
