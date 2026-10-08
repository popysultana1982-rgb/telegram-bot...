import io
import json
import os
import re
import shutil
import sqlite3
import threading
from datetime import datetime, timedelta
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
DEFAULT_REWARD_PER_EMAIL = 25.0
DAILY_FILE_LIMIT = 5
MAX_EMAILS_PER_FILE = 5
DUPLICATE_CHECK_DAYS = 3

REFERRAL_BONUS = 5.0
USDT_RATE = 124.0

DB_NAME = "bot.db"
PORT = int(os.environ.get("PORT", 10000))

# =========================================================
# CHAT IDS & LINKS
# =========================================================
FORCE_GROUP_CHAT_ID = -1004471047712
FORCE_GROUP_LINK = "https://t.me/+rVP6CkmqrnFlNzA1"

FORCE_CHANNEL_CHAT_ID = -1003991468184
FORCE_CHANNEL_LINK = "https://t.me/fast_payment_proof_chanel"


# =========================================================
# RENDER WEB SERVER
# =========================================================

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Telegram bot is running.")

    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, format, *args):
        return


def run_web_server():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), HealthHandler)
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
            in_review_json TEXT DEFAULT '[]',
            accepted_count INTEGER DEFAULT 0,
            rejected_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'pending',
            handled_by TEXT DEFAULT NULL,
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS submitted_emails (
            email TEXT,
            submitted_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS helpers (
            user_id INTEGER PRIMARY KEY,
            admin_name TEXT DEFAULT 'Admin',
            added_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)


# =========================================================
# ADMIN & ROLE MANAGEMENT
# =========================================================

def get_owner_id():
    res = db_execute("SELECT value FROM settings WHERE key='owner_id'", fetchone=True)
    if res:
        try:
            return int(res["value"])
        except ValueError:
            return None
    return None


def set_owner_id(user_id):
    db_execute("""
        INSERT INTO settings (key, value) VALUES ('owner_id', ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (str(user_id),))


def is_admin(user_id):
    if user_id == get_owner_id():
        return True
    helper = db_execute("SELECT user_id FROM helpers WHERE user_id=?", (user_id,), fetchone=True)
    return helper is not None


def get_admin_name(user_id):
    if user_id == get_owner_id():
        return "👑 Owner"
    helper = db_execute("SELECT admin_name FROM helpers WHERE user_id=?", (user_id,), fetchone=True)
    if helper and helper["admin_name"]:
        return helper["admin_name"]
    return f"Admin ({user_id})"


def add_helper(user_id, name="Admin"):
    db_execute("""
        INSERT INTO helpers (user_id, admin_name, added_at) VALUES (?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET admin_name=excluded.admin_name
    """, (user_id, name, datetime.now().isoformat()))


def remove_helper(user_id):
    db_execute("DELETE FROM helpers WHERE user_id=?", (user_id,))


def set_helper_name(user_id, name):
    db_execute("UPDATE helpers SET admin_name=? WHERE user_id=?", (name, user_id))


def get_all_admins():
    owner = get_owner_id()
    helpers = db_execute("SELECT user_id, admin_name FROM helpers", fetchall=True)
    return owner, helpers


def get_email_rate():
    res = db_execute("SELECT value FROM settings WHERE key='email_rate'", fetchone=True)
    if res:
        try:
            return float(res["value"])
        except ValueError:
            return DEFAULT_REWARD_PER_EMAIL
    return DEFAULT_REWARD_PER_EMAIL


def set_email_rate(rate):
    db_execute("""
        INSERT INTO settings (key, value) VALUES ('email_rate', ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (str(rate),))


# =========================================================
# EXCEL GENERATOR HELPER
# =========================================================

def create_rejected_excel(emails):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rejected"
    ws.append(["Index", "Rejected Gmail Account", "Reason"])
    for i, em in enumerate(emails, 1):
        ws.append([i, em, "Login issue or invalid on check"])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# =========================================================
# SYSTEM & MESSAGE HELPERS
# =========================================================

DEFAULT_MESSAGES = {
    "welcome": (
        "আসসালামু আলাইকুম, {name}! 🌸\n\n"
        "💙 আপনাকে স্বাগতম আমাদের Gmail Sell Bot-এ!\n"
        "এখানে আপনি আপনার তৈরি করা Valid Gmail Account সেল/সাবমিট করে নিরাপদভাবে টাকা ইনকাম করতে পারবেন।\n\n"
        "📌 **৩টি ধাপে সাবমিশন যাচাই পদ্ধতি:**\n"
        "🔹 ধাপ ১: ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক রিসিভ করবেন এবং যে জিমেইলগুলো লগইন করা যায় সেগুলো পর্যালোচনায় রাখবেন। যেগুলোতে লগইন সমস্যা থাকবে সেগুলো ১ম ধাপেই বাতিল ও এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
        "🔹 ধাপ ২: পর্যালোচনায় রাখা জিমেইলগুলো পরবর্তী ২৪ থেকে ৪৮ ঘণ্টা অ্যাডমিনের পর্যবেক্ষণে থাকবে। এই সময়ে ব্যালেন্স যোগ হবে না।\n"
        "🔹 ধাপ ৩: ২৪ থেকে ৪৮ ঘণ্টা পর যে জিমেইলগুলো ঠিক থাকবে, সেগুলোর প্রতিটির জন্য ৳{rate} টাকা আপনার ব্যালেন্সে যোগ হয়ে যাবে! 💰\n"
        "🔹 আর শেষ ধাপে কোনো জিমেইল নষ্ট হলে তা আপনাকে এক্সেল ফাইলে ফেরত দেওয়া হবে।\n\n"
        "⚠️ প্রতিদিন সর্বোচ্চ ৫টি ফাইল এবং প্রতি ফাইলে সর্বোচ্চ ৫টি ভ্যালিড Gmail দিতে পারবেন। একবার সাবমিট করা জিমেইল আগামী ৩ দিনের মধ্যে পুনরায় দেওয়া যাবে না।\n\n"
        "👇 **বটের কার্যক্রম শুরু করতে নিচের চ্যানেল ও গ্রুপে জয়েন হয়ে ভেরিফাই বাটনে ক্লিক করুন:**"
    ),
    "rules": (
        "📜 **আমাদের জিমেইল সাবমিশন নিয়মাবলী:**\n\n"
        "🔹 ধাপ ১: ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক বাছাই করবেন।\n"
        "🔹 ধাপ ২: পর্যালোচনায় থাকা জিমেইলগুলো ২৪ থেকে ৪৮ ঘণ্টা পর্যবেক্ষণে থাকবে।\n"
        "🔹 ধাপ ৩: অক্ষত প্রতি জিমেইলে ৳{rate} টাকা পাবেন এবং নষ্ট জিমেইল এক্সেল ফাইলে ফেরত যাবে।\n"
        "⚠️ এক ফাইলে সর্বোচ্চ ৫টি @gmail.com এবং দিনে সর্বোচ্চ ৫টি ফাইল সাবমিট করা যাবে।"
    ),
    "sell": (
        "📤 আপনার ফ্রেশ জিমেইল সম্বলিত এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।\n\n"
        "💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate} BDT\n"
        "⚠️ এক ফাইলে সর্বোচ্চ ৫টি @gmail.com থাকতে হবে। বিগত ৩ দিনের মধ্যে জমা দেওয়া কোনো মেইল গ্রহণ করা হবে না।"
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
    "welcome": "🎉 Welcome Message & Rules",
    "rules": "📜 Rules Message",
    "sell": "📤 Sell Gmail Instruction",
    "referral": "👥 Referral Message",
    "support": "📞 Support Text",
    "maintenance": "🛠 Maintenance Notice"
}


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
    rate = get_email_rate()
    text = res["value"] if (res and res["value"]) else DEFAULT_MESSAGES.get(key, "")
    return text.replace("{rate}", f"{rate:.2f}")


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
# FORCE JOIN CHAT-ID VERIFICATION SYSTEM
# =========================================================

async def is_user_joined_all(bot, user_id):
    if is_admin(user_id):
        return True

    try:
        member = await bot.get_chat_member(chat_id=FORCE_CHANNEL_CHAT_ID, user_id=user_id)
        if member.status in ["left", "kicked"]:
            return False
    except Exception:
        return False

    try:
        member = await bot.get_chat_member(chat_id=FORCE_GROUP_CHAT_ID, user_id=user_id)
        if member.status in ["left", "kicked"]:
            return False
    except Exception:
        return False

    return True


def get_force_join_markup():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📢 ১. পেমেন্ট প্রুফ চ্যানেল", url=FORCE_CHANNEL_LINK)],
        [InlineKeyboardButton("👥 ২. অফিসিয়াল গ্রুপ", url=FORCE_GROUP_LINK)],
        [InlineKeyboardButton("✅ ৩. ভেরিফাই করুন (Verify)", callback_data="check_joined")]
    ])


# =========================================================
# VALIDATIONS (DUPLICATES, PHONE, BINANCE & GMAIL)
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

    valid_gmails = [m.lower() for m in raw_entries if GMAIL_REGEX.match(m)]
    unique_gmails = []
    for g in valid_gmails:
        if g not in unique_gmails:
            unique_gmails.append(g)

    return unique_gmails


def check_recent_duplicate_emails(emails):
    cutoff_date = (datetime.now() - timedelta(days=DUPLICATE_CHECK_DAYS)).isoformat()
    placeholders = ",".join(["?"] * len(emails))
    query = f"""
        SELECT email FROM submitted_emails 
        WHERE email IN ({placeholders}) AND submitted_at >= ?
    """
    params = list(emails) + [cutoff_date]
    rows = db_execute(query, params, fetchall=True)
    if rows:
        return [r["email"] for r in rows]
    return []


def save_submitted_emails(emails):
    now_str = datetime.now().isoformat()
    cur = db.cursor()
    cur.executemany("INSERT INTO submitted_emails (email, submitted_at) VALUES (?, ?)", [(e, now_str) for e in emails])
    db.commit()


# =========================================================
# KEYBOARD & MENUS
# =========================================================

def get_bottom_keyboard():
    btn_sell = get_button_title("sell", "📤 SELL FRESH GMAIL ACCOUNT")
    btn_bal = get_button_title("balance", "💰 BALANCE")
    btn_wd = get_button_title("withdraw", "💸 WITHDRAW")
    btn_hist = get_button_title("history", "📜 হিস্ট্রি")
    btn_ref = get_button_title("referral", "👥 REFERRAL")
    btn_rules = get_button_title("rules", "📜 RULES")
    btn_sup = get_button_title("support", "📞 SUPPORT")

    keyboard = [
        [KeyboardButton(btn_sell)],
        [KeyboardButton(btn_bal), KeyboardButton(btn_wd)],
        [KeyboardButton(btn_hist), KeyboardButton(btn_ref)],
        [KeyboardButton(btn_rules), KeyboardButton(btn_sup)]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


async def show_main_menu(update, context):
    user = update.effective_user
    add_user(user)

    msg_text = "🎉 **ধন্যবাদ! আপনি সফলভাবে যুক্ত আছেন।**\n\nনিচের বাটনগুলো চেপে আপনার কাঙ্ক্ষিত অপশন বেছে নিন:"

    buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
    inline_kb = None
    if buttons:
        kb_list = [[InlineKeyboardButton(b["title"], url=b["url"])] for b in buttons]
        inline_kb = InlineKeyboardMarkup(kb_list)

    if update.callback_query:
        await update.callback_query.message.reply_text(msg_text, reply_markup=get_bottom_keyboard(), parse_mode="Markdown")
        if inline_kb:
            await update.callback_query.message.reply_text("Important Links:", reply_markup=inline_kb)
    else:
        await update.message.reply_text(msg_text, reply_markup=get_bottom_keyboard(), parse_mode="Markdown")
        if inline_kb:
            await update.message.reply_text("Important Links:", reply_markup=inline_kb)


# =========================================================
# COMMAND HANDLERS
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return

    user = update.effective_user
    if is_maintenance_mode() and not is_admin(user.id):
        await update.message.reply_text(get_custom_msg("maintenance"))
        return

    context.user_data.clear()

    # রেফারেল ট্র্যাকিং
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
        try:
            await context.bot.send_message(
                chat_id=referrer_id,
                text=f"🎉 আপনার একটি নতুন রেফার সফল হয়েছে! বোনাস: +৳{REFERRAL_BONUS:.2f} BDT"
            )
        except Exception:
            pass

    joined = await is_user_joined_all(context.bot, user.id)
    if joined:
        await show_main_menu(update, context)
        return

    name = user.first_name or "User"
    welcome_text = get_custom_msg("welcome").replace("{name}", name)

    await update.message.reply_text(
        welcome_text,
        reply_markup=get_force_join_markup(),
        parse_mode="Markdown"
    )


async def set_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/setadmin <secret_key>`")
        return

    if context.args[0] == ADMIN_SECRET_KEY:
        set_owner_id(user_id)
        await update.message.reply_text(f"👑 অভিনন্দন! আপনার আইডি ({user_id}) মূল মালিক (Owner) হিসেবে স্থায়ীভাবে সেট করা হয়েছে।\n/admin লিখে প্যানেল খুলুন।")
    else:
        await update.message.reply_text("❌ পাসওয়ার্ড ভুল!")


async def add_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if user_id != get_owner_id():
        await update.message.reply_text("❌ শুধুমাত্র মূল মালিক (Owner) নতুন অ্যাডমিন যোগ করতে পারেন।")
        return

    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/addadmin <user_id> [Admin_Name]`\nউদাহরণ: `/addadmin 123456789 Admin 1`")
        return

    try:
        target_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন।")
        return

    admin_name = " ".join(context.args[1:]) if len(context.args) > 1 else "Admin"
    add_helper(target_id, admin_name)
    await update.message.reply_text(f"✅ নতুন সহযোগী অ্যাডমিন যোগ করা হয়েছে!\n👤 নাম: **{admin_name}**\n🆔 আইডি: `{target_id}`", parse_mode="Markdown")


async def remove_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if user_id != get_owner_id():
        await update.message.reply_text("❌ শুধুমাত্র মূল মালিক (Owner) অ্যাডমিন বাদ দিতে পারেন।")
        return

    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/removeadmin <user_id>`")
        return

    try:
        target_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন।")
        return

    remove_helper(target_id)
    await update.message.reply_text(f"🗑 সহযোগী অ্যাডমিন ক্ষমতা বাতিল করা হয়েছে!\nআইডি: `{target_id}`", parse_mode="Markdown")


async def admin_list_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
        return

    owner, helpers = get_all_admins()
    msg = f"👑 **মূল মালিক (Owner):** `{owner}`\n\n👥 **সহযোগী অ্যাডমিন তালিকা ({len(helpers)} জন):**\n"
    if helpers:
        for idx, h in enumerate(helpers, 1):
            msg += f"{idx}. **{h['admin_name']}** (`{h['user_id']}`)\n"
    else:
        msg += "কোনো সহযোগী অ্যাডমিন যুক্ত নেই।"

    await update.message.reply_text(msg, parse_mode="Markdown")


async def admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ আপনার অ্যাডমিন অ্যাক্সেস নেই।")
        return

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👨‍💼 Open Admin Panel", callback_data="admin_panel")]
    ])
    await update.message.reply_text("🔐 Admin Access Granted.", reply_markup=keyboard)


async def add_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
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
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
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
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
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
# ADMIN CONTROL PANEL (OWNER EXCLUSIVE BUTTON INCLUDED)
# =========================================================

async def show_admin_panel(query, user_id):
    m_status = "🔴 বট বন্ধ (Maintenance ON)" if is_maintenance_mode() else "🟢 বট চালু (Active)"
    current_rate = get_email_rate()
    is_owner = (user_id == get_owner_id())

    keyboard = [
        [InlineKeyboardButton(f"🛠 স্ট্যাটাস: {m_status}", callback_data="toggle_maintenance")],
        [InlineKeyboardButton(f"💵 জিমেইল রেট এডিট (বর্তমান: ৳{current_rate:.2f})", callback_data="admin_edit_rate")],
        [InlineKeyboardButton("📂 ফাইল হিস্ট্রি ও রিপোর্ট (Users Files)", callback_data="admin_file_history:0")],
        [InlineKeyboardButton("💬 বাটন মেসেজ কন্ট্রোল (Add/Edit/Delete)", callback_data="admin_messages_menu")],
        [InlineKeyboardButton("✏️ মেনু বাটন নাম এডিট", callback_data="admin_edit_buttons_menu")],
        [InlineKeyboardButton("🧹 ক্লিয়ার ক্যাশ / আবর্জনা মুছুন", callback_data="admin_clear_cache")],
        [InlineKeyboardButton("➕ নতুন লিংক বাটন", callback_data="admin_add_button"),
         InlineKeyboardButton("🗑 লিংক বাটন মুছুন", callback_data="admin_remove_button")],
        [InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast"),
         InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message")],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")]
    ]

    # শুধুমাত্র মূল মালিকের (Owner) জন্য স্পেশাল এক্সট্রা বাটন
    if is_owner:
        keyboard.append([InlineKeyboardButton("👥 অ্যাডমিন তালিকা ও নাম এডিট (Owner Only)", callback_data="owner_manage_admins")])

    keyboard.append([InlineKeyboardButton("⬅️ প্যানেল বন্ধ করুন", callback_data="admin_close")])

    role_text = "👑 OWNER CONTROL PANEL" if is_owner else "👨‍💼 ADMIN CONTROL PANEL"
    await query.edit_message_text(f"{role_text}\n\nএকটি অপশন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))


# =========================================================
# CALLBACK HANDLER
# =========================================================

admin_stage1_selections = {}
admin_stage3_selections = {}


def build_stage_keyboard(sub_id, emails, selected_indices, stage_num):
    rate = get_email_rate()
    num_buttons = []
    for idx, _ in enumerate(emails):
        icon = "✅" if idx in selected_indices else "❌"
        num_buttons.append(InlineKeyboardButton(f"[{idx+1}] {icon}", callback_data=f"sub_tg:{stage_num}:{sub_id}:{idx}"))

    chunked = [num_buttons[i:i + 3] for i in range(0, len(num_buttons), 3)]

    if stage_num == 1:
        action_rows = [
            [InlineKeyboardButton(f"🚀 ধাপ ১ নিশ্চিত করুন ({len(selected_indices)}টি পর্যালোচনায় / {len(emails)-len(selected_indices)}টি বাতিল)", callback_data=f"sub_c1:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল (Reject All)", callback_data=f"sub_reject_all:{sub_id}")]
        ]
    else:
        action_rows = [
            [InlineKeyboardButton(f"💰 চূড়ান্ত অনুমোদন দিন ({len(selected_indices)}টি বৈধ - ৳{len(selected_indices)*rate:.2f})", callback_data=f"sub_c3:{sub_id}")],
            [InlineKeyboardButton("✅ সবগুলোই ঠিক আছে", callback_data=f"sub_accept_all_s3:{sub_id}"),
             InlineKeyboardButton("❌ সবগুলোই নষ্ট (বাতিল)", callback_data=f"sub_reject_all_s3:{sub_id}")]
        ]

    return InlineKeyboardMarkup(chunked + action_rows)


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    # ভেরিফাই বাটন
    if data == "check_joined":
        joined = await is_user_joined_all(context.bot, user_id)
        if joined:
            await query.message.delete()
            await show_main_menu(update, context)
        else:
            await query.answer("❌ আপনি এখনো জয়েন করেননি! দুটি গ্রুপ ও চ্যানেলে জয়েন হয়ে আবার ভেরিফাই চাপুন।", show_alert=True)
        return

    # অ্যাডমিন পারমিশন চেক
    if not is_admin(user_id):
        return

    if data == "admin_panel":
        await show_admin_panel(query, user_id)

    elif data == "admin_close":
        await query.message.delete()

    elif data == "toggle_maintenance":
        new_state = toggle_maintenance_mode()
        msg = "বট বন্ধ করা হয়েছে (ইউজাররা রক্ষণাবেক্ষণ নোটিশ পাবে)!" if new_state else "বট সফলভাবে চালু করা হয়েছে!"
        await query.answer(msg, show_alert=True)
        await show_admin_panel(query, user_id)

    # -------------------------------------------------------------
    # ওনার স্পেশাল: অ্যাডমিন ম্যানেজমেন্ট ও নাম এডিট
    # -------------------------------------------------------------
    elif data == "owner_manage_admins":
        if user_id != get_owner_id():
            await query.answer("❌ শুধুমাত্র মূল মালিকের এক্সেস আছে!", show_alert=True)
            return

        owner, helpers = get_all_admins()
        if not helpers:
            await query.edit_message_text(
                "👥 **বর্তমানে কোনো সহযোগী অ্যাডমিন যুক্ত নেই।**\n\nনতুন অ্যাডমিন যোগ করতে চ্যাটে কমান্ড পাঠান:\n`/addadmin <user_id> [Admin_Name]`",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back to Panel", callback_data="admin_panel")]]),
                parse_mode="Markdown"
            )
            return

        kb = []
        for h in helpers:
            kb.append([InlineKeyboardButton(f"✏️ {h['admin_name']} (ID: {h['user_id']})", callback_data=f"edit_admin_name:{h['user_id']}")])

        kb.append([InlineKeyboardButton("⬅️ Back to Panel", callback_data="admin_panel")])
        await query.edit_message_text("👥 **যে অ্যাডমিনের নাম এডিট করতে চান তাকে নির্বাচন করুন:**", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif data.startswith("edit_admin_name:"):
        if user_id != get_owner_id():
            return
        target_helper_id = int(data.split(":")[1])
        context.user_data["target_helper_id"] = target_helper_id
        context.user_data["state"] = "waiting_helper_new_name"

        current_n = get_admin_name(target_helper_id)
        await query.edit_message_text(
            f"👤 **অ্যাডমিন নাম পরিবর্তন**\n\n"
            f"🆔 আইডি: `{target_helper_id}`\n"
            f"বর্তমান নাম: **{current_n}**\n\n"
            f"👉 নতুন যে নাম দিতে চান তা লিখে পাঠান (যেমন: Admin 1, Helper Joy):\n(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_edit_rate":
        context.user_data["state"] = "waiting_new_email_rate"
        current_rate = get_email_rate()
        await query.edit_message_text(
            f"💵 **জিমেইল রেট পরিবর্তন**\n\n"
            f"🔹 বর্তমান প্রতি ভ্যালিড জিমেইল রেট: ৳{current_rate:.2f} BDT\n\n"
            f"নতুন রেট কত টাকা নির্ধারণ করতে চান? সংখ্যায় লিখে পাঠান (যেমন: 25 বা 30):\n(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_clear_cache":
        deleted = clear_junk_cache()
        await query.answer(f"ক্যাশ ক্লিয়ার সম্পন্ন! {deleted}টি ফাইল মুছে ফেলা হয়েছে।", show_alert=True)
        await show_admin_panel(query, user_id)

    elif data == "admin_messages_menu":
        keyboard = [
            [InlineKeyboardButton(f"{v}", callback_data=f"msg_view:{k}")] for k, v in MESSAGE_NAMES.items()
        ]
        keyboard.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text("💬 কোন বাটনের মেসেজ কাস্টমাইজ বা এডিট/ডিলিট করতে চান? নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("msg_view:"):
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
        msg_key = data.split(":")[1]
        context.user_data["editing_msg_key"] = msg_key
        context.user_data["state"] = "waiting_new_msg_text"
        tips = ""
        if msg_key in ["rules", "welcome"]:
            tips = "\n💡 টিপস: লেখার মধ্যে `{name}` দিলে ইউজার নাম এবং `{rate}` দিলে বর্তমান রেট বসবে।"
        elif msg_key == "referral":
            tips = "\n💡 টিপস: লেখার মধ্যে `{link}` দিলে রেফারেল লিংক এবং `{bonus}` দিলে বোনাসের টাকা বসবে।"
        await query.edit_message_text(f"✍️ **{MESSAGE_NAMES.get(msg_key)}** বাটনের জন্য নতুন মেসেজটি লিখে পাঠিয়ে দিন:{tips}\n\n(ক্যানসেল করতে /cancel পাঠান)")

    elif data.startswith("msg_del:"):
        msg_key = data.split(":")[1]
        delete_custom_msg(msg_key)
        await query.answer("মেসেজ সফলভাবে রিসেট করা হয়েছে!", show_alert=True)
        await button_handler(update, context)

    elif data == "admin_edit_buttons_menu":
        keyboard = [
            [InlineKeyboardButton("✏️ Sell বাটন", callback_data="edit_btn:sell"),
             InlineKeyboardButton("✏️ Balance বাটন", callback_data="edit_btn:balance")],
            [InlineKeyboardButton("✏️ Withdraw বাটন", callback_data="edit_btn:withdraw"),
             InlineKeyboardButton("✏️ History বাটন", callback_data="edit_btn:history")],
            [InlineKeyboardButton("✏️ Referral বাটন", callback_data="edit_btn:referral"),
             InlineKeyboardButton("✏️ Rules বাটন", callback_data="edit_btn:rules")],
            [InlineKeyboardButton("✏️ Support বাটন", callback_data="edit_btn:support")],
            [InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]
        ]
        await query.edit_message_text("✏️ কোন বাটনটির নাম পরিবর্তন করতে চান?", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("edit_btn:"):
        btn_key = data.split(":")[1]
        context.user_data["editing_btn_key"] = btn_key
        context.user_data["state"] = "waiting_new_button_title"
        await query.edit_message_text(f"বাটনটির নতুন নাম লিখে পাঠান:\n(কী: {btn_key})")

    elif data == "admin_add_button":
        context.user_data["state"] = "admin_add_button_title"
        await query.edit_message_text("➕ Add Link Button\n\nপ্রথমে বাটনের নাম লিখে পাঠান:")

    elif data == "admin_remove_button":
        buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
        if not buttons:
            await query.edit_message_text("কোনো কাস্টম লিংক বাটন নেই।", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")]]))
            return

        keyboard = [[InlineKeyboardButton(f"❌ {b['title']}", callback_data=f"delete_button:{b['id']}")] for b in buttons]
        keyboard.append([InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")])
        await query.edit_message_text("মুছে ফেলতে বাটন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_broadcast":
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 ব্রডকাস্ট করার মেসেজটি লিখে পাঠান:")

    elif data == "admin_single_message":
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ যাকে মেসেজ পাঠাবেন তার টেলিগ্রাম আইডি পাঠান:")

    elif data == "admin_stats":
        users = db_execute("SELECT COUNT(*) AS c FROM users", fetchone=True)["c"]
        total_balance = db_execute("SELECT COALESCE(SUM(balance),0) AS total FROM users", fetchone=True)["total"]
        total_usdt = float(total_balance) / USDT_RATE
        pending = db_execute("SELECT COUNT(*) AS c FROM withdrawals WHERE status='pending'", fetchone=True)["c"]
        pending_files = db_execute("SELECT COUNT(*) AS c FROM submissions WHERE status IN ('pending', 'stage2_review')", fetchone=True)["c"]

        await query.edit_message_text(
            f"📊 Admin Statistics\n\n"
            f"👤 মোট ইউজার: {users}\n"
            f"💰 মোট ব্যালেন্স: ৳{float(total_balance):.2f} BDT (${total_usdt:.2f} USDT)\n"
            f"💸 পেন্ডিং উইথড্র: {pending}\n"
            f"📁 চলমান ফাইল (পেন্ডিং/পর্যবেক্ষণে): {pending_files}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]])
        )

    # -------------------------------------------------------------
    # ফাইল হিস্ট্রি ও ডিলিট
    # -------------------------------------------------------------
    elif data.startswith("admin_file_history:"):
        page = int(data.split(":")[1])
        limit = 5
        offset = page * limit
        subs = db_execute("SELECT * FROM submissions ORDER BY id DESC LIMIT ? OFFSET ?", (limit, offset), fetchall=True)
        total_subs = db_execute("SELECT COUNT(*) AS c FROM submissions", fetchone=True)["c"]

        if not subs:
            await query.edit_message_text("📂 কোনো সাবমিশন ফাইল পাওয়া যায়নি।", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]]))
            return

        kb = []
        for s in subs:
            emails = json.loads(s["emails_json"])
            st_text = "🟡 নতুন" if s["status"] == "pending" else ("🟠 পর্যবেক্ষণে" if s["status"] == "stage2_review" else ("🟢 সম্পন্ন" if s["status"] == "completed" else "🔴 বাতিল"))
            btn_text = f"#{s['id']} | User: {s['user_id']} | {len(emails)}টি ({st_text})"
            kb.append([InlineKeyboardButton(btn_text, callback_data=f"adm_view_sub:{s['id']}")])

        nav_buttons = []
        if page > 0:
            nav_buttons.append(InlineKeyboardButton("⬅️ Previous", callback_data=f"admin_file_history:{page-1}"))
        if offset + limit < total_subs:
            nav_buttons.append(InlineKeyboardButton("Next ➡️", callback_data=f"admin_file_history:{page+1}"))

        if nav_buttons:
            kb.append(nav_buttons)
        kb.append([InlineKeyboardButton("⬅️ Back to Panel", callback_data="admin_panel")])

        await query.edit_message_text(f"📂 সকল ইউজারদের ফাইল তালিকা (মোট: {total_subs} টি):\nফাইল দেখতে ক্লিক করুন:", reply_markup=InlineKeyboardMarkup(kb))

    elif data.startswith("adm_view_sub:"):
        sub_id = int(data.split(":")[1])
        s = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not s:
            await query.answer("ফাইল পাওয়া যায়নি!")
            return

        emails = json.loads(s["emails_json"])
        review_emails = json.loads(s["in_review_json"]) if s["in_review_json"] else []
        mail_list = "\n".join([f"{i+1}. {m}" for i, m in enumerate(emails)])
        rate = get_email_rate()

        st_map = {
            "pending": "🟡 ধাপ ১: নতুন পেন্ডিং",
            "stage2_review": "🟠 ধাপ ২: ২৪-৪৮ ঘণ্টা পর্যবেক্ষণে",
            "completed": "🟢 ধাপ ৩: চূড়ান্ত অনুমোদিত (সম্পন্ন)",
            "rejected": "🔴 সম্পূর্ণ বাতিল"
        }

        handled_info = f"\n🔒 রিসিভ করেছেন: **{s['handled_by']}**\n" if s["handled_by"] else ""

        info_text = (
            f"📄 সাবমিশন বিস্তারিত (#{s['id']})\n\n"
            f"👤 ইউজার আইডি: `{s['user_id']}`\n"
            f"📅 তারিখ: {s['created_at'][:19]}\n"
            f"⚡ স্ট্যাটাস: {st_map.get(s['status'], s['status'])}{handled_info}\n"
            f"🔹 মূল জিমেইল: {len(emails)} টি\n"
            f"🔹 পর্যালোচনায় নেওয়া হয়েছিল: {len(review_emails)} টি\n"
            f"🔹 চূড়ান্ত অনুমোদিত: {s['accepted_count']} টি (৳{s['accepted_count']*rate:.2f})\n\n"
            f"📋 জিমেইল তালিকা:\n{mail_list}"
        )

        action_kb = [
            [InlineKeyboardButton("📥 ডাউনলোড এক্সেল ফাইল (.xlsx)", callback_data=f"adm_dl_excel:{sub_id}")],
            [InlineKeyboardButton("🗑 এই ফাইলটি ডিলিট করুন", callback_data=f"adm_del_sub:{sub_id}")],
            [InlineKeyboardButton("⬅️ ফাইল তালিকায় ফিরুন", callback_data="admin_file_history:0")]
        ]
        await query.edit_message_text(info_text, reply_markup=InlineKeyboardMarkup(action_kb), parse_mode="Markdown")

    elif data.startswith("adm_del_sub:"):
        sub_id = int(data.split(":")[1])
        db_execute("DELETE FROM submissions WHERE id=?", (sub_id,))
        await query.answer("ফাইলটি সফলভাবে ডিলিট করা হয়েছে!", show_alert=True)
        query.data = "admin_file_history:0"
        await button_handler(update, context)

    elif data.startswith("adm_dl_excel:"):
        sub_id = int(data.split(":")[1])
        s = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not s:
            await query.answer("ফাইল পাওয়া যায়নি!")
            return

        emails = json.loads(s["emails_json"])
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Submission_{sub_id}"
        ws.append(["Index", "Gmail Account", "Submission ID", "User ID"])

        for i, em in enumerate(emails, 1):
            ws.append([i, em, sub_id, s["user_id"]])

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        bio.name = f"user_{s['user_id']}_sub_{sub_id}.xlsx"

        await context.bot.send_document(
            chat_id=user_id,
            document=bio,
            caption=f"📂 ইউজার `{s['user_id']}` এর সাবমিশন (#{sub_id}) এক্সেল ফাইল প্রস্তুত।"
        )
        await query.answer("এক্সেল ফাইল পাঠানো হয়েছে!")

    # -------------------------------------------------------------
    # ইউজার হিস্ট্রি
    # -------------------------------------------------------------
    elif data == "user_hist_live":
        subs = db_execute(
            "SELECT * FROM submissions WHERE user_id=? AND status IN ('pending', 'stage2_review') ORDER BY id DESC", 
            (user_id,), 
            fetchall=True
        )
        if not subs:
            back_kb = InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ হিস্ট্রি মেনুতে ফিরুন", callback_data="user_hist_menu")]])
            await query.edit_message_text("⏳ **বর্তমানে আপনার কোনো ফাইল প্রসেসিং বা পর্যবেক্ষণে নেই।**\n\nনতুন ফাইল জমা দিলে তার লাইভ স্ট্যাটাস এখানে দেখা যাবে।", reply_markup=back_kb, parse_mode="Markdown")
            return

        hist_msg = "⏳ **আপনার লাইভ ও চলমান ফাইলগুলোর স্ট্যাটাস:**\n\n"
        for s in subs:
            emails = json.loads(s["emails_json"])
            if s["status"] == "pending":
                st_text = "🟡 নতুন জমা (অ্যাডমিন পর্যালোচনার অপেক্ষায়)"
            else:
                review_count = len(json.loads(s["in_review_json"])) if s["in_review_json"] else 0
                st_text = f"🟠 ২৪-৪৮ ঘণ্টা পর্যবেক্ষণে রয়েছে ({review_count}টি জিমেইল)"

            hist_msg += (
                f"📁 **ফাইল ID: #{s['id']}**\n"
                f"📅 তারিখ: {s['created_at'][:10]}\n"
                f"✉️ মোট জিমেইল: {len(emails)} টি\n"
                f"⚡ স্ট্যাটাস: {st_text}\n"
                f"-----------------------------\n"
            )
        back_kb = InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ হিস্ট্রি মেনুতে ফিরুন", callback_data="user_hist_menu")]])
        await query.edit_message_text(hist_msg, reply_markup=back_kb, parse_mode="Markdown")

    elif data == "user_hist_all":
        subs = db_execute("SELECT * FROM submissions WHERE user_id=? ORDER BY id DESC LIMIT 15", (user_id,), fetchall=True)
        if not subs:
            back_kb = InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ হিস্ট্রি মেনুতে ফিরুন", callback_data="user_hist_menu")]])
            await query.edit_message_text("📜 আপনি এখনও কোনো ফাইল সাবমিট করেননি।", reply_markup=back_kb)
            return

        hist_msg = "📜 **আপনার অতীতের সকল সাবমিশন হিস্ট্রি (সর্বশেষ ১৫টি):**\n\n"
        for s in subs:
            emails = json.loads(s["emails_json"])
            if s["status"] == "pending":
                st_text = "🟡 নতুন জমা (পর্যালোচনার অপেক্ষায়)"
            elif s["status"] == "stage2_review":
                st_text = "🟠 ২৪-৪৮ ঘণ্টা পর্যবেক্ষণে"
            elif s["status"] == "completed":
                st_text = f"🟢 সম্পন্ন (অনুমোদিত: {s['accepted_count']}টি | বাতিল: {s['rejected_count']}টি)"
            else:
                st_text = f"🔴 সম্পূর্ণ বাতিল ({s['rejected_count']}টি নষ্ট)"

            hist_msg += (
                f"📁 **ফাইল ID: #{s['id']}** | {s['created_at'][:10]}\n"
                f"✉️ মোট: {len(emails)}টি | ফলাফল: {st_text}\n"
                f"-----------------------------\n"
            )
        back_kb = InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ হিস্ট্রি মেনুতে ফিরুন", callback_data="user_hist_menu")]])
        await query.edit_message_text(hist_msg, reply_markup=back_kb, parse_mode="Markdown")

    elif data == "user_hist_menu":
        hist_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("⏳ লাইভ চলমান সাবমিশন (Review)", callback_data="user_hist_live")],
            [InlineKeyboardButton("📜 অতীতের সকল সাবমিশন হিস্ট্রি", callback_data="user_hist_all")]
        ])
        await query.edit_message_text("📜 **আপনার সাবমিশন হিস্ট্রি নির্বাচন করুন:**\n\nনিচের যেকোনো অপশনে ক্লিক করে বিস্তারিত দেখুন:", reply_markup=hist_kb, parse_mode="Markdown")

    # -------------------------------------------------------------
    # ধাপ ১: রিসিভ ও লাইভ অ্যাডমিন নোটিফিকেশন সিস্টেম
    # -------------------------------------------------------------
    elif data.startswith("sub_s1_open:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("ফাইলটি অলরেডি প্রসেস করা হয়েছে!", show_alert=True)
            return

        # যিনি রিসিভ করলেন তার নাম নেওয়া
        rec_name = get_admin_name(user_id)

        # চেক করা অন্য কেউ অলরেডি হাত দিয়েছে কি না
        if sub["handled_by"] and sub["handled_by"] != rec_name:
            await query.answer(f"⚠️ এই ফাইলটি অলরেডি {sub['handled_by']} রিসিভ করেছেন!", show_alert=True)
            return

        # ডাটাবেজে রেকর্ড করা যে ইনি রিসিভ করলেন
        db_execute("UPDATE submissions SET handled_by=? WHERE id=?", (rec_name, sub_id))

        # অন্য সকল অ্যাডমিনকে সাথে সাথে জানিয়ে দেওয়া
        owner, helpers = get_all_admins()
        all_admin_ids = [owner] + [h["user_id"] for h in helpers] if owner else [h["user_id"] for h in helpers]

        for a_id in all_admin_ids:
            if a_id != user_id:
                try:
                    await context.bot.send_message(
                        chat_id=a_id,
                        text=f"📢 **টিম আপডেট (#Submission_{sub_id}):**\nফাইলটি রিসিভ করে প্রাথমিক যাচাই শুরু করেছেন: **{rec_name}**",
                        parse_mode="Markdown"
                    )
                except Exception:
                    pass

        emails = json.loads(sub["emails_json"])
        if sub_id not in admin_stage1_selections:
            admin_stage1_selections[sub_id] = set(range(len(emails)))

        kb = build_stage_keyboard(sub_id, emails, admin_stage1_selections[sub_id], stage_num=1)
        await query.edit_message_caption(
            caption=(
                f"📥 **ধাপ ১: প্রাথমিক বাছাই (#{sub_id})**\n"
                f"👤 ইউজার: {sub['user_id']}\n"
                f"🔒 রিসিভ করেছেন: **{rec_name}**\n\n"
                f"যেসব জিমেইল লগইন করা যায় সেগুলো ✅ এবং লগইন সমস্যাযুক্ত জিমেইল ❌ করে নিচে কনফার্ম করুন:"
            ),
            reply_markup=kb,
            parse_mode="Markdown"
        )

    elif data.startswith("sub_tg:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        idx = int(parts[3])

        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub:
            return

        if stage_num == 1:
            emails = json.loads(sub["emails_json"])
            if sub_id not in admin_stage1_selections:
                admin_stage1_selections[sub_id] = set(range(len(emails)))
            if idx in admin_stage1_selections[sub_id]:
                admin_stage1_selections[sub_id].remove(idx)
            else:
                admin_stage1_selections[sub_id].add(idx)
            kb = build_stage_keyboard(sub_id, emails, admin_stage1_selections[sub_id], stage_num=1)
        else:
            emails = json.loads(sub["in_review_json"])
            if sub_id not in admin_stage3_selections:
                admin_stage3_selections[sub_id] = set(range(len(emails)))
            if idx in admin_stage3_selections[sub_id]:
                admin_stage3_selections[sub_id].remove(idx)
            else:
                admin_stage3_selections[sub_id].add(idx)
            kb = build_stage_keyboard(sub_id, emails, admin_stage3_selections[sub_id], stage_num=3)

        await query.edit_message_reply_markup(reply_markup=kb)

    elif data.startswith("sub_c1:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            return

        emails = json.loads(sub["emails_json"])
        selected = admin_stage1_selections.get(sub_id, set(range(len(emails))))

        in_review_emails = [emails[i] for i in sorted(list(selected))]
        s1_rejected = [emails[i] for i in range(len(emails)) if i not in selected]

        if not in_review_emails:
            await query.answer("কমপক্ষে একটি জিমেইল পর্যালোচনায় রাখতে হবে!", show_alert=True)
            return

        db_execute("""
            UPDATE submissions 
            SET status='stage2_review', in_review_json=?, rejected_count=?
            WHERE id=?
        """, (json.dumps(in_review_emails), len(s1_rejected), sub_id))

        admin_stage1_selections.pop(sub_id, None)
        rate = get_email_rate()

        next_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton(f"🔍 ধাপ ৩: চূড়ান্ত অনুমোদন দিন ({len(in_review_emails)}টি)", callback_data=f"sub_s3_open:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_reject_all:{sub_id}")]
        ])

        h_by = sub["handled_by"] or get_admin_name(user_id)
        await query.edit_message_caption(
            caption=(
                f"⏳ **ধাপ ১ সম্পন্ন - ধাপ ২ (২৪-৪৮ ঘণ্টা পর্যবেক্ষণ চলছে)** (#{sub_id})\n\n"
                f"👤 ইউজার: {sub['user_id']}\n"
                f"🔒 রিসিভ করেছিলেন: **{h_by}**\n"
                f"🔹 পর্যালোচনায় নেওয়া হয়েছে: {len(in_review_emails)} টি\n"
                f"🔹 লগইন না হওয়ায় বাতিল: {len(s1_rejected)} টি\n\n"
                f"২৪ থেকে ৪৮ ঘণ্টা পর নিচের বাটনে চাপ দিয়ে চূড়ান্ত যাচাই করুন।"
            ),
            reply_markup=next_kb,
            parse_mode="Markdown"
        )

        user_msg = (
            f"📋 **আপনার জিমেইল সাবমিশনের ১ম ধাপের আপডেট!** (#{sub_id})\n\n"
            f"🔹 পর্যালোচনায় রাখা হয়েছে: {len(in_review_emails)} টি\n"
            f"❌ লগইন না হওয়ায় বাতিল: {len(s1_rejected)} টি\n\n"
            f"📌 **জরুরি তথ্য:** পর্যালোচনায় থাকা {len(in_review_emails)}টি জিমেইল আগামী **২৪ থেকে ৪৮ ঘণ্টা** পর্যবেক্ষণে থাকবে। "
            f"পর্যবেক্ষণ শেষে যে জিমেইলগুলো ঠিক থাকবে, প্রতিটির জন্য ৳{rate:.2f} টাকা আপনার ব্যালেন্সে যোগ হবে।"
        )
        try:
            await context.bot.send_message(chat_id=sub["user_id"], text=user_msg)
        except Exception:
            pass

        if s1_rejected:
            bio = create_rejected_excel(s1_rejected)
            bio.name = f"rejected_stage1_{sub_id}.xlsx"
            try:
                await context.bot.send_document(
                    chat_id=sub["user_id"],
                    document=bio,
                    caption=f"⚠️ {len(s1_rejected)}টি জিমেইলে লগইন সমস্যা থাকায় এক্সেল ফাইলে ফেরত দেওয়া হলো।"
                )
            except Exception:
                pass

    elif data.startswith("sub_s3_open:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            await query.answer("Already processed.", show_alert=True)
            return

        review_emails = json.loads(sub["in_review_json"])
        if sub_id not in admin_stage3_selections:
            admin_stage3_selections[sub_id] = set(range(len(review_emails)))

        kb = build_stage_keyboard(sub_id, review_emails, admin_stage3_selections[sub_id], stage_num=3)
        await query.edit_message_caption(
            caption=(
                f"💰 **ধাপ ৩: চূড়ান্ত অনুমোদন (#{sub_id})**\n"
                f"👤 ইউজার: {sub['user_id']}\n"
                f"🔒 দায়িত্বপ্রাপ্ত: **{sub['handled_by'] or 'Admin'}**\n\n"
                f"পর্যালোচনায় থাকা {len(review_emails)}টি জিমেইল চেক করুন। অক্ষত জিমেইল ✅ রাখুন এবং নষ্টগুলো ❌ করে অনুমোদন দিন:"
            ),
            reply_markup=kb,
            parse_mode="Markdown"
        )

    elif data.startswith("sub_c3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            return

        rate = get_email_rate()
        review_emails = json.loads(sub["in_review_json"])
        selected = admin_stage3_selections.get(sub_id, set(range(len(review_emails))))

        final_accepted = [review_emails[i] for i in sorted(list(selected))]
        final_rejected = [review_emails[i] for i in range(len(review_emails)) if i not in selected]

        total_accepted = len(final_accepted)
        total_rejected = sub["rejected_count"] + len(final_rejected)
        reward = total_accepted * rate

        db_execute("""
            UPDATE submissions 
            SET status='completed', accepted_count=?, rejected_count=?
            WHERE id=?
        """, (total_accepted, total_rejected, sub_id))

        if total_accepted > 0:
            update_balance(sub["user_id"], reward)

        admin_stage3_selections.pop(sub_id, None)

        await query.edit_message_caption(
            caption=(
                f"✅ **সাবমিশন চূড়ান্তভাবে সম্পন্ন ও ক্লোজ হয়েছে! (#{sub_id})**\n\n"
                f"🔹 চূড়ান্ত ভ্যালিড: {total_accepted} টি\n"
                f"🔹 মোট বাতিল: {total_rejected} টি\n"
                f"💰 ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT\n"
                f"🔒 সম্পন্ন করেছেন: **{get_admin_name(user_id)}**"
            ),
            parse_mode="Markdown"
        )

        user_msg = (
            f"🎉 **অভিনন্দন! আপনার জিমেইল অ্যাকাউন্টের চূড়ান্ত পর্যবেক্ষণ সম্পন্ন হয়েছে!**\n\n"
            f"🔹 সাবমিশন আইডি: #{sub_id}\n"
            f"✅ সফলভাবে অ্যাপ্রুভ হয়েছে: {total_accepted} টি\n"
            f"💰 প্রতিটিতে ৳{rate:.2f} হারে আপনার ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT!\n"
        )
        if final_rejected:
            user_msg += f"⚠️ পর্যবেক্ষণে {len(final_rejected)}টি নষ্ট হওয়ায় বাতিল করা হয়েছে।"

        try:
            await context.bot.send_message(chat_id=sub["user_id"], text=user_msg)
        except Exception:
            pass

        if final_rejected:
            bio = create_rejected_excel(final_rejected)
            bio.name = f"rejected_stage3_{sub_id}.xlsx"
            try:
                await context.bot.send_document(
                    chat_id=sub["user_id"],
                    document=bio,
                    caption=f"⚠️ {len(final_rejected)}টি জিমেইল নষ্ট হওয়ায় এক্সেল ফাইলে ফেরত দেওয়া হলো।"
                )
            except Exception:
                pass

    elif data.startswith("sub_accept_all_s3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            return

        rate = get_email_rate()
        review_emails = json.loads(sub["in_review_json"])
        total_accepted = len(review_emails)
        reward = total_accepted * rate

        db_execute("UPDATE submissions SET status='completed', accepted_count=? WHERE id=?", (total_accepted, sub_id))
        update_balance(sub["user_id"], reward)
        admin_stage3_selections.pop(sub_id, None)

        await query.edit_message_caption(caption=f"✅ সবকটি ({total_accepted}টি) জিমেইল সফলভাবে অনুমোদিত! যোগ হয়েছে: ৳{reward:.2f} BDT।")
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=f"🎉 আপনার পর্যবেক্ষণে থাকা সবকটি ({total_accepted}টি) জিমেইল সফলভাবে অনুমোদিত হয়েছে!\n💰 ব্যালেন্সে যোগ হয়েছে ৳{reward:.2f} BDT।"
            )
        except Exception:
            pass

    elif data.startswith("sub_reject_all_s3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            return

        review_emails = json.loads(sub["in_review_json"])
        total_rej = sub["rejected_count"] + len(review_emails)
        db_execute("UPDATE submissions SET status='rejected', accepted_count=0, rejected_count=? WHERE id=?", (total_rej, sub_id))
        admin_stage3_selections.pop(sub_id, None)

        await query.edit_message_caption(caption=f"❌ পর্যবেক্ষণে থাকা সবকটি ({len(review_emails)}টি) জিমেইল নষ্ট হওয়ায় বাতিল।")

        bio = create_rejected_excel(review_emails)
        bio.name = f"rejected_all_stage3_{sub_id}.xlsx"
        try:
            await context.bot.send_document(
                chat_id=sub["user_id"],
                document=bio,
                caption="❌ পর্যবেক্ষণে আপনার সবগুলো জিমেইল নষ্ট হওয়ায় এক্সেল ফাইলে ফেরত দেওয়া হলো।"
            )
        except Exception:
            pass

    elif data.startswith("sub_reject_all:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub:
            return

        emails = json.loads(sub["emails_json"])
        db_execute("UPDATE submissions SET status='rejected', accepted_count=0, rejected_count=? WHERE id=?", (len(emails), sub_id))

        await query.edit_message_caption(caption=f"❌ পুরো ফাইল ({len(emails)}টি জিমেইল) বাতিল করা হয়েছে।")
        bio = create_rejected_excel(emails)
        bio.name = f"all_rejected_{sub_id}.xlsx"
        try:
            await context.bot.send_document(chat_id=sub["user_id"], document=bio, caption="❌ আপনার ফাইলের সবগুলো জিমেইল বাতিল করে এক্সেল ফাইলে ফেরত দেওয়া হলো।")
        except Exception:
            pass

    # Withdrawals
    elif data.startswith("approve_withdraw:"):
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
        b_id = int(data.split(":")[1])
        db_execute("DELETE FROM buttons WHERE id=?", (b_id,))
        await query.answer("বাটন মুছে ফেলা হয়েছে।")
        await show_admin_panel(query, user_id)


# =========================================================
# MESSAGE HANDLER (PRIVATE ONLY)
# =========================================================

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return

    user = update.effective_user
    if not user:
        return

    if is_maintenance_mode() and not is_admin(user.id):
        await update.message.reply_text(get_custom_msg("maintenance"))
        return

    add_user(user)
    text = update.message.text.strip() if update.message.text else ""
    state = context.user_data.get("state")
    rate = get_email_rate()

    btn_sell = get_button_title("sell", "📤 SELL FRESH GMAIL ACCOUNT")
    btn_bal = get_button_title("balance", "💰 BALANCE")
    btn_wd = get_button_title("withdraw", "💸 WITHDRAW")
    btn_hist = get_button_title("history", "📜 হিস্ট্রি")
    btn_ref = get_button_title("referral", "👥 REFERRAL")
    btn_rules = get_button_title("rules", "📜 RULES")
    btn_sup = get_button_title("support", "📞 SUPPORT")

    if text == "/cancel":
        context.user_data.clear()
        await update.message.reply_text("বাতিল করা হয়েছে।")
        return

    # ১. সাপোর্ট বাটন
    if text == btn_sup:
        sup_text = get_custom_msg("support")
        keyboard = InlineKeyboardMarkup([[InlineKeyboardButton("💬 মেসেজ পাঠান", url="https://t.me/Talha_juba098")]])
        await update.message.reply_text(sup_text, reply_markup=keyboard)
        return

    # ২. অন্যান্য বাটনের ক্ষেত্রে ফোর্স জয়েন যাচাই
    if not is_admin(user.id):
        joined = await is_user_joined_all(context.bot, user.id)
        if not joined:
            await update.message.reply_text(
                "⚠️ **দুঃখিত!** বটের সুবিধাসমূহ ব্যবহার করতে হলে আপনাকে আমাদের গ্রুপ ও চ্যানেলে জয়েন থাকতে হবে।\n\n"
                "নিচের লিংকে ক্লিক করে যুক্ত হয়ে 'ভেরিফাই করুন' চাপুন:",
                reply_markup=get_force_join_markup(),
                parse_mode="Markdown"
            )
            return

    # ওনার কর্তৃক সহযোগী অ্যাডমিনের নাম পরিবর্তনের টেক্সট ইনপুট
    if user.id == get_owner_id() and state == "waiting_helper_new_name":
        target_id = context.user_data.get("target_helper_id")
        if text:
            set_helper_name(target_id, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! অ্যাডমিন `{target_id}` এর নাম পরিবর্তন করে **{text}** রাখা হয়েছে।", parse_mode="Markdown")
        return

    # Admin Settings Handlers
    if is_admin(user.id) and state == "waiting_new_email_rate":
        try:
            new_r = float(text)
            if new_r <= 0:
                await update.message.reply_text("❌ রেট অবশ্যই শূন্যের চেয়ে বেশি হতে হবে।")
                return
            set_email_rate(new_r)
            context.user_data.clear()
            await update.message.reply_text(f"✅ জিমেইলের নতুন রেট সফলভাবে ৳{new_r:.2f} BDT নির্ধারণ করা হয়েছে!")
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যায় রেট লিখে পাঠান (যেমন: 25 বা 30)।")
        return

    if is_admin(user.id) and state == "waiting_new_msg_text":
        key = context.user_data.get("editing_msg_key")
        if text:
            set_custom_msg(key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! '{MESSAGE_NAMES.get(key)}' মেসেজটি সেভ হয়েছে।")
        return

    if is_admin(user.id) and state == "waiting_new_button_title":
        btn_key = context.user_data.get("editing_btn_key")
        if text:
            set_button_title(btn_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ বাটন নাম আপডেট হয়েছে:\n{text}")
        return

    # ইউজার হিস্ট্রি মেনু বাটন
    if text == btn_hist:
        hist_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("⏳ লাইভ চলমান সাবমিশন (Review)", callback_data="user_hist_live")],
            [InlineKeyboardButton("📜 অতীতের সকল সাবমিশন হিস্ট্রি", callback_data="user_hist_all")]
        ])
        await update.message.reply_text(
            "📜 **আপনার সাবমিশন হিস্ট্রি নির্বাচন করুন:**\n\n"
            "🔹 **লাইভ চলমান সাবমিশন:** বর্তমানে রিভিউ বা ২৪-৪৮ ঘণ্টা পর্যালোচনায় থাকা ফাইল দেখতে পাবেন।\n"
            "🔹 **সকল সাবমিশন হিস্ট্রি:** শুরু থেকে আপনার জমা দেওয়া সকল ফাইলের ফলাফল দেখতে পাবেন।",
            reply_markup=hist_kb,
            parse_mode="Markdown"
        )
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
            f"💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate:.2f} BDT\n"
            f"⚠️ ফাইলে সর্বোচ্চ {MAX_EMAILS_PER_FILE}টি জিমেইল থাকতে পারবে।"
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

    # Admin actions
    if is_admin(user.id) and state == "admin_add_button_title":
        context.user_data["button_title"] = text
        context.user_data["state"] = "admin_add_button_url"
        await update.message.reply_text("বাটনের URL দিন:")
        return

    if is_admin(user.id) and state == "admin_add_button_url":
        if not text.startswith(("http://", "https://", "tg://")):
            await update.message.reply_text("❌ সঠিক URL দিন।")
            return
        t = context.user_data.get("button_title")
        db_execute("INSERT INTO buttons (title, url) VALUES (?, ?)", (t, text))
        context.user_data.clear()
        await update.message.reply_text(f"✅ বাটন যুক্ত হয়েছে: {t} -> {text}")
        return

    if is_admin(user.id) and state == "admin_broadcast":
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

    if is_admin(user.id) and state == "admin_single_user":
        try:
            context.user_data["target_user"] = int(text)
            context.user_data["state"] = "admin_single_text"
            await update.message.reply_text("মেসেজটি লিখে পাঠান:")
        except ValueError:
            await update.message.reply_text("❌ সঠিক নিউমেরিক আইডি দিন।")
        return

    if is_admin(user.id) and state == "admin_single_text":
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

        owner, helpers = get_all_admins()
        all_admin_ids = [owner] + [h["user_id"] for h in helpers] if owner else [h["user_id"] for h in helpers]

        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")]
        ])
        for a_id in all_admin_ids:
            try:
                await context.bot.send_message(
                    chat_id=a_id,
                    text=f"🔔 নতুন উইথড্র রিকোয়েস্ট (#{w_id})\nইউজার: {user.id}\nটাকা: ৳{rec:.2f} (${rec_usdt:.2f})\nমেথড: {method}\nঅ্যাকাউন্ট: {text}",
                    reply_markup=keyboard
                )
            except Exception:
                pass
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

        # ডুপ্লিকেট জিমেইল চেক
        duplicates = check_recent_duplicate_emails(valid_emails)
        if duplicates:
            dup_list = "\n".join([f"• {d}" for d in duplicates])
            await update.message.reply_text(
                f"❌ **ফাইল গ্রহণ করা হয়নি!**\n\n"
                f"নিচের জিমেইলগুলো বিগত ৩ দিনের মধ্যে জমা দেওয়া হয়েছিল:\n{dup_list}\n\n"
                f"⚠️ একবার জমা দেওয়া জিমেইল ৩ দিন অতিবাহিত না হওয়া পর্যন্ত আর সাবমিট করা যাবে না।"
            )
            return

        save_submitted_emails(valid_emails)

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
            f"⏳ অ্যাডমিন প্রাথমিক বাছাইয়ের পর জিমেইলগুলো ২৪ থেকে ৪৮ ঘণ্টা পর্যবেক্ষণ করবেন।"
        )

        owner, helpers = get_all_admins()
        all_admin_ids = [owner] + [h["user_id"] for h in helpers] if owner else [h["user_id"] for h in helpers]

        mail_preview = "\n".join([f"{i+1}. {m}" for i, m in enumerate(valid_emails)])
        admin_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 ধাপ ১: প্রাথমিক বাছাই ও রিসিভ", callback_data=f"sub_s1_open:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_reject_all:{sub_id}")]
        ])

        for a_id in all_admin_ids:
            try:
                await context.bot.send_document(
                    chat_id=a_id,
                    document=doc.file_id,
                    caption=(
                        f"📁 নতুন জিমেইল সাবমিশন (#{sub_id})\n"
                        f"👤 ইউজার: {user.id}\n"
                        f"✉️ জিমেইল সংখ্যা: {len(valid_emails)} টি\n\n"
                        f"📋 তালিকা:\n{mail_preview}"
                    ),
                    reply_markup=admin_kb
                )
            except Exception:
                pass
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

    # Commands (Strictly Private Chat)
    application.add_handler(CommandHandler("start", start, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("admin", admin, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("setadmin", set_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("addadmin", add_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("removeadmin", remove_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("adminlist", admin_list_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("addbalance", add_balance_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("cutbalance", cut_balance_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("resetbalance", reset_balance_cmd, filters=filters.ChatType.PRIVATE))

    # Handlers
    application.add_handler(CallbackQueryHandler(withdrawal_method_handler, pattern=r"^method_"))
    application.add_handler(CallbackQueryHandler(button_handler))
    
    # Message Handler (Strictly Private Chat)
    application.add_handler(MessageHandler(filters.ChatType.PRIVATE & ~filters.COMMAND, message_handler))

    print("Gmail Sell Bot is running strictly in private chat mode...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
