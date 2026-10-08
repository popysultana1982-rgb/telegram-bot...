import asyncio
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
# CONFIG & SECRETS
# =========================================================

BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN environment variable is not set.")

ADMIN_SECRET_KEY = os.getenv("ADMIN_SECRET_KEY", "mysecretadmin123")

DB_NAME = "bot.db"
PORT = int(os.environ.get("PORT", 10000))

# অফিসিয়াল গ্রুপ (যেখানে ব্রডকাস্ট পোস্ট হবে)
FORCE_GROUP_CHAT_ID = -1004471047712
FORCE_GROUP_LINK = "https://t.me/+rVP6CkmqrnFlNzA1"

# পেমেন্ট প্রুফ চ্যানেল (যেখানে ব্রডকাস্ট যাবে না)
FORCE_CHANNEL_CHAT_ID = -1003991468184
FORCE_CHANNEL_LINK = "https://t.me/fast_payment_proof_chanel"

SUPPORT_URL = "https://t.me/Talha_juba098"


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
# DATABASE & SCHEMA
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
            seen_rules INTEGER DEFAULT 0,
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS dynamic_buttons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_id INTEGER DEFAULT 0,
            title TEXT NOT NULL,
            btn_type TEXT NOT NULL,
            content TEXT NOT NULL,
            show_in_reply INTEGER DEFAULT 1
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
        CREATE TABLE IF NOT EXISTS submission_admin_messages (
            sub_id INTEGER,
            admin_id INTEGER,
            message_id INTEGER
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
# CONFIG HELPERS
# =========================================================

CONFIG_DEFAULTS = {
    "min_withdraw": "50.0",
    "withdraw_fee_percent": "4.0",
    "usdt_rate": "124.0",
    "referral_bonus": "5.0",
    "email_rate": "25.0",
    "daily_file_limit": "5",
    "max_emails_per_file": "5",
    "duplicate_check_days": "3",
    "maintenance": "0"
}


def get_config(key, as_type=float):
    res = db_execute("SELECT value FROM settings WHERE key=?", (key,), fetchone=True)
    val = res["value"] if res else CONFIG_DEFAULTS.get(key, "0")
    try:
        return as_type(val)
    except (ValueError, TypeError):
        return as_type(CONFIG_DEFAULTS.get(key, "0"))


def set_config(key, val):
    db_execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (key, str(val)))


# =========================================================
# ADMIN & ROLE MANAGEMENT
# =========================================================

def get_owner_id():
    res = db_execute("SELECT value FROM settings WHERE key='owner_id'", fetchone=True)
    if res and res["value"]:
        try:
            return int(res["value"])
        except ValueError:
            return None
    return None


def set_owner_id(user_id):
    set_config("owner_id", str(user_id))


def is_admin(user_id):
    owner = get_owner_id()
    if owner and int(user_id) == int(owner):
        return True
    helper = db_execute("SELECT user_id FROM helpers WHERE user_id=?", (user_id,), fetchone=True)
    return helper is not None


def get_admin_name(user_id):
    owner = get_owner_id()
    if owner and int(user_id) == int(owner):
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


async def notify_all_admins(context, message, exclude_user_id=None):
    """টিমের সকল এডমিন ও ওনারের ইনবক্সে ১০০% নির্ভরযোগ্যভাবে লাইভ আপডেট পাঠানো"""
    owner, helpers = get_all_admins()
    admin_list = []
    
    if owner:
        try:
            admin_list.append(int(owner))
        except (ValueError, TypeError):
            pass

    for h in helpers:
        try:
            admin_list.append(int(h["user_id"]))
        except (ValueError, TypeError):
            pass

    unique_admins = set(admin_list)
    ex_id = int(exclude_user_id) if exclude_user_id else None

    for a_id in unique_admins:
        if ex_id and a_id == ex_id:
            continue
        try:
            await context.bot.send_message(
                chat_id=a_id, 
                text=message, 
                parse_mode="Markdown"
            )
        except Exception as e:
            print(f"Failed to send admin notify to {a_id}: {e}")


def is_maintenance_mode():
    return get_config("maintenance", int) == 1


def toggle_maintenance_mode():
    current = is_maintenance_mode()
    set_config("maintenance", 0 if current else 1)
    return not current


# =========================================================
# EXCEL GENERATOR HELPER
# =========================================================

def create_rejected_excel(rejected_items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rejected"
    ws.append(["Index", "Rejected Gmail Account", "Reason"])
    for i, item in enumerate(rejected_items, 1):
        if isinstance(item, tuple):
            em, r_text = item
        else:
            em, r_text = item, "Invalid on check"
        ws.append([i, em, r_text])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# =========================================================
# MESSAGES & BUTTON TITLES
# =========================================================

DEFAULT_MESSAGES = {
    "welcome": (
        "আসসালামু আলাইকুম, {name}! 🌸\n\n"
        "💙 আপনাকে স্বাগতম আমাদের Gmail Sell Bot-এ!\n"
        "এখানে আপনি আপনার তৈরি করা Valid Gmail Account সেল/সাবমিট করে নিরাপদভাবে টাকা ইনকাম করতে পারবেন।\n\n"
        "📜 **আমাদের জিমেইল সাবমিশন ও কাজের নিয়মাবলী:**\n"
        "১. ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক বাছাই করবেন এবং লগইন করা যায় এমন মেইলগুলো পর্যালোচনায় রাখবেন।\n"
        "২. প্রাথমিক বাছায়ে যেসব জিমেইলে সমস্যা থাকবে, সেগুলো কারণসহ সাথে সাথে বাতিল করে এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
        "৩. পর্যালোচনায় রাখা জিমেইলগুলো পরবর্তী ২৪ থেকে ৪৮ ঘণ্টা অ্যাডমিনের পর্যবেক্ষণে থাকবে।\n"
        "৪. ২৪ থেকে ৪৮ ঘণ্টার পর্যবেক্ষণ সময়ে ব্যালেন্স যোগ হবে না।\n"
        "৫. পর্যবেক্ষণ শেষে যেসব জিমেইল অক্ষত থাকবে, প্রতিটির জন্য ৳{rate} টাকা ব্যালেন্সে যোগ হবে! 💰\n"
        "৬. শেষ ধাপে কোনো জিমেইল নষ্ট হলে তা আপনাকে কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
        "৭. প্রতিদিন সর্বোচ্চ {daily_limit}টি ফাইল এবং প্রতি ফাইলে সর্বোচ্চ {email_limit}টি ভ্যালিড Gmail দিতে পারবেন।\n"
        "৮. একবার জমা দেওয়া জিমেইল আগামী {cooldown} দিনের মধ্যে পুনরায় সাবমিট করা যাবে না।"
    ),
    "rules": (
        "📜 **আমাদের কাজের নিয়মাবলী ও শর্তাবলী:**\n\n"
        "১. ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক বাছাই সম্পন্ন করবেন।\n"
        "২. সমস্যা থাকা জিমেইল ১ম ধাপেই নির্দিষ্ট কারণসহ বাতিল ও এক্সেল ফাইলে ফেরত যাবে।\n"
        "৩. নির্বাচিত জিমেইলগুলো ২৪ থেকে ৪৮ ঘণ্টা পর্যবেক্ষণে থাকবে।\n"
        "৪. পর্যবেক্ষণ শেষে সঠিক প্রতি জিমেইলে ৳{rate} টাকা ব্যালেন্সে যোগ হবে।\n"
        "৫. শেষ ধাপে নষ্ট হওয়া মেইলগুলো কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
        "৬. এক ফাইলে সর্বোচ্চ {email_limit}টি @gmail.com এবং দিনে সর্বোচ্চ {daily_limit}টি ফাইল দেওয়া যাবে।\n"
        "৭. বিগত {cooldown} দিনের মধ্যে জমা দেওয়া জিমেইল পুনরায় গ্রহণযোগ্য নয়।"
    ),
    "sell": (
        "📤 আপনার ফ্রেশ জিমেইল সম্বলিত এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।\n\n"
        "💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate} BDT\n"
        "⚠️ এক ফাইলে সর্বোচ্চ {email_limit}টি @gmail.com থাকতে হবে। বিগত {cooldown} দিনের মধ্যে জমা দেওয়া কোনো মেইল গ্রহণ করা হবে না।"
    ),
    "support": "যেকোনো সমস্যা, প্রশ্ন বা সহায়তার জন্য নিচে থাকা বাটনে ক্লিক করে সাপোর্টে মেসেজ পাঠান:",
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


def get_custom_msg(key):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"msg_{key}",), fetchone=True)
    rate = get_config("email_rate", float)
    d_lim = get_config("daily_file_limit", int)
    e_lim = get_config("max_emails_per_file", int)
    cd = get_config("duplicate_check_days", int)

    text = res["value"] if (res and res["value"]) else DEFAULT_MESSAGES.get(key, "")
    return (
        text.replace("{rate}", f"{rate:.2f}")
        .replace("{daily_limit}", str(d_lim))
        .replace("{email_limit}", str(e_lim))
        .replace("{cooldown}", str(cd))
    )


def set_custom_msg(key, text):
    set_config(f"msg_{key}", text)


def delete_custom_msg(key):
    db_execute("DELETE FROM settings WHERE key=?", (f"msg_{key}",))


def get_button_title(btn_key, default_title):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"btn_{btn_key}",), fetchone=True)
    return res["value"] if res else default_title


def set_button_title(btn_key, new_title):
    set_config(f"btn_{btn_key}", new_title)


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
            INSERT INTO users (user_id, username, balance, files_today, last_file_date, referred_by, seen_rules, created_at)
            VALUES (?, ?, 0, 0, ?, ?, 0, ?)
        """, (user.id, user.username or "", datetime.now().strftime("%Y-%m-%d"), referrer_id, datetime.now().isoformat()))
        return True
    else:
        db_execute("UPDATE users SET username=? WHERE user_id=?", (user.username or "", user.id))
        return False


def get_user(user_id):
    return db_execute("SELECT * FROM users WHERE user_id=?", (user_id,), fetchone=True)


def mark_rules_seen(user_id):
    db_execute("UPDATE users SET seen_rules=1 WHERE user_id=?", (user_id,))


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
# ASYNC SAFE MEMBERSHIP CHECKER
# =========================================================

async def is_user_joined_all(bot, user_id):
    if is_admin(user_id):
        return True

    async def check_channel():
        try:
            m = await bot.get_chat_member(chat_id=FORCE_CHANNEL_CHAT_ID, user_id=user_id)
            return m.status not in ["left", "kicked"]
        except Exception:
            return False

    async def check_group():
        try:
            m = await bot.get_chat_member(chat_id=FORCE_GROUP_CHAT_ID, user_id=user_id)
            return m.status not in ["left", "kicked"]
        except Exception:
            return False

    try:
        ch_joined, gr_joined = await asyncio.wait_for(
            asyncio.gather(check_channel(), check_group()),
            timeout=2.5
        )
        return ch_joined and gr_joined
    except Exception:
        return False


def get_first_time_markup():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📢 ১. পেমেন্ট প্রুফ চ্যানেল", url=FORCE_CHANNEL_LINK)],
        [InlineKeyboardButton("👥 ২. অফিসিয়াল গ্রুপ", url=FORCE_GROUP_LINK)],
        [InlineKeyboardButton("✅ ৩. ভেরিফাই করুন (Verify)", callback_data="check_joined")],
        [InlineKeyboardButton("💬 সরাসরি মেসেজ পাঠান / Support", url=SUPPORT_URL)]
    ])


def get_rejoin_markup():
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
GMAIL_REGEX = re.compile(r"^[\w.+-]+@gmail\.com$", re.IGNORECASE)


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
    cd_days = get_config("duplicate_check_days", int)
    cutoff_date = (datetime.now() - timedelta(days=cd_days)).isoformat()
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
# DYNAMIC BIG KEYBOARD & INLINE BUILDERS
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

    custom_reply_btns = db_execute("SELECT title FROM dynamic_buttons WHERE parent_id=0 AND show_in_reply=1 ORDER BY id ASC", fetchall=True)
    if custom_reply_btns:
        row = []
        for b in custom_reply_btns:
            row.append(KeyboardButton(b["title"]))
            if len(row) == 2:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)

    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


def build_dynamic_sub_markup(parent_id):
    btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=? ORDER BY id ASC", (parent_id,), fetchall=True)
    if not btns:
        return None

    kb = []
    for b in btns:
        if b["btn_type"] == "url":
            kb.append([InlineKeyboardButton(b["title"], url=b["content"])])
        else:
            kb.append([InlineKeyboardButton(b["title"], callback_data=f"dyn_click:{b['id']}")])

    return InlineKeyboardMarkup(kb)


async def show_main_menu(update, context):
    user = update.effective_user
    add_user(user)

    msg_text = "🎉 **আপনি আমাদের সাথে সফলভাবে যুক্ত আছেন।**\n\nনিচের বাটনগুলো চেপে আপনার অপশন বেছে নিন:"

    url_btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=0 AND btn_type='url' ORDER BY id ASC", fetchall=True)
    inline_kb = None
    if url_btns:
        inline_kb = InlineKeyboardMarkup([[InlineKeyboardButton(b["title"], url=b["content"])] for b in url_btns])

    if update.callback_query:
        await update.callback_query.message.reply_text(msg_text, reply_markup=get_bottom_keyboard(), parse_mode="Markdown")
        if inline_kb:
            await update.callback_query.message.reply_text("📌 **গুরুত্বপূর্ণ লিংকসমূহ:**", reply_markup=inline_kb)
    else:
        await update.message.reply_text(msg_text, reply_markup=get_bottom_keyboard(), parse_mode="Markdown")
        if inline_kb:
            await update.message.reply_text("📌 **গুরুত্বপূর্ণ লিংকসমূহ:**", reply_markup=inline_kb)


# =========================================================
# COMMAND HANDLERS
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return

    user = update.effective_user
    if not user:
        return

    try:
        if is_maintenance_mode() and not is_admin(user.id):
            await update.message.reply_text(get_custom_msg("maintenance"))
            return

        context.user_data.clear()

        ref_bonus = get_config("referral_bonus", float)
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
            update_balance(referrer_id, ref_bonus)
            try:
                await context.bot.send_message(
                    chat_id=referrer_id,
                    text=f"🎉 আপনার একটি নতুন রেফার সফল হয়েছে! বোনাস: +৳{ref_bonus:.2f} BDT"
                )
            except Exception:
                pass

        u_data = get_user(user.id)
        seen_rules = u_data["seen_rules"] if u_data else 0

        joined = await is_user_joined_all(context.bot, user.id)

        if joined:
            await show_main_menu(update, context)
            return

        if not seen_rules:
            name = user.first_name or "User"
            welcome_text = get_custom_msg("welcome").replace("{name}", name)

            await update.message.reply_text(
                f"{welcome_text}\n\n👇 **বট চালু করতে নিচের চ্যানেল ও গ্রুপে যুক্ত হয়ে ভেরিফাই বাটনে চাপ দিন:**",
                reply_markup=get_first_time_markup(),
                parse_mode="Markdown"
            )
            return

        await update.message.reply_text(
            "⚠️ আপনি আমাদের চ্যানেল বা গ্রুপে যুক্ত নেই!\n\n👇 দয়া করে জয়েন হয়ে ভেরিফাই বাটনে চাপ দিন:",
            reply_markup=get_rejoin_markup()
        )

    except Exception as e:
        print(f"Error in start: {e}")
        try:
            await update.message.reply_text(
                "স্বাগতম! নিচের বাটনগুলো চেপে আমাদের চ্যানেল ও গ্রুপে যুক্ত হয়ে ভেরিফাই করুন:",
                reply_markup=get_rejoin_markup()
            )
        except Exception:
            pass


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
# ADMIN CONTROL PANEL
# =========================================================

async def show_admin_panel(query, user_id):
    m_status = "🔴 বট বন্ধ (Maintenance ON)" if is_maintenance_mode() else "🟢 বট চালু (Active)"
    owner = get_owner_id()
    is_owner = (owner and int(user_id) == int(owner))

    keyboard = [
        [InlineKeyboardButton(f"🛠 স্ট্যাটাস: {m_status}", callback_data="toggle_maintenance")],
        [InlineKeyboardButton("⚙️ আর্থিক লেনদেন ও সিস্টেম লিমিট কন্ট্রোল", callback_data="admin_limits_menu")],
        [InlineKeyboardButton("🎛 বড় কীবোর্ড বাটন ও সাব-মেনু বিল্ডার", callback_data="dyn_manage:0")],
        [InlineKeyboardButton("📂 ফাইল হিস্ট্রি ও রিপোর্ট (Users Files)", callback_data="admin_file_history:0")],
        [InlineKeyboardButton("💬 বাটন মেসেজ কন্ট্রোল (Add/Edit/Delete)", callback_data="admin_messages_menu")],
        [InlineKeyboardButton("✏️ ডিফল্ট বাটন নাম এডিট", callback_data="admin_edit_buttons_menu")],
        [InlineKeyboardButton("🧹 ক্লিয়ার ক্যাশ / আবর্জনা মুছুন", callback_data="admin_clear_cache")],
        [InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast"),
         InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message")],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")]
    ]

    if is_owner:
        keyboard.append([InlineKeyboardButton("👥 অ্যাডমিন তালিকা ও নাম এডিট (Owner Only)", callback_data="owner_manage_admins")])

    keyboard.append([InlineKeyboardButton("⬅️ প্যানেল বন্ধ করুন", callback_data="admin_close")])

    role_text = "👑 OWNER CONTROL PANEL" if is_owner else "👨‍💼 ADMIN CONTROL PANEL"
    await query.edit_message_text(f"{role_text}\n\nএকটি অপশন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))


# =========================================================
# STAGE KEYBOARDS & REASON SELECTOR
# =========================================================

admin_stage1_selections = {}
admin_stage3_selections = {}


def build_stage_keyboard(sub_id, emails, selected_indices, stage_num):
    rate = get_config("email_rate", float)
    num_buttons = []
    for idx, _ in enumerate(emails):
        icon = "✅" if idx in selected_indices else "❌"
        num_buttons.append(InlineKeyboardButton(f"[{idx+1}] {icon}", callback_data=f"sub_tg:{stage_num}:{sub_id}:{idx}"))

    chunked = [num_buttons[i:i + 3] for i in range(0, len(num_buttons), 3)]

    if stage_num == 1:
        action_rows = [
            [InlineKeyboardButton(f"🚀 ধাপ ১ নিশ্চিত করুন ({len(selected_indices)}টি পর্যালোচনায় / {len(emails)-len(selected_indices)}টি বাতিল)", callback_data=f"sub_c1:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল (Reject All)", callback_data=f"sub_rj_all_prompt:{stage_num}:{sub_id}")]
        ]
    else:
        action_rows = [
            [InlineKeyboardButton(f"💰 চূড়ান্ত অনুমোদন দিন ({len(selected_indices)}টি বৈধ - ৳{len(selected_indices)*rate:.2f})", callback_data=f"sub_c3:{sub_id}")],
            [InlineKeyboardButton("✅ সবগুলোই ঠিক আছে", callback_data=f"sub_accept_all_s3:{sub_id}"),
             InlineKeyboardButton("❌ পুরো ফাইল বাতিল (Reject All)", callback_data=f"sub_rj_all_prompt:{stage_num}:{sub_id}")]
        ]

    return InlineKeyboardMarkup(chunked + action_rows)


def build_reason_selection_keyboard(stage_num, sub_id, current_email_idx, total_rejected_count):
    kb = [
        [InlineKeyboardButton("🔴 পাসওয়ার্ড ভুল / ইনভ্যালিড ক্রেডেনশিয়াল", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt1")],
        [InlineKeyboardButton("🔴 অ্যাকাউন্ট ডিজেবল্ড অথবা সাসপেন্ডেড", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt2")],
        [InlineKeyboardButton("✍️ কাস্টম কারণ লিখুন", callback_data=f"sub_rs_custom:{stage_num}:{sub_id}:{current_email_idx}")]
    ]
    return InlineKeyboardMarkup(kb)


# =========================================================
# CALLBACK HANDLER
# =========================================================

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    data = query.data

    # ১. মেম্বারশিপ ভেরিফাই
    if data == "check_joined":
        joined = await is_user_joined_all(context.bot, user_id)
        if joined:
            mark_rules_seen(user_id)
            await query.message.delete()
            await show_main_menu(update, context)
        else:
            await query.answer("❌ আপনি এখনো জয়েন করেননি! চ্যানেল ও গ্রুপে জয়েন হয়ে আবার ভেরিফাই চাপুন।", show_alert=True)
        return

    # ২. ডায়নামিক সাব-বাটন ও মাল্টিপল মেসেজ ক্লিক (সকল ইউজারের জন্য)
    if data.startswith("dyn_click:"):
        btn_id = int(data.split(":")[1])
        btn = db_execute("SELECT * FROM dynamic_buttons WHERE id=?", (btn_id,), fetchone=True)
        if not btn:
            await query.answer("বাটন পাওয়া যায়নি!")
            return

        messages = btn["content"].split("---SPLIT---")
        for m in messages:
            cleaned = m.strip()
            if cleaned:
                await context.bot.send_message(chat_id=user_id, text=cleaned, parse_mode="Markdown")

        sub_markup = build_dynamic_sub_markup(btn_id)
        if sub_markup:
            await context.bot.send_message(chat_id=user_id, text=f"📂 **{btn['title']}** এর সাব-মেনু:", reply_markup=sub_markup, parse_mode="Markdown")
        return

    # ৩. ইউজারের সাবমিশন হিস্ট্রি দেখা (সকল সাধারণ ইউজার ও অ্যাডমিনের জন্য উন্মুক্ত)
    if data == "user_hist_live":
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
        return

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
        return

    elif data == "user_hist_menu":
        hist_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("⏳ লাইভ চলমান সাবমিশন (Review)", callback_data="user_hist_live")],
            [InlineKeyboardButton("📜 অতীতের সকল সাবমিশন হিস্ট্রি", callback_data="user_hist_all")]
        ])
        await query.edit_message_text("📜 **আপনার সাবমিশন হিস্ট্রি নির্বাচন করুন:**\n\nনিচের যেকোনো অপশনে ক্লিক করে বিস্তারিত দেখুন:", reply_markup=hist_kb, parse_mode="Markdown")
        return

    # -------------------------------------------------------------
    # এডমিন গার্ড (নিচের একশনগুলো শুধুমাত্র অ্যাডমিন/ওনার করতে পারবে)
    # -------------------------------------------------------------
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
    # আর্থিক লেনদেন ও সিস্টেম লিমিট মেনু
    # -------------------------------------------------------------
    elif data == "admin_limits_menu":
        min_w = get_config("min_withdraw", float)
        w_fee = get_config("withdraw_fee_percent", float)
        u_rate = get_config("usdt_rate", float)
        ref_b = get_config("referral_bonus", float)
        e_rate = get_config("email_rate", float)
        d_lim = get_config("daily_file_limit", int)
        m_lim = get_config("max_emails_per_file", int)
        cd_days = get_config("duplicate_check_days", int)

        kb = [
            [InlineKeyboardButton(f"💵 জিমেইল রেট: ৳{e_rate:.2f}", callback_data="cfg_edit:email_rate"),
             InlineKeyboardButton(f"💸 মিনিমাম উইথড্র: ৳{min_w:.2f}", callback_data="cfg_edit:min_withdraw")],
            [InlineKeyboardButton(f"💳 উইথড্র ফি: {w_fee:.1f}%", callback_data="cfg_edit:withdraw_fee_percent"),
             InlineKeyboardButton(f"💲 USDT রেট: ৳{u_rate:.2f}", callback_data="cfg_edit:usdt_rate")],
            [InlineKeyboardButton(f"👥 রেফার বোনাস: ৳{ref_b:.2f}", callback_data="cfg_edit:referral_bonus"),
             InlineKeyboardButton(f"📅 ডুপ্লিকেট কুলডাউন: {cd_days} দিন", callback_data="cfg_edit:duplicate_check_days")],
            [InlineKeyboardButton(f"📁 দৈনিক ফাইল লিমিট: {d_lim} টি", callback_data="cfg_edit:daily_file_limit"),
             InlineKeyboardButton(f"✉️ ফাইলে জিমেইল লিমিট: {m_lim} টি", callback_data="cfg_edit:max_emails_per_file")],
            [InlineKeyboardButton("⬅️ Back to Panel", callback_data="admin_panel")]
        ]
        await query.edit_message_text("⚙️ **আর্থিক লেনদেন ও সিস্টেম লিমিট কন্ট্রোল:**\n\nযেটি পরিবর্তন করতে চান সেটিতে ক্লিক করুন:", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif data.startswith("cfg_edit:"):
        cfg_key = data.split(":")[1]
        context.user_data["editing_cfg_key"] = cfg_key
        context.user_data["state"] = "waiting_cfg_val"
        names = {
            "min_withdraw": "মিনিমাম উইথড্র (টাকায়)",
            "withdraw_fee_percent": "উইথড্র ফি (শতাংশে, যেমন 4 বা 5)",
            "usdt_rate": "USDT ডলার রেট (টাকায়)",
            "referral_bonus": "রেফারেল বোনাস (টাকায়)",
            "email_rate": "প্রতি ভ্যালিড জিমেইল রেট (টাকায়)",
            "daily_file_limit": "ইউজারের দৈনিক ফাইল লিমিট (সংখ্যা)",
            "max_emails_per_file": "এক ফাইলে সর্বোচ্চ জিমেইল লিমিট (সংখ্যা)",
            "duplicate_check_days": "ডুপ্লিকেট জিমেইল চেক দিন (সংখ্যা)"
        }
        await query.edit_message_text(f"✍️ **{names.get(cfg_key, cfg_key)}** এর নতুন মান লিখে পাঠিয়ে দিন:\n\n(বাতিল করতে /cancel পাঠান)")

    # -------------------------------------------------------------
    # বড় কীবোর্ড ও ডায়নামিক বাটন বিল্ডার
    # -------------------------------------------------------------
    elif data.startswith("dyn_manage:"):
        p_id = int(data.split(":")[1])
        btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=? ORDER BY id ASC", (p_id,), fetchall=True)

        kb = []
        if btns:
            for b in btns:
                t_label = "🔗" if b["btn_type"] == "url" else "💬"
                kb.append([InlineKeyboardButton(f"{t_label} {b['title']}", callback_data=f"dyn_item_opts:{b['id']}")])

        kb.append([InlineKeyboardButton("➕ নতুন বাটন যোগ করুন (বড় কীবোর্ড)", callback_data=f"dyn_add_init:{p_id}")])
        if p_id != 0:
            parent_row = db_execute("SELECT parent_id FROM dynamic_buttons WHERE id=?", (p_id,), fetchone=True)
            prev_p = parent_row["parent_id"] if parent_row else 0
            kb.append([InlineKeyboardButton("⬅️ আগের মেনুতে ফিরুন", callback_data=f"dyn_manage:{prev_p}")])
        else:
            kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])

        lvl_name = "মূল বড় কীবোর্ড (Main Keyboard)" if p_id == 0 else f"সাব-মেনু স্তর (Parent: {p_id})"
        await query.edit_message_text(
            f"🎛 **বড় কীবোর্ড বাটন ও সাব-মেনু বিল্ডার**\n📍 অবস্থান: {lvl_name}\n\n"
            f"যেকোনো বাটনে ক্লিক করে সেটির সাব-বাটন তৈরি করতে পারেন অথবা এক এক করে ডিলিট করতে পারেন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("dyn_item_opts:"):
        b_id = int(data.split(":")[1])
        b = db_execute("SELECT * FROM dynamic_buttons WHERE id=?", (b_id,), fetchone=True)
        if not b:
            await query.answer("বাটন পাওয়া যায়নি!")
            return

        kb = [
            [InlineKeyboardButton("➕ এর ভেতরে সাব-বাটন যোগ করুন", callback_data=f"dyn_add_init:{b['id']}")],
            [InlineKeyboardButton("📂 এর সাব-বাটনগুলো দেখুন ও ম্যানেজ করুন", callback_data=f"dyn_manage:{b['id']}")],
            [InlineKeyboardButton("🗑 শুধুমাত্র এই বাটনটি ডিলিট করুন", callback_data=f"dyn_del_one:{b['id']}")],
            [InlineKeyboardButton("⬅️ তালিকায় ফিরুন", callback_data=f"dyn_manage:{b['parent_id']}")]
        ]

        desc = f"🔗 লিংক: `{b['content']}`" if b["btn_type"] == "url" else f"💬 বার্তা প্রিভিউ:\n`{b['content'][:150]}`"
        await query.edit_message_text(
            f"📌 **বাটন সেটিংস:** {b['title']}\n"
            f"ধরন: {b['btn_type'].upper()}\n{desc}\n\n"
            f"একটি অপশন নির্বাচন করুন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("dyn_del_one:"):
        b_id = int(data.split(":")[1])
        b = db_execute("SELECT * FROM dynamic_buttons WHERE id=?", (b_id,), fetchone=True)
        if b:
            parent_id = b["parent_id"]
            db_execute("DELETE FROM dynamic_buttons WHERE id=?", (b_id,))
            await query.answer("বাটনটি সফলভাবে ডিলিট করা হয়েছে!", show_alert=True)
            query.data = f"dyn_manage:{parent_id}"
            await button_handler(update, context)

    elif data.startswith("dyn_add_init:"):
        p_id = int(data.split(":")[1])
        context.user_data["dyn_parent_id"] = p_id
        context.user_data["state"] = "dyn_waiting_title"
        await query.edit_message_text("✍️ নতুন বাটনের নাম লিখে পাঠান (যেমন: 🎁 অফার বা 📚 ভিডিও গাইড):\n(বাতিল করতে /cancel পাঠান)")

    # -------------------------------------------------------------
    # ওনার কর্তৃক অ্যাডমিন ম্যানেজমেন্ট
    # -------------------------------------------------------------
    elif data == "owner_manage_admins":
        owner = get_owner_id()
        if not owner or int(user_id) != int(owner):
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
        owner = get_owner_id()
        if not owner or int(user_id) != int(owner):
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
            tips = "\n💡 টিপস: লেখার মধ্যে `{name}` দিলে ইউজার নাম, `{rate}` দিলে রেট, `{daily_limit}` দিলে ফাইল লিমিট এবং `{cooldown}` দিলে ডুপ্লিকেট দিন বসবে।"
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

    elif data == "admin_broadcast":
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 ব্রডকাস্ট করার মেসেজটি লিখে পাঠান:\n(এটি সকল ইউজারের ইনবক্সে এবং অফিসিয়াল গ্রুপে চলে যাবে)")

    elif data == "admin_single_message":
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ যাকে মেসেজ পাঠাবেন তার টেলিগ্রাম আইডি পাঠান:")

    elif data == "admin_stats":
        users = db_execute("SELECT COUNT(*) AS c FROM users", fetchone=True)["c"]
        total_balance = db_execute("SELECT COALESCE(SUM(balance),0) AS total FROM users", fetchone=True)["total"]
        u_rate = get_config("usdt_rate", float)
        total_usdt = float(total_balance) / u_rate if u_rate > 0 else 0
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
    # অ্যাডমিন ফাইল হিস্ট্রি ও রিপোর্ট
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
        rate = get_config("email_rate", float)

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
    # ধাপ ১: রিসিভ করা এবং বাকি সবার কাছে মেসেজ যাওয়া + ফাইল রিমুভ
    # -------------------------------------------------------------
    elif data.startswith("sub_s1_open:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("ফাইলটি অলরেডি অন্য একজন প্রসেস করছেন!", show_alert=True)
            return

        rec_name = get_admin_name(user_id)

        if sub["handled_by"] and sub["handled_by"] != rec_name:
            await query.answer(f"⚠️ এই ফাইলটি অলরেডি {sub['handled_by']} রিসিভ করেছেন!", show_alert=True)
            return

        db_execute("UPDATE submissions SET handled_by=? WHERE id=?", (rec_name, sub_id))

        # ১. সবার আগে ওনার ও বাকি অ্যাডমিনদের ইনবক্সে নিশ্চিত নোটিফিকেশন পাঠানো
        team_msg = (
            f"📢 **টিম আপডেট (#Submission_{sub_id})**\n\n"
            f"👤 ফাইল রিসিভ করেছেন: **{rec_name}**\n"
            f"🆔 ইউজার: `{sub['user_id']}`\n"
            f"⚡ প্রাথমিক বাছাই ও ভেরিফিকেশন শুরু হয়েছে।"
        )
        await notify_all_admins(context, team_msg, exclude_user_id=user_id)

        # ২. এরপর বাকি অ্যাডমিনদের চ্যাট থেকে ফাইল মুছে ফেলা
        other_msgs = db_execute("SELECT admin_id, message_id FROM submission_admin_messages WHERE sub_id=?", (sub_id,), fetchall=True)
        for om in other_msgs:
            if int(om["admin_id"]) != int(user_id):
                try:
                    await context.bot.delete_message(chat_id=int(om["admin_id"]), message_id=int(om["message_id"]))
                except Exception:
                    pass

        db_execute("DELETE FROM submission_admin_messages WHERE sub_id=?", (sub_id,))

        # ৩. যিনি রিসিভ করেছেন তার চ্যাটে চেকিং বাটনগুলো দেখানো
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

    # -------------------------------------------------------------
    # ধাপ ১ কনফার্মেশন ও রিজেকশন কারণ উইজার্ড
    # -------------------------------------------------------------
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

        if s1_rejected:
            context.user_data["pending_rejection_flow"] = {
                "stage": 1,
                "sub_id": sub_id,
                "rejected_list": s1_rejected,
                "in_review_list": in_review_emails,
                "current_idx": 0,
                "reasons": {}
            }
            cur_mail = s1_rejected[0]
            kb = build_reason_selection_keyboard(1, sub_id, 0, len(s1_rejected))
            await query.edit_message_caption(
                caption=(
                    f"⚠️ **ধাপ ১: রিজেকশন কারণ নির্বাচন (১/{len(s1_rejected)})**\n\n"
                    f"📧 জিমেইল: `{cur_mail}`\n\n"
                    f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান? নিচের একটি কারণ বেছে নিন:"
                ),
                reply_markup=kb,
                parse_mode="Markdown"
            )
            return

        await finalize_stage1(context, query, sub_id, in_review_emails, [], user_id)

    # -------------------------------------------------------------
    # ধাপ ৩ চূড়ান্ত কনফার্মেশন ও রিজেকশন কারণ উইজার্ড
    # -------------------------------------------------------------
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

        review_emails = json.loads(sub["in_review_json"])
        selected = admin_stage3_selections.get(sub_id, set(range(len(review_emails))))

        final_accepted = [review_emails[i] for i in sorted(list(selected))]
        final_rejected = [review_emails[i] for i in range(len(review_emails)) if i not in selected]

        if final_rejected:
            context.user_data["pending_rejection_flow"] = {
                "stage": 3,
                "sub_id": sub_id,
                "rejected_list": final_rejected,
                "accepted_list": final_accepted,
                "current_idx": 0,
                "reasons": {}
            }
            cur_mail = final_rejected[0]
            kb = build_reason_selection_keyboard(3, sub_id, 0, len(final_rejected))
            await query.edit_message_caption(
                caption=(
                    f"⚠️ **ধাপ ৩: রিজেকশন কারণ নির্বাচন (১/{len(final_rejected)})**\n\n"
                    f"📧 জিমেইল: `{cur_mail}`\n\n"
                    f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান? নিচের একটি কারণ বেছে নিন:"
                ),
                reply_markup=kb,
                parse_mode="Markdown"
            )
            return

        await finalize_stage3(context, query, sub_id, final_accepted, [], user_id)

    # -------------------------------------------------------------
    # রিজেকশন কারণ হ্যান্ডলিং
    # -------------------------------------------------------------
    elif data.startswith("sub_rs_set:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        cur_idx = int(parts[3])
        opt = parts[4]

        flow = context.user_data.get("pending_rejection_flow")
        if not flow or flow["sub_id"] != sub_id:
            await query.answer("সেশন এক্সপায়ার হয়েছে!")
            return

        reason_text = "পাসওয়ার্ড ভুল / ইনভ্যালিড ক্রেডেনশিয়াল" if opt == "opt1" else "অ্যাকাউন্ট ডিজেবল্ড অথবা সাসপেন্ডেড"
        email = flow["rejected_list"][cur_idx]
        flow["reasons"][email] = reason_text

        next_idx = cur_idx + 1
        if next_idx < len(flow["rejected_list"]):
            flow["current_idx"] = next_idx
            next_mail = flow["rejected_list"][next_idx]
            kb = build_reason_selection_keyboard(stage_num, sub_id, next_idx, len(flow["rejected_list"]))
            await query.edit_message_caption(
                caption=(
                    f"⚠️ **ধাপ {stage_num}: রিজেকশন কারণ নির্বাচন ({next_idx+1}/{len(flow['rejected_list'])})**\n\n"
                    f"📧 জিমেইল: `{next_mail}`\n\n"
                    f"এই জিমেইলটির কারণ বেছে নিন:"
                ),
                reply_markup=kb,
                parse_mode="Markdown"
            )
        else:
            rejected_items = [(em, flow["reasons"].get(em, "Invalid")) for em in flow["rejected_list"]]
            if stage_num == 1:
                await finalize_stage1(context, query, sub_id, flow["in_review_list"], rejected_items, user_id)
            elif stage_num == 3:
                await finalize_stage3(context, query, sub_id, flow["accepted_list"], rejected_items, user_id)
            elif stage_num == 99:
                await finalize_reject_all(context, query, sub_id, rejected_items, user_id)
            context.user_data.pop("pending_rejection_flow", None)

    elif data.startswith("sub_rs_custom:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        cur_idx = int(parts[3])

        flow = context.user_data.get("pending_rejection_flow")
        if not flow:
            return

        cur_mail = flow["rejected_list"][cur_idx]
        context.user_data["state"] = "waiting_custom_reason_text"
        context.user_data["custom_reason_info"] = {
            "stage": stage_num,
            "sub_id": sub_id,
            "cur_idx": cur_idx
        }
        await query.edit_message_caption(
            caption=(
                f"✍️ **কাস্টম কারণ লিখুন**\n\n"
                f"📧 জিমেইল: `{cur_mail}`\n\n"
                f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান তা চ্যাটে লিখে পাঠান:\n(বাতিল করতে /cancel পাঠান)"
            ),
            parse_mode="Markdown"
        )

    # -------------------------------------------------------------
    # পুরো ফাইল বাতিল
    # -------------------------------------------------------------
    elif data.startswith("sub_rj_all_prompt:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub:
            return

        emails = json.loads(sub["emails_json"]) if stage_num == 1 else json.loads(sub["in_review_json"])

        context.user_data["pending_rejection_flow"] = {
            "stage": 99,
            "sub_id": sub_id,
            "rejected_list": emails,
            "current_idx": 0,
            "reasons": {}
        }
        cur_mail = emails[0]
        kb = build_reason_selection_keyboard(99, sub_id, 0, len(emails))
        await query.edit_message_caption(
            caption=(
                f"❌ **পুরো ফাইল বাতিল: কারণ নির্বাচন (১/{len(emails)})**\n\n"
                f"📧 জিমেইল: `{cur_mail}`\n\n"
                f"কারণ নির্বাচন করুন:"
            ),
            reply_markup=kb,
            parse_mode="Markdown"
        )

    elif data.startswith("sub_accept_all_s3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            return

        review_emails = json.loads(sub["in_review_json"])
        await finalize_stage3(context, query, sub_id, review_emails, [], user_id)

    # Withdrawals (শুধুমাত্র ওনারের জন্য অনুমোদন/বাতিল)
    elif data.startswith("approve_withdraw:"):
        owner = get_owner_id()
        if not owner or int(user_id) != int(owner):
            await query.answer("❌ শুধুমাত্র বটের মূল মালিক (Owner) উইথড্র অনুমোদন করতে পারেন!", show_alert=True)
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
        u_rate = get_config("usdt_rate", float)
        rec_bdt = float(w['receive_amount'])
        rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0
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
        owner = get_owner_id()
        if not owner or int(user_id) != int(owner):
            await query.answer("❌ শুধুমাত্র বটের মূল মালিক (Owner) উইথড্র বাতিল করতে পারেন!", show_alert=True)
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


# =========================================================
# FINALIZATION LOGICS (STAGE 1, 3 & REJECT ALL)
# =========================================================

async def finalize_stage1(context, query, sub_id, in_review_emails, rejected_items, admin_user_id):
    sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
    rate = get_config("email_rate", float)
    db_execute("""
        UPDATE submissions 
        SET status='stage2_review', in_review_json=?, rejected_count=?
        WHERE id=?
    """, (json.dumps(in_review_emails), len(rejected_items), sub_id))

    admin_stage1_selections.pop(sub_id, None)

    h_by = sub["handled_by"] or get_admin_name(admin_user_id)
    next_kb = InlineKeyboardMarkup([
        [InlineKeyboardButton(f"🔍 ধাপ ৩: চূড়ান্ত অনুমোদন দিন ({len(in_review_emails)}টি)", callback_data=f"sub_s3_open:{sub_id}")],
        [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:2:{sub_id}")]
    ])

    await query.edit_message_caption(
        caption=(
            f"⏳ **ধাপ ১ সম্পন্ন - ধাপ ২ (২৪-৪৮ ঘণ্টা পর্যবেক্ষণ চলছে)** (#{sub_id})\n\n"
            f"👤 ইউজার: {sub['user_id']}\n"
            f"🔒 রিসিভ করেছিলেন: **{h_by}**\n"
            f"🔹 পর্যালোচনায় নেওয়া হয়েছে: {len(in_review_emails)} টি\n"
            f"🔹 সমস্যা থাকায় বাতিল: {len(rejected_items)} টি\n\n"
            f"২৪ থেকে ৪৮ ঘণ্টা পর নিচের বাটনে চাপ দিয়ে চূড়ান্ত যাচাই করুন।"
        ),
        reply_markup=next_kb,
        parse_mode="Markdown"
    )

    user_msg = (
        f"📋 **আপনার জিমেইল সাবমিশনের ১ম ধাপের ফলাফল!** (#{sub_id})\n\n"
        f"🔹 পর্যালোচনায় রাখা হয়েছে: {len(in_review_emails)} টি\n"
        f"❌ সমস্যা থাকায় বাতিল: {len(rejected_items)} টি\n\n"
        f"📌 **জরুরি তথ্য:** পর্যালোচনায় থাকা {len(in_review_emails)}টি জিমেইল আগামী **২৪ থেকে ৪৮ ঘণ্টা** পর্যবেক্ষণে থাকবে। "
        f"পর্যবেক্ষণ শেষে যে জিমেইলগুলো ঠিক থাকবে, প্রতিটির জন্য ৳{rate:.2f} টাকা আপনার ব্যালেন্সে যোগ হবে।"
    )
    if rejected_items:
        user_msg += "\n\n⚠️ **বাতিলকৃত জিমেইল ও সুনির্দিষ্ট কারণ:**\n"
        for i, (em, r_text) in enumerate(rejected_items, 1):
            user_msg += f"{i}. `{em}` ➔ {r_text}\n"

    try:
        await context.bot.send_message(chat_id=sub["user_id"], text=user_msg, parse_mode="Markdown")
    except Exception:
        pass

    if rejected_items:
        bio = create_rejected_excel(rejected_items)
        bio.name = f"rejected_stage1_{sub_id}.xlsx"
        try:
            await context.bot.send_document(
                chat_id=sub["user_id"],
                document=bio,
                caption=f"⚠️ {len(rejected_items)}টি জিমেইলে সমস্যা থাকায় কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হলো।"
            )
        except Exception:
            pass


async def finalize_stage3(context, query, sub_id, accepted_emails, rejected_items, admin_user_id):
    sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
    rate = get_config("email_rate", float)
    total_acc = len(accepted_emails)
    total_rej = sub["rejected_count"] + len(rejected_items)
    reward = total_acc * rate

    db_execute("""
        UPDATE submissions 
        SET status='completed', accepted_count=?, rejected_count=?
        WHERE id=?
    """, (total_acc, total_rej, sub_id))

    if total_acc > 0:
        update_balance(sub["user_id"], reward)

    admin_stage3_selections.pop(sub_id, None)
    approver = get_admin_name(admin_user_id)

    team_msg = (
        f"🎉 **টিম অ্যালার্ট: ফাইল চূড়ান্ত সম্পন্ন হয়েছে! (#Submission_{sub_id})**\n\n"
        f"👤 সম্পন্নকারী: **{approver}**\n"
        f"🆔 ইউজার: `{sub['user_id']}`\n"
        f"✅ অনুমোদিত: {total_acc} টি\n"
        f"❌ মোট বাতিল: {total_rej} টি\n"
        f"💰 ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT"
    )
    await notify_all_admins(context, team_msg, exclude_user_id=admin_user_id)

    await query.edit_message_caption(
        caption=(
            f"✅ **সাবমিশন চূড়ান্তভাবে সম্পন্ন ও ক্লোজ হয়েছে! (#{sub_id})**\n\n"
            f"🔹 চূড়ান্ত ভ্যালিড: {total_acc} টি\n"
            f"🔹 মোট বাতিল: {total_rej} টি\n"
            f"💰 ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT\n"
            f"🔒 সম্পন্ন করেছেন: **{approver}**"
        ),
        parse_mode="Markdown"
    )

    user_msg = (
        f"🎉 **অভিনন্দন! আপনার জিমেইল অ্যাকাউন্টের চূড়ান্ত পর্যবেক্ষণ সম্পন্ন হয়েছে!**\n\n"
        f"🔹 সাবমিশন আইডি: #{sub_id}\n"
        f"✅ সফলভাবে অনুমোদিত: {total_acc} টি\n"
        f"💰 প্রতিটিতে ৳{rate:.2f} হারে আপনার ব্যালেন্সে যোগ হয়েছে: ৳{reward:.2f} BDT!\n"
    )
    if rejected_items:
        user_msg += f"\n⚠️ পর্যবেক্ষণে {len(rejected_items)}টি নষ্ট হওয়ায় বাতিল করা হয়েছে:\n"
        for i, (em, r_text) in enumerate(rejected_items, 1):
            user_msg += f"{i}. `{em}` ➔ {r_text}\n"

    try:
        await context.bot.send_message(chat_id=sub["user_id"], text=user_msg, parse_mode="Markdown")
    except Exception:
        pass

    if rejected_items:
        bio = create_rejected_excel(rejected_items)
        bio.name = f"rejected_stage3_{sub_id}.xlsx"
        try:
            await context.bot.send_document(
                chat_id=sub["user_id"],
                document=bio,
                caption=f"⚠️ {len(rejected_items)}টি জিমেইল নষ্ট হওয়ায় কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হলো।"
            )
        except Exception:
            pass


async def finalize_reject_all(context, query, sub_id, rejected_items, admin_user_id):
    sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
    admin_name = get_admin_name(admin_user_id)
    total_rej = len(rejected_items)

    db_execute("UPDATE submissions SET status='rejected', accepted_count=0, rejected_count=? WHERE id=?", (total_rej, sub_id))

    team_msg = (
        f"❌ **টিম অ্যালার্ট: ফাইল বাতিল করা হয়েছে! (#Submission_{sub_id})**\n\n"
        f"👤 বাতিলকারী: **{admin_name}**\n"
        f"🆔 ইউজার: `{sub['user_id']}`\n"
        f"❌ বাতিল জিমেইল: {total_rej} টি"
    )
    await notify_all_admins(context, team_msg, exclude_user_id=admin_user_id)

    await query.edit_message_caption(caption=f"❌ পুরো ফাইল ({total_rej}টি জিমেইল) কারণসহ বাতিল করা হয়েছে।")

    user_msg = f"❌ **আপনার সাবমিশন (#{sub_id}) বাতিল করা হয়েছে!**\n\nবাতিলকৃত জিমেইল ও সুনির্দিষ্ট কারণ:\n"
    for i, (em, r_text) in enumerate(rejected_items, 1):
        user_msg += f"{i}. `{em}` ➔ {r_text}\n"

    try:
        await context.bot.send_message(chat_id=sub["user_id"], text=user_msg, parse_mode="Markdown")
    except Exception:
        pass

    bio = create_rejected_excel(rejected_items)
    bio.name = f"all_rejected_{sub_id}.xlsx"
    try:
        await context.bot.send_document(chat_id=sub["user_id"], document=bio, caption="❌ আপনার ফাইলের সবগুলো জিমেইল কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হলো।")
    except Exception:
        pass


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
    rate = get_config("email_rate", float)

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

    # ১. কাস্টম রিজেকশন কারণ টাইপ করে সেভ করা
    if is_admin(user.id) and state == "waiting_custom_reason_text":
        info = context.user_data.get("custom_reason_info")
        flow = context.user_data.get("pending_rejection_flow")
        if info and flow:
            cur_idx = info["cur_idx"]
            stage_num = info["stage"]
            sub_id = info["sub_id"]
            email = flow["rejected_list"][cur_idx]
            flow["reasons"][email] = text

            context.user_data["state"] = None
            next_idx = cur_idx + 1

            if next_idx < len(flow["rejected_list"]):
                flow["current_idx"] = next_idx
                next_mail = flow["rejected_list"][next_idx]
                kb = build_reason_selection_keyboard(stage_num, sub_id, next_idx, len(flow["rejected_list"]))
                await update.message.reply_text(
                    f"✅ কারণ সেভ হয়েছে!\n\n"
                    f"⚠️ **পরবর্তী জিমেইল ({next_idx+1}/{len(flow['rejected_list'])})**\n"
                    f"📧 জিমেইল: `{next_mail}`\n\n"
                    f"কারণ নির্বাচন করুন:",
                    reply_markup=kb,
                    parse_mode="Markdown"
                )
            else:
                rejected_items = [(em, flow["reasons"].get(em, "Invalid")) for em in flow["rejected_list"]]
                class DummyQuery:
                    async def edit_message_caption(self, caption, reply_markup=None, parse_mode=None):
                        await update.message.reply_text(caption, reply_markup=reply_markup, parse_mode=parse_mode)
                dq = DummyQuery()

                if stage_num == 1:
                    await finalize_stage1(context, dq, sub_id, flow["in_review_list"], rejected_items, user.id)
                elif stage_num == 3:
                    await finalize_stage3(context, dq, sub_id, flow["accepted_list"], rejected_items, user.id)
                elif stage_num == 99:
                    await finalize_reject_all(context, dq, sub_id, rejected_items, user.id)
                context.user_data.pop("pending_rejection_flow", None)
        return

    # ২. ডায়নামিক কনফিগ ভ্যালু সেভ করা
    if is_admin(user.id) and state == "waiting_cfg_val":
        cfg_key = context.user_data.get("editing_cfg_key")
        try:
            val = float(text)
            if val < 0:
                await update.message.reply_text("❌ ঋণাত্মক মান গ্রহণযোগ্য নয়।")
                return
            set_config(cfg_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! **{cfg_key}** এর নতুন মান **{text}** সেভ হয়েছে।", parse_mode="Markdown")
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যায় মান লিখে পাঠান।")
        return

    # ৩. ডায়নামিক বাটন উইজার্ড
    if is_admin(user.id) and state == "dyn_waiting_title":
        context.user_data["dyn_title"] = text
        context.user_data["state"] = "dyn_waiting_content"
        await update.message.reply_text(
            f"✅ বাটনের নাম: **{text}**\n\n"
            "এবার এই বাটনের তথ্য পাঠিয়ে দিন:\n\n"
            "🔹 **যদি লিংক দিতে চান:** সরাসরি URL লিখে পাঠান (যেমন: `https://t.me/...`)\n"
            "🔹 **যদি মেসেজ দিতে চান:** মেসেজটি লিখে দিন।\n"
            "💡 **টিপস (একাধিক মেসেজ):** আপনি যদি চান এই বাটনে চাপ দিলে ২-৩টি মেসেজ পর পর আসবে, তবে মেসেজগুলোর মাঝখানে `---SPLIT---` লিখে দিন!\n\n"
            "উদাহরণ:\n"
            "`স্বাগতম আমাদের নিয়মে!---SPLIT---প্রতি জিমেইলে ২৫ টাকা পাবেন।---SPLIT---ধন্যবাদ সাথে থাকার জন্য।`",
            parse_mode="Markdown"
        )
        return

    if is_admin(user.id) and state == "dyn_waiting_content":
        parent_id = context.user_data.get("dyn_parent_id", 0)
        title = context.user_data.get("dyn_title")

        b_type = "url" if text.startswith(("http://", "https://", "tg://")) else "message"

        db_execute(
            "INSERT INTO dynamic_buttons (parent_id, title, btn_type, content, show_in_reply) VALUES (?, ?, ?, ?, 1)",
            (parent_id, title, b_type, text)
        )
        context.user_data.clear()
        await update.message.reply_text(f"🎉 **বাটন সফলভাবে তৈরি হয়েছে!**\n\nনাম: {title}\nধরন: {b_type.upper()}\n(এটি নিচের বড় কীবোর্ডে স্বয়ংক্রিয়ভাবে যুক্ত হয়েছে)", reply_markup=get_bottom_keyboard())
        return

    # ৪. বড় কীবোর্ডের ডায়নামিক বাটন ক্লিক হ্যান্ডলিং
    dyn_btn_row = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=0 AND title=?", (text,), fetchone=True)
    if dyn_btn_row:
        if dyn_btn_row["btn_type"] == "url":
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔗 ওপেন করুন", url=dyn_btn_row["content"])]])
            await update.message.reply_text(f"👉 **{dyn_btn_row['title']}** লিংকে যেতে নিচের বাটনে চাপ দিন:", reply_markup=kb, parse_mode="Markdown")
        else:
            messages = dyn_btn_row["content"].split("---SPLIT---")
            for m in messages:
                cleaned = m.strip()
                if cleaned:
                    await update.message.reply_text(cleaned, parse_mode="Markdown")
            
            sub_markup = build_dynamic_sub_markup(dyn_btn_row["id"])
            if sub_markup:
                await update.message.reply_text(f"📂 **{dyn_btn_row['title']}** এর সাব-মেনু:", reply_markup=sub_markup, parse_mode="Markdown")
        return

    # ৫. ব্রডকাস্ট পোস্ট হ্যান্ডলিং (ইউজার ইনবক্স + অফিসিয়াল গ্রুপে)
    if is_admin(user.id) and state == "admin_broadcast":
        users = db_execute("SELECT user_id FROM users", fetchall=True)
        user_count = 0
        for u in users:
            try:
                await context.bot.send_message(chat_id=u["user_id"], text=text, parse_mode="Markdown")
                user_count += 1
            except Exception:
                pass

        group_posted = False
        try:
            await context.bot.send_message(chat_id=FORCE_GROUP_CHAT_ID, text=text, parse_mode="Markdown")
            group_posted = True
        except Exception as e:
            print(f"Failed to post broadcast to official group: {e}")

        context.user_data.clear()
        
        report_msg = (
            f"📢 **ব্রডকাস্ট সম্পন্ন হয়েছে!**\n\n"
            f"👤 ইউজার ইনবক্সে পৌঁছেছে: {user_count}/{len(users)} জনের কাছে\n"
            f"👥 অফিসিয়াল গ্রুপে পোস্ট: {'✅ সফল' if group_posted else '❌ ব্যর্থ'}\n"
            f"📢 পেমেন্ট প্রুফ চ্যানেল: 🚫 বাদ রাখা হয়েছে (নিয়মানুযায়ী পোস্ট হয়নি)"
        )
        await update.message.reply_text(report_msg, parse_mode="Markdown")
        return

    # ৬. সাপোর্ট বাটন
    if text == btn_sup:
        sup_text = get_custom_msg("support")
        keyboard = InlineKeyboardMarkup([[InlineKeyboardButton("💬 সরাসরি মেসেজ পাঠান / Support", url=SUPPORT_URL)]])
        await update.message.reply_text(sup_text, reply_markup=keyboard)
        return

    # ৭. ফোর্স জয়েন যাচাই
    if not is_admin(user.id):
        joined = await is_user_joined_all(context.bot, user.id)
        if not joined:
            await update.message.reply_text(
                "⚠️ আপনি আমাদের চ্যানেল বা গ্রুপে যুক্ত নেই!\n\n👇 দয়া করে জয়েন হয়ে ভেরিফাই বাটনে চাপ দিন:",
                reply_markup=get_rejoin_markup()
            )
            return

    # ওনার কর্তৃক অ্যাডমিন নাম এডিট
    owner = get_owner_id()
    if owner and int(user.id) == int(owner) and state == "waiting_helper_new_name":
        target_id = context.user_data.get("target_helper_id")
        if text:
            set_helper_name(target_id, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! অ্যাডমিন `{target_id}` এর নাম পরিবর্তন করে **{text}** রাখা হয়েছে।", parse_mode="Markdown")
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
            await update.message.reply_text(f"✅ বাটন নাম আপডেট হয়েছে:\n{text}", reply_markup=get_bottom_keyboard())
        return

    # হিস্ট্রি বাটন (কীবোর্ড থেকে চাপলে মেনু আসবে)
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
        d_lim = get_config("daily_file_limit", int)
        e_lim = get_config("max_emails_per_file", int)
        if count >= d_lim:
            await update.message.reply_text(f"❌ দৈনিক লিমিট শেষ ({d_lim}/{d_lim})।")
            return

        context.user_data["state"] = "waiting_file"
        sell_msg = get_custom_msg("sell")
        await update.message.reply_text(
            f"{sell_msg}\n\n"
            f"📊 আজকের সাবমিশন: {count}/{d_lim}\n"
            f"💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate:.2f} BDT\n"
            f"⚠️ ফাইলে সর্বোচ্চ {e_lim}টি জিমেইল থাকতে পারবে।"
        )
        return

    elif text == btn_bal:
        b_bdt = get_balance(user.id)
        u_rate = get_config("usdt_rate", float)
        b_usdt = b_bdt / u_rate if u_rate > 0 else 0
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
        u_rate = get_config("usdt_rate", float)
        min_w = get_config("min_withdraw", float)
        w_fee = get_config("withdraw_fee_percent", float)
        bal_usdt = bal / u_rate if u_rate > 0 else 0
        if bal < min_w:
            await update.message.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.2f} BDT। আপনার আছে: ৳{bal:.2f} BDT")
            return

        context.user_data["state"] = "withdraw_amount"
        await update.message.reply_text(
            f"💸 WITHDRAWAL SYSTEM\n\n"
            f"💰 ব্যালেন্স: ৳{bal:.2f} BDT (${bal_usdt:.2f} USDT)\n"
            f"🔹 সর্বনিম্ন উইথড্র: ৳{min_w:.2f} BDT\n"
            f"🔹 ফি: {w_fee:.1f}%\n"
            f"🔹 ডলার রেট: $1 = ৳{u_rate:.2f}\n\n"
            f"কত টাকা উইথড্র করতে চান? সংখ্যায় লিখে পাঠান:"
        )
        return

    elif text == btn_ref:
        b_info = await context.bot.get_me()
        ref_link = f"https://t.me/{b_info.username}?start={user.id}"
        refs = get_referral_count(user.id)
        ref_bonus = get_config("referral_bonus", float)
        earned = refs * ref_bonus

        ref_template = get_custom_msg("referral")
        custom_ref = ref_template.replace("{link}", ref_link).replace("{bonus}", str(ref_bonus))

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

    if state == "withdraw_amount":
        try:
            amount = float(text)
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যা লিখুন।")
            return

        min_w = get_config("min_withdraw", float)
        w_fee = get_config("withdraw_fee_percent", float)
        u_rate = get_config("usdt_rate", float)

        if amount < min_w:
            await update.message.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.2f} BDT।")
            return

        if amount > get_balance(user.id):
            await update.message.reply_text("❌ ব্যালেন্স পর্যাপ্ত নেই।")
            return

        fee = amount * w_fee / 100
        rec = amount - fee
        rec_usdt = rec / u_rate if u_rate > 0 else 0

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
        u_rate = get_config("usdt_rate", float)
        rec_usdt = rec / u_rate if u_rate > 0 else 0

        db_execute("""
            INSERT INTO withdrawals (user_id, amount, fee, receive_amount, method, account, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (user.id, amt, fee, rec, method, text, datetime.now().isoformat()))

        w_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        context.user_data.clear()

        await update.message.reply_text(f"✅ উইথড্র রিকোয়েস্ট সফল হয়েছে!\nমেথড: {method}\nঅ্যাকাউন্ট: {text}\nপাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)")

        owner_id = get_owner_id()
        if owner_id:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")]
            ])
            try:
                await context.bot.send_message(
                    chat_id=int(owner_id),
                    text=(
                        f"🔔 **নতুন উইথড্র রিকোয়েস্ট (Owner Only)** (#{w_id})\n\n"
                        f"👤 ইউজার: `{user.id}`\n"
                        f"💰 তোলার পরিমাণ: ৳{amt:.2f} BDT\n"
                        f"💸 পেমেন্ট দিতে হবে: ৳{rec:.2f} (${rec_usdt:.2f})\n"
                        f"💳 মেথড: {method}\n"
                        f"📌 অ্যাকাউন্ট/UID: `{text}`"
                    ),
                    reply_markup=keyboard,
                    parse_mode="Markdown"
                )
            except Exception:
                pass
        return

    # ফাইল সাবমিট প্রসেসিং
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

        m_lim = get_config("max_emails_per_file", int)
        if len(valid_emails) > m_lim:
            await update.message.reply_text(
                f"❌ ফাইলে {len(valid_emails)}টি জিমেইল রয়েছে।\n"
                f"একটি ফাইলে সর্বোচ্চ {m_lim}টি জিমেইল দেওয়া অনুমোদিত।"
            )
            return

        duplicates = check_recent_duplicate_emails(valid_emails)
        if duplicates:
            cd_days = get_config("duplicate_check_days", int)
            dup_list = "\n".join([f"• {d}" for d in duplicates])
            await update.message.reply_text(
                f"❌ **ফাইল গ্রহণ করা হয়নি!**\n\n"
                f"নিচের জিমেইলগুলো বিগত {cd_days} দিনের মধ্যে জমা দেওয়া হয়েছিল:\n{dup_list}\n\n"
                f"⚠️ একবার জমা দেওয়া জিমেইল {cd_days} দিন অতিবাহিত না হওয়া পর্যন্ত আর সাবমিট করা যাবে না।"
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
        admin_list = []
        if owner:
            try:
                admin_list.append(int(owner))
            except (ValueError, TypeError):
                pass
        for h in helpers:
            try:
                admin_list.append(int(h["user_id"]))
            except (ValueError, TypeError):
                pass
        all_admin_ids = list(set(admin_list))

        mail_preview = "\n".join([f"{i+1}. {m}" for i, m in enumerate(valid_emails)])
        admin_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 ধাপ ১: প্রাথমিক বাছাই ও রিসিভ", callback_data=f"sub_s1_open:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:1:{sub_id}")]
        ])

        for a_id in all_admin_ids:
            try:
                sent_msg = await context.bot.send_document(
                    chat_id=int(a_id),
                    document=doc.file_id,
                    caption=(
                        f"📁 নতুন জিমেইল সাবমিশন (#{sub_id})\n"
                        f"👤 ইউজার: {user.id}\n"
                        f"✉️ জিমেইল সংখ্যা: {len(valid_emails)} টি\n\n"
                        f"📋 তালিকা:\n{mail_preview}"
                    ),
                    reply_markup=admin_kb
                )
                db_execute(
                    "INSERT INTO submission_admin_messages (sub_id, admin_id, message_id) VALUES (?, ?, ?)",
                    (sub_id, int(a_id), sent_msg.message_id)
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

    application.add_handler(CommandHandler("start", start, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("admin", admin, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("setadmin", set_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("addadmin", add_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("removeadmin", remove_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("adminlist", admin_list_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("addbalance", add_balance_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("cutbalance", cut_balance_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("resetbalance", reset_balance_cmd, filters=filters.ChatType.PRIVATE))

    application.add_handler(CallbackQueryHandler(withdrawal_method_handler, pattern=r"^method_"))
    application.add_handler(CallbackQueryHandler(button_handler))
    
    application.add_handler(MessageHandler(filters.ChatType.PRIVATE & ~filters.COMMAND, message_handler))

    print("Gmail Sell Bot is running safely in private chat mode...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
