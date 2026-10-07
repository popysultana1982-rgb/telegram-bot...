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
REWARD_PER_FILE = 25
DAILY_FILE_LIMIT = 5

REFERRAL_BONUS = 5.0
USDT_RATE = 124.0

DB_NAME = "bot.db"
PORT = int(os.getenv("PORT", "10000"))


# =========================================================
# RENDER WEB SERVER
# =========================================================

class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"Telegram bot is running.")

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
            valid_emails INTEGER DEFAULT 0,
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
# ADMIN & SYSTEM HELPERS
# =========================================================

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


def get_button_title(btn_key, default_title):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"btn_{btn_key}",), fetchone=True)
    return res["value"] if res else default_title


def set_button_title(btn_key, new_title):
    db_execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (f"btn_{btn_key}", new_title))


def get_custom_message(msg_key, default_text):
    res = db_execute("SELECT value FROM settings WHERE key=?", (f"msg_{msg_key}",), fetchone=True)
    return res["value"] if res else default_text


def set_custom_message(msg_key, text):
    db_execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value
    """, (f"msg_{msg_key}", text))


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
    unique_gmails = list(set(valid_gmails))
    return len(raw_entries), len(unique_gmails), unique_gmails[:5]


# =========================================================
# DEFAULT MESSAGES & TEMPLATES
# =========================================================

DEFAULT_RULES = (
    "আসসালামু আলাইকুম, {name}! 🌸\n\n"
    "💙 আপনাকে স্বাগতম আমাদের Gmail Sell Bot-এ!\n"
    "আপনি আমাদের বটে নতুন এসেছেন। এখানে আপনি আপনার তৈরি করা Valid Gmail Account সেল/সাবমিট করতে পারবেন।\n\n"
    "📌 গুরুত্বপূর্ণ নিয়মাবলি:\n"
    "🔹 প্রতিদিন সর্বোচ্চ ৫টি Gmail Account সাবমিট করতে পারবেন।\n"
    "🔹 প্রতিটি Gmail অবশ্যই Valid এবং ব্যবহারযোগ্য হতে হবে।\n"
    "🔹 Gmail সাবমিট করার পর সর্বোচ্চ ২৪ ঘণ্টার মধ্যে আপনার Gmail রিসিভ করা হবে।\n"
    "🔹 Gmail রিসিভ হওয়ার পর অ্যাডমিন আপনাকে মেসেজের মাধ্যমে জানিয়ে দেবে যে আপনার Gmail রিসিভ হয়েছে।\n"
    "🔹 অ্যাডমিন রিসিভ করার সময় থেকে পরবর্তী ২৪ ঘণ্টা Gmail-এর স্ট্যাটাস পর্যবেক্ষণ করা হবে।\n"
    "🔹 ২৪ ঘণ্টা পর যে Gmail Accountগুলো ঠিক থাকবে/নষ্ট হবে না, সেই Gmailগুলোর টাকা আপনার Balance-এ অটোমেটিক যোগ হয়ে যাবে। 💰\n"
    "🔹 আর যে Gmail Accountগুলো নষ্ট হয়ে যাবে, সেগুলোর জন্য টাকা যোগ হবে না এবং সেই Gmail Accountগুলো আপনাকে ফেরত দেওয়া হবে।\n\n"
    "⚠️ দয়া করে শুধু Valid Gmail Account সাবমিট করুন এবং উপরের নিয়মগুলো মেনে চলুন।\n\n"
    "💚 ধন্যবাদ আমাদের সাথে থাকার জন্য।\n"
    "সুন্দর ও নিরাপদ লেনদেনের শুভকামনা!"
)

DEFAULT_SUPPORT = "যে কোনো সমস্যা বা সহযোগিতার জন্য সরাসরি সাপোর্টে যোগাযোগ করুন:"
DEFAULT_MAINTENANCE = (
    "⚠️ **বট আপডেটের কাজ চলছে!** 🛠\n\n"
    "সম্মানিত ইউজার, আমাদের সিস্টেমে জরুরি আপডেটের কাজ চলছে। "
    "সাময়িকভাবে বটের কার্যক্রম স্থগিত রয়েছে। কাজ সম্পন্ন হওয়ামাত্রই বট পুনরায় চালু হবে।"
)


def get_rules_text(user_name):
    template = get_custom_message("rules_welcome", DEFAULT_RULES)
    return template.replace("{name}", user_name)


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
    welcome_msg = get_rules_text(name)

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
        m_msg = get_custom_message("maintenance_text", DEFAULT_MAINTENANCE)
        await update.message.reply_text(m_msg)
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

    provided_key = context.args[0]
    if provided_key == ADMIN_SECRET_KEY:
        set_admin_id(user_id)
        await update.message.reply_text(
            f"✅ সফল হয়েছে! আপনার আইডি ({user_id}) এখন অ্যাডমিন হিসেবে সেট করা হয়েছে।\nএখন /admin লিখে প্যানেল ওপেন করুন।"
        )
    else:
        await update.message.reply_text("❌ গোপন পাসওয়ার্ড ভুল!")


async def admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if not admin_id or update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.\nপ্রথমে নিজেকে অ্যাডমিন করতে `/setadmin <key>` পাঠান।")
        return

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👨‍💼 Open Admin Panel", callback_data="admin_panel")]
    ])
    await update.message.reply_text("🔐 Admin access granted.", reply_markup=keyboard)


async def add_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.")
        return

    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার নিয়ম: `/addbalance <user_id> <amount>`")
        return

    try:
        target_user_id = int(context.args[0])
        amount = float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ সংখ্যায় আইডি ও টাকার পরিমাণ দিন।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ এই আইডির কোনো ইউজার নেই।")
        return

    update_balance(target_user_id, amount)
    new_bdt = get_balance(target_user_id)
    new_usdt = new_bdt / USDT_RATE

    await update.message.reply_text(
        f"✅ ব্যালেন্স যোগ হয়েছে!\n👤 ইউজার: {target_user_id}\n➕ যোগ: ৳{amount:.2f} BDT\n💰 মোট: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
    )

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text=f"🎁 আপনার অ্যাকাউন্টে ৳{amount:.2f} BDT যোগ করা হয়েছে!\nআপনার বর্তমান ব্যালেন্স: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
        )
    except Exception:
        pass


async def cut_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.")
        return

    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার নিয়ম: `/cutbalance <user_id> <amount>`")
        return

    try:
        target_user_id = int(context.args[0])
        amount = float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ সংখ্যায় আইডি ও টাকার পরিমাণ দিন।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ এই আইডির কোনো ইউজার নেই।")
        return

    update_balance(target_user_id, -amount)
    new_bdt = get_balance(target_user_id)
    new_usdt = new_bdt / USDT_RATE

    await update.message.reply_text(
        f"✂️ ব্যালেন্স কেটে নেওয়া হয়েছে!\n👤 ইউজার: {target_user_id}\n➖ কাটা হয়েছে: ৳{amount:.2f} BDT\n💰 বর্তমান: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
    )

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text=f"⚠️ আপনার অ্যাকাউন্ট থেকে ৳{amount:.2f} BDT কেটে নেওয়া হয়েছে।\nবর্তমান ব্যালেন্স: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
        )
    except Exception:
        pass


async def reset_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.")
        return

    if len(context.args) < 1:
        await update.message.reply_text("ব্যবহার নিয়ম: `/resetbalance <user_id>`")
        return

    try:
        target_user_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ ইউজার আইডি সংখ্যায় হতে হবে।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ ইউজার পাওয়া যায়নি।")
        return

    db_execute("UPDATE users SET balance=0 WHERE user_id=?", (target_user_id,))
    await update.message.reply_text(f"🔄 ব্যালেন্স রিসেট সফল হয়েছে!\n👤 ইউজার: {target_user_id}\n💰 বর্তমান: ৳0.00 BDT")

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text="⚠️ আপনার অ্যাকাউন্টের ব্যালেন্স রিসেট করে ৳0.00 করা হয়েছে।"
        )
    except Exception:
        pass


# =========================================================
# ADMIN PANEL (WITH MESSAGES & BUTTONS EDIT)
# =========================================================

async def show_admin_panel(query):
    m_status = "🔴 বট বন্ধ (Maintenance ON)" if is_maintenance_mode() else "🟢 বট চালু (Active)"

    keyboard = [
        [InlineKeyboardButton(f"🛠 স্ট্যাটাস: {m_status}", callback_data="toggle_maintenance")],
        [InlineKeyboardButton("📝 মেসেজ টেক্সট এডিট করুন", callback_data="admin_edit_messages_menu")],
        [InlineKeyboardButton("✏️ মেনু বাটন নাম এডিট করুন", callback_data="admin_edit_buttons_menu")],
        [InlineKeyboardButton("🧹 ক্লিয়ার ক্যাশ ও জ্যাঙ্ক ফাইলস", callback_data="admin_clear_cache")],
        [InlineKeyboardButton("➕ নতুন লিংক বাটন যোগ", callback_data="admin_add_button"),
         InlineKeyboardButton("🗑 লিংক বাটন মুছুন", callback_data="admin_remove_button")],
        [InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast"),
         InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message")],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")],
        [InlineKeyboardButton("⬅️ প্যানেল বন্ধ করুন", callback_data="admin_close")]
    ]
    await query.edit_message_text("👨‍💼 ADMIN CONTROL PANEL\n\nএকটি অপশন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))


# =========================================================
# CALLBACK HANDLER
# =========================================================

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    admin_id = get_admin_id()
    add_user(query.from_user)
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
        msg = "বট বন্ধ করা হয়েছে (ইউজারদের জন্য রক্ষণাবেক্ষণ মোড চালু)!" if new_state else "বট সফলভাবে সবার জন্য চালু করা হয়েছে!"
        await query.answer(msg, show_alert=True)
        await show_admin_panel(query)

    elif data == "admin_clear_cache":
        if user_id != admin_id:
            return
        deleted = clear_junk_cache()
        await query.answer(f"সফল হয়েছে! {deleted}টি টেম্প ফাইল মুছে ক্যাশ পরিষ্কার করা হয়েছে।", show_alert=True)
        await show_admin_panel(query)

    # --- মেসেজ এডিট সাব-মেনু ---
    elif data == "admin_edit_messages_menu":
        if user_id != admin_id:
            return
        keyboard = [
            [InlineKeyboardButton("📝 ওয়েলকাম ও রুলস মেসেজ", callback_data="edit_msg:rules_welcome")],
            [InlineKeyboardButton("📝 সাপোর্ট মেসেজ", callback_data="edit_msg:support_text")],
            [InlineKeyboardButton("📝 মেইনটেনেন্স মেসেজ", callback_data="edit_msg:maintenance_text")],
            [InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]
        ]
        await query.edit_message_text("📝 কোন মেসেজটি এডিট করতে চান? নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("edit_msg:"):
        if user_id != admin_id:
            return
        msg_key = data.split(":")[1]
        context.user_data["editing_msg_key"] = msg_key
        context.user_data["state"] = "waiting_new_message_text"

        hint = ""
        if msg_key == "rules_welcome":
            hint = "\n\n💡 টিপস: লেখার মধ্যে `{name}` রাখলে সেখানে স্বয়ংক্রিয়ভাবে ইউজারের নাম বসে যাবে।"

        await query.edit_message_text(f"এই মেসেজের জন্য নতুন সম্পূর্ণ টেক্সটটি লিখে পাঠান:{hint}\n(কী: {msg_key})")

    # --- বাটন নাম এডিট সাব-মেনু ---
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
        await query.edit_message_text("✏️ কোন বাটনটির নাম পরিবর্তন করতে চান? নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("edit_btn:"):
        if user_id != admin_id:
            return
        btn_key = data.split(":")[1]
        context.user_data["editing_btn_key"] = btn_key
        context.user_data["state"] = "waiting_new_button_title"
        await query.edit_message_text(f"বাটনটির নতুন নাম কী রাখতে চান? লিখে মেসেজ পাঠান:\n(বাটনের কী: {btn_key})")

    elif data == "admin_add_button":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_add_button_title"
        await query.edit_message_text("➕ Add New Button\n\nপ্রথমে বাটনের নাম লিখে পাঠান।\n\nউদাহরণ:\n📢 Our Channel")

    elif data == "admin_remove_button":
        if user_id != admin_id:
            return
        buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
        if not buttons:
            await query.edit_message_text(
                "কোনো কাস্টম বাটন তৈরি করা নেই।",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")]])
            )
            return

        keyboard = [[InlineKeyboardButton(f"❌ {b['title']}", callback_data=f"delete_button:{b['id']}")] for b in buttons]
        keyboard.append([InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")])
        await query.edit_message_text("🗑 মুছে ফেলতে বাটন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_broadcast":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 Broadcast Post\n\nসকল ইউজারের কাছে যে বার্তাটি পাঠাতে চান তা লিখে পাঠান:")

    elif data == "admin_single_message":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ Single User Message\n\nপ্রথমে ইউজারের নিউমেরিক আইডি পাঠান:")

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

    # Submissions
    elif data.startswith("approve_sub:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        db_execute("UPDATE submissions SET status='approved' WHERE id=?", (sub_id,))
        update_balance(sub["user_id"], REWARD_PER_FILE)
        await query.edit_message_text(f"✅ Submission #{sub_id} Approved! ৳{REWARD_PER_FILE} BDT যোগ হয়েছে।")
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=(
                    f"🎉 অভিনন্দন! আপনার জমা দেওয়া ফ্রেশ জিমেইল ফাইলটি অনুমোদিত হয়েছে।\n"
                    f"💰 আপনার অ্যাকাউন্টে ৳{REWARD_PER_FILE} যোগ করা হয়েছে।"
                )
            )
        except Exception:
            pass

    elif data.startswith("reject_sub:"):
        if user_id != admin_id:
            return
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=?", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        db_execute("UPDATE submissions SET status='rejected' WHERE id=?", (sub_id,))
        await query.edit_message_text(f"❌ Submission #{sub_id} Rejected.")
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=(
                    "❌ আপনার জমা দেওয়া জিমেইল ফাইলটি বাতিল (Rejected) করা হয়েছে।\n\n"
                    "⚠️ কারণ: ফাইলে ভুল মেইল ছিল অথবা ফ্রেশ জিমেইল পাওয়া যায়নি। দয়া করে সঠিক ফ্রেশ জিমেইল ফাইল জমা দিন।"
                )
            )
        except Exception:
            pass

    # Withdrawals
    elif data.startswith("approve_withdraw:"):
        if user_id != admin_id:
            return
        withdrawal_id = int(data.split(":")[1])
        withdrawal = db_execute("SELECT * FROM withdrawals WHERE id=?", (withdrawal_id,), fetchone=True)
        if not withdrawal or withdrawal["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        balance = get_balance(withdrawal["user_id"])
        if balance < float(withdrawal["amount"]):
            db_execute("UPDATE withdrawals SET status='rejected' WHERE id=?", (withdrawal_id,))
            await query.edit_message_text("❌ Withdrawal rejected: ব্যালেন্স পর্যাপ্ত নেই।")
            return

        update_balance(withdrawal["user_id"], -float(withdrawal["amount"]))
        db_execute("UPDATE withdrawals SET status='approved' WHERE id=?", (withdrawal_id,))

        receive_bdt = float(withdrawal['receive_amount'])
        receive_usdt = receive_bdt / USDT_RATE

        await query.edit_message_text(
            f"✅ Withdrawal Approved.\nইউজার পাবে: ৳{receive_bdt:.2f} BDT (${receive_usdt:.2f} USDT)"
        )

        try:
            msg_text = (
                f"✅ আপনার উইথড্রয়াল সফল হয়েছে!\n\n"
                f"💸 পাঠানো হয়েছে: ৳{receive_bdt:.2f} BDT"
            )
            if withdrawal["method"] == "Binance":
                msg_text += f" (${receive_usdt:.2f} USDT)"
            msg_text += f"\n💳 মেথড: {withdrawal['method']}\n📌 একাউন্ট/UID: {withdrawal['account']}"

            await context.bot.send_message(chat_id=withdrawal["user_id"], text=msg_text)
        except Exception:
            pass

    elif data.startswith("reject_withdraw:"):
        if user_id != admin_id:
            return
        withdrawal_id = int(data.split(":")[1])
        withdrawal = db_execute("SELECT * FROM withdrawals WHERE id=?", (withdrawal_id,), fetchone=True)
        if not withdrawal or withdrawal["status"] != "pending":
            await query.answer("Already processed.", show_alert=True)
            return

        db_execute("UPDATE withdrawals SET status='rejected' WHERE id=?", (withdrawal_id,))
        await query.edit_message_text(f"❌ Withdrawal #{withdrawal_id} Rejected.")

        try:
            await context.bot.send_message(
                chat_id=withdrawal["user_id"],
                text=(
                    "❌ আপনার উইথড্রয়াল রিকোয়েস্টটি বাতিল (Rejected) করা হয়েছে।\n\n"
                    f"💸 অ্যামাউন্ট: ৳{withdrawal['amount']:.2f}\n"
                    f"💳 মেথড: {withdrawal['method']}\n"
                    f"📞 বিস্তারিত জানতে সাপোর্টে যোগাযোগ করুন।"
                )
            )
        except Exception:
            pass

    elif data.startswith("delete_button:"):
        if user_id != admin_id:
            return
        button_id = int(data.split(":")[1])
        db_execute("DELETE FROM buttons WHERE id=?", (button_id,))
        await query.answer("Button deleted.")
        await show_admin_panel(query)


# =========================================================
# MESSAGE HANDLER
# =========================================================

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not user:
        return

    admin_id = get_admin_id()

    # Maintenance Check
    if is_maintenance_mode() and user.id != admin_id:
        m_msg = get_custom_message("maintenance_text", DEFAULT_MAINTENANCE)
        await update.message.reply_text(m_msg)
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

    # মেসেজ এডিটের ইনপুট নেওয়া
    if user.id == admin_id and state == "waiting_new_message_text":
        msg_key = context.user_data.get("editing_msg_key")
        if text:
            set_custom_message(msg_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! '{msg_key}' মেসেজ টেক্সট আপডেট করা হয়েছে।")
        return

    # বাটন নাম এডিটের ইনপুট নেওয়া
    if user.id == admin_id and state == "waiting_new_button_title":
        btn_key = context.user_data.get("editing_btn_key")
        if text:
            set_button_title(btn_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! '{btn_key}' বাটনের নতুন নাম রাখা হয়েছে:\n{text}")
        return

    # বাটন ক্লিকে রেসপন্স
    if text == btn_sell:
        count = get_today_file_count(user.id)
        if count >= DAILY_FILE_LIMIT:
            await update.message.reply_text(
                f"❌ দৈনিক লিমিট শেষ।\n\nআপনি প্রতিদিন সর্বোচ্চ {DAILY_FILE_LIMIT}টি ফাইল পাঠাতে পারবেন।"
            )
            return

        context.user_data["state"] = "waiting_file"
        await update.message.reply_text(
            f"📤 আপনার ফ্রেশ জিমেইল সম্বলিত এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।\n\n"
            f"⚠️ সতর্কতা: শুধুমাত্র @gmail.com ফরম্যাটের মেইল গ্রহণ করা হবে। অন্য কোনো মেইল বা টেক্সট গ্রহণযোগ্য নয়।\n\n"
            f"আজকের সাবমিশন: {count}/{DAILY_FILE_LIMIT}\n"
            f"অনুমোদিত ফাইল রিওয়ার্ড: ৳{REWARD_PER_FILE} BDT"
        )
        return

    elif text == btn_bal:
        balance_bdt = get_balance(user.id)
        balance_usdt = balance_bdt / USDT_RATE
        ref_count = get_referral_count(user.id)
        await update.message.reply_text(
            f"💰 YOUR ACCOUNT BALANCE\n\n"
            f"🔹 BDT Balance: ৳{balance_bdt:.2f} BDT\n"
            f"🔹 USDT Balance: ${balance_usdt:.2f} USDT\n"
            f"👥 Total Referrals: {ref_count} users"
        )
        return

    elif text == btn_wd:
        balance_bdt = get_balance(user.id)
        balance_usdt = balance_bdt / USDT_RATE
        if balance_bdt < MIN_WITHDRAW:
            await update.message.reply_text(
                f"❌ সর্বনিম্ন উইথড্র ৳{MIN_WITHDRAW} BDT।\n\n"
                f"আপনার বর্তমান ব্যালেন্স: ৳{balance_bdt:.2f} BDT (${balance_usdt:.2f} USDT)"
            )
            return

        context.user_data["state"] = "withdraw_amount"
        await update.message.reply_text(
            f"💸 WITHDRAWAL SYSTEM\n\n"
            f"💰 Available Balance: ৳{balance_bdt:.2f} BDT (${balance_usdt:.2f} USDT)\n"
            f"🔹 Minimum Withdraw: ৳{MIN_WITHDRAW} BDT\n"
            f"🔹 Fee: {WITHDRAW_FEE_PERCENT}%\n"
            f"🔹 Dollar Rate: $1 = ৳{USDT_RATE:.2f}\n\n"
            f"কত টাকা (BDT) উইথড্র করতে চান? টাকার পরিমাণ লিখে পাঠান:"
        )
        return

    elif text == btn_ref:
        bot_info = await context.bot.get_me()
        bot_username = bot_info.username
        referral_link = f"https://t.me/{bot_username}?start={user.id}"
        ref_count = get_referral_count(user.id)
        total_earned = ref_count * REFERRAL_BONUS

        msg = (
            f"👥 রেফার করে ইনকাম করুন!\n\n"
            f"আপনি প্রতি সফল রেফারে পাবেন ৳{REFERRAL_BONUS:.2f} BDT বোনাস।\n"
            f"আপনার রেফার লিংকটি শেয়ার করুন:\n\n"
            f"🔗 রেফারেল লিংক:\n`{referral_link}`\n\n"
            f"📊 রেফারেল পরিসংখ্যান:\n"
            f"🔹 মোট রেফার: {ref_count} জন\n"
            f"🔹 মোট আয়: ৳{total_earned:.2f} BDT"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")
        return

    elif text == btn_rules:
        name = user.first_name or "User"
        rules_msg = get_rules_text(name)
        await update.message.reply_text(rules_msg)
        return

    elif text == btn_sup:
        sup_text = get_custom_message("support_text", DEFAULT_SUPPORT)
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("💬 মেসেজ পাঠান", url="https://t.me/Talha_juba098")]
        ])
        await update.message.reply_text(sup_text, reply_markup=keyboard)
        return

    # Admin Actions
    if user.id == admin_id and state == "admin_add_button_title":
        context.user_data["button_title"] = text
        context.user_data["state"] = "admin_add_button_url"
        await update.message.reply_text("এখন বাটনের URL দিন (যেমন: https://t.me/yourchannel):")
        return

    if user.id == admin_id and state == "admin_add_button_url":
        url = text
        if not url.startswith(("http://", "https://", "tg://")):
            await update.message.reply_text("❌ সঠিক URL লিখে পাঠান।")
            return
        title = context.user_data.get("button_title")
        db_execute("INSERT INTO buttons (title, url) VALUES (?, ?)", (title, url))
        context.user_data.clear()
        await update.message.reply_text(f"✅ বাটন যুক্ত হয়েছে:\n{title} -> {url}")
        return

    if user.id == admin_id and state == "admin_broadcast":
        users = db_execute("SELECT user_id FROM users", fetchall=True)
        success = 0
        for u in users:
            try:
                await context.bot.send_message(chat_id=u["user_id"], text=text)
                success += 1
            except Exception:
                pass
        context.user_data.clear()
        await update.message.reply_text(f"📢 ব্রডকাস্ট সম্পন্ন: {success}/{len(users)}")
        return

    if user.id == admin_id and state == "admin_single_user":
        try:
            context.user_data["target_user"] = int(text)
            context.user_data["state"] = "admin_single_text"
            await update.message.reply_text("এখন বার্তাটি লিখে পাঠান যা আপনি পাঠাতে চান:")
        except ValueError:
            await update.message.reply_text("❌ সঠিক নিউমেরিক ইউজার আইডি পাঠান।")
        return

    if user.id == admin_id and state == "admin_single_text":
        target_id = context.user_data.get("target_user")
        try:
            await context.bot.send_message(chat_id=target_id, text=text)
            await update.message.reply_text("✅ বার্তা সফলভাবে পাঠানো হয়েছে।")
        except Exception as e:
            await update.message.reply_text(f"❌ বার্তা পাঠানো যায়নি: {e}")
        context.user_data.clear()
        return

    # Withdraw Handling
    if state == "withdraw_amount":
        try:
            amount = float(text)
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যায় টাকার পরিমাণ লিখুন।")
            return

        if amount < MIN_WITHDRAW:
            await update.message.reply_text(f"❌ সর্বনিম্ন উইথড্র ৳{MIN_WITHDRAW} BDT।")
            return

        balance = get_balance(user.id)
        if amount > balance:
            await update.message.reply_text(f"❌ অপর্যাপ্ত ব্যালেন্স। আপনার আছে: ৳{balance:.2f} BDT")
            return

        fee = amount * WITHDRAW_FEE_PERCENT / 100
        receive = amount - fee
        receive_usdt = receive / USDT_RATE

        context.user_data.update({
            "withdraw_amount": amount,
            "withdraw_fee": fee,
            "withdraw_receive": receive
        })
        context.user_data["state"] = "withdraw_method"

        keyboard = [
            [InlineKeyboardButton("💳 bKash", callback_data="method_bkash")],
            [InlineKeyboardButton("💳 Nagad", callback_data="method_nagad")],
            [InlineKeyboardButton("💰 Binance (UID)", callback_data="method_binance")]
        ]
        await update.message.reply_text(
            f"💸 WITHDRAWAL SUMMARY\n\n"
            f"🔹 উত্তোলনের পরিমাণ: ৳{amount:.2f} BDT\n"
            f"🔹 ফি ({WITHDRAW_FEE_PERCENT}%): ৳{fee:.2f} BDT\n"
            f"🔹 আপনি পাবেন: ৳{receive:.2f} BDT (${receive_usdt:.2f} USDT)\n\n"
            f"পেমেন্ট মেথড নির্বাচন করুন:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    if state == "withdraw_account":
        account = text
        method = context.user_data.get("withdraw_method")

        if method in ["Bkash", "Nagad"]:
            if not BD_PHONE_REGEX.match(account):
                await update.message.reply_text(
                    f"❌ ভুল {method} নম্বর!\n\n"
                    f"দয়া করে ১১ ডিজিটের সঠিক বাংলাদেশি মোবাইল নম্বর দিন (যেমন: 017xxxxxxxx, 018xxxxxxxx)।"
                )
                return

        elif method == "Binance":
            if not BINANCE_UID_REGEX.match(account):
                await update.message.reply_text(
                    "❌ ভুল Binance UID!\n\n"
                    "Binance UID শুধুমাত্র ৭ থেকে ১০ ডিজিটের খাঁটি সংখ্যা (যেমন: 123456789) হয়ে থাকে।"
                )
                return

        amt = context.user_data["withdraw_amount"]
        fee = context.user_data["withdraw_fee"]
        rec = context.user_data["withdraw_receive"]
        rec_usdt = rec / USDT_RATE

        db_execute("""
            INSERT INTO withdrawals (user_id, amount, fee, receive_amount, method, account, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (user.id, amt, fee, rec, method, account, datetime.now().isoformat()))

        w_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        context.user_data.clear()

        await update.message.reply_text(
            f"✅ উইথড্র রিকোয়েস্ট সফলভাবে জমা হয়েছে!\n\n"
            f"🔹 মেথড: {method}\n"
            f"🔹 অ্যাকাউন্ট/UID: {account}\n"
            f"🔹 আপনি পাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\n"
            f"⏳ অ্যাডমিন ভেরিফাই করে পেমেন্ট পাঠিয়ে দেবে।"
        )

        if admin_id:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")]
            ])
            await context.bot.send_message(
                chat_id=admin_id,
                text=(
                    f"🔔 NEW WITHDRAWAL REQUEST (#{w_id})\n\n"
                    f"👤 ইউজার: @{user.username or 'No username'} ({user.id})\n"
                    f"💰 রিকোয়েস্ট: ৳{amt:.2f} BDT\n"
                    f"💸 প্রদানযোগ্য: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n"
                    f"💳 মেথড: {method}\n"
                    f"📌 অ্যাকাউন্ট/UID: {account}"
                ),
                reply_markup=keyboard
            )
        return

    # File Submission
    if state == "waiting_file":
        doc = update.message.document
        if not doc:
            await update.message.reply_text("❌ অনুগ্রহ করে এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।")
            return

        file_name = doc.file_name.lower()
        if not (file_name.endswith(".xlsx") or file_name.endswith(".xls") or file_name.endswith(".csv")):
            await update.message.reply_text("❌ শুধুমাত্র Excel (.xlsx, .xls) অথবা CSV (.csv) ফাইল গ্রহণ করা হয়।")
            return

        os.makedirs("downloads", exist_ok=True)
        local_path = os.path.join("downloads", f"{user.id}_{doc.file_name}")
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(local_path)

        try:
            total_found, valid_count, sample_list = validate_gmail_file(local_path)
        except Exception as e:
            if os.path.exists(local_path):
                os.remove(local_path)
            await update.message.reply_text(f"❌ ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        if os.path.exists(local_path):
            os.remove(local_path)

        if valid_count == 0:
            await update.message.reply_text(
                "❌ ফাইলে কোনো সঠিক Gmail অ্যাকাউন্ট পাওয়া যায়নি!\n"
                "মনে রাখবেন: শুধু `@gmail.com` সাপোর্টেড। ইয়াহু বা অন্য কোনো মেইল বা সাধারণ টেক্সট গ্রহণযোগ্য নয়।"
            )
            return

        db_execute("""
            INSERT INTO submissions (user_id, file_id, valid_emails, status, created_at)
            VALUES (?, ?, ?, 'pending', ?)
        """, (user.id, doc.file_id, valid_count, datetime.now().isoformat()))

        sub_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        increase_file_count(user.id)
        context.user_data.clear()

        await update.message.reply_text(
            f"✅ ফাইল সফলভাবে জমা হয়েছে!\n\n"
            f"🔍 যাচাইকৃত ফ্রেশ Gmail পাওয়া গেছে: {valid_count} টি\n"
            f"⏳ অ্যাডমিন ভেরিফাই করলেই আপনার অ্যাকাউন্টে ৳{REWARD_PER_FILE} BDT যোগ হবে।"
        )

        if admin_id:
            sample_str = "\n".join(sample_list)
            admin_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_sub:{sub_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_sub:{sub_id}")]
            ])
            caption = (
                f"📁 NEW GMAIL SUBMISSION (#{sub_id})\n\n"
                f"👤 ইউজার: @{user.username or 'No username'} ({user.id})\n"
                f"✉️ ভ্যালিড Gmail সংখ্যা: {valid_count}\n"
                f"📋 কিছু মেইল নমুনা:\n{sample_str}"
            )
            await context.bot.send_document(
                chat_id=admin_id,
                document=doc.file_id,
                caption=caption,
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
            method = "Bkash"
            hint = "আপনার ১১ ডিজিটের বিকাশ নম্বর দিন (যেমন: 017xxxxxxxx):"
        elif raw_m == "nagad":
            method = "Nagad"
            hint = "আপনার ১১ ডিজিটের নগদ নম্বর দিন (যেমন: 018xxxxxxxx):"
        else:
            method = "Binance"
            hint = "আপনার বাইনান্স ইউআইডি (Binance UID) দিন (৭-১০ ডিজিটের সংখ্যা):"

        context.user_data["withdraw_method"] = method
        context.user_data["state"] = "withdraw_account"
        await query.edit_message_text(f"💳 নির্বাচিত মেথড: {method}\n\n👉 {hint}")


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
