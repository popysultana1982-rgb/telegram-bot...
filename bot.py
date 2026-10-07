import os
import re
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

# রেফারেল বোনাস এবং ডলার রেট
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
# ADMIN & USER HELPERS
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


def add_user(user, referrer_id=None):
    existing = db_execute("SELECT * FROM users WHERE user_id=?", (user.id,), fetchone=True)
    if not existing:
        db_execute("""
            INSERT INTO users (user_id, username, balance, files_today, last_file_date, referred_by, created_at)
            VALUES (?, ?, 0, 0, ?, ?, ?)
        """, (user.id, user.username or "", datetime.now().strftime("%Y-%m-%d"), referrer_id, datetime.now().isoformat()))
        return True  # নতুন ইউজার
    else:
        db_execute("UPDATE users SET username=? WHERE user_id=?", (user.username or "", user.id))
        return False  # পুরাতন ইউজার


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
# RULES & WELCOME TEXT
# =========================================================

def get_rules_text(user_name):
    return (
        f"আসসালামু আলাইকুম, {user_name}! 🌸\n\n"
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


# =========================================================
# MAIN MENU (BOTTOM KEYBOARD WITH RULES)
# =========================================================

def get_bottom_keyboard():
    keyboard = [
        [KeyboardButton("📤 SELL FRESH GMAIL ACCOUNT")],
        [KeyboardButton("💰 BALANCE"), KeyboardButton("💸 WITHDRAW")],
        [KeyboardButton("👥 REFERRAL"), KeyboardButton("📜 RULES")],
        [KeyboardButton("📞 SUPPORT")]
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
    context.user_data.clear()

    # রেফারেল ট্র্যাকিং
    referrer_id = None
    if context.args and len(context.args) > 0:
        try:
            potential_ref = int(context.args[0])
            if potential_ref != user.id:
                if get_user(potential_ref):
                    referrer_id = potential_ref
        except ValueError:
            referrer_id = None

    is_new = add_user(user, referrer_id)

    # নতুন ইউজার রেফারেল বোনাস
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
        await update.message.reply_text("ব্যবহার নিয়ম: `/addbalance <user_id> <amount>`\nউদাহরণ: `/addbalance 8919985167 50`")
        return

    try:
        target_user_id = int(context.args[0])
        amount = float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ ইউজার আইডি এবং টাকার পরিমাণ সংখ্যায় হতে হবে।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ এই আইডির কোনো ইউজার পাওয়া যায়নি।")
        return

    update_balance(target_user_id, amount)
    new_bdt = get_balance(target_user_id)
    new_usdt = new_bdt / USDT_RATE

    await update.message.reply_text(
        f"✅ ব্যালেন্স যোগ হয়েছে!\n\n"
        f"👤 ইউজার ID: {target_user_id}\n"
        f"➕ যোগ: ৳{amount:.2f} BDT\n"
        f"💰 মোট ব্যালেন্স: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
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
        await update.message.reply_text("ব্যবহার নিয়ম: `/cutbalance <user_id> <amount>`\nউদাহরণ: `/cutbalance 8919985167 20`")
        return

    try:
        target_user_id = int(context.args[0])
        amount = float(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ ইউজার আইডি এবং টাকার পরিমাণ সংখ্যায় হতে হবে।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ এই আইডির কোনো ইউজার পাওয়া যায়নি।")
        return

    update_balance(target_user_id, -amount)
    new_bdt = get_balance(target_user_id)
    new_usdt = new_bdt / USDT_RATE

    await update.message.reply_text(
        f"✂️ ব্যালেন্স কেটে নেওয়া হয়েছে!\n\n"
        f"👤 ইউজার ID: {target_user_id}\n"
        f"➖ কাটা হয়েছে: ৳{amount:.2f} BDT\n"
        f"💰 বর্তমান ব্যালেন্স: ৳{new_bdt:.2f} BDT (${new_usdt:.2f} USDT)"
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
        await update.message.reply_text("ব্যবহার নিয়ম: `/resetbalance <user_id>`\nউদাহরণ: `/resetbalance 8919985167`")
        return

    try:
        target_user_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ ইউজার আইডি সংখ্যায় হতে হবে।")
        return

    user = get_user(target_user_id)
    if not user:
        await update.message.reply_text("❌ এই আইডির কোনো ইউজার পাওয়া যায়নি।")
        return

    db_execute("UPDATE users SET balance=0 WHERE user_id=?", (target_user_id,))

    await update.message.reply_text(
        f"🔄 ব্যালেন্স রিসেট সফল হয়েছে!\n\n"
        f"👤 ইউজার ID: {target_user_id}\n"
        f"💰 বর্তমান ব্যালেন্স: ৳0.00 BDT ($0.00 USDT)"
    )

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text="⚠️ আপনার অ্যাকাউন্টের ব্যালেন্স রিসেট করে ৳0.00 ($0.00 USDT) করা হয়েছে।"
        )
    except Exception:
        pass


async def show_admin_panel(query):
    keyboard = [
        [InlineKeyboardButton("➕ Add Button", callback_data="admin_add_button")],
        [InlineKeyboardButton("🗑 Remove Button", callback_data="admin_remove_button")],
        [InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast")],
        [InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message")],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")],
        [InlineKeyboardButton("⬅️ Close Panel", callback_data="admin_close")]
    ]
    await query.edit_message_text("👨‍💼 ADMIN PANEL\n\nএকটি অপশন নির্বাচন করুন:", reply_markup=InlineKeyboardMarkup(keyboard))


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

    add_user(user)
    text = update.message.text.strip() if update.message.text else ""
    state = context.user_data.get("state")
    admin_id = get_admin_id()

    # --- বাটন অ্যাকশনসমূহ ---
    if text == "📤 SELL FRESH GMAIL ACCOUNT":
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

    elif text == "💰 BALANCE":
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

    elif text == "💸 WITHDRAW":
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

    elif text == "👥 REFERRAL":
        bot_info = await context.bot.get_me()
        bot_username = bot_info.username
        referral_link = f"https://t.me/{bot_username}?start={user.id}"
        ref_count = get_referral_count(user.id)
        total_earned = ref_count * REFERRAL_BONUS

        msg = (
            f"👥 রেফার করে ইনকাম করুন!\n\n"
            f"আপনি প্রতি সফল রেফারে পাবেন ৳{REFERRAL_BONUS:.2f} BDT বোনাস।\n"
            f"আপনার বন্ধু বা অন্যদের কাছে আপনার রেফার লিংকটি শেয়ার করুন। কেউ এই লিংকে ক্লিক করে বটে যুক্ত হওয়ার সাথে সাথে আপনার অ্যাকাউন্টে টাকা যোগ হয়ে যাবে।\n\n"
            f"🔗 আপনার রেফারেল লিংক:\n`{referral_link}`\n\n"
            f"📊 আপনার রেফারেল পরিসংখ্যান:\n"
            f"🔹 মোট রেফার করেছেন: {ref_count} জন\n"
            f"🔹 রেফারেল থেকে আয়: ৳{total_earned:.2f} BDT"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")
        return

    elif text == "📜 RULES":
        name = user.first_name or "User"
        rules_msg = get_rules_text(name)
        await update.message.reply_text(rules_msg)
        return

    elif text == "📞 SUPPORT":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("💬 মেসেজ পাঠান", url="https://t.me/Talha_juba098")]
        ])
        await update.message.reply_text("যে কোনো সমস্যা বা সহযোগিতার জন্য সরাসরি সাপোর্টে যোগাযোগ করুন:", reply_markup=keyboard)
        return

    # --- অ্যাডমিন বাটন তৈরি ---
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

    # --- অ্যাডমিন ব্রডকাস্ট ---
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

    # --- অ্যাডমিন সিঙ্গেল মেসেজ ---
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

    # --- উইথড্র পরিমাণ ইনপুট ---
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

    # --- উইথড্র একাউন্ট নম্বর / UID ইনপুট ও ভ্যালিডেশন ---
    if state == "withdraw_account":
        account = text
        method = context.user_data.get("withdraw_method")

        # ১. বিকাশ ও নগদ: ১১ ডিজিটের বাংলাদেশি নম্বর ভ্যালিডেশন
        if method in ["Bkash", "Nagad"]:
            if not BD_PHONE_REGEX.match(account):
                await update.message.reply_text(
                    f"❌ ভুল {method} নম্বর!\n\n"
                    f"দয়া করে ১১ ডিজিটের সঠিক বাংলাদেশি মোবাইল নম্বর দিন (যেমন: 017xxxxxxxx, 018xxxxxxxx)।\n"
                    f"কোনো স্পেস বা বাড়তি অক্ষর গ্রহণযোগ্য নয়।"
                )
                return

        # ২. বাইনান্স: সঠিক UID ভ্যালিডেশন (৭-১০ ডিজিটের সংখ্যা)
        elif method == "Binance":
            if not BINANCE_UID_REGEX.match(account):
                await update.message.reply_text(
                    "❌ ভুল Binance UID!\n\n"
                    "Binance UID শুধুমাত্র ৭ থেকে ১০ ডিজিটের খাঁটি সংখ্যা (যেমন: 123456789) হয়ে থাকে।\n"
                    "কোনো অক্ষর, স্পেস বা আলফানিউমেরিক টেক্সট দেওয়া যাবে না।"
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

    # --- ফাইল সাবমিশন ভ্যালিডেশন ---
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
