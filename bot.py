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

# এই পাসওয়ার্ডটি দিয়ে টেলিগ্রাম থেকে অ্যাডমিন হবেন
ADMIN_SECRET_KEY = os.getenv("ADMIN_SECRET_KEY", "mysecretadmin123")

MIN_WITHDRAW = 50
WITHDRAW_FEE_PERCENT = 4
REWARD_PER_FILE = 25
DAILY_FILE_LIMIT = 5

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


def add_user(user):
    existing = db_execute("SELECT user_id FROM users WHERE user_id=?", (user.id,), fetchone=True)
    if not existing:
        db_execute("""
            INSERT INTO users (user_id, username, balance, files_today, last_file_date, created_at)
            VALUES (?, ?, 0, 0, ?, ?)
        """, (user.id, user.username or "", datetime.now().strftime("%Y-%m-%d"), datetime.now().isoformat()))
    else:
        db_execute("UPDATE users SET username=? WHERE user_id=?", (user.username or "", user.id))


def get_user(user_id):
    return db_execute("SELECT * FROM users WHERE user_id=?", (user_id,), fetchone=True)


def get_balance(user_id):
    user = get_user(user_id)
    return float(user["balance"]) if user else 0


def update_balance(user_id, amount):
    db_execute("UPDATE users SET balance = balance + ? WHERE user_id=?", (amount, user_id))


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
# EMAIL & EXCEL VALIDATION
# =========================================================

GMAIL_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@gmail\.com$", re.IGNORECASE)


def validate_gmail_file(file_path):
    emails = []
    if file_path.endswith(".csv"):
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                for token in line.strip().split(","):
                    val = token.strip()
                    if "@" in val:
                        emails.append(val)
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell and isinstance(cell, str) and "@" in cell:
                    emails.append(cell.strip())

    valid_gmails = [m for m in emails if GMAIL_REGEX.match(m)]
    return len(emails), len(valid_gmails), valid_gmails[:5]


# =========================================================
# MAIN MENU
# =========================================================

async def show_main_menu(update, context):
    user = update.effective_user
    add_user(user)
    balance = get_balance(user.id)

    keyboard = [
        [
            InlineKeyboardButton("📤 Sell", callback_data="sell"),
            InlineKeyboardButton("💰 Balance", callback_data="balance")
        ],
        [
            InlineKeyboardButton("💸 Withdraw", callback_data="withdraw"),
            InlineKeyboardButton("📞 Support", url="https://t.me/Talha_juba098")
        ]
    ]

    buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
    for button in buttons:
        keyboard.append([InlineKeyboardButton(button["title"], url=button["url"])])

    text = f"🤖 Welcome!\n\n💰 Balance: ৳{balance:.2f}\n\nChoose an option below:"

    if update.callback_query:
        await update.callback_query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard))
    else:
        await update.message.reply_text(text, reply_markup=InlineKeyboardMarkup(keyboard))


# =========================================================
# COMMAND HANDLERS
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    add_user(update.effective_user)
    context.user_data.clear()
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
        await update.message.reply_text("❌ গোপন পাসওয়ার্ডটি ভুল!")


async def admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if not admin_id or update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.\nপ্রথমে নিজেকে অ্যাডমিন করতে `/setadmin <key>` কমান্ড পাঠান।")
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
        await update.message.reply_text("ব্যবহার নিয়ম: `/addbalance <user_id> <amount>`\nউদাহরণ: `/addbalance 123456789 50`")
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
    new_balance = get_balance(target_user_id)

    await update.message.reply_text(
        f"✅ ব্যালেন্স যোগ হয়েছে!\n\n"
        f"👤 ইউজার ID: {target_user_id}\n"
        f"➕ যোগ: ৳{amount:.2f}\n"
        f"💰 বর্তমান মোট ব্যালেন্স: ৳{new_balance:.2f}"
    )

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text=f"🎁 আপনার অ্যাকাউন্টে ৳{amount:.2f} যোগ করা হয়েছে!\nআপনার বর্তমান ব্যালেন্স: ৳{new_balance:.2f}"
        )
    except Exception:
        pass


async def cut_balance_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    admin_id = get_admin_id()
    if update.effective_user.id != admin_id:
        await update.message.reply_text("❌ You are not authorized.")
        return

    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার নিয়ম: `/cutbalance <user_id> <amount>`\nউদাহরণ: `/cutbalance 123456789 20`")
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
    new_balance = get_balance(target_user_id)

    await update.message.reply_text(
        f"✂️ ব্যালেন্স কেটে নেওয়া হয়েছে!\n\n"
        f"👤 ইউজার ID: {target_user_id}\n"
        f"➖ কাটা হয়েছে: ৳{amount:.2f}\n"
        f"💰 বর্তমান মোট ব্যালেন্স: ৳{new_balance:.2f}"
    )

    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text=f"⚠️ আপনার অ্যাকাউন্ট থেকে ৳{amount:.2f} কেটে নেওয়া হয়েছে।\nআপনার বর্তমান ব্যালেন্স: ৳{new_balance:.2f}"
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
        [InlineKeyboardButton("⬅️ Main Menu", callback_data="home")]
    ]
    await query.edit_message_text("👨‍💼 ADMIN PANEL\n\nChoose an action:", reply_markup=InlineKeyboardMarkup(keyboard))


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

    if data == "balance":
        balance = get_balance(user_id)
        await query.edit_message_text(
            f"💰 Your Balance\n\nAvailable Balance: ৳{balance:.2f}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="home")]])
        )

    elif data == "sell":
        count = get_today_file_count(user_id)
        if count >= DAILY_FILE_LIMIT:
            await query.edit_message_text(
                f"❌ Daily limit reached.\n\nYou can submit maximum {DAILY_FILE_LIMIT} files per day.",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="home")]])
            )
            return

        context.user_data["state"] = "waiting_file"
        await query.edit_message_text(
            f"📤 Send your Excel file (.xlsx, .xls, .csv) containing Gmail addresses now.\n\n"
            f"Today's submissions: {count}/{DAILY_FILE_LIMIT}\n\n"
            f"Approved file reward: ৳{REWARD_PER_FILE}"
        )

    elif data == "withdraw":
        balance = get_balance(user_id)
        if balance < MIN_WITHDRAW:
            await query.edit_message_text(
                f"❌ Minimum withdrawal is ৳{MIN_WITHDRAW}.\n\nYour current balance: ৳{balance:.2f}",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="home")]])
            )
            return

        context.user_data["state"] = "withdraw_amount"
        await query.edit_message_text(
            f"💸 Withdrawal\n\nAvailable Balance: ৳{balance:.2f}\nMinimum Withdrawal: ৳{MIN_WITHDRAW}\nFee: {WITHDRAW_FEE_PERCENT}%\n\nEnter the amount you want to withdraw:"
        )

    elif data == "home":
        context.user_data.clear()
        await show_main_menu(update, context)

    elif data == "admin_panel":
        if user_id != admin_id:
            return
        await show_admin_panel(query)

    elif data == "admin_add_button":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_add_button_title"
        await query.edit_message_text("➕ Add New Button\n\nFirst send the button name.\n\nExample:\n📢 Our Channel")

    elif data == "admin_remove_button":
        if user_id != admin_id:
            return
        buttons = db_execute("SELECT * FROM buttons ORDER BY id DESC", fetchall=True)
        if not buttons:
            await query.edit_message_text(
                "There are no custom buttons.",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")]])
            )
            return

        keyboard = [[InlineKeyboardButton(f"❌ {b['title']}", callback_data=f"delete_button:{b['id']}")] for b in buttons]
        keyboard.append([InlineKeyboardButton("⬅️ Back", callback_data="admin_panel")])
        await query.edit_message_text("🗑 Select a button to remove:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data == "admin_broadcast":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 Broadcast Post\n\nSend the message you want to send to all users.")

    elif data == "admin_single_message":
        if user_id != admin_id:
            return
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ Single User Message\n\nSend the Telegram User ID first.")

    elif data == "admin_stats":
        if user_id != admin_id:
            return
        users = db_execute("SELECT COUNT(*) AS c FROM users", fetchone=True)["c"]
        total_balance = db_execute("SELECT COALESCE(SUM(balance),0) AS total FROM users", fetchone=True)["total"]
        pending = db_execute("SELECT COUNT(*) AS c FROM withdrawals WHERE status='pending'", fetchone=True)["c"]
        pending_files = db_execute("SELECT COUNT(*) AS c FROM submissions WHERE status='pending'", fetchone=True)["c"]

        await query.edit_message_text(
            f"📊 Admin Statistics\n\n👤 Total Users: {users}\n💰 Total Balance: ৳{float(total_balance):.2f}\n💸 Pending Withdrawals: {pending}\n📁 Pending Files: {pending_files}",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]])
        )

    # Submission Approval
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
        await query.edit_message_text(f"✅ Submission #{sub_id} Approved! ৳{REWARD_PER_FILE} added to user.")
        try:
            await context.bot.send_message(
                chat_id=sub["user_id"],
                text=f"🎉 Your submitted file has been approved!\n৳{REWARD_PER_FILE} has been added to your balance."
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
            await context.bot.send_message(chat_id=sub["user_id"], text="❌ Your submitted file was rejected by admin.")
        except Exception:
            pass

    # Withdrawal Approval
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
            await query.edit_message_text("❌ Withdrawal rejected: User balance insufficient.")
            return

        update_balance(withdrawal["user_id"], -float(withdrawal["amount"]))
        db_execute("UPDATE withdrawals SET status='approved' WHERE id=?", (withdrawal_id,))
        await query.edit_message_text(f"✅ Withdrawal Approved.\nUser receives: ৳{withdrawal['receive_amount']:.2f}")

        try:
            await context.bot.send_message(
                chat_id=withdrawal["user_id"],
                text=f"✅ Your withdrawal has been approved.\nAmount: ৳{withdrawal['receive_amount']:.2f}"
            )
        except Exception:
            pass

    elif data.startswith("reject_withdraw:"):
        if user_id != admin_id:
            return
        withdrawal_id = int(data.split(":")[1])
        db_execute("UPDATE withdrawals SET status='rejected' WHERE id=?", (withdrawal_id,))
        await query.edit_message_text("❌ Withdrawal Rejected.")

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
    state = context.user_data.get("state")
    admin_id = get_admin_id()

    # Admin Add Button
    if user.id == admin_id and state == "admin_add_button_title":
        context.user_data["button_title"] = update.message.text.strip()
        context.user_data["state"] = "admin_add_button_url"
        await update.message.reply_text("Now send the button URL (e.g., https://t.me/yourchannel):")
        return

    if user.id == admin_id and state == "admin_add_button_url":
        url = update.message.text.strip()
        if not url.startswith(("http://", "https://", "tg://")):
            await update.message.reply_text("❌ Invalid URL. Send a valid URL.")
            return
        title = context.user_data.get("button_title")
        db_execute("INSERT INTO buttons (title, url) VALUES (?, ?)", (title, url))
        context.user_data.clear()
        await update.message.reply_text(f"✅ Button added:\n{title} -> {url}")
        return

    # Admin Broadcast
    if user.id == admin_id and state == "admin_broadcast":
        users = db_execute("SELECT user_id FROM users", fetchall=True)
        success = 0
        for u in users:
            try:
                await context.bot.send_message(chat_id=u["user_id"], text=update.message.text)
                success += 1
            except Exception:
                pass
        context.user_data.clear()
        await update.message.reply_text(f"📢 Broadcast completed: {success}/{len(users)}")
        return

    # Admin Single Message
    if user.id == admin_id and state == "admin_single_user":
        try:
            context.user_data["target_user"] = int(update.message.text.strip())
            context.user_data["state"] = "admin_single_text"
            await update.message.reply_text("Now send the message you want to send.")
        except ValueError:
            await update.message.reply_text("❌ Send a valid numeric User ID.")
        return

    if user.id == admin_id and state == "admin_single_text":
        target_id = context.user_data.get("target_user")
        try:
            await context.bot.send_message(chat_id=target_id, text=update.message.text)
            await update.message.reply_text("✅ Message sent.")
        except Exception as e:
            await update.message.reply_text(f"❌ Failed to send: {e}")
        context.user_data.clear()
        return

    # Withdraw Handling
    if state == "withdraw_amount":
        try:
            amount = float(update.message.text.strip())
        except ValueError:
            await update.message.reply_text("❌ Enter a valid number.")
            return

        if amount < MIN_WITHDRAW:
            await update.message.reply_text(f"❌ Minimum withdrawal is ৳{MIN_WITHDRAW}.")
            return

        balance = get_balance(user.id)
        if amount > balance:
            await update.message.reply_text(f"❌ Insufficient balance. Available: ৳{balance:.2f}")
            return

        fee = amount * WITHDRAW_FEE_PERCENT / 100
        receive = amount - fee
        context.user_data.update({"withdraw_amount": amount, "withdraw_fee": fee, "withdraw_receive": receive})
        context.user_data["state"] = "withdraw_method"

        keyboard = [
            [InlineKeyboardButton("💳 bKash", callback_data="method_bkash")],
            [InlineKeyboardButton("💳 Nagad", callback_data="method_nagad")],
            [InlineKeyboardButton("💰 Binance", callback_data="method_binance")]
        ]
        await update.message.reply_text(
            f"💸 Summary:\nAmount: ৳{amount:.2f}\nFee: ৳{fee:.2f}\nYou Receive: ৳{receive:.2f}\n\nSelect method:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    if state == "withdraw_account":
        account = update.message.text.strip()
        amt = context.user_data["withdraw_amount"]
        fee = context.user_data["withdraw_fee"]
        rec = context.user_data["withdraw_receive"]
        mth = context.user_data["withdraw_method"]

        db_execute("""
            INSERT INTO withdrawals (user_id, amount, fee, receive_amount, method, account, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (user.id, amt, fee, rec, mth, account, datetime.now().isoformat()))

        w_id = db_execute("SELECT last_insert_rowid() AS id", fetchone=True)["id"]
        context.user_data.clear()
        await update.message.reply_text("✅ Withdrawal request submitted! Waiting for admin approval.")

        if admin_id:
            keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")]
            ])
            await context.bot.send_message(
                chat_id=admin_id,
                text=f"🔔 NEW WITHDRAWAL REQUEST\n\nUser: @{user.username or 'No username'} ({user.id})\nAmount: ৳{amt:.2f}\nReceive: ৳{rec:.2f}\nMethod: {mth}\nAccount: {account}",
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
                "❌ ফাইলে কোনো ভ্যালিড Gmail অ্যাড্রেস (@gmail.com) পাওয়া যায়নি। ফরম্যাট ঠিক করে আবার পাঠান।"
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
            f"✅ ফাইল গৃহীত হয়েছে!\n\n"
            f"🔍 মোট জিমেইল পাওয়া গেছে: {valid_count} টি\n"
            f"⏳ অ্যাডমিন ভেরিফাই করার পর আপনার ব্যালেন্সে ৳{REWARD_PER_FILE} যোগ হবে।"
        )

        if admin_id:
            sample_str = "\n".join(sample_list)
            admin_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Approve", callback_data=f"approve_sub:{sub_id}"),
                 InlineKeyboardButton("❌ Reject", callback_data=f"reject_sub:{sub_id}")]
            ])
            caption = (
                f"📁 NEW GMAIL FILE SUBMISSION (#{sub_id})\n\n"
                f"👤 User: @{user.username or 'No username'} ({user.id})\n"
                f"✉️ ভ্যালিড Gmail সংখ্যা: {valid_count}\n"
                f"📋 নমুনা (Sample):\n{sample_str}"
            )
            await context.bot.send_document(
                chat_id=admin_id,
                document=doc.file_id,
                caption=caption,
                reply_markup=admin_kb
            )
        return

    await update.message.reply_text("Please use the buttons in the menu.\nSend /start to open the menu.")


async def withdrawal_method_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data.startswith("method_"):
        method = query.data.replace("method_", "").capitalize()
        context.user_data["withdraw_method"] = method
        context.user_data["state"] = "withdraw_account"
        await query.edit_message_text(f"💳 Selected: {method}\n\nNow send your account number / wallet address.")


# =========================================================
# MAIN
# =========================================================

def main():
    init_db()

    web_thread = threading.Thread(target=run_web_server, daemon=True)
    web_thread.start()

    application = Application.builder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("admin", admin))
    application.add_handler(CommandHandler("setadmin", set_admin_cmd))
    application.add_handler(CommandHandler("addbalance", add_balance_cmd))
    application.add_handler(CommandHandler("cutbalance", cut_balance_cmd))

    application.add_handler(CallbackQueryHandler(withdrawal_method_handler, pattern=r"^method_"))
    application.add_handler(CallbackQueryHandler(button_handler))
    application.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, message_handler))

    print("Telegram bot is running...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
