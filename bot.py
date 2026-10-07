import os
import sqlite3
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

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

ADMIN_ID = 8919985167

MIN_WITHDRAW = 50
WITHDRAW_FEE_PERCENT = 4
REWARD_PER_FILE = 25
DAILY_FILE_LIMIT = 5

DB_NAME = "bot.db"

# Render automatically provides PORT
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
    server = ThreadingHTTPServer(
        ("0.0.0.0", PORT),
        HealthHandler
    )

    print(f"Web server running on port {PORT}")

    server.serve_forever()


# =========================================================
# DATABASE
# =========================================================

db = sqlite3.connect(
    DB_NAME,
    check_same_thread=False
)

db.row_factory = sqlite3.Row


def db_execute(
    query,
    params=(),
    fetchone=False,
    fetchall=False
):

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
            status TEXT DEFAULT 'pending',
            created_at TEXT
        )
    """)


# =========================================================
# USER FUNCTIONS
# =========================================================

def add_user(user):

    existing = db_execute(
        "SELECT user_id FROM users WHERE user_id=?",
        (user.id,),
        fetchone=True
    )

    if not existing:

        db_execute("""
            INSERT INTO users
            (
                user_id,
                username,
                balance,
                files_today,
                last_file_date,
                created_at
            )
            VALUES (?, ?, 0, 0, ?, ?)
        """, (
            user.id,
            user.username or "",
            datetime.now().strftime("%Y-%m-%d"),
            datetime.now().isoformat()
        ))

    else:

        db_execute(
            """
            UPDATE users
            SET username=?
            WHERE user_id=?
            """,
            (
                user.username or "",
                user.id
            )
        )


def get_user(user_id):

    return db_execute(
        "SELECT * FROM users WHERE user_id=?",
        (user_id,),
        fetchone=True
    )


def get_balance(user_id):

    user = get_user(user_id)

    if not user:
        return 0

    return float(user["balance"])


def update_balance(user_id, amount):

    db_execute(
        """
        UPDATE users
        SET balance = balance + ?
        WHERE user_id=?
        """,
        (
            amount,
            user_id
        )
    )


def get_today_file_count(user_id):

    user = get_user(user_id)

    today = datetime.now().strftime("%Y-%m-%d")

    if not user:
        return 0

    if user["last_file_date"] != today:

        db_execute("""
            UPDATE users
            SET files_today=0,
                last_file_date=?
            WHERE user_id=?
        """, (
            today,
            user_id
        ))

        return 0

    return user["files_today"]


def increase_file_count(user_id):

    today = datetime.now().strftime("%Y-%m-%d")

    db_execute("""
        UPDATE users
        SET files_today = files_today + 1,
            last_file_date = ?
        WHERE user_id=?
    """, (
        today,
        user_id
    ))


# =========================================================
# MAIN MENU
# =========================================================

async def show_main_menu(update, context):

    user = update.effective_user

    add_user(user)

    balance = get_balance(user.id)

    keyboard = [
        [
            InlineKeyboardButton(
                "📤 Sell",
                callback_data="sell"
            ),
            InlineKeyboardButton(
                "💰 Balance",
                callback_data="balance"
            )
        ],
        [
            InlineKeyboardButton(
                "💸 Withdraw",
                callback_data="withdraw"
            ),
            InlineKeyboardButton(
                "📞 Support",
                url="https://t.me/Talha_juba098"
            )
        ]
    ]

    buttons = db_execute(
        "SELECT * FROM buttons ORDER BY id DESC",
        fetchall=True
    )

    for button in buttons:

        keyboard.append([
            InlineKeyboardButton(
                button["title"],
                url=button["url"]
            )
        ])

    text = (
        "🤖 Welcome!\n\n"
        f"💰 Balance: ৳{balance:.2f}\n\n"
        "Choose an option below:"
    )

    if update.callback_query:

        await update.callback_query.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    else:

        await update.message.reply_text(
            text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


# =========================================================
# START
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    add_user(update.effective_user)

    context.user_data.clear()

    await show_main_menu(update, context)


# =========================================================
# ADMIN PANEL
# =========================================================

async def show_admin_panel(query):

    keyboard = [

        [
            InlineKeyboardButton(
                "➕ Add Button",
                callback_data="admin_add_button"
            )
        ],

        [
            InlineKeyboardButton(
                "🗑 Remove Button",
                callback_data="admin_remove_button"
            )
        ],

        [
            InlineKeyboardButton(
                "📢 Broadcast Post",
                callback_data="admin_broadcast"
            )
        ],

        [
            InlineKeyboardButton(
                "✉️ Single Message",
                callback_data="admin_single_message"
            )
        ],

        [
            InlineKeyboardButton(
                "📊 Statistics",
                callback_data="admin_stats"
            )
        ],

        [
            InlineKeyboardButton(
                "⬅️ Main Menu",
                callback_data="home"
            )
        ]
    ]

    await query.edit_message_text(
        "👨‍💼 ADMIN PANEL\n\n"
        "Choose an action:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


# =========================================================
# CALLBACK HANDLER
# =========================================================

async def button_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_id = query.from_user.id

    add_user(query.from_user)

    data = query.data

    # =====================================================
    # BALANCE
    # =====================================================

    if data == "balance":

        balance = get_balance(user_id)

        await query.edit_message_text(
            "💰 Your Balance\n\n"
            f"Available Balance: ৳{balance:.2f}",
            reply_markup=InlineKeyboardMarkup([
                [
                    InlineKeyboardButton(
                        "⬅️ Back",
                        callback_data="home"
                    )
                ]
            ])
        )

    # =====================================================
    # SELL
    # =====================================================

    elif data == "sell":

        count = get_today_file_count(user_id)

        if count >= DAILY_FILE_LIMIT:

            await query.edit_message_text(
                "❌ Daily limit reached.\n\n"
                f"You can submit maximum "
                f"{DAILY_FILE_LIMIT} files per day.",
                reply_markup=InlineKeyboardMarkup([
                    [
                        InlineKeyboardButton(
                            "⬅️ Back",
                            callback_data="home"
                        )
                    ]
                ])
            )

            return

        context.user_data["state"] = "waiting_file"

        await query.edit_message_text(
            "📤 Send your file now.\n\n"
            f"Today's submissions: "
            f"{count}/{DAILY_FILE_LIMIT}\n\n"
            f"Approved file reward: ৳{REWARD_PER_FILE}"
        )

    # =====================================================
    # WITHDRAW
    # =====================================================

    elif data == "withdraw":

        balance = get_balance(user_id)

        if balance < MIN_WITHDRAW:

            await query.edit_message_text(
                "❌ Minimum withdrawal is ৳50.\n\n"
                f"Your current balance: ৳{balance:.2f}",
                reply_markup=InlineKeyboardMarkup([
                    [
                        InlineKeyboardButton(
                            "⬅️ Back",
                            callback_data="home"
                        )
                    ]
                ])
            )

            return

        context.user_data["state"] = "withdraw_amount"

        await query.edit_message_text(
            "💸 Withdrawal\n\n"
            f"Available Balance: ৳{balance:.2f}\n"
            f"Minimum Withdrawal: ৳{MIN_WITHDRAW}\n"
            f"Fee: {WITHDRAW_FEE_PERCENT}%\n\n"
            "Enter the amount you want to withdraw:"
        )

    # =====================================================
    # HOME
    # =====================================================

    elif data == "home":

        context.user_data.clear()

        await show_main_menu(update, context)

    # =====================================================
    # ADMIN PANEL
    # =====================================================

    elif data == "admin_panel":

        if user_id != ADMIN_ID:
            return

        await show_admin_panel(query)

    # =====================================================
    # ADMIN ADD BUTTON
    # =====================================================

    elif data == "admin_add_button":

        if user_id != ADMIN_ID:
            return

        context.user_data["state"] = "admin_add_button_title"

        await query.edit_message_text(
            "➕ Add New Button\n\n"
            "First send the button name.\n\n"
            "Example:\n"
            "📢 Our Channel"
        )

    # =====================================================
    # ADMIN REMOVE BUTTON
    # =====================================================

    elif data == "admin_remove_button":

        if user_id != ADMIN_ID:
            return

        buttons = db_execute(
            "SELECT * FROM buttons ORDER BY id DESC",
            fetchall=True
        )

        if not buttons:

            await query.edit_message_text(
                "There are no custom buttons.",
                reply_markup=InlineKeyboardMarkup([
                    [
                        InlineKeyboardButton(
                            "⬅️ Back",
                            callback_data="admin_panel"
                        )
                    ]
                ])
            )

            return

        keyboard = []

        for b in buttons:

            keyboard.append([
                InlineKeyboardButton(
                    f"❌ {b['title']}",
                    callback_data=f"delete_button:{b['id']}"
                )
            ])

        keyboard.append([
            InlineKeyboardButton(
                "⬅️ Back",
                callback_data="admin_panel"
            )
        ])

        await query.edit_message_text(
            "🗑 Select a button to remove:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    # =====================================================
    # ADMIN BROADCAST
    # =====================================================

    elif data == "admin_broadcast":

        if user_id != ADMIN_ID:
            return

        context.user_data["state"] = "admin_broadcast"

        await query.edit_message_text(
            "📢 Broadcast Post\n\n"
            "Send the message you want to send to all users."
        )

    # =====================================================
    # ADMIN SINGLE MESSAGE
    # =====================================================

    elif data == "admin_single_message":

        if user_id != ADMIN_ID:
            return

        context.user_data["state"] = "admin_single_user"

        await query.edit_message_text(
            "✉️ Single User Message\n\n"
            "Send the Telegram User ID first."
        )

    # =====================================================
    # ADMIN STATS
    # =====================================================

    elif data == "admin_stats":

        if user_id != ADMIN_ID:
            return

        users = db_execute(
            "SELECT COUNT(*) AS c FROM users",
            fetchone=True
        )["c"]

        total_balance = db_execute(
            """
            SELECT COALESCE(SUM(balance),0) AS total
            FROM users
            """,
            fetchone=True
        )["total"]

        pending = db_execute(
            """
            SELECT COUNT(*) AS c
            FROM withdrawals
            WHERE status='pending'
            """,
            fetchone=True
        )["c"]

        pending_files = db_execute(
            """
            SELECT COUNT(*) AS c
            FROM submissions
            WHERE status='pending'
            """,
            fetchone=True
        )["c"]

        await query.edit_message_text(
            "📊 Admin Statistics\n\n"
            f"👤 Total Users: {users}\n"
            f"💰 Total Balance: ৳{float(total_balance):.2f}\n"
            f"💸 Pending Withdrawals: {pending}\n"
            f"📁 Pending Files: {pending_files}",
            reply_markup=InlineKeyboardMarkup([
                [
                    InlineKeyboardButton(
                        "⬅️ Admin Panel",
                        callback_data="admin_panel"
                    )
                ]
            ])
        )

    # =====================================================
    # APPROVE WITHDRAWAL
    # =====================================================

    elif data.startswith("approve_withdraw:"):

        if user_id != ADMIN_ID:
            return

        withdrawal_id = int(
            data.split(":")[1]
        )

        withdrawal = db_execute(
            """
            SELECT *
            FROM withdrawals
            WHERE id=?
            """,
            (withdrawal_id,),
            fetchone=True
        )

        if not withdrawal:

            await query.answer(
                "Withdrawal not found.",
                show_alert=True
            )

            return

        if withdrawal["status"] != "pending":

            await query.answer(
                "Already processed.",
                show_alert=True
            )

            return

        balance = get_balance(
            withdrawal["user_id"]
        )

        if balance < float(withdrawal["amount"]):

            db_execute(
                """
                UPDATE withdrawals
                SET status='rejected'
                WHERE id=?
                """,
                (withdrawal_id,)
            )

            await query.edit_message_text(
                "❌ Withdrawal rejected automatically.\n\n"
                "User does not have enough balance."
            )

            return

        update_balance(
            withdrawal["user_id"],
            -float(withdrawal["amount"])
        )

        db_execute(
            """
            UPDATE withdrawals
            SET status='approved'
            WHERE id=?
            """,
            (withdrawal_id,)
        )

        await query.edit_message_text(
            "✅ Withdrawal Approved\n\n"
            f"Amount: ৳{withdrawal['amount']:.2f}\n"
            f"Fee: ৳{withdrawal['fee']:.2f}\n"
            f"User receives: "
            f"৳{withdrawal['receive_amount']:.2f}"
        )

        try:

            await context.bot.send_message(
                chat_id=withdrawal["user_id"],
                text=(
                    "✅ Your withdrawal has been approved.\n\n"
                    f"Requested: "
                    f"৳{withdrawal['amount']:.2f}\n"
                    f"Fee: ৳{withdrawal['fee']:.2f}\n"
                    f"Receive: "
                    f"৳{withdrawal['receive_amount']:.2f}"
                )
            )

        except Exception:
            pass

    # =====================================================
    # REJECT WITHDRAWAL
    # =====================================================

    elif data.startswith("reject_withdraw:"):

        if user_id != ADMIN_ID:
            return

        withdrawal_id = int(
            data.split(":")[1]
        )

        withdrawal = db_execute(
            """
            SELECT *
            FROM withdrawals
            WHERE id=?
            """,
            (withdrawal_id,),
            fetchone=True
        )

        if not withdrawal:
            return

        if withdrawal["status"] != "pending":

            await query.answer(
                "Already processed.",
                show_alert=True
            )

            return

        db_execute(
            """
            UPDATE withdrawals
            SET status='rejected'
            WHERE id=?
            """,
            (withdrawal_id,)
        )

        await query.edit_message_text(
            "❌ Withdrawal Rejected."
        )

        try:

            await context.bot.send_message(
                chat_id=withdrawal["user_id"],
                text=(
                    "❌ Your withdrawal request "
                    "was rejected."
                )
            )

        except Exception:
            pass

    # =====================================================
    # DELETE BUTTON
    # =====================================================

    elif data.startswith("delete_button:"):

        if user_id != ADMIN_ID:
            return

        button_id = int(
            data.split(":")[1]
        )

        db_execute(
            "DELETE FROM buttons WHERE id=?",
            (button_id,)
        )

        await query.answer(
            "Button deleted."
        )

        await show_admin_panel(query)


# =========================================================
# ADMIN COMMAND
# =========================================================

async def admin(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if update.effective_user.id != ADMIN_ID:

        await update.message.reply_text(
            "❌ You are not authorized."
        )

        return

    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "👨‍💼 Open Admin Panel",
                callback_data="admin_panel"
            )
        ]
    ])

    await update.message.reply_text(
        "🔐 Admin access granted.",
        reply_markup=keyboard
    )


# =========================================================
# MESSAGE HANDLER
# =========================================================

async def message_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = update.effective_user

    if not user:
        return

    add_user(user)

    state = context.user_data.get("state")

    # =====================================================
    # ADMIN ADD BUTTON - TITLE
    # =====================================================

    if (
        user.id == ADMIN_ID
        and state == "admin_add_button_title"
    ):

        if not update.message.text:
            await update.message.reply_text(
                "❌ Please send the button name as text."
            )
            return

        context.user_data["button_title"] = (
            update.message.text.strip()
        )

        context.user_data["state"] = (
            "admin_add_button_url"
        )

        await update.message.reply_text(
            "Now send the button URL.\n\n"
            "Example:\n"
            "https://t.me/yourchannel"
        )

        return

    # =====================================================
    # ADMIN ADD BUTTON - URL
    # =====================================================

    if (
        user.id == ADMIN_ID
        and state == "admin_add_button_url"
    ):

        if not update.message.text:
            await update.message.reply_text(
                "❌ Please send a valid URL."
            )
            return

        title = context.user_data.get(
            "button_title"
        )

        url = update.message.text.strip()

        if not url.startswith(
            ("http://", "https://", "tg://")
        ):

            await update.message.reply_text(
                "❌ Invalid URL.\n\n"
                "Please send a valid "
                "https:// or Telegram URL."
            )

            return

        db_execute(
            """
            INSERT INTO buttons
            (title, url)
            VALUES (?, ?)
            """,
            (
                title,
                url
            )
        )

        context.user_data.clear()

        await update.message.reply_text(
            "✅ Button added successfully.\n\n"
            f"Button: {title}\n"
            f"URL: {url}"
        )

        return

    # =====================================================
    # ADMIN BROADCAST
    # =====================================================

    if (
        user.id == ADMIN_ID
        and state == "admin_broadcast"
    ):

        if not update.message.text:

            await update.message.reply_text(
                "❌ Please send a text message."
            )

            return

        users = db_execute(
            "SELECT user_id FROM users",
            fetchall=True
        )

        success = 0

        for u in users:

            try:

                await context.bot.send_message(
                    chat_id=u["user_id"],
                    text=update.message.text
                )

                success += 1

            except Exception:
                pass

        context.user_data.clear()

        await update.message.reply_text(
            "📢 Broadcast completed.\n\n"
            f"Sent successfully: "
            f"{success}/{len(users)}"
        )

        return

    # =====================================================
    # ADMIN SINGLE MESSAGE - USER ID
    # =====================================================

    if (
        user.id == ADMIN_ID
        and state == "admin_single_user"
    ):

        try:

            target_id = int(
                update.message.text.strip()
            )

        except (ValueError, AttributeError):

            await update.message.reply_text(
                "❌ Please send a valid "
                "numeric Telegram User ID."
            )

            return

        context.user_data["target_user"] = target_id

        context.user_data["state"] = (
            "admin_single_text"
        )

        await update.message.reply_text(
            "Now send the message you want to send."
        )

        return

    # =====================================================
    # ADMIN SINGLE MESSAGE - TEXT
    # =====================================================

    if (
        user.id == ADMIN_ID
        and state == "admin_single_text"
    ):

        target_id = context.user_data.get(
            "target_user"
        )

        try:

            await context.bot.send_message(
                chat_id=target_id,
                text=update.message.text
            )

            await update.message.reply_text(
                "✅ Message sent successfully."
            )

        except Exception as e:

            await update.message.reply_text(
                "❌ Could not send message.\n\n"
                f"{e}"
            )

        context.user_data.clear()

        return

    # =====================================================
    # WITHDRAW AMOUNT
    # =====================================================

    if state == "withdraw_amount":

        if not update.message.text:

            await update.message.reply_text(
                "❌ Please enter the amount."
            )

            return

        text = update.message.text.strip()

        try:

            amount = float(text)

        except ValueError:

            await update.message.reply_text(
                "❌ Please enter a valid amount.\n\n"
                "Example: 50, 56, 57, 100"
            )

            return

        if amount < MIN_WITHDRAW:

            await update.message.reply_text(
                f"❌ Minimum withdrawal is "
                f"৳{MIN_WITHDRAW}."
            )

            return

        balance = get_balance(user.id)

        if amount > balance:

            await update.message.reply_text(
                "❌ Insufficient balance.\n\n"
                f"Your balance: ৳{balance:.2f}\n"
                f"You requested: ৳{amount:.2f}"
            )

            return

        fee = amount * WITHDRAW_FEE_PERCENT / 100

        receive_amount = amount - fee

        context.user_data["withdraw_amount"] = amount
        context.user_data["withdraw_fee"] = fee
        context.user_data["withdraw_receive"] = (
            receive_amount
        )

        context.user_data["state"] = (
            "withdraw_method"
        )

        keyboard = [

            [
                InlineKeyboardButton(
                    "💳 bKash",
                    callback_data="method_bkash"
                )
            ],

            [
                InlineKeyboardButton(
                    "💳 Nagad",
                    callback_data="method_nagad"
                )
            ],

            [
                InlineKeyboardButton(
                    "💰 Binance",
                    callback_data="method_binance"
                )
            ]
        ]

        await update.message.reply_text(
            "💸 Withdrawal Summary\n\n"
            f"Requested: ৳{amount:.2f}\n"
            f"Fee (4%): ৳{fee:.2f}\n"
            f"You receive: ৳{receive_amount:.2f}\n\n"
            "Choose withdrawal method:",
            reply_markup=InlineKeyboardMarkup(
                keyboard
            )
        )

        return

    # =====================================================
    # WITHDRAW ACCOUNT
    # =====================================================

    if state == "withdraw_account":

        if not update.message.text:

            await update.message.reply_text(
                "❌ Please send your account "
                "number or wallet address."
            )

            return

        account = update.message.text.strip()

        amount = context.user_data[
            "withdraw_amount"
        ]

        fee = context.user_data[
            "withdraw_fee"
        ]

        receive_amount = context.user_data[
            "withdraw_receive"
        ]

        method = context.user_data[
            "withdraw_method"
        ]

        db_execute("""
            INSERT INTO withdrawals
            (
                user_id,
                amount,
                fee,
                receive_amount,
                method,
                account,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        """, (
            user.id,
            amount,
            fee,
            receive_amount,
            method,
            account,
            datetime.now().isoformat()
        ))

        withdrawal_id = db_execute(
            """
            SELECT last_insert_rowid() AS id
            """,
            fetchone=True
        )["id"]

        context.user_data.clear()

        await update.message.reply_text(
            "✅ Withdrawal request submitted!\n\n"
            f"Amount: ৳{amount:.2f}\n"
            f"Fee: ৳{fee:.2f}\n"
            f"You receive: "
            f"৳{receive_amount:.2f}\n"
            f"Method: {method}\n\n"
            "⏳ Waiting for admin approval."
        )

        keyboard = InlineKeyboardMarkup([

            [
                InlineKeyboardButton(
                    "✅ Approve",
                    callback_data=(
                        f"approve_withdraw:{withdrawal_id}"
                    )
                ),

                InlineKeyboardButton(
                    "❌ Reject",
                    callback_data=(
                        f"reject_withdraw:{withdrawal_id}"
                    )
                )
            ]

        ])

        await context.bot.send_message(

            chat_id=ADMIN_ID,

            text=(
                "🔔 NEW WITHDRAWAL REQUEST\n\n"

                f"User: "
                f"@{user.username or 'No username'}\n"

                f"User ID: {user.id}\n\n"

                f"Amount: ৳{amount:.2f}\n"

                f"Fee: ৳{fee:.2f}\n"

                f"User receives: "
                f"৳{receive_amount:.2f}\n\n"

                f"Method: {method}\n"

                f"Account: {account}"
            ),

            reply_markup=keyboard
        )

        return

    # =====================================================
    # FILE SUBMISSION
    # =====================================================

    if state == "waiting_file":

        count = get_today_file_count(user.id)

        if count >= DAILY_FILE_LIMIT:

            context.user_data.clear()

            await update.message.reply_text(
                "❌ Daily file limit reached."
            )

            return

        file_id = None

        if update.message.document:

            file_id = update.message.document.file_id

        elif update.message.photo:

            file_id = (
                update.message.photo[-1].file_id
            )

        elif update.message.video:

            file_id = update.message.video.file_id

        else:

            await update.message.reply_text(
                "❌ Please send a file."
            )

            return

        db_execute("""
            INSERT INTO submissions
            (
                user_id,
                file_id,
                status,
                created_at
            )
            VALUES (?, ?, 'pending', ?)
        """, (
            user.id,
            file_id,
            datetime.now().isoformat()
        ))

        increase_file_count(user.id)

        context.user_data.clear()

        await update.message.reply_text(
            "✅ File submitted successfully.\n\n"
            "⏳ Waiting for admin review."
        )

        return

    # =====================================================
    # DEFAULT
    # =====================================================

    await update.message.reply_text(
        "Please use the buttons in the menu.\n\n"
        "Send /start to open the menu."
    )


# =========================================================
# WITHDRAW METHOD
# =========================================================

async def withdrawal_method_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    if query.data.startswith("method_"):

        method = (
            query.data
            .replace("method_", "")
            .capitalize()
        )

        context.user_data["withdraw_method"] = method

        context.user_data["state"] = (
            "withdraw_account"
        )

        await query.edit_message_text(
            f"💳 Selected: {method}\n\n"
            "Now send your account number / "
            "wallet address."
        )


# =========================================================
# MAIN
# =========================================================

def main():

    init_db()

    # Start Render HTTP server
    web_thread = threading.Thread(
        target=run_web_server,
        daemon=True
    )

    web_thread.start()

    # Start Telegram bot
    application = (
        Application.builder()
        .token(BOT_TOKEN)
        .build()
    )

    application.add_handler(
        CommandHandler(
            "start",
            start
        )
    )

    application.add_handler(
        CommandHandler(
            "admin",
            admin
        )
    )

    application.add_handler(
        CallbackQueryHandler(
            withdrawal_method_handler,
            pattern=r"^method_"
        )
    )

    application.add_handler(
        CallbackQueryHandler(
            button_handler
        )
    )

    application.add_handler(
        MessageHandler(
            filters.ALL & ~filters.COMMAND,
            message_handler
        )
    )

    print("Telegram bot is running...")

    application.run_polling(
        allowed_updates=Update.ALL_TYPES
    )


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":
    main()
