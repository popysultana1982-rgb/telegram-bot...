# =========================================================
# FILE: bot.py
# PROJECT: Automated Multi-Stage Gmail Buy/Sell Bot (Health-Check & Super Fast)
# =========================================================

import asyncio
import io
import json
import os
import re
import shutil
import threading
from datetime import datetime, timedelta

import aiohttp
from aiohttp import web
import openpyxl
import psycopg2
from psycopg2.extras import RealDictCursor
from psycopg2.pool import SimpleConnectionPool
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
HARDCODED_OWNER_ID = 8919985167

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL environment variable is not set.")

PORT = int(os.environ.get("PORT", 10000))
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL")

FORCE_GROUP_CHAT_ID = -1004471047712
FORCE_GROUP_LINK = "https://t.me/+rVP6CkmqrnFlNzA1"

FORCE_CHANNEL_CHAT_ID = -1003991468184
FORCE_CHANNEL_LINK = "https://t.me/fast_payment_proof_chanel"

SUPPORT_URL = "https://t.me/Talha_juba098"


# =========================================================
# POSTGRESQL CONNECTION POOL (SUPERFAST)
# =========================================================

db_pool = None
pool_lock = threading.Lock()

def get_db_pool():
    global db_pool
    if db_pool is None:
        with pool_lock:
            if db_pool is None:
                db_pool = SimpleConnectionPool(1, 20, DATABASE_URL)
    return db_pool

def db_execute(query, params=(), fetchone=False, fetchall=False):
    pool = get_db_pool()
    conn = pool.getconn()
    conn.autocommit = True
    pg_query = query.replace("?", "%s")
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(pg_query, params)
            if fetchone:
                return cur.fetchone()
            if fetchall:
                return cur.fetchall()
            return None
    except Exception as e:
        print(f"Database query error: {e}")
        raise e
    finally:
        pool.putconn(conn)


def init_db():
    db_execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id BIGINT PRIMARY KEY,
            username TEXT,
            balance DOUBLE PRECISION DEFAULT 0,
            manual_ref_count INTEGER DEFAULT 0,
            files_today INTEGER DEFAULT 0,
            last_file_date TEXT,
            referred_by BIGINT DEFAULT NULL,
            ref_bonus_rewarded INTEGER DEFAULT 0,
            seen_rules INTEGER DEFAULT 0,
            language TEXT DEFAULT 'bn',
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS dynamic_buttons (
            id SERIAL PRIMARY KEY,
            parent_id INTEGER DEFAULT 0,
            title TEXT NOT NULL,
            btn_type TEXT NOT NULL,
            content TEXT NOT NULL,
            show_in_reply INTEGER DEFAULT 1
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS withdrawals (
            id SERIAL PRIMARY KEY,
            user_id BIGINT,
            amount DOUBLE PRECISION,
            fee DOUBLE PRECISION,
            receive_amount DOUBLE PRECISION,
            method TEXT,
            account TEXT,
            status TEXT DEFAULT 'pending',
            created_at TEXT
        )
    """)

    db_execute("""
        CREATE TABLE IF NOT EXISTS submissions (
            id SERIAL PRIMARY KEY,
            user_id BIGINT,
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
            admin_id BIGINT,
            message_id BIGINT
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
            user_id BIGINT PRIMARY KEY,
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

    try:
        db_execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS ref_bonus_rewarded INTEGER DEFAULT 0")
        db_execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS language TEXT DEFAULT 'bn'")
    except Exception:
        pass

    for k, v in CONFIG_DEFAULTS.items():
        db_execute("""
            INSERT INTO settings (key, value) VALUES (%s, %s)
            ON CONFLICT (key) DO NOTHING
        """, (k, v))


# =========================================================
# CONFIG HELPERS & SWITCH CACHE
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
    "maintenance": "0",
    "switch_sell": "1",
    "switch_withdraw": "1",
    "switch_referral": "1",
    "switch_balance": "1",
    "switch_history": "1",
    "switch_rules": "1",
    "switch_tutorial": "1",
    "switch_support": "1",
    "show_btn_sell": "1",
    "show_btn_balance": "1",
    "show_btn_withdraw": "1",
    "show_btn_history": "1",
    "show_btn_referral": "1",
    "show_btn_rules": "1",
    "show_btn_tutorial": "1",
    "show_btn_support": "1",
    "notice_sell": "⚠️ বর্তমানে আমাদের জিমেইল ফাইল রিসিভ কার্যক্রম সাময়িকভাবে বন্ধ আছে। পরবর্তীতে খোলার সময় নোটিশের মাধ্যমে জানিয়ে দেওয়া হবে।",
    "notice_withdraw": "⚠️ বর্তমানে পেমেন্ট ও উইথড্র সিস্টেম সাময়িকভাবে মেইনটেন্যান্সে রয়েছে। খুব শীঘ্রই চালু হবে।",
    "notice_referral": "⚠️ রেফারেল সিস্টেম সাময়িকভাবে স্থগিত রয়েছে।",
    "notice_balance": "⚠️ ব্যালেন্স সার্ভারে কাজ চলছে। অনুগ্রহ করে কিছুক্ষণ পর চেষ্টা করুন।",
    "notice_history": "⚠️ হিস্ট্রি রেকর্ড বর্তমানে আপডেট করা হচ্ছে।",
    "notice_rules": "⚠️ নিয়মাবলী সেকশন বর্তমানে আপডেট হচ্ছে।",
    "notice_tutorial": "⚠️ টিউটোরিয়াল ভিডিও শীঘ্রই নতুন নিয়মে আপডেট করা হবে।",
    "notice_support": "⚠️ সাপোর্ট টিম বর্তমানে ব্যস্ত রয়েছে। অনুগ্রহ করে কিছুক্ষণ পর মেসেজ দিন."
}

CONFIG_CACHE = {}

def get_config(key, as_type=float):
    now = datetime.now()
    if key in CONFIG_CACHE:
        val, expire_at = CONFIG_CACHE[key]
        if now < expire_at:
            try:
                return as_type(val)
            except (ValueError, TypeError):
                return as_type(CONFIG_DEFAULTS.get(key, "0"))

    res = db_execute("SELECT value FROM settings WHERE key=%s", (key,), fetchone=True)
    val = res["value"] if res else CONFIG_DEFAULTS.get(key, "0")
    CONFIG_CACHE[key] = (val, now + timedelta(seconds=120))
    try:
        return as_type(val)
    except (ValueError, TypeError):
        return as_type(CONFIG_DEFAULTS.get(key, "0"))

def set_config(key, val):
    CONFIG_CACHE[key] = (str(val), datetime.now() + timedelta(seconds=120))
    db_execute("""
        INSERT INTO settings (key, value) VALUES (%s, %s)
        ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value
    """, (key, str(val)))

def is_button_active(btn_key):
    return get_config(f"switch_{btn_key}", int) == 1

def toggle_button_switch(btn_key):
    cur = is_button_active(btn_key)
    new_val = 0 if cur else 1
    set_config(f"switch_{btn_key}", new_val)
    return new_val == 1

def get_button_closed_notice(btn_key, lang='bn'):
    res = db_execute("SELECT value FROM settings WHERE key=%s", (f"notice_{btn_key}",), fetchone=True)
    notice = res["value"] if (res and res["value"]) else CONFIG_DEFAULTS.get(f"notice_{btn_key}", "⚠️ এই অপশনটি বর্তমানে সাময়িকভাবে বন্ধ আছে।")
    if lang == 'en':
        if "বন্ধ আছে" in notice or "বন্ধ" in notice:
            return "⚠️ This option is temporarily closed."
        elif "মেইনটেন্যান্সে" in notice:
            return "⚠️ Payment & withdrawal system is currently under maintenance. Will be available soon."
    return notice

def is_button_visible(btn_key):
    return get_config(f"show_btn_{btn_key}", int) == 1

def toggle_button_visibility(btn_key):
    cur = is_button_visible(btn_key)
    new_val = 0 if cur else 1
    set_config(f"show_btn_{btn_key}", new_val)
    return new_val == 1


# =========================================================
# BUTTON PARSER & SAFE SENDER HELPER
# =========================================================

BUTTON_PATTERN = re.compile(r"\[([^\|\]]+)\|([^\]]+)\]")

def parse_text_and_buttons(raw_text):
    matches = BUTTON_PATTERN.findall(raw_text)
    clean_text = BUTTON_PATTERN.sub("", raw_text).strip()

    if not matches:
        return clean_text, None

    keyboard = []
    for btn_title, btn_target in matches:
        b_title = btn_title.strip()
        b_target = btn_target.strip()
        if b_target.startswith("@"):
            b_target = f"https://t.me/{b_target[1:]}"
        elif not b_target.startswith("http://") and not b_target.startswith("https://") and not b_target.startswith("tg://"):
            b_target = f"https://{b_target}"

        if b_title and b_target:
            keyboard.append([InlineKeyboardButton(b_title, url=b_target)])

    markup = InlineKeyboardMarkup(keyboard) if keyboard else None
    return clean_text, markup

async def safe_send_message(bot, chat_id, text, reply_markup=None):
    try:
        return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup, parse_mode="Markdown")
    except Exception:
        try:
            return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup)
        except Exception as e:
            print(f"Failed to send message to {chat_id}: {e}")
            return None

def get_user_chat_url(target_user_id, username=None):
    if username:
        clean_user = username.replace("@", "").strip()
        if clean_user:
            return f"https://t.me/{clean_user}"
    return f"tg://user?id={target_user_id}"


# =========================================================
# ADMIN & ROLE MANAGEMENT
# =========================================================

def get_owner_id():
    return HARDCODED_OWNER_ID

def is_admin(user_id):
    if int(user_id) == HARDCODED_OWNER_ID:
        return True
    helper = db_execute("SELECT user_id FROM helpers WHERE user_id=%s", (user_id,), fetchone=True)
    return helper is not None

def get_admin_name(user_id):
    if int(user_id) == HARDCODED_OWNER_ID:
        return "👑 Owner"
    helper = db_execute("SELECT admin_name FROM helpers WHERE user_id=%s", (user_id,), fetchone=True)
    if helper and helper["admin_name"]:
        return helper["admin_name"]
    return f"Admin ({user_id})"

def add_helper(user_id, name="Admin"):
    db_execute("""
        INSERT INTO helpers (user_id, admin_name, added_at) VALUES (%s, %s, %s)
        ON CONFLICT(user_id) DO UPDATE SET admin_name=EXCLUDED.admin_name
    """, (user_id, name, datetime.now().isoformat()))

def remove_helper(user_id):
    db_execute("DELETE FROM helpers WHERE user_id=%s", (user_id,))

def set_helper_name(user_id, name):
    db_execute("UPDATE helpers SET admin_name=%s WHERE user_id=%s", (name, user_id))

def get_all_admins():
    owner = HARDCODED_OWNER_ID
    helpers = db_execute("SELECT user_id, admin_name FROM helpers", fetchall=True)
    return owner, (helpers or [])

async def notify_all_admins(context, message, exclude_user_id=None):
    owner, helpers = get_all_admins()
    admin_list = [int(owner)]

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
        await safe_send_message(context.bot, chat_id=a_id, text=message)

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
# MESSAGES & BUTTON TITLES (MULTI-LANGUAGE SUPPORT)
# =========================================================

DEFAULT_MESSAGES = {
    "bn": {
        "welcome": (
            "আসসালামু আলাইকুম, {name}! 🌸\n\n"
            "💙 আপনাকে স্বাগতম আমাদের Gmail Sell Bot-এ!\n"
            "এখানে আপনি আপনার তৈরি করা Valid Gmail Account সেল/সাবমিট করে নিরাপদভাবে টাকা ইনকাম করতে পারবেন।\n\n"
            "📜 **আমাদের জিমেইল সাবমিশন ও কাজের নিয়মাবলী:**\n"
            "১. ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক বাছাই করবেন এবং লগইন করা যায় এমন মেইলগুলো পর্যালোচনায় রাখবেন।\n"
            "২. প্রাথমিক বাছায়ে যেসব জিমেইলে সমস্যা থাকবে, সেগুলো কারণসহ সাথে সাথে বাতিল করে এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
            "৩. পর্যালোচনায় রাখা জিমেইলগুলো পরবর্তী ২৪ থেকে ৪৮ ঘণ্টা অ্যাডমিনের পর্যবেক্ষণে থাকবে।\n"
            "৪. ২৪ থেকে ৪৮ ঘণ্টার পর্যবেক্ষণ সময়ে ব্যালেন্স যোগ হবে না।\n"
            "৫. পর্যবেক্ষণ শেষে যেসব জিমেইল অক্ষত থাকবে, প্রতিটির জন্য ৳{rate} টাকা ব্যালেন্সে যোগ হবে! 💰\n"
            "৬. শেষ ধাপে কোনো জিমেইল নষ্ট হলে তা আপনাকে কারণসহ এক্সেল ফাইলে ফেরত দেওয়া হবে।\n"
            "৭. প্রতিদিন সর্বোচ্চ {daily_limit}টি ফাইল এবং প্রতি ফাইলে সর্বোচ্চ {email_limit}টি ভ্যালিড Gmail দিতে পারবেন।\n"
            "৮. একবার জমা দেওয়া জিমেইল আগামী {cooldown} দিনের মধ্যে পুনরায় সাবমিট করা যাবে না।"
        ),
        "rules": (
            "📜 **আমাদের কাজের নিয়মাবলী ও শর্তাবলী:**\n\n"
            "১️⃣ ফাইল সাবমিটের পর অ্যাডমিন প্রাথমিক বাছাই সম্পন্ন করবেন।\n"
            "২️⃣ সমস্যা থাকা Gmail প্রথম ধাপেই নির্দিষ্ট কারণসহ বাতিল করা হবে এবং Excel ফাইলে ফেরত দেওয়া হবে।\n"
            "৩️⃣ নির্বাচিত Gmail-গুলো ২৪ থেকে ৪৮ ঘণ্টা পর্যবেক্ষণে থাকবে।\n"
            "৪️⃣ পর্যবেক্ষণ শেষে সঠিক ও অনুমোদিত প্রতিটি Gmail-এর জন্য ৳{rate} টাকা ব্যালেন্সে যোগ করা হবে।\n"
            "৫️⃣ পর্যবেক্ষণের শেষ ধাপে সমস্যা পাওয়া Gmail-গুলো কারণসহ Excel ফাইলে ফেরত দেওয়া হবে।\n"
            "৬️⃣ এক ফাইলে সর্বোচ্চ {email_limit}টি @gmail.com এবং প্রতিদিন সর্বোচ্চ {daily_limit}টি ফাইল সাবমিট করা যাবে।\n"
            "৭️⃣ বিগত {cooldown} দিনের মধ্যে জমা দেওয়া Gmail পুনরায় গ্রহণযোগ্য নয়।\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "⚠️ 📌 **বিশেষ দ্রষ্টব্য (IMPORTANT NOTICE)**\n"
            "✅ **শুধুমাত্র নিজের ফোনে, নিজের স্বাভাবিক নেটওয়ার্ক/IP ব্যবহার করে নিয়মিত ও বৈধভাবে তৈরি করা আসল Gmail অ্যাকাউন্ট সাবমিট করবেন।**\n"
            "❌ **কোনো ধরনের পদ্ধতি ব্যবহার করে তৈরি করা, অস্বাভাবিক, ভুয়া বা সন্দেহজনক Gmail অ্যাকাউন্ট সাবমিট করবেন না।**\n"
            "🚫 **অন্যের অ্যাকাউন্ট, অননুমোদিত পদ্ধতিতে তৈরি অ্যাকাউন্ট কিংবা যেসব Gmail-এ স্বাভাবিকভাবে লগইন করতে সমস্যা হয়, সেগুলো সাবমিট করা থেকে বিরত থাকুন।**\n"
            "📌 **নিয়ম না মেনে Gmail সাবমিট করলে তা বাতিল হতে পারে। তাই ফাইল জমা দেওয়ার আগে সব নিয়ম ভালোভাবে পড়ে নিন।**\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "❤️ সবার সহযোগিতা কামনা করছি। ধন্যবাদ।\n"
            "🤖 BOT ADMIN TEAM"
        ),
        "sell": (
            "📤 **আপনার ফ্রেশ জিমেইল সম্বলিত এক্সেল বা সিএসভি ফাইল (.xlsx, .xls, .csv) পাঠান।**\n\n"
            "💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate} BDT\n"
            "⚠️ এক ফাইলে সর্বোচ্চ {email_limit}টি @gmail.com থাকতে হবে। বিগত {cooldown} দিনের মধ্যে জমা দেওয়া কোনো মেইল গ্রহণ করা হবে না।\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "🚨 **জরুরি নির্দেশনা (MUST READ):**\n"
            "⚠️ **ফাইলে থাকা জিমেইলগুলো বটে সাবমিট করার আগে অবশ্যই আপনার ফোন থেকে সম্পূর্ণ রিমুভ (Logout/Remove) করতে হবে!**\n\n"
            "⚠️ 📌 **বিশেষ দ্রষ্টব্য (IMPORTANT NOTICE):**\n"
            "✅ শুধুমাত্র নিজের ফোনে, নিজের স্বাভাবিক নেটওয়ার্ক/IP ব্যবহার করে নিয়মিত ও বৈধভাবে তৈরি করা আসল Gmail অ্যাকাউন্ট সাবমিট করবেন।\n"
            "❌ কোনো ধরনের ক্লোন, সফটওয়্যার বা অস্বাভাবিক পদ্ধতিতে তৈরি করা ভুয়া বা সন্দেহজনক Gmail অ্যাকাউন্ট সাবমিট করবেন না।\n"
            "🚫 অন্যের অ্যাকাউন্ট বা লগইন সমস্যাযুক্ত জিমেইল জমা দিলে সম্পূর্ণ ফাইল বাতিল হয়ে যাবে।\n"
            "━━━━━━━━━━━━━━━━━━"
        ),
        "support": "যেকোনো সমস্যা, প্রশ্ন বা সহায়তার জন্য নিচে থাকা বাটনে ক্লিক করে সাপোর্টে মেসেজ পাঠান:",
        "tutorial": (
            "🎥 **ভিডিও টিউটোরিয়াল ও কাজের গাইড:**\n\n"
            "বটে কীভাবে কাজ করবেন, কীভাবে জিমেইল ফাইল তৈরি করবেন এবং কীভাবে টাকা উইথড্র করবেন তা বিস্তারিত দেখতে নিচের বাটনে ক্লিক করুন:"
        ),
        "maintenance": (
            "⚠️ **বট আপডেটের কাজ চলছে!** 🛠\n\n"
            "সম্মানিত ইউজার, আমাদের সিস্টেমে জরুরি আপডেটের কাজ চলছে। "
            "সাময়িকভাবে বটের কার্যক্রম স্থগিত রয়েছে। কাজ সম্পন্ন হওয়ামাত্রই বট পুনরায় চালু হবে।"
        ),
        "referral": (
            "👥 **রেফার করে ইনকাম করুন!**\n\n"
            "আপনার রেফারেল লিংকে জয়েন করা কোনো ইউজার যখন তার প্রথম জিমেইল ফাইল সফলভাবে জমা দিয়ে অ্যাপ্রুভ পাবেন, তখন আপনি পাবেন ৳{bonus} BDT বোনাস。\n\n"
            "আপনার রেফারেল লিংকটি শেয়ার করুন:\n"
            "🔗 `{link}`"
        )
    },
    "en": {
        "welcome": (
            "Assalamu Alaikum, {name}! 🌸\n\n"
            "💙 Welcome to our Gmail Sell Bot!\n"
            "Here you can earn money securely by selling/submitting valid Gmail accounts created by you.\n\n"
            "📜 **Our Gmail Submission & Work Rules:**\n"
            "1. After file submission, admin will do initial screening and keep login-working emails in review.\n"
            "2. Problematic emails in initial screening will be rejected instantly with reason and returned in an Excel file.\n"
            "3. Reviewed emails will remain under admin observation for the next 24 to 48 hours.\n"
            "4. Balance will not be added during the 24-48 hours observation period.\n"
            "5. After observation, for each intact Gmail, ৳{rate} BDT will be added to your balance! 💰\n"
            "6. If any Gmail fails in the final step, it will be returned to you in an Excel file with reason.\n"
            "7. You can submit a maximum of {daily_limit} files per day and up to {email_limit} valid Gmails per file.\n"
            "8. Once submitted, a Gmail cannot be re-submitted within the next {cooldown} days."
        ),
        "rules": (
            "📜 **Our Rules & Terms:**\n\n"
            "1️⃣ Admin will complete initial screening after file submission.\n"
            "2️⃣ Problematic Gmails will be rejected in the first step with specific reasons and returned in an Excel file.\n"
            "3️⃣ Selected Gmails will remain under observation for 24 to 48 hours.\n"
            "4️⃣ After observation, ৳{rate} BDT will be added to the balance for each correct and approved Gmail.\n"
            "5️⃣ Gmails found with issues at the final observation step will be returned in an Excel file with reasons.\n"
            "6️⃣ Maximum {email_limit} @gmail.com per file and {daily_limit} files per day can be submitted.\n"
            "7️⃣ Gmail submitted within the last {cooldown} days is not acceptable again.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "⚠️ 📌 **IMPORTANT NOTICE**\n"
            "✅ **Submit only genuine Gmail accounts created regularly and legally on your own phone using your normal network/IP.**\n"
            "❌ **Do not submit fake, unusual, or suspicious Gmail accounts created using automated methods.**\n"
            "🚫 **Avoid submitting other people's accounts, unauthorized accounts, or Gmails with login issues.**\n"
            "📌 **Submitting Gmails bypassing rules may lead to cancellation. Read all rules carefully before submitting.**\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "❤️ Thank you for your cooperation.\n"
            "🤖 BOT ADMIN TEAM"
        ),
        "sell": (
            "📤 **Send your Excel or CSV file (.xlsx, .xls, .csv) containing fresh Gmails.**\n\n"
            "💰 Rate per valid Gmail: ৳{rate} BDT\n"
            "⚠️ Maximum {email_limit} @gmail.com per file. Mails submitted within the last {cooldown} days will not be accepted.\n\n"
            "━━━━━━━━━━━━━━━━━━\n"
            "🚨 **MUST READ:**\n"
            "⚠️ **Before submitting Gmails from the file to the bot, you must completely remove (Logout/Remove) them from your phone!**\n\n"
            "⚠️ 📌 **IMPORTANT NOTICE:**\n"
            "✅ Submit only genuine Gmail accounts created legally on your own device.\n"
            "❌ Do not submit fake or clone-based Gmails.\n"
            "🚫 Submitting unauthorized or problem-locked accounts will cancel the entire file.\n"
            "━━━━━━━━━━━━━━━━━━"
        ),
        "support": "For any issues, questions or support, click the button below to message support:",
        "tutorial": (
            "🎥 **Video Tutorial & Guide:**\n\n"
            "To see in detail how to work in the bot, create Gmail files, and withdraw money, click the button below:"
        ),
        "maintenance": (
            "⚠️ **Bot Update in Progress!** 🛠\n\n"
            "Dear user, emergency system updates are underway. "
            "Bot operations are temporarily suspended. It will be back online shortly."
        ),
        "referral": (
            "👥 **Earn by Referring!**\n\n"
            "When a user who joined via your referral link successfully submits their first Gmail file and gets approved, you receive a ৳{bonus} BDT bonus.\n\n"
            "Share your referral link:\n"
            "🔗 `{link}`"
        )
    }
}

DEFAULT_BUTTON_TITLES = {
    "bn": {
        "sell": "📩 জিমেইল সেল করুন",
        "balance": "💰 ব্যালেন্স",
        "withdraw": "💸 উইথড্র",
        "history": "📜 হিস্ট্রি",
        "referral": "🤝 আমার রেফারেল",
        "rules": "📜 নিয়মাবলী",
        "tutorial": "▶️ টিউটোরিয়াল",
        "support": "📞 সাপোর্ট",
        "lang": "🌐 CHANGE LANGUAGE"
    },
    "en": {
        "sell": "📩 SELL FRESH GMAIL ACCOUNT",
        "balance": "💰 BALANCE",
        "withdraw": "💸 WITHDRAW",
        "history": "📜 HISTORY",
        "referral": "🤝 MY REFERRAL",
        "rules": "📜 RULES",
        "tutorial": "▶️ TUTORIAL",
        "support": "📞 SUPPORT",
        "lang": "🌐 CHANGE LANGUAGE"
    }
}

MESSAGE_NAMES = {
    "welcome": "🎉 Welcome Message & Rules",
    "rules": "📜 Rules Message",
    "sell": "📤 Sell Gmail Instruction",
    "tutorial": "▶️ Tutorial Message",
    "referral": "👥 Referral Message",
    "support": "📞 Support Text",
    "maintenance": "🛠 Maintenance Notice"
}

def get_user_lang(user_id):
    u = get_user(user_id)
    if u and u.get("language"):
        return u["language"]
    return "bn"

def set_user_lang(user_id, lang):
    db_execute("UPDATE users SET language=%s WHERE user_id=%s", (lang, user_id))

def get_custom_msg(key, lang='bn'):
    res = db_execute("SELECT value FROM settings WHERE key=%s", (f"msg_{key}_{lang}",), fetchone=True)
    if not res:
        res = db_execute("SELECT value FROM settings WHERE key=%s", (f"msg_{key}",), fetchone=True)
    
    rate = get_config("email_rate", float)
    d_lim = get_config("daily_file_limit", int)
    e_lim = get_config("max_emails_per_file", int)
    cd = get_config("duplicate_check_days", int)

    lang_dict = DEFAULT_MESSAGES.get(lang, DEFAULT_MESSAGES["bn"])
    text = res["value"] if (res and res["value"]) else lang_dict.get(key, DEFAULT_MESSAGES["bn"].get(key, ""))
    return (
        text.replace("{rate}", f"{rate:.2f}")
        .replace("{daily_limit}", str(d_lim))
        .replace("{email_limit}", str(e_lim))
        .replace("{cooldown}", str(cd))
    )

def set_custom_msg(key, text, lang='bn'):
    set_config(f"msg_{key}_{lang}", text)

def delete_custom_msg(key, lang='bn'):
    db_execute("DELETE FROM settings WHERE key=%s", (f"msg_{key}_{lang}",))

def get_button_title(btn_key, lang='bn'):
    res = db_execute("SELECT value FROM settings WHERE key=%s", (f"btn_{btn_key}_{lang}",), fetchone=True)
    if res and res["value"]:
        return res["value"]
    return DEFAULT_BUTTON_TITLES.get(lang, DEFAULT_BUTTON_TITLES["bn"]).get(btn_key, DEFAULT_BUTTON_TITLES["bn"].get(btn_key, btn_key))

def set_button_title(btn_key, new_title, lang='bn'):
    set_config(f"btn_{btn_key}_{lang}", new_title)

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
    return deleted_files


# =========================================================
# USER & BALANCE OPERATIONS
# =========================================================

def add_user(user, referrer_id=None):
    existing = db_execute("SELECT * FROM users WHERE user_id=%s", (user.id,), fetchone=True)
    if not existing:
        db_execute("""
            INSERT INTO users (user_id, username, balance, manual_ref_count, files_today, last_file_date, referred_by, ref_bonus_rewarded, seen_rules, language, created_at)
            VALUES (%s, %s, 0, 0, 0, %s, %s, 0, 0, 'bn', %s)
        """, (user.id, user.username or "", datetime.now().strftime("%Y-%m-%d"), referrer_id, datetime.now().isoformat()))
        return True
    else:
        db_execute("UPDATE users SET username=%s WHERE user_id=%s", (user.username or "", user.id))
        return False

def get_user(user_id):
    return db_execute("SELECT * FROM users WHERE user_id=%s", (user_id,), fetchone=True)

def mark_rules_seen(user_id):
    db_execute("UPDATE users SET seen_rules=1 WHERE user_id=%s", (user_id,))

def get_balance(user_id):
    user = get_user(user_id)
    return float(user["balance"]) if user else 0.0

def update_balance(user_id, amount):
    db_execute("UPDATE users SET balance = balance + %s WHERE user_id=%s", (amount, user_id))

def get_referral_count(user_id):
    u = get_user(user_id)
    manual = int(u["manual_ref_count"]) if (u and u.get("manual_ref_count")) else 0
    res = db_execute("SELECT COUNT(*) AS c FROM users WHERE referred_by=%s AND ref_bonus_rewarded=1", (user_id,), fetchone=True)
    organic = res["c"] if res else 0
    return organic + manual

def add_manual_referral(user_id, count):
    db_execute("UPDATE users SET manual_ref_count = manual_ref_count + %s WHERE user_id=%s", (count, user_id))

def get_referral_details(user_id):
    return db_execute("""
        SELECT user_id, username, ref_bonus_rewarded, created_at 
        FROM users 
        WHERE referred_by=%s 
        ORDER BY created_at DESC
    """, (user_id,), fetchall=True)

def get_top_referrers(limit=25):
    return db_execute(f"""
        SELECT u.user_id, u.username, u.manual_ref_count,
               COUNT(CASE WHEN r.ref_bonus_rewarded = 1 THEN 1 END) as organic_count,
               (COUNT(CASE WHEN r.ref_bonus_rewarded = 1 THEN 1 END) + COALESCE(u.manual_ref_count, 0)) as total_refs
        FROM users u
        LEFT JOIN users r ON u.user_id = r.referred_by
        GROUP BY u.user_id, u.username, u.manual_ref_count
        HAVING (COUNT(CASE WHEN r.ref_bonus_rewarded = 1 THEN 1 END) + COALESCE(u.manual_ref_count, 0)) > 0
        ORDER BY total_refs DESC
        LIMIT {limit}
    """, fetchall=True)

def get_all_users_list(limit=50):
    return db_execute(f"SELECT user_id, username, balance, created_at FROM users ORDER BY user_id DESC LIMIT {limit}", fetchall=True)

def get_today_file_count(user_id):
    user = get_user(user_id)
    today = datetime.now().strftime("%Y-%m-%d")
    if not user:
        return 0
    if user["last_file_date"] != today:
        db_execute("UPDATE users SET files_today=0, last_file_date=%s WHERE user_id=%s", (today, user_id))
        return 0
    return user["files_today"]

def increase_file_count(user_id):
    today = datetime.now().strftime("%Y-%m-%d")
    db_execute("UPDATE users SET files_today = files_today + 1, last_file_date = %s WHERE user_id=%s", (today, user_id))


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
            timeout=2.0
        )
        return ch_joined and gr_joined
    except Exception:
        return False

def get_first_time_markup(lang='bn'):
    if lang == 'en':
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📢 1. Payment Proof Channel", url=FORCE_CHANNEL_LINK)],
            [InlineKeyboardButton("👥 2. Official Group", url=FORCE_GROUP_LINK)],
            [InlineKeyboardButton("✅ 3. Verify", callback_data="check_joined")],
            [InlineKeyboardButton("💬 Direct Message / Support", url=SUPPORT_URL)]
        ])
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📢 ১. পেমেন্ট প্রুফ চ্যানেল", url=FORCE_CHANNEL_LINK)],
        [InlineKeyboardButton("👥 ২. অফিসিয়াল গ্রুপ", url=FORCE_GROUP_LINK)],
        [InlineKeyboardButton("✅ ৩. ভেরিফাই করুন (Verify)", callback_data="check_joined")],
        [InlineKeyboardButton("💬 সরাসরি মেসেজ পাঠান / Support", url=SUPPORT_URL)]
    ])

def get_rejoin_markup(lang='bn'):
    if lang == 'en':
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📢 1. Payment Proof Channel", url=FORCE_CHANNEL_LINK)],
            [InlineKeyboardButton("👥 2. Official Group", url=FORCE_GROUP_LINK)],
            [InlineKeyboardButton("✅ 3. Verify", callback_data="check_joined")]
        ])
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📢 ১. পেমেন্ট প্রুফ চ্যানেল", url=FORCE_CHANNEL_LINK)],
        [InlineKeyboardButton("👥 ২. অফিসিয়াল গ্রুপ", url=FORCE_GROUP_LINK)],
        [InlineKeyboardButton("✅ ৩. ভেরিফাই করুন (Verify)", callback_data="check_joined")]
    ])


# =========================================================
# VALIDATIONS & DUPLICATE CHECK
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
    placeholders = ",".join(["%s"] * len(emails))
    query = f"""
        SELECT email FROM submitted_emails 
        WHERE email IN ({placeholders}) AND submitted_at >= %s
    """
    params = list(emails) + [cutoff_date]
    rows = db_execute(query, params, fetchall=True)
    if rows:
        return [r["email"] for r in rows]
    return []

def save_submitted_emails(emails):
    now_str = datetime.now().isoformat()
    for e in emails:
        db_execute("INSERT INTO submitted_emails (email, submitted_at) VALUES (%s, %s)", (e, now_str))


# =========================================================
# DYNAMIC BIG KEYBOARD (INCL. ADMIN SHORTCUT & LANGUAGE TOGGLE)
# =========================================================

def get_bottom_keyboard(user_id=None, lang='bn'):
    keyboard = []

    if is_button_visible("sell"):
        btn_sell = get_button_title("sell", lang)
        keyboard.append([KeyboardButton(btn_sell)])

    row1 = []
    if is_button_visible("balance"):
        row1.append(KeyboardButton(get_button_title("balance", lang)))
    if is_button_visible("withdraw"):
        row1.append(KeyboardButton(get_button_title("withdraw", lang)))
    if row1:
        keyboard.append(row1)

    row2 = []
    if is_button_visible("history"):
        row2.append(KeyboardButton(get_button_title("history", lang)))
    if is_button_visible("referral"):
        row2.append(KeyboardButton(get_button_title("referral", lang)))
    if row2:
        keyboard.append(row2)

    row3 = []
    if is_button_visible("rules"):
        row3.append(KeyboardButton(get_button_title("rules", lang)))
    if is_button_visible("tutorial"):
        row3.append(KeyboardButton(get_button_title("tutorial", lang)))
    if row3:
        keyboard.append(row3)

    row4 = []
    if is_button_visible("support"):
        row4.append(KeyboardButton(get_button_title("support", lang)))
    row4.append(KeyboardButton(get_button_title("lang", lang)))
    keyboard.append(row4)

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

    if user_id and is_admin(user_id):
        keyboard.append([KeyboardButton("🔐 ADMIN PANEL")])

    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


def build_dynamic_sub_markup(parent_id):
    btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=%s ORDER BY id ASC", (parent_id,), fetchall=True)
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
    lang = get_user_lang(user.id)

    if lang == 'en':
        msg_text = "🎉 **You are successfully connected with us.**\n\nChoose your option by pressing the buttons below:"
        link_title = "📌 **Important Links:**"
    else:
        msg_text = "🎉 **আপনি আমাদের সাথে সফলভাবে যুক্ত আছেন।**\n\nনিচের বাটনগুলো চেপে আপনার অপশন বেছে নিন:"
        link_title = "📌 **গুরুত্বপূর্ণ লিংকসমূহ:**"

    url_btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=0 AND btn_type='url' ORDER BY id ASC", fetchall=True)
    inline_kb = None
    if url_btns:
        inline_kb = InlineKeyboardMarkup([[InlineKeyboardButton(b["title"], url=b["content"])] for b in url_btns])

    target_msg = update.callback_query.message if update.callback_query else update.message
    await target_msg.reply_text(msg_text, reply_markup=get_bottom_keyboard(user.id, lang), parse_mode="Markdown")
    if inline_kb:
        await target_msg.reply_text(link_title, reply_markup=inline_kb)


# =========================================================
# HELPER: DISPATCH USER SUBMISSION TO ALL ADMINS
# =========================================================

async def dispatch_submission_to_admins(context, sub_id, user_id, username, doc_file_id, valid_emails):
    owner, helpers = get_all_admins()
    admin_list = [int(owner)]
    for h in helpers:
        try:
            admin_list.append(int(h["user_id"]))
        except (ValueError, TypeError):
            pass
    all_admin_ids = list(set(admin_list))

    chat_url = get_user_chat_url(user_id, username)
    mail_preview = "\n".join([f"{i+1}. {m}" for i, m in enumerate(valid_emails)])
    admin_kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("💬 ইউজারের সাথে চ্যাট", url=chat_url)],
        [InlineKeyboardButton("📥 ধাপ ১: প্রাথমিক বাছাই ও রিসিভ", callback_data=f"sub_s1_open:{sub_id}")],
        [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:1:{sub_id}")]
    ])

    for a_id in all_admin_ids:
        try:
            sent_msg = await context.bot.send_document(
                chat_id=int(a_id),
                document=doc_file_id,
                caption=(
                    f"📁 **নতুন জিমেইল সাবমিশন (#{sub_id})**\n"
                    f"🆔 ইউজার: `{user_id}`\n"
                    f"✉️ জিমেইল সংখ্যা: {len(valid_emails)} টি\n\n"
                    f"📋 তালিকা:\n{mail_preview}"
                ),
                reply_markup=admin_kb
            )
            db_execute(
                "INSERT INTO submission_admin_messages (sub_id, admin_id, message_id) VALUES (%s, %s, %s)",
                (sub_id, int(a_id), sent_msg.message_id)
            )
        except Exception as e:
            print(f"Failed to dispatch to admin {a_id}: {e}")


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
        lang = get_user_lang(user.id)
        if is_maintenance_mode() and not is_admin(user.id):
            clean_m, m_kb = parse_text_and_buttons(get_custom_msg("maintenance", lang))
            await update.message.reply_text(clean_m, reply_markup=m_kb, parse_mode="Markdown")
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

        add_user(user, referrer_id)
        u_data = get_user(user.id)
        seen_rules = u_data["seen_rules"] if u_data else 0

        if not seen_rules:
            name = user.first_name or "User"
            welcome_text = get_custom_msg("welcome", lang).replace("{name}", name)
            clean_w, w_kb = parse_text_and_buttons(welcome_text)

            prompt_text = (
                f"{clean_w}\n\n👇 **বট চালু করতে নিচের চ্যানেল ও গ্রুপে যুক্ত হয়ে ভেরিফাই বাটনে চাপ দিন:**"
                if lang == 'bn' else
                f"{clean_w}\n\n👇 **To start the bot, join the channel & group below and click Verify:**"
            )
            link_title = "📌 **গুরুত্বপূর্ণ লিংক:**" if lang == 'bn' else "📌 **Important Links:**"

            await update.message.reply_text(
                prompt_text,
                reply_markup=get_first_time_markup(lang),
                parse_mode="Markdown"
            )
            if w_kb:
                await update.message.reply_text(link_title, reply_markup=w_kb)
            return

        joined = await is_user_joined_all(context.bot, user.id)
        if not joined:
            rejoin_txt = (
                "⚠️ আপনি আমাদের চ্যানেল বা গ্রুপে যুক্ত নেই!\n\n👇 দয়া করে জয়েন হয়ে ভেরিফাই বাটনে চাপ দিন:"
                if lang == 'bn' else
                "⚠️ You are not joined in our channel or group!\n\n👇 Please join and click Verify:"
            )
            await update.message.reply_text(
                rejoin_txt,
                reply_markup=get_rejoin_markup(lang)
            )
            return

        await show_main_menu(update, context)

    except Exception as e:
        print(f"Error in start: {e}")
        try:
            await update.message.reply_text(
                "স্বাগতম! নিচের বাটনগুলো চেপে আমাদের চ্যানেল ও গ্রুপে যুক্ত হয়ে ভেরিফাই করুন:",
                reply_markup=get_rejoin_markup('bn')
            )
        except Exception:
            pass


async def set_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/setadmin <secret_key>`")
        return

    if context.args[0] == ADMIN_SECRET_KEY:
        lang = get_user_lang(user_id)
        await update.message.reply_text(f"👑 আপনার আইডি ({user_id}) ওনার হিসেবে সেট করা আছে। নিচে বড় কিবোর্ডে 🔐 ADMIN PANEL পাবেন।", reply_markup=get_bottom_keyboard(user_id, lang))
    else:
        await update.message.reply_text("❌ পাসওয়ার্ড ভুল!")


async def add_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if user_id != get_owner_id():
        await update.message.reply_text("❌ শুধুমাত্র মূল মালিক (Owner) নতুন অ্যাডমিন যোগ করতে পারেন।")
        return

    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/addadmin <user_id> [Admin_Name]`\nউদাহরণ: `/addadmin 123456789 Admin 1`")
        return

    try:
        target_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন।")
        return

    admin_name = " ".join(context.args[1:]) if len(context.args) > 1 else "Admin"
    add_helper(target_id, admin_name)
    await update.message.reply_text(f"✅ নতুন সহযোগী অ্যাডমিন যোগ করা হয়েছে!\n👤 নাম: **{admin_name}**\n🆔 আইডি: `{target_id}`", parse_mode="Markdown")


async def remove_admin_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    user_id = update.effective_user.id
    if user_id != get_owner_id():
        await update.message.reply_text("❌ শুধুমাত্র মূল মালিক (Owner) অ্যাডমিন বাদ দিতে পারেন।")
        return

    if not context.args:
        await update.message.reply_text("ব্যবহার নিয়ম: `/removeadmin <user_id>`")
        return

    try:
        target_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন।")
        return

    remove_helper(target_id)
    await update.message.reply_text(f"🗑 সহযোগী অ্যাডমিন ক্ষমতা বাতিল করা হয়েছে!\nআইডি: `{target_id}`", parse_mode="Markdown")


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
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("❌ আপনার অ্যাডমিন অ্যাক্সেস নেই।")
        return

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("👨‍💼 Open Admin Panel", callback_data="admin_panel")]
    ])
    await update.message.reply_text("🔐 Admin Access Granted.", reply_markup=keyboard)


# =========================================================
# ADMIN CONTROL PANEL
# =========================================================

async def show_admin_panel(query, user_id):
    m_status = "🔴 বট বন্ধ (Maintenance ON)" if is_maintenance_mode() else "🟢 বট চালু (Active)"
    owner = get_owner_id()
    is_owner = (owner and int(user_id) == int(owner))

    keyboard = [
        [InlineKeyboardButton(f"🛠 গ্লোবাল স্ট্যাটাস: {m_status}", callback_data="toggle_maintenance")],
        [
            InlineKeyboardButton("📥 ইনবক্স ফাইল আপলোড", callback_data="adm_wizard_inbox_start"),
            InlineKeyboardButton("📨 পেন্ডিং ফাইলস", callback_data="admin_pending_files")
        ],
        [
            InlineKeyboardButton("💸 পেন্ডিং উইথড্র", callback_data="admin_pending_withdraws"),
            InlineKeyboardButton("💰 ব্যালেন্স ম্যানেজমেন্ট", callback_data="adm_balance_menu")
        ],
        [
            InlineKeyboardButton("🎛️ বাটন অন/অফ ও নোটিশ", callback_data="admin_switches_menu"),
            InlineKeyboardButton("✏️ কিবোর্ড বাটন প্রদর্শন ও নাম", callback_data="admin_default_btns_menu")
        ],
        [
            InlineKeyboardButton("⚙️ আর্থিক লেনদেন ও লিমিট", callback_data="admin_limits_menu"),
            InlineKeyboardButton("👥 ইউজার ডিরেক্টরি", callback_data="admin_all_users:0")
        ],
        [
            InlineKeyboardButton("👥 রেফারেল ট্র্যাকার", callback_data="admin_ref_tracker"),
            InlineKeyboardButton("📂 ফাইল হিস্ট্রি ও রিপোর্ট", callback_data="admin_file_history:0")
        ],
        [
            InlineKeyboardButton("➕ ম্যানুয়াল রেফার", callback_data="admin_add_ref_prompt"),
            InlineKeyboardButton("🎛 বড় কীবোর্ড বাটন ও সাব-মেনু", callback_data="dyn_manage:0")
        ],
        [
            InlineKeyboardButton("💬 বাটন মেসেজ কন্ট্রোল", callback_data="admin_messages_menu"),
            InlineKeyboardButton("📢 Broadcast Post", callback_data="admin_broadcast")
        ],
        [
            InlineKeyboardButton("✉️ Single Message", callback_data="admin_single_message"),
            InlineKeyboardButton("🧹 ক্লিয়ার ক্যাশ", callback_data="admin_clear_cache")
        ],
        [
            InlineKeyboardButton("🗑️ সম্পূর্ণ বট ডাটা ক্লিয়ার", callback_data="admin_clear_all_data_prompt"),
            InlineKeyboardButton("👤 নির্দিষ্ট ইউজারের ডাটা মুছুন", callback_data="admin_clear_user_data_prompt")
        ],
        [InlineKeyboardButton("📊 Statistics", callback_data="admin_stats")]
    ]

    if is_owner:
        keyboard.append([InlineKeyboardButton("👥 অ্যাডমিন তালিকা ও নাম এডিট (Owner Only)", callback_data="owner_manage_admins")])

    keyboard.append([InlineKeyboardButton("❌ প্যানেল বন্ধ করুন", callback_data="admin_close")])

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
            [InlineKeyboardButton(f"🚀 ধাপ ১ নিশ্চিত করুন ({len(selected_indices)}টি পর্যালোচনায় / {len(emails)-len(selected_indices)}টি বাতিল)", callback_data=f"sub_c1:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল (Reject All)", callback_data=f"sub_rj_all_prompt:1:{sub_id}")]
        ]
    else:
        action_rows = [
            [InlineKeyboardButton(f"💰 চূড়ান্ত অনুমোদন দিন ({len(selected_indices)}টি বৈধ - ৳{len(selected_indices)*rate:.2f})", callback_data=f"sub_c3:{sub_id}")],
            [InlineKeyboardButton("✅ সবগুলোই ঠিক আছে", callback_data=f"sub_accept_all_s3:{sub_id}"),
             InlineKeyboardButton("❌ পুরো ফাইল বাতিল (Reject All)", callback_data=f"sub_rj_all_prompt:{stage_num}:{sub_id}")]
        ]

    return InlineKeyboardMarkup(chunked + action_rows)


def build_reason_selection_keyboard(stage_num, sub_id, current_email_idx, total_rejected_count):
    kb = [
        [InlineKeyboardButton("🔑 পাসওয়ার্ড ভুল / ইনভ্যালিড ক্রেডেনশিয়াল", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt1")],
        [InlineKeyboardButton("📱 2-Step ভেরিফিকেশন অন", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt2")],
        [InlineKeyboardButton("⚠️ ভেরিফিকেশন প্রবলেম", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt3")],
        [InlineKeyboardButton("🚫 অ্যাকাউন্ট ডিজেবল্ড অথবা সাসপেন্ডেড", callback_data=f"sub_rs_set:{stage_num}:{sub_id}:{current_email_idx}:opt4")],
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

    if data == "dismiss_msg":
        try:
            await query.message.delete()
        except Exception:
            pass
        return

    if data.startswith("set_lang:"):
        new_lang = data.split(":")[1]
        set_user_lang(user_id, new_lang)
        try:
            await query.message.delete()
        except Exception:
            pass

        if new_lang == "en":
            success_msg = "✅ Language successfully changed to English!"
        else:
            success_msg = "✅ ভাষা সফলভাবে বাংলায় পরিবর্তন করা হয়েছে!"

        await context.bot.send_message(
            chat_id=user_id,
            text=success_msg,
            reply_markup=get_bottom_keyboard(user_id, new_lang)
        )
        return

    if data == "check_joined":
        joined = await is_user_joined_all(context.bot, user_id)
        lang = get_user_lang(user_id)
        if joined:
            mark_rules_seen(user_id)
            await query.message.delete()
            await show_main_menu(update, context)
        else:
            alert_txt = "❌ আপনি এখনো জয়েন করেননি! চ্যানেল ও গ্রুপে জয়েন হয়ে আবার ভেরিফাই চাপুন।" if lang == 'bn' else "❌ You haven't joined yet! Please join channel & group and click Verify again."
            await query.answer(alert_txt, show_alert=True)
        return

    if data.startswith("dyn_click:"):
        btn_id = int(data.split(":")[1])
        btn = db_execute("SELECT * FROM dynamic_buttons WHERE id=%s", (btn_id,), fetchone=True)
        if not btn:
            await query.answer("বাটন পাওয়া যায়নি!")
            return

        messages = btn["content"].split("---SPLIT---")
        for m in messages:
            cleaned = m.strip()
            if cleaned:
                cl, kb = parse_text_and_buttons(cleaned)
                await safe_send_message(context.bot, chat_id=user_id, text=cl, reply_markup=kb)

        sub_markup = build_dynamic_sub_markup(btn_id)
        if sub_markup:
            await safe_send_message(context.bot, chat_id=user_id, text=f"📂 **{btn['title']}** এর সাব-মেনু:", reply_markup=sub_markup)
        return

    # ইউজার হিস্ট্রি মেনু
    if data == "user_hist_live":
        lang = get_user_lang(user_id)
        subs = db_execute(
            "SELECT * FROM submissions WHERE user_id=%s AND status IN ('pending', 'stage2_review') ORDER BY id DESC", 
            (user_id,), 
            fetchall=True
        )
        if lang == 'en':
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ History Menu", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ Close", callback_data="dismiss_msg")]
            ])
            if not subs:
                await query.edit_message_text("⏳ **You currently have no files processing or under review.**\n\nWhen you submit a new file, its live status will appear here.", reply_markup=back_kb, parse_mode="Markdown")
                return

            hist_msg = "⏳ **Live & Ongoing Files Status:**\n\n"
            for s in subs:
                emails = json.loads(s["emails_json"])
                if s["status"] == "pending":
                    st_text = "🟡 New Submission (Waiting for admin review)"
                else:
                    review_count = len(json.loads(s["in_review_json"])) if s["in_review_json"] else 0
                    st_text = f"🟠 Under 24-48 hours observation ({review_count} Gmails)"

                hist_msg += (
                    f"📁 **File ID: #{s['id']}**\n"
                    f"📅 Date: {s['created_at'][:10]}\n"
                    f"✉️ Total Gmails: {len(emails)}\n"
                    f"⚡ Status: {st_text}\n"
                    f"-----------------------------\n"
                )
        else:
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ হিস্ট্রি মেনু", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ বন্ধ করুন", callback_data="dismiss_msg")]
            ])
            if not subs:
                await query.edit_message_text("⏳ **বর্তমানে আপনার কোনো ফাইল প্রসেসিং বা পর্যালোচনায় নেই।**\n\nনতুন ফাইল জমা দিলে তার লাইভ স্ট্যাটাস এখানে দেখা যাবে।", reply_markup=back_kb, parse_mode="Markdown")
                return

            hist_msg = "⏳ **আপনার লাইভ ও চলমান ফাইলগুলোর স্ট্যাটাস:**\n\n"
            for s in subs:
                emails = json.loads(s["emails_json"])
                if s["status"] == "pending":
                    st_text = "🟡 নতুন জমা (অ্যাডমিন পর্যালোচনার অপেক্ষায়)"
                else:
                    review_count = len(json.loads(s["in_review_json"])) if s["in_review_json"] else 0
                    st_text = f"🟠 ২৪-৪৮ ঘণ্টা পর্যবেক্ষণে রয়েছে ({review_count}টি জিমেইল)"

                hist_msg += (
                    f"📁 **ফাইল ID: #{s['id']}**\n"
                    f"📅 তারিখ: {s['created_at'][:10]}\n"
                    f"✉️ মোট জিমেইল: {len(emails)} টি\n"
                    f"⚡ স্ট্যাটাস: {st_text}\n"
                    f"-----------------------------\n"
                )
        await query.edit_message_text(hist_msg, reply_markup=back_kb, parse_mode="Markdown")
        return

    elif data == "user_hist_all":
        lang = get_user_lang(user_id)
        subs = db_execute("SELECT * FROM submissions WHERE user_id=%s ORDER BY id DESC LIMIT 15", (user_id,), fetchall=True)
        if lang == 'en':
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ History Menu", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ Close", callback_data="dismiss_msg")]
            ])
            if not subs:
                await query.edit_message_text("📜 You have not submitted any files yet.", reply_markup=back_kb)
                return

            hist_msg = "📜 **All Past Submission History (Latest 15):**\n\n"
            for s in subs:
                emails = json.loads(s["emails_json"])
                if s["status"] == "pending":
                    st_text = "🟡 New Submission (Pending review)"
                elif s["status"] == "stage2_review":
                    st_text = "🟠 Under 24-48h observation"
                elif s["status"] == "completed":
                    st_text = f"🟢 Completed (Accepted: {s['accepted_count']} | Rejected: {s['rejected_count']})"
                else:
                    st_text = f"🔴 Fully Rejected ({s['rejected_count']} bad)"

                hist_msg += (
                    f"📁 **File ID: #{s['id']}** | {s['created_at'][:10]}\n"
                    f"✉️ Total: {len(emails)} | Result: {st_text}\n"
                    f"-----------------------------\n"
                )
        else:
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ হিস্ট্রি মেনু", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ বন্ধ করুন", callback_data="dismiss_msg")]
            ])
            if not subs:
                await query.edit_message_text("📜 আপনি এখনও কোনো ফাইল সাবমিট করেননি।", reply_markup=back_kb)
                return

            hist_msg = "📜 **আপনার অতীতের সকল সাবমিশন হিস্ট্রি (সর্বশেষ ১৫টি):**\n\n"
            for s in subs:
                emails = json.loads(s["emails_json"])
                if s["status"] == "pending":
                    st_text = "🟡 নতুন জমা (পর্যালোচনার অপেক্ষায়)"
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
        await query.edit_message_text(hist_msg, reply_markup=back_kb, parse_mode="Markdown")
        return

    elif data == "user_hist_withdrawals":
        lang = get_user_lang(user_id)
        wds = db_execute("SELECT * FROM withdrawals WHERE user_id=%s ORDER BY id DESC LIMIT 20", (user_id,), fetchall=True)
        if lang == 'en':
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ History Menu", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ Close", callback_data="dismiss_msg")]
            ])
            if not wds:
                await query.edit_message_text("💸 **No withdrawal history found.**", reply_markup=back_kb, parse_mode="Markdown")
                return

            u_rate = get_config("usdt_rate", float)
            wd_msg = "💸 **Your Withdrawal History (Live & Past):**\n\n"
            for w in wds:
                status_map = {
                    "pending": "🟡 Pending (Processing)",
                    "approved": "🟢 Success / Paid (Approved)",
                    "rejected": "🔴 Rejected"
                }
                st_text = status_map.get(w["status"], w["status"])
                rec_bdt = float(w["receive_amount"])
                rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0

                wd_msg += (
                    f"🆔 **Withdraw ID: #{w['id']}**\n"
                    f"📅 Date: {w['created_at'][:10]}\n"
                    f"💳 Method: **{w['method']}** (`{w['account']}`)\n"
                    f"💰 Amount: ৳{float(w['amount']):.2f} (Fee: ৳{float(w['fee']):.2f})\n"
                    f"💵 Payout: ৳{rec_bdt:.2f} BDT"
                )
                if w["method"] == "Binance":
                    wd_msg += f" (${rec_usdt:.2f} USDT)"
                wd_msg += f"\n⚡ Status: **{st_text}**\n-----------------------------\n"
        else:
            back_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⬅️ হিস্ট্রি মেনু", callback_data="user_hist_menu"),
                 InlineKeyboardButton("❌ বন্ধ করুন", callback_data="dismiss_msg")]
            ])
            if not wds:
                await query.edit_message_text("💸 **আপনার কোনো উইথড্রয়াল হিস্ট্রি পাওয়া যায়নি।**", reply_markup=back_kb, parse_mode="Markdown")
                return

            u_rate = get_config("usdt_rate", float)
            wd_msg = "💸 **আপনার উইথড্রয়াল হিস্ট্রি (লাইভ ও অতীত):**\n\n"
            for w in wds:
                status_map = {
                    "pending": "🟡 পেন্ডিং (প্রসেসিং চলছে)",
                    "approved": "🟢 সফল / পেইড (Approved)",
                    "rejected": "🔴 বাতিল (Rejected)"
                }
                st_text = status_map.get(w["status"], w["status"])
                rec_bdt = float(w["receive_amount"])
                rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0

                wd_msg += (
                    f"🆔 **উইথড্র ID: #{w['id']}**\n"
                    f"📅 তারিখ: {w['created_at'][:10]}\n"
                    f"💳 মেথড: **{w['method']}** (`{w['account']}`)\n"
                    f"💰 উত্তোলিত: ৳{float(w['amount']):.2f} (ফি: ৳{float(w['fee']):.2f})\n"
                    f"💵 পেমেন্ট পরিমাণ: ৳{rec_bdt:.2f} BDT"
                )
                if w["method"] == "Binance":
                    wd_msg += f" (${rec_usdt:.2f} USDT)"
                wd_msg += f"\n⚡ বর্তমান অবস্থা: **{st_text}**\n-----------------------------\n"

        await query.edit_message_text(wd_msg, reply_markup=back_kb, parse_mode="Markdown")
        return

    elif data == "user_hist_menu":
        lang = get_user_lang(user_id)
        if lang == 'en':
            hist_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⏳ Live Ongoing Submissions (Review)", callback_data="user_hist_live")],
                [InlineKeyboardButton("📜 All Past Submission History", callback_data="user_hist_all")],
                [InlineKeyboardButton("💸 Withdrawal History (Live & Past)", callback_data="user_hist_withdrawals")],
                [InlineKeyboardButton("❌ Close", callback_data="dismiss_msg")]
            ])
            prompt_text = "📜 **Select your history category:**"
        else:
            hist_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⏳ লাইভ চলমান সাবমিশন (Review)", callback_data="user_hist_live")],
                [InlineKeyboardButton("📜 অতীতের সকল সাবমিশন হিস্ট্রি", callback_data="user_hist_all")],
                [InlineKeyboardButton("💸 উইথড্র হিস্ট্রি (লাইভ ও অতীত)", callback_data="user_hist_withdrawals")],
                [InlineKeyboardButton("❌ বন্ধ করুন", callback_data="dismiss_msg")]
            ])
            prompt_text = "📜 **আপনার হিস্ট্রি ক্যাটাগরি নির্বাচন করুন:**"

        await query.edit_message_text(
            prompt_text,
            reply_markup=hist_kb,
            parse_mode="Markdown"
        )
        return

    # অ্যাডমিন গার্ড
    if not is_admin(user_id):
        return

    if data == "admin_panel":
        context.user_data.pop("state", None)
        await show_admin_panel(query, user_id)

    elif data == "admin_close":
        await query.message.delete()

    elif data == "toggle_maintenance":
        new_state = toggle_maintenance_mode()
        msg = "বট বন্ধ করা হয়েছে (ইউজাররা রক্ষণাবেক্ষণ নোটিশ পাবে)!" if new_state else "বট সফলভাবে চালু করা হয়েছে!"
        await query.answer(msg, show_alert=True)
        await show_admin_panel(query, user_id)

    elif data == "admin_clear_all_data_prompt":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("⚠️ হ্যাঁ, সব ডাটা চিরতরে মুছুন", callback_data="admin_clear_all_data_confirm")],
            [InlineKeyboardButton("❌ বাতিল", callback_data="admin_panel")]
        ])
        await query.edit_message_text(
            "⚠️ **সতর্কতা: সম্পূর্ণ বট ডাটাবেজ ক্লিয়ার করা!**\n\n"
            "আপনি কি নিশ্চিত যে আপনি সকল ইউজার, সাবমিশন ফাইল, উইথড্র রিকোয়েস্ট এবং হিস্ট্রি চিরতরে মুছে ফেলতে চান? এটি আর ফিরিয়ে আনা যাবে না!",
            reply_markup=kb,
            parse_mode="Markdown"
        )

    elif data == "admin_clear_all_data_confirm":
        db_execute("DELETE FROM users")
        db_execute("DELETE FROM withdrawals")
        db_execute("DELETE FROM submissions")
        db_execute("DELETE FROM submission_admin_messages")
        db_execute("DELETE FROM submitted_emails")
        db_execute("DELETE FROM helpers")
        init_db()
        await query.answer("বটের সকল ডাটা সফলভাবে সম্পূর্ণ ক্লিয়ার করা হয়েছে!", show_alert=True)
        await show_admin_panel(query, user_id)

    elif data == "admin_clear_user_data_prompt":
        context.user_data["state"] = "waiting_clear_user_id"
        await query.edit_message_text(
            "👤 **নির্দিষ্ট ইউজারের সমস্ত ডাটা মুছুন:**\n\n"
            "যে ইউজারের ডাটা (প্রোফাইল, ব্যালেন্স, সাবমিশন, উইথড্র) মুছে ফেলতে চান তার **টেলিগ্রাম ইউজার আইডি (Numeric ID)** লিখে পাঠান:\n\n"
            "(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_pending_withdraws":
        wds = db_execute("SELECT * FROM withdrawals WHERE status='pending' ORDER BY id ASC", fetchall=True)
        if not wds:
            await query.edit_message_text(
                "💸 **বর্তমানে কোনো পেন্ডিং উইথড্র রিকোয়েস্ট নেই!**",
                reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")]]),
                parse_mode="Markdown"
            )
            return

        kb = []
        u_rate = get_config("usdt_rate", float)
        for w in wds:
            rec_bdt = float(w['receive_amount'])
            rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0
            btn_txt = f"#{w['id']} | User: {w['user_id']} | ৳{rec_bdt:.2f} ({w['method']})"
            kb.append([InlineKeyboardButton(btn_txt, callback_data=f"adm_view_wd:{w['id']}")])

        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text(
            f"💸 **পেন্ডিং উইথড্র তালিকা (মোট: {len(wds)}টি):**\n\nঅনুমোদন বা তথ্য দেখতে সিলেক্ট করুন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("adm_view_wd:"):
        w_id = int(data.split(":")[1])
        w = db_execute("SELECT * FROM withdrawals WHERE id=%s", (w_id,), fetchone=True)
        if not w:
            await query.answer("উইথড্র রিকোয়েস্টটি পাওয়া যায়নি!")
            return

        u_rate = get_config("usdt_rate", float)
        rec_bdt = float(w['receive_amount'])
        rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0
        u_info = get_user(w['user_id'])
        c_url = get_user_chat_url(w['user_id'], u_info['username'] if u_info else None)

        text = (
            f"💸 **উইথড্র রিকোয়েস্ট বিস্তারিত (#{w['id']})**\n\n"
            f"👤 ইউজার আইডি: `{w['user_id']}`\n"
            f"📅 তারিখ: {w['created_at'][:19]}\n"
            f"💰 মূল আবেদন: ৳{float(w['amount']):.2f} (ফি: ৳{float(w['fee']):.2f})\n"
            f"💵 পেমেন্ট পরিমাণ: **৳{rec_bdt:.2f} BDT**"
        )
        if w['method'] == "Binance":
            text += f" (${rec_usdt:.2f} USDT)"
        text += f"\n💳 মেথড: **{w['method']}**\n📌 অ্যাকাউন্ট/নম্বর: `{w['account']}`\n⚡ স্ট্যাটাস: **{w['status'].upper()}**"

        kb = []
        if w['status'] == 'pending':
            kb.append([
                InlineKeyboardButton("✅ Approve / Paid", callback_data=f"approve_withdraw:{w['id']}"),
                InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w['id']}")
            ])
        kb.append([InlineKeyboardButton("💬 ইউজারের সাথে চ্যাট", url=c_url)])
        kb.append([InlineKeyboardButton("⬅️ পেন্ডিং তালিকায় ফিরুন", callback_data="admin_pending_withdraws")])

        await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif data == "adm_balance_menu":
        kb = [
            [InlineKeyboardButton("➕ ব্যালেন্স যোগ করুন", callback_data="bal_op:add")],
            [InlineKeyboardButton("✂️ ব্যালেন্স কেটে নিন", callback_data="bal_op:cut")],
            [InlineKeyboardButton("🔄 ব্যালেন্স রিসেট করুন (0)", callback_data="bal_op:reset")],
            [InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")]
        ]
        await query.edit_message_text(
            "💰 **ব্যালেন্স ম্যানেজমেন্ট প্যানেল:**\n\n"
            "টাইপ করে কমান্ড দেওয়ার প্রয়োজন নেই। নিচের বাটন চাপুন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("bal_op:"):
        op = data.split(":")[1]
        context.user_data["bal_op_type"] = op
        context.user_data["state"] = "waiting_bal_uid"
        action_names = {"add": "যোগ", "cut": "কাটা", "reset": "রিসেট"}
        await query.edit_message_text(
            f"👤 **ব্যালেন্স {action_names.get(op)} করার জন্য ইউজারের টেলিগ্রাম আইডি (Numeric ID) লিখে পাঠান:**\n\n"
            f"(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_switches_menu":
        buttons = [
            ("sell", "📩 Sell Gmail"),
            ("withdraw", "💸 Withdraw"),
            ("referral", "👥 Referral"),
            ("balance", "💰 Balance"),
            ("history", "📜 History"),
            ("rules", "📜 Rules"),
            ("tutorial", "▶️ Tutorial"),
            ("support", "📞 Support")
        ]
        kb = []
        for b_key, b_title in buttons:
            state_icon = "🟢 অন (Active)" if is_button_active(b_key) else "🔴 অফ (Closed)"
            kb.append([
                InlineKeyboardButton(f"{b_title}: {state_icon}", callback_data=f"btn_switch_toggle:{b_key}"),
                InlineKeyboardButton("✍️ নোটিশ এডিট", callback_data=f"btn_notice_edit_prompt:{b_key}")
            ])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])

        await query.edit_message_text(
            "🎛️ **বাটন অন/অফ ও নোটিশ কন্ট্রোল:**\n\n"
            "যেকোনো বাটনের কার্যকারিতা অন/অফ করতে ক্লিক করুন। বাটন অফ থাকলে ইউজাররা যে নোটিশ পাবে তা '✍️ নোটিশ এডিট' বাটন থেকে সেট করুন:\n\n"
            "💡 *টিপস: নোটিশের নিচে লিংক বাটন যুক্ত করতে লেখার শেষে দিন:* `[বাটন | লিংক]`",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("btn_switch_toggle:"):
        b_key = data.split(":")[1]
        new_st = toggle_button_switch(b_key)
        status_txt = "চালু (Active)" if new_st else "বন্ধ (Disabled)"
        await query.answer(f"বাটনটি {status_txt} করা হয়েছে!", show_alert=True)
        query.data = "admin_switches_menu"
        await button_handler(update, context)

    elif data.startswith("btn_notice_edit_prompt:"):
        b_key = data.split(":")[1]
        context.user_data["editing_notice_bkey"] = b_key
        context.user_data["state"] = "waiting_btn_notice_text"
        current_n = get_button_closed_notice(b_key)
        await query.edit_message_text(
            f"✍️ **'{b_key.upper()}' বাটন অফ নোটিশ এডিট**\n\n"
            f"📝 বর্তমান নোটিশ:\n`{current_n}`\n\n"
            "👉 বাটন অফ থাকা অবস্থায় ইউজারদের জন্য নতুন নোটিশটি লিখে পাঠিয়ে দিন:\n"
            "💡 বাটন দিতে চাইলে লেখার শেষে লিখুন: `[বাটনের নাম | লিংক]`\n\n"
            "(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_default_btns_menu":
        btn_items = [
            ("sell", "📩 Sell Gmail"),
            ("balance", "💰 Balance"),
            ("withdraw", "💸 Withdraw"),
            ("history", "📜 History"),
            ("referral", "👥 Referral"),
            ("rules", "📜 Rules"),
            ("tutorial", "▶️ Tutorial"),
            ("support", "📞 Support")
        ]
        kb = []
        for b_key, b_label in btn_items:
            vis = is_button_visible(b_key)
            vis_text = "👁️ দৃশ্যমান" if vis else "🗑️ হাইড/মুছে ফেলা"
            kb.append([
                InlineKeyboardButton(f"{b_label}: {vis_text}", callback_data=f"toggle_vis:{b_key}"),
                InlineKeyboardButton("✏️ নাম এডিট", callback_data=f"edit_btn:{b_key}")
            ])
        kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text(
            "✏️ **কিবোর্ড বাটন প্রদর্শন ও নাম এডিট:**\n\n"
            "• যেকোনো বাটনে চাপ দিয়ে বড় কিবোর্ড থেকে তা **হাইড/রিমুভ** করতে পারেন বা আবার ফিরিয়ে আনতে পারেন。\n"
            "• '✏️ নাম এডিট' চেপে বাটনের নাম পরিবর্তন করতে পারেন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("toggle_vis:"):
        b_key = data.split(":")[1]
        now_vis = toggle_button_visibility(b_key)
        state_label = "কিবোর্ডে যুক্ত করা হয়েছে" if now_vis else "কিবোর্ড থেকে রিমুভ/হাইড করা হয়েছে"
        await query.answer(f"বাটনটি {state_label}!", show_alert=True)
        query.data = "admin_default_btns_menu"
        await button_handler(update, context)

    elif data.startswith("edit_btn:"):
        btn_key = data.split(":")[1]
        context.user_data["editing_btn_key"] = btn_key
        context.user_data["state"] = "waiting_new_button_title"
        await query.edit_message_text(f"✍️ **'{btn_key.upper()}'** বাটনটির নতুন নাম লিখে চ্যাটে পাঠিয়ে দিন:\n(বাতিল করতে /cancel পাঠান)")

    elif data == "adm_wizard_inbox_start":
        context.user_data["state"] = "wizard_waiting_uid"
        await query.edit_message_text(
            "📥 **ইনবক্স ফাইল সাবমিশন উইজার্ড:**\n\n"
            "যে ইউজার আপনাকে ইনবক্সে ফাইল পাঠিয়েছে, তার **টেলিগ্রাম ইউজার আইডি (Numeric ID)** লিখে পাঠান:\n\n"
            "(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "adm_clear_all_subs_prompt":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("⚠️ হ্যাঁ, সব ফাইল মুছুন", callback_data="adm_clear_all_subs_confirm")],
            [InlineKeyboardButton("❌ বাতিল", callback_data="admin_file_history:0")]
        ])
        await query.edit_message_text(
            "⚠️ **সতর্কতা: সকল সাবমিশন ফাইল হিস্ট্রি মুছে ফেলা!**\n\n"
            "আপনি কি নিশ্চিত যে আপনি ডাটাবেজ থেকে অতীতের সকল ফাইলের হিস্ট্রি চিরতরে মুছে ফেলতে চান?",
            reply_markup=kb,
            parse_mode="Markdown"
        )

    elif data == "adm_clear_all_subs_confirm":
        db_execute("DELETE FROM submissions")
        db_execute("DELETE FROM submission_admin_messages")
        await query.answer("সকল সাবমিশন ফাইল হিস্ট্রি সম্পূর্ণ মুছে ফেলা হয়েছে!", show_alert=True)
        query.data = "admin_file_history:0"
        await button_handler(update, context)

    elif data.startswith("approve_withdraw:"):
        w_id = int(data.split(":")[1])
        w = db_execute("SELECT * FROM withdrawals WHERE id=%s", (w_id,), fetchone=True)
        if not w or w["status"] != "pending":
            await query.answer("উইথড্র রিকোয়েস্টটি ইতিমধ্যে প্রক্রিয়াজাত হয়েছে!", show_alert=True)
            return

        bal = get_balance(w["user_id"])
        if bal < float(w["amount"]):
            db_execute("UPDATE withdrawals SET status='rejected' WHERE id=%s", (w_id,))
            await query.edit_message_text("❌ পর্যাপ্ত ব্যালেন্স না থাকায় অটো-রিজেক্ট করা হলো।")
            return

        update_balance(w["user_id"], -float(w["amount"]))
        db_execute("UPDATE withdrawals SET status='approved' WHERE id=%s", (w_id,))
        u_rate = get_config("usdt_rate", float)
        rec_bdt = float(w['receive_amount'])
        rec_usdt = rec_bdt / u_rate if u_rate > 0 else 0

        await query.edit_message_text(f"✅ Withdrawal #{w_id} Approved! পেমেন্ট পাঠানো সম্পন্ন। ৳{rec_bdt:.2f} (${rec_usdt:.2f})")

        msg = f"✅ উইথড্র সফল হয়েছে!\n💸 পেয়েছেন: ৳{rec_bdt:.2f} BDT"
        if w["method"] == "Binance":
            msg += f" (${rec_usdt:.2f} USDT)"
        msg += f"\n💳 মেথড: {w['method']}\n📌 একাউন্ট/UID: `{w['account']}`"
        await safe_send_message(context.bot, chat_id=w["user_id"], text=msg)

    elif data.startswith("reject_withdraw:"):
        w_id = int(data.split(":")[1])
        w = db_execute("SELECT * FROM withdrawals WHERE id=%s", (w_id,), fetchone=True)
        if not w or w["status"] != "pending":
            await query.answer("উইথড্র রিকোয়েস্টটি ইতিমধ্যে প্রক্রিয়াজাত হয়েছে!", show_alert=True)
            return

        db_execute("UPDATE withdrawals SET status='rejected' WHERE id=%s", (w_id,))
        await query.edit_message_text(f"❌ Withdrawal #{w_id} Rejected.")
        await safe_send_message(context.bot, chat_id=w["user_id"], text=f"❌ আপনার উইথড্রয়াল রিকোয়েস্ট (৳{w['amount']:.2f}) বাতিল করা হয়েছে।")

    elif data == "admin_pending_files":
        subs = db_execute("SELECT * FROM submissions WHERE status IN ('pending', 'stage2_review') ORDER BY id DESC LIMIT 15", fetchall=True)
        if not subs:
            await query.edit_message_text("ℹ️ বর্তমানে কোনো পেন্ডিং বা পর্যালোচনায় থাকা ফাইল নেই।", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]]))
            return

        kb = []
        for s in subs:
            emails = json.loads(s["emails_json"])
            st_text = "🟡 ধাপ ১" if s["status"] == "pending" else "🟠 ধাপ ৩ (পর্যবেক্ষণে)"
            kb.append([InlineKeyboardButton(f"📁 #{s['id']} | User: {s['user_id']} ({len(emails)}টি) - {st_text}", callback_data=f"adm_open_sub_review:{s['id']}")])
        kb.append([InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text("📨 **পেন্ডিং ফাইল ম্যানেজার:**\nযে ফাইলটি প্রসেস করতে চান সিলেক্ট করুন:", reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif data.startswith("adm_open_sub_review:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub:
            await query.answer("ফাইল পাওয়া যায়নি!")
            return

        emails = json.loads(sub["emails_json"])
        u_info = get_user(sub["user_id"])
        chat_url = get_user_chat_url(sub["user_id"], u_info["username"] if u_info else None)

        kb = [[InlineKeyboardButton("💬 ইউজারের সাথে চ্যাট", url=chat_url)]]
        if sub["status"] == "pending":
            kb.append([InlineKeyboardButton("📥 ধাপ ১: প্রাথমিক বাছাই শুরু করুন", callback_data=f"sub_s1_open:{sub_id}")])
            kb.append([InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:1:{sub_id}")])
        elif sub["status"] == "stage2_review":
            kb.append([InlineKeyboardButton("🔍 ধাপ ৩: চূড়ান্ত অনুমোদন শুরু করুন", callback_data=f"sub_s3_open:{sub_id}")])
            kb.append([InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:2:{sub_id}")] )
        kb.append([InlineKeyboardButton("⬅️ পেন্ডিং লিস্ট", callback_data="admin_pending_files")])

        await query.edit_message_text(
            f"📄 **সাবমিশন #{sub_id}**\n\n"
            f"🆔 ইউজার আইডি: `{sub['user_id']}`\n"
            f"✉️ মোট জিমেইল: {len(emails)} টি\n"
            f"⚡ বর্তমান স্ট্যাটাস: {sub['status']}\n\n"
            "নিচের বাটন চেপে কাজ শুরু করুন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("admin_file_history:"):
        page = int(data.split(":")[1])
        limit = 5
        offset = page * limit
        subs = db_execute("SELECT * FROM submissions ORDER BY id DESC LIMIT %s OFFSET %s", (limit, offset), fetchall=True)
        total_subs = db_execute("SELECT COUNT(*) AS c FROM submissions", fetchone=True)["c"]

        if not subs:
            kb = [
                [InlineKeyboardButton("➕ ইনবক্স ফাইল আপলোড", callback_data="adm_wizard_inbox_start")],
                [InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]
            ]
            await query.edit_message_text("📂 কোনো সাবমিশন ফাইল পাওয়া যায়নি।", reply_markup=InlineKeyboardMarkup(kb))
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

        kb.append([InlineKeyboardButton("🗑 সকল ফাইল হিস্ট্রি মুছুন (Clear All)", callback_data="adm_clear_all_subs_prompt")])
        kb.append([InlineKeyboardButton("⬅️ Back to Panel", callback_data="admin_panel")])

        await query.edit_message_text(f"📂 সকল ইউজারদের ফাইল তালিকা (মোট: {total_subs} টি):\nফাইল দেখতে ক্লিক করুন:", reply_markup=InlineKeyboardMarkup(kb))

    elif data.startswith("adm_view_sub:"):
        sub_id = int(data.split(":")[1])
        s = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not s:
            await query.answer("ফাইল পাওয়া যায়নি!")
            return

        emails = json.loads(s["emails_json"])
        review_emails = json.loads(s["in_review_json"]) if s["in_review_json"] else []
        mail_list = "\n".join([f"{i+1}. {m}" for i, m in enumerate(emails)])
        rate = get_config("email_rate", float)

        st_map = {
            "pending": "🟡 ধাপ ১: নতুন পেন্ডিং",
            "stage2_review": "🟠 ধাপ ২: ২৪-৪৮ ঘণ্টা পর্যবেক্ষণে",
            "completed": "🟢 ধাপ ৩: চূড়ান্ত অনুমোদিত (সম্পন্ন)",
            "rejected": "🔴 সম্পূর্ণ বাতিল"
        }

        handled_info = f"\n🔒 দায়িত্বপ্রাপ্ত: **{s['handled_by']}**\n" if s["handled_by"] else ""

        info_text = (
            f"📄 সাবমিশন বিস্তারিত (#{s['id']})\n\n"
            f"👤 ইউজার আইডি: `{s['user_id']}`\n"
            f"📅 তারিখ: {s['created_at'][:19]}\n"
            f"⚡ স্ট্যাটাস: {st_map.get(s['status'], s['status'])}{handled_info}\n"
            f"🔹 মূল জিমেইল: {len(emails)} টি\n"
            f"🔹 পর্যালোচনায় নেওয়া হয়েছিল: {len(review_emails)} টি\n"
            f"🔹 চূড়ান্ত অনুমোদিত: {s['accepted_count']} টি (৳{s['accepted_count']*rate:.2f})\n\n"
            f"📋 জিমেইল তালিকা:\n{mail_list}"
        )

        u_info = get_user(s["user_id"])
        c_url = get_user_chat_url(s["user_id"], u_info["username"] if u_info else None)

        action_kb = [
            [InlineKeyboardButton("💬 ইউজারের সাথে চ্যাট", url=c_url)],
            [InlineKeyboardButton("📥 ডাউনলোড এক্সেল ফাইল (.xlsx)", callback_data=f"adm_dl_excel:{sub_id}")]
        ]

        if s["status"] == "rejected":
            action_kb.append([InlineKeyboardButton("🔄 আন-রিজেক্ট / পুনরায় সক্রিয় করুন", callback_data=f"adm_unreject_sub:{sub_id}")])

        action_kb.append([InlineKeyboardButton("🗑 এই ফাইলটি ডিলিট করুন", callback_data=f"adm_del_sub:{sub_id}")])
        action_kb.append([InlineKeyboardButton("⬅️ ফাইল তালিকায় ফিরুন", callback_data="admin_file_history:0")])

        await query.edit_message_text(info_text, reply_markup=InlineKeyboardMarkup(action_kb), parse_mode="Markdown")

    elif data.startswith("adm_unreject_sub:"):
        sub_id = int(data.split(":")[1])
        s = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not s:
            await query.answer("ফাইল পাওয়া যায়নি!")
            return

        db_execute("UPDATE submissions SET status='pending', rejected_count=0 WHERE id=%s", (sub_id,))
        admin_name = get_admin_name(user_id)

        await safe_send_message(
            context.bot,
            chat_id=s["user_id"],
            text=f"🔄 **আপনার সাবমিশন (#{sub_id}) ভুলবশত বাতিল করা হয়েছিল।**\nএটি পুনরায় পর্যালোচনার জন্য গ্রহণ করা হয়েছে।"
        )

        await notify_all_admins(
            context,
            f"🔄 **ফাইল আন-রিজেক্ট অ্যালার্ট (#Submission_{sub_id})**\n\n"
            f"👤 আন-রিজেক্ট করেছেন: **{admin_name}**\n"
            f"🆔 ইউজার: `{s['user_id']}`\n"
            f"⚡ ফাইলটি পুনরায় পেন্ডিং তালিকায় পাঠানো হয়েছে।"
        )

        await query.answer("ফাইলটি পুনরায় সক্রিয় করা হয়েছে!", show_alert=True)
        query.data = f"adm_view_sub:{sub_id}"
        await button_handler(update, context)

    elif data.startswith("adm_del_sub:"):
        sub_id = int(data.split(":")[1])
        db_execute("DELETE FROM submissions WHERE id=%s", (sub_id,))
        await query.answer("ফাইলটি সফলভাবে ডিলিট করা হয়েছে!", show_alert=True)
        query.data = "admin_file_history:0"
        await button_handler(update, context)

    elif data.startswith("adm_dl_excel:"):
        sub_id = int(data.split(":")[1])
        s = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not s:
            await query.answer("ফাইল পাওয়া যায়নি!")
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
        await query.answer("এক্সেল ফাইল পাঠানো হয়েছে!")

    elif data.startswith("sub_s1_open:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            await query.answer("ফাইলটি অলরেডি অন্য একজন প্রসেস করছেন!", show_alert=True)
            return

        rec_name = get_admin_name(user_id)

        if sub["handled_by"] and sub["handled_by"] != rec_name:
            await query.answer(f"⚠️ এই ফাইলটি অলরেডি {sub['handled_by']} রিসিভ করেছেন!", show_alert=True)
            return

        db_execute("UPDATE submissions SET handled_by=%s WHERE id=%s", (rec_name, sub_id))

        other_msgs = db_execute("SELECT admin_id, message_id FROM submission_admin_messages WHERE sub_id=%s", (sub_id,), fetchall=True)
        if other_msgs:
            for om in other_msgs:
                target_admin_id = int(om["admin_id"])
                target_msg_id = int(om["message_id"])
                if target_admin_id != int(user_id):
                    try:
                        await context.bot.delete_message(chat_id=target_admin_id, message_id=target_msg_id)
                    except Exception:
                        pass

        db_execute("DELETE FROM submission_admin_messages WHERE sub_id=%s", (sub_id,))

        team_msg = (
            f"📢 **টিম আপডেট (#Submission_{sub_id})**\n\n"
            f"👤 ফাইল রিসিভ করেছেন: **{rec_name}**\n"
            f"🆔 ইউজার: `{sub['user_id']}`\n"
            f"⚡ প্রাথমিক বাছাই ও ভেরিফিকেশন শুরু হয়েছে।"
        )
        await notify_all_admins(context, team_msg, exclude_user_id=user_id)

        emails = json.loads(sub["emails_json"])
        if sub_id not in admin_stage1_selections:
            admin_stage1_selections[sub_id] = set(range(len(emails)))

        kb = build_stage_keyboard(sub_id, emails, admin_stage1_selections[sub_id], stage_num=1)
        caption_txt = (
            f"📥 **ধাপ ১: প্রাথমিক বাছাই (#{sub_id})**\n"
            f"🆔 ইউজার: `{sub['user_id']}`\n"
            f"🔒 দায়িত্বপ্রাপ্ত: **{rec_name}**\n\n"
            f"যেসব জিমেইল লগইন করা যায় সেগুলো ✅ এবং সমস্যাযুক্ত জিমেইল ❌ করে নিচে কনফার্ম করুন:"
        )

        try:
            await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
        except Exception:
            await query.message.reply_text(caption_txt, reply_markup=kb)

    elif data.startswith("sub_tg:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        idx = int(parts[3])

        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
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
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "pending":
            return

        emails = json.loads(sub["emails_json"])
        selected = admin_stage1_selections.get(sub_id, set(range(len(emails))))

        in_review_emails = [emails[i] for i in sorted(list(selected))]
        s1_rejected = [emails[i] for i in range(len(emails)) if i not in selected]

        if not in_review_emails:
            await query.answer("কমপক্ষে একটি জিমেইল পর্যালোচনায় রাখতে হবে!", show_alert=True)
            return

        if s1_rejected:
            context.user_data["pending_rejection_flow"] = {
                "stage": 1,
                "sub_id": sub_id,
                "rejected_list": s1_rejected,
                "in_review_list": in_review_emails,
                "current_idx": 0,
                "reasons": {},
                "chat_id": query.message.chat_id,
                "message_id": query.message.message_id
            }
            cur_mail = s1_rejected[0]
            kb = build_reason_selection_keyboard(1, sub_id, 0, len(s1_rejected))
            caption_txt = (
                f"⚠️ **ধাপ ১: রিজেকশন কারণ নির্বাচন (১/{len(s1_rejected)})**\n\n"
                f"📧 জিমেইল: `{cur_mail}`\n\n"
                f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান? নিচের একটি কারণ বেছে নিন:"
            )
            try:
                await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
            except Exception:
                await query.message.reply_text(caption_txt, reply_markup=kb)
            return

        await finalize_stage1(context, query, sub_id, in_review_emails, [], user_id)

    elif data.startswith("sub_s3_open:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            await query.answer("Already processed.", show_alert=True)
            return

        review_emails = json.loads(sub["in_review_json"])
        if sub_id not in admin_stage3_selections:
            admin_stage3_selections[sub_id] = set(range(len(review_emails)))

        kb = build_stage_keyboard(sub_id, review_emails, admin_stage3_selections[sub_id], stage_num=3)
        caption_txt = (
            f"💰 **ধাপ ৩: চূড়ান্ত অনুমোদন (#{sub_id})**\n"
            f"🆔 ইউজার: `{sub['user_id']}`\n"
            f"🔒 দায়িত্বপ্রাপ্ত: **{sub['handled_by'] or 'Admin'}**\n\n"
            f"পর্যালোচনায় থাকা {len(review_emails)}টি জিমেইল চেক করুন। অক্ষত জিমেইল ✅ রাখুন এবং নষ্টগুলো ❌ করে অনুমোদন দিন:"
        )
        try:
            await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
        except Exception:
            await query.message.reply_text(caption_txt, reply_markup=kb)

    elif data.startswith("sub_c3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
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
                "reasons": {},
                "chat_id": query.message.chat_id,
                "message_id": query.message.message_id
            }
            cur_mail = final_rejected[0]
            kb = build_reason_selection_keyboard(3, sub_id, 0, len(final_rejected))
            caption_txt = (
                f"⚠️ **ধাপ ৩: রিজেকশন কারণ নির্বাচন (১/{len(final_rejected)})**\n\n"
                f"📧 জিমেইল: `{cur_mail}`\n\n"
                f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান? নিচের একটি কারণ বেছে নিন:"
            )
            try:
                await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
            except Exception:
                await query.message.reply_text(caption_txt, reply_markup=kb)
            return

        await finalize_stage3(context, query, sub_id, final_accepted, [], user_id)

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

        reason_map = {
            "opt1": "পাসওয়ার্ড ভুল / ইনভ্যালিড ক্রেডেনশিয়াল",
            "opt2": "2-Step ভেরিফিকেশন অন",
            "opt3": "ভেরিফিকেশন প্রবলেম",
            "opt4": "অ্যাকাউন্ট ডিজেবল্ড অথবা সাসপেন্ডেড"
        }
        reason_text = reason_map.get(opt, "সমস্যাযুক্ত জিমেইল")

        email = flow["rejected_list"][cur_idx]
        flow["reasons"][email] = reason_text

        next_idx = cur_idx + 1
        if next_idx < len(flow["rejected_list"]):
            flow["current_idx"] = next_idx
            next_mail = flow["rejected_list"][next_idx]
            kb = build_reason_selection_keyboard(stage_num, sub_id, next_idx, len(flow["rejected_list"]))
            caption_txt = (
                f"⚠️ **ধাপ {stage_num}: রিজেকশন কারণ নির্বাচন ({next_idx+1}/{len(flow['rejected_list'])})**\n\n"
                f"📧 জিমেইল: `{next_mail}`\n\n"
                f"এই জিমেইলটির কারণ বেছে নিন:"
            )
            try:
                await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
            except Exception:
                await query.message.reply_text(caption_txt, reply_markup=kb)
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
        caption_txt = (
            f"✍️ **কাস্টম কারণ লিখুন**\n\n"
            f"📧 জিমেইল: `{cur_mail}`\n\n"
            f"এই জিমেইলটি কী কারণে রিজেক্ট করতে চান তা চ্যাটে লিখে পাঠান:\n(বাতিল করতে /cancel পাঠান)"
        )
        try:
            await query.edit_message_caption(caption=caption_txt, parse_mode="Markdown")
        except Exception:
            await query.message.reply_text(caption_txt)

    elif data.startswith("sub_rj_all_prompt:"):
        parts = data.split(":")
        stage_num = int(parts[1])
        sub_id = int(parts[2])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub:
            return

        emails = json.loads(sub["emails_json"]) if stage_num == 1 else json.loads(sub["in_review_json"])

        context.user_data["pending_rejection_flow"] = {
            "stage": 99,
            "sub_id": sub_id,
            "rejected_list": emails,
            "current_idx": 0,
            "reasons": {},
            "chat_id": query.message.chat_id,
            "message_id": query.message.message_id
        }
        cur_mail = emails[0]
        kb = build_reason_selection_keyboard(99, sub_id, 0, len(emails))
        caption_txt = (
            f"❌ **পুরো ফাইল বাতিল: কারণ নির্বাচন (১/{len(emails)})**\n\n"
            f"📧 জিমেইল: `{cur_mail}`\n\n"
            f"কারণ নির্বাচন করুন:"
        )
        try:
            await query.edit_message_caption(caption=caption_txt, reply_markup=kb, parse_mode="Markdown")
        except Exception:
            await query.message.reply_text(caption_txt, reply_markup=kb)

    elif data.startswith("sub_accept_all_s3:"):
        sub_id = int(data.split(":")[1])
        sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
        if not sub or sub["status"] != "stage2_review":
            return

        review_emails = json.loads(sub["in_review_json"])
        await finalize_stage3(context, query, sub_id, review_emails, [], user_id)

    elif data.startswith("admin_all_users:"):
        users = get_all_users_list(limit=25)
        if not users:
            await query.edit_message_text("ℹ️ কোনো ইউজার পাওয়া যায়নি।", reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]]))
            return

        text = "👥 **সকল ইউজার তালিকা (সর্বশেষ ২৫ জন):**\n\n"
        kb = []
        for u in users:
            u_id = u["user_id"]
            bal = u["balance"]
            chat_url = get_user_chat_url(u_id, u["username"])
            text += f"🆔 `{u_id}` | ব্যালেন্স: ৳{bal:.2f}\n"
            kb.append([InlineKeyboardButton(f"💬 চ্যাট ({u_id})", url=chat_url)])

        kb.append([InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")])
        await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif data == "admin_ref_tracker":
        top = get_top_referrers()
        if not top:
            text = "ℹ️ এখনো কেউ রেফার করেনি।"
        else:
            text = "👥 **শীর্ষ রেফারেল ট্র্যাকার (ভেরিফায়েড):**\n\n"
            for r in top:
                u_name = f"@{r['username']}" if r['username'] else "No Username"
                text += f"👤 আইডি: `{r['user_id']}` ({u_name})\n"
                text += f"   ┗ মোট ভেরিফায়েড রেফার: **{r['total_refs']}** (রিয়েল: {r['organic_count']}, ম্যানুয়াল: {r['manual_ref_count']})\n\n"
            text += "💡 ইউজারের বিস্তারিত রেফার দেখতে চ্যাটে লিখুন:\n`/refinfo <user_id>`"

        kb = InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Admin Panel", callback_data="admin_panel")]])
        await query.edit_message_text(text, reply_markup=kb, parse_mode="Markdown")

    elif data == "admin_add_ref_prompt":
        context.user_data["state"] = "waiting_add_ref_val"
        await query.edit_message_text(
            "➕ **ইউজারকে রেফার যোগ করুন:**\n\n"
            "ফরম্যাট: `<user_id> <কত_রেফার_যোগ_করবেন>`\n"
            "উদাহরণ: `123456789 5`\n\n(বাতিল করতে /cancel পাঠান)",
            reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ বাতিল", callback_data="admin_panel")]]),
            parse_mode="Markdown"
        )

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
            "min_withdraw": "মিনিমাম উইথড্র (টাকায়)",
            "withdraw_fee_percent": "উইথড্র ফি (শতাংশে, যেমন 4 বা 5)",
            "usdt_rate": "USDT ডলার রেট (টাকায়)",
            "referral_bonus": "রেফারেল বোনাস (টাকায়)",
            "email_rate": "প্রতি ভ্যালিড জিমেইল রেট (টাকায়)",
            "daily_file_limit": "ইউজারের দৈনিক ফাইল লিমিট (সংখ্যা)",
            "max_emails_per_file": "এক ফাইলে সর্বোচ্চ জিমেইল লিমিট (সংখ্যা)",
            "duplicate_check_days": "ডুপ্লিকেট জিমেইল চেক দিন (সংখ্যা)"
        }
        await query.edit_message_text(f"✍️ **{names.get(cfg_key, cfg_key)}** এর নতুন মান লিখে পাঠিয়ে দিন:\n\n(বাতিল করতে /cancel পাঠান)")

    elif data.startswith("dyn_manage:"):
        p_id = int(data.split(":")[1])
        btns = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=%s ORDER BY id ASC", (p_id,), fetchall=True)

        kb = []
        if btns:
            for b in btns:
                t_label = "🔗" if b["btn_type"] == "url" else "💬"
                kb.append([InlineKeyboardButton(f"{t_label} {b['title']}", callback_data=f"dyn_item_opts:{b['id']}")])

        kb.append([InlineKeyboardButton("➕ নতুন বাটন যোগ করুন (বড় কীবোর্ড)", callback_data=f"dyn_add_init:{p_id}")])
        if p_id != 0:
            parent_row = db_execute("SELECT parent_id FROM dynamic_buttons WHERE id=%s", (p_id,), fetchone=True)
            prev_p = parent_row["parent_id"] if parent_row else 0
            kb.append([InlineKeyboardButton("⬅️ আগের মেনুতে ফিরুন", callback_data=f"dyn_manage:{prev_p}")])
        else:
            kb.append([InlineKeyboardButton("⬅️ Back to Admin Panel", callback_data="admin_panel")])

        lvl_name = "মূল বড় কীবোর্ড (Main Keyboard)" if p_id == 0 else f"সাব-মেনু স্তর (Parent: {p_id})"
        await query.edit_message_text(
            f"🎛 **বড় কীবোর্ড বাটন ও সাব-মেনু বিল্ডার**\n📍 অবস্থান: {lvl_name}\n\n"
            f"যেকোনো বাটনে ক্লিক করে সেটির সাব-বাটন তৈরি করতে পারেন অথবা ডিলিট করতে পারেন:",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode="Markdown"
        )

    elif data.startswith("dyn_item_opts:"):
        b_id = int(data.split(":")[1])
        b = db_execute("SELECT * FROM dynamic_buttons WHERE id=%s", (b_id,), fetchone=True)
        if not b:
            await query.answer("বাটন পাওয়া যায়নি!")
            return

        kb = [
            [InlineKeyboardButton("➕ এর ভেতরে সাব-বাটন যোগ করুন", callback_data=f"dyn_add_init:{b['id']}")],
            [InlineKeyboardButton("📂 এর সাব-বাটনগুলো দেখুন ও ম্যানেজ করুন", callback_data=f"dyn_manage:{b['id']}")],
            [InlineKeyboardButton("🗑 শুধুমাত্র এই বাটনটি ডিলিট করুন", callback_data=f"dyn_del_one:{b['id']}")],
            [InlineKeyboardButton("⬅️ তালিকায় ফিরুন", callback_data=f"dyn_manage:{b['parent_id']}")]
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
        b = db_execute("SELECT * FROM dynamic_buttons WHERE id=%s", (b_id,), fetchone=True)
        if b:
            parent_id = b["parent_id"]
            db_execute("DELETE FROM dynamic_buttons WHERE id=%s", (b_id,))
            await query.answer("বাটনটি সফলভাবে ডিলিট করা হয়েছে!", show_alert=True)
            query.data = f"dyn_manage:{parent_id}"
            await button_handler(update, context)

    elif data.startswith("dyn_add_init:"):
        p_id = int(data.split(":")[1])
        context.user_data["dyn_parent_id"] = p_id
        context.user_data["state"] = "dyn_waiting_title"
        await query.edit_message_text("✍️ নতুন বাটনের নাম লিখে পাঠান (যেমন: 🎁 অফার বা 📚 ভিডিও গাইড):\n(বাতিল করতে /cancel পাঠান)")

    elif data == "owner_manage_admins":
        owner = get_owner_id()
        if not owner or int(user_id) != int(owner):
            await query.answer("❌ শুধুমাত্র মূল মালিকের এক্সেস আছে!", show_alert=True)
            return

        owner, helpers = get_all_admins()
        if not helpers:
            await query.edit_message_text(
                "👥 **currently কোনো সহযোগী অ্যাডমিন যুক্ত নেই।**\n\nনতুন অ্যাডমিন যোগ করতে চ্যাটে কমান্ড পাঠান:\n`/addadmin <user_id> [Admin_Name]`",
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
            f"👉 নতুন যে নাম দিতে চান তা লিখে পাঠান:\n(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )

    elif data == "admin_clear_cache":
        deleted = clear_junk_cache()
        await query.answer(f"ক্যাশ ক্লিয়ার সম্পন্ন! {deleted}টি ফাইল মুছে ফেলা হয়েছে।", show_alert=True)
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
        tips = "\n💡 নিচে বাটন যুক্ত করতে চাইলে লেখার শেষে দিন: `[বাটনের নাম | লিংক]`"
        await query.edit_message_text(f"✍️ **{MESSAGE_NAMES.get(msg_key)}** বাটনের জন্য নতুন মেসেজটি লিখে পাঠিয়ে দিন:{tips}\n\n(ক্যানসেল করতে /cancel পাঠান)")

    elif data.startswith("msg_del:"):
        msg_key = data.split(":")[1]
        delete_custom_msg(msg_key)
        await query.answer("মেসেজ সফলভাবে রিসেট করা হয়েছে!", show_alert=True)
        await show_admin_panel(query, user_id)

    elif data == "admin_broadcast":
        context.user_data["state"] = "admin_broadcast"
        await query.edit_message_text("📢 ব্রডকাস্ট মেসেজ লিখে পাঠান:\n\n(নিচে বাটন দিতে চাইলে লেখার সাথে `[বাটন | লিংক]` দিয়ে দিন)\n(বাতিল করতে /cancel পাঠান)")

    elif data == "admin_single_message":
        context.user_data["state"] = "admin_single_user"
        await query.edit_message_text("✉️ যাকে মেসেজ পাঠাবেন তার টেলিগ্রাম আইডি পাঠান:\n(বাতিল করতে /cancel পাঠান)")

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


async def finalize_stage1(context, query, sub_id, in_review_list, rejected_list, admin_user_id):
    db_execute(
        "UPDATE submissions SET status='stage2_review', in_review_json=%s WHERE id=%s",
        (json.dumps(in_review_list), sub_id)
    )
    sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
    admin_name = get_admin_name(admin_user_id)

    await safe_send_message(
        context.bot,
        chat_id=sub["user_id"],
        text=(
            f"🔍 **আপনার সাবমিশনের প্রাথমিক বাছাই সম্পন্ন হয়েছে (#{sub_id})!**\n\n"
            f"✅ পর্যালোচনার জন্য গৃহীত: {len(in_review_list)} টি জিমেইল\n"
            f"⏳ এই জিমেইলগুলো এখন ২৪ থেকে ৪৮ ঘণ্টা অ্যাডমিনের পর্যবেক্ষণে থাকবে।"
            f"{f' (বাতিল: {len(rejected_list)} টি)' if rejected_list else ''}"
        )
    )

    if rejected_list:
        bio = create_rejected_excel(rejected_items=rejected_list)
        bio.name = f"rejected_stage1_sub_{sub_id}.xlsx"
        await context.bot.send_document(
            chat_id=sub["user_id"],
            document=bio,
            caption=f"📁 **সাবমিশন #{sub_id}: প্রথম ধাপে বাতিল হওয়া জিমেইলের তালিকা ও কারণসমূহ।**"
        )

    await notify_all_admins(
        context,
        f"✅ **ধাপ ১ সম্পন্ন (#Submission_{sub_id})**\n\n"
        f"👤 সম্পন্ন করেছেন: **{admin_name}**\n"
        f"🆔 ইউজার: `{sub['user_id']}`\n"
        f"✉️ পর্যালোচনায় আছে: {len(in_review_list)}টি | বাতিল: {len(rejected_list)}টি"
    )

    try:
        if hasattr(query, "edit_message_caption"):
            await query.edit_message_caption(caption=f"✅ সাবমিশন #{sub_id} - ধাপ ১ সফলভাবে সম্পন্ন হয়েছে!", reply_markup=None)
        elif hasattr(query, "edit_text"):
            await query.edit_text(text=f"✅ সাবমিশন #{sub_id} - ধাপ ১ সফলভাবে সম্পন্ন হয়েছে!")
    except Exception:
        pass


async def finalize_stage3(context, query, sub_id, accepted_list, rejected_list, admin_user_id):
    rate = get_config("email_rate", float)
    total_reward = len(accepted_list) * rate

    sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
    db_execute(
        "UPDATE submissions SET status='completed', accepted_count=%s, rejected_count=%s WHERE id=%s",
        (len(accepted_list), len(rejected_list), sub_id)
    )

    if total_reward > 0:
        update_balance(sub["user_id"], total_reward)

        ref_user = get_user(sub["user_id"])
        if ref_user and ref_user.get("referred_by") and ref_user.get("ref_bonus_rewarded") == 0:
            referrer_id = ref_user["referred_by"]
            ref_bonus = get_config("referral_bonus", float)
            update_balance(referrer_id, ref_bonus)
            db_execute("UPDATE users SET ref_bonus_rewarded=1 WHERE user_id=%s", (sub["user_id"],))
            await safe_send_message(
                context.bot,
                chat_id=referrer_id,
                text=f"🎉 **অভিনন্দন!** আপনার রেফার করা ইউজার (`{sub['user_id']}`) তার প্রথম জিমেইল ফাইল সফলভাবে সম্পন্ন করায় আপনি ৳{ref_bonus:.2f} BDT রেফারেল বোনাস পেয়েছেন!"
            )

    admin_name = get_admin_name(admin_user_id)

    await safe_send_message(
        context.bot,
        chat_id=sub["user_id"],
        text=(
            f"🎉 **আপনার সাবমিশন সম্পূর্ণ অনুমোদিত হয়েছে (#{sub_id})!**\n\n"
            f"✅ চূড়ান্ত অনুমোদিত: {len(accepted_list)} টি জিমেইল\n"
            f"💰 আপনার অ্যাকাউন্টে যোগ হয়েছে: **৳{total_reward:.2f} BDT**"
            f"{f' (নষ্ট বাতিল: {len(rejected_list)} টি)' if rejected_list else ''}"
        )
    )

    if rejected_list:
        bio = create_rejected_excel(rejected_items=rejected_list)
        bio.name = f"rejected_stage3_sub_{sub_id}.xlsx"
        await context.bot.send_document(
            chat_id=sub["user_id"],
            document=bio,
            caption=f"📁 **সাবমিশন #{sub_id}: চূড়ান্ত ধাপে নষ্ট বা বাতিল হওয়া জিমেইলের তালিকা ও কারণসমূহ।**"
        )

    await notify_all_admins(
        context,
        f"💰 **ধাপ ৩ (সম্পন্ন) (#Submission_{sub_id})**\n\n"
        f"👤 অনুমোদন করেছেন: **{admin_name}**\n"
        f"🆔 ইউজার: `{sub['user_id']}`\n"
        f"✅ অনুমোদিত: {len(accepted_list)}টি (৳{total_reward:.2f}) | বাতিল: {len(rejected_list)}টি"
    )

    try:
        if hasattr(query, "edit_message_caption"):
            await query.edit_message_caption(caption=f"💰 সাবমিশন #{sub_id} - চূড়ান্ত অনুমোদন ও পেমেন্ট সম্পন্ন!", reply_markup=None)
        elif hasattr(query, "edit_text"):
            await query.edit_text(text=f"💰 সাবমিশন #{sub_id} - চূড়ান্ত অনুমোদন ও পেমেন্ট সম্পন্ন!")
    except Exception:
        pass


async def finalize_reject_all(context, query, sub_id, rejected_items, admin_user_id):
    sub = db_execute("SELECT * FROM submissions WHERE id=%s", (sub_id,), fetchone=True)
    db_execute(
        "UPDATE submissions SET status='rejected', rejected_count=%s WHERE id=%s",
        (len(rejected_items), sub_id)
    )
    admin_name = get_admin_name(admin_user_id)

    await safe_send_message(
        context.bot,
        chat_id=sub["user_id"],
        text=f"❌ **দুঃখিত! আপনার সাবমিশন ফাইলটি (#{sub_id}) সম্পূর্ণ বাতিল করা হয়েছে।**"
    )

    bio = create_rejected_excel(rejected_items=rejected_items)
    bio.name = f"fully_rejected_sub_{sub_id}.xlsx"
    await context.bot.send_document(
        chat_id=sub["user_id"],
        document=bio,
        caption=f"📁 **সাবমিশন #{sub_id}: সম্পূর্ণ বাতিল হওয়া ফাইলের কারণসহ তালিকা।**"
    )

    await notify_all_admins(
        context,
        f"❌ **ফাইল সম্পূর্ণ বাতিল (#Submission_{sub_id})**\n\n"
        f"👤 বাতিল করেছেন: **{admin_name}**\n"
        f"🆔 ইউজার: `{sub['user_id']}`\n"
        f"🚫 মোট বাতিল: {len(rejected_items)}টি জিমেইল"
    )

    try:
        if hasattr(query, "edit_message_caption"):
            await query.edit_message_caption(caption=f"❌ সাবমিশন #{sub_id} - সম্পূর্ণ ফাইল বাতিল করা হয়েছে।", reply_markup=None)
        elif hasattr(query, "edit_text"):
            await query.edit_text(text=f"❌ সাবমিশন #{sub_id} - সম্পূর্ণ ফাইল বাতিল করা হয়েছে।")
    except Exception:
        pass


# =========================================================
# MESSAGE HANDLER
# =========================================================

async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return

    user = update.effective_user
    if not user:
        return

    lang = get_user_lang(user.id)

    if is_maintenance_mode() and not is_admin(user.id):
        clean_m, m_kb = parse_text_and_buttons(get_custom_msg("maintenance", lang))
        await safe_send_message(context.bot, chat_id=user.id, text=clean_m, reply_markup=m_kb)
        return

    add_user(user)
    text = update.message.text.strip() if update.message.text else ""
    state = context.user_data.get("state")
    rate = get_config("email_rate", float)

    # ভাষা পরিবর্তন মেনু (Language Selection Menu)
    btn_lang_bn = DEFAULT_BUTTON_TITLES["bn"]["lang"]
    btn_lang_en = DEFAULT_BUTTON_TITLES["en"]["lang"]
    if text in [btn_lang_bn, btn_lang_en] or "Change Language" in text or "ভাষা" in text or "Language" in text:
        lang_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🇧🇩 বাংলা (Bangla)", callback_data="set_lang:bn")],
            [InlineKeyboardButton("🇺🇸 English", callback_data="set_lang:en")]
        ])
        await update.message.reply_text(
            "আপনার ভাষা নির্বাচন করুন:\nSelect your language:",
            reply_markup=lang_kb
        )
        return

    if text == "🔐 ADMIN PANEL" and is_admin(user.id):
        await admin(update, context)
        return

    btn_sell = get_button_title("sell", lang)
    btn_bal = get_button_title("balance", lang)
    btn_wd = get_button_title("withdraw", lang)
    btn_hist = get_button_title("history", lang)
    btn_ref = get_button_title("referral", lang)
    btn_rules = get_button_title("rules", lang)
    btn_tut = get_button_title("tutorial", lang)
    btn_sup = get_button_title("support", lang)

    if text == "/cancel":
        context.user_data.clear()
        cancel_txt = "Cancelled." if lang == 'en' else "বাতিল করা হয়েছে।"
        await update.message.reply_text(cancel_txt, reply_markup=get_bottom_keyboard(user.id, lang))
        return

    # নির্দিষ্ট ইউজারের ডাটা ডিলিট করার প্রম্পট হ্যান্ডলার
    if is_admin(user.id) and state == "waiting_clear_user_id":
        if not text.isdigit():
            await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন:")
            return
        t_uid = int(text)
        t_usr = get_user(t_uid)
        if not t_usr:
            await update.message.reply_text("❌ এই ইউজার আইডি বট ডাটাবেজে পাওয়া যায়নি!")
            return

        db_execute("DELETE FROM users WHERE user_id=%s", (t_uid,))
        db_execute("DELETE FROM withdrawals WHERE user_id=%s", (t_uid,))
        db_execute("DELETE FROM submissions WHERE user_id=%s", (t_uid,))
        context.user_data.clear()

        await update.message.reply_text(
            f"✅ ইউজার ID `{t_uid}` এর সমস্ত ডাটা (প্রোফাইল, ব্যালেন্স, সাবমিশন ও উইথড্র হিস্ট্রি) সফলভাবে মুছে ফেলা হয়েছে!",
            parse_mode="Markdown"
        )
        return

    # বাটনভিত্তিক ব্যালেন্স ম্যানেজমেন্ট
    if is_admin(user.id) and state == "waiting_bal_uid":
        if not text.isdigit():
            await update.message.reply_text("❌ সঠিক সংখ্যায় ইউজার আইডি দিন:")
            return
        t_uid = int(text)
        t_usr = get_user(t_uid)
        if not t_usr:
            await update.message.reply_text("❌ এই ইউজার আইডি বট ডাটাবেজে পাওয়া যায়নি!")
            return

        op = context.user_data.get("bal_op_type")
        context.user_data["target_bal_uid"] = t_uid

        if op == "reset":
            db_execute("UPDATE users SET balance=0 WHERE user_id=%s", (t_uid,))
            context.user_data.clear()
            await update.message.reply_text(f"🔄 ব্যালেন্স রিসেট সম্পন্ন! ইউজার: `{t_uid}`, বর্তমান ব্যালেন্স: ৳0.00", parse_mode="Markdown")
            return
        else:
            context.user_data["state"] = "waiting_bal_amount"
            action_names = {"add": "যোগ", "cut": "কাটার"}
            await update.message.reply_text(
                f"👤 ইউজার: `{t_uid}` (বর্তমান ব্যালেন্স: ৳{t_usr['balance']:.2f})\n\n"
                f"👉 কত টাকা {action_names.get(op)} করতে চান তা সংখ্যায় লিখে পাঠান:\n(বাতিল করতে /cancel পাঠান)",
                parse_mode="Markdown"
            )
            return

    if is_admin(user.id) and state == "waiting_bal_amount":
        try:
            amt = float(text)
            if amt <= 0:
                await update.message.reply_text("❌ সঠিক পজিটিভ সংখ্যা দিন:")
                return
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যায় টাকার পরিমাণ লিখুন:")
            return

        t_uid = context.user_data.get("target_bal_uid")
        op = context.user_data.get("bal_op_type")

        if op == "add":
            update_balance(t_uid, amt)
            new_b = get_balance(t_uid)
            await update.message.reply_text(f"✅ ব্যালেন্স যোগ হয়েছে!\nইউজার: `{t_uid}`\nযোগকৃত: +৳{amt:.2f}\nনতুন ব্যালেন্স: ৳{new_b:.2f}", parse_mode="Markdown")
            await safe_send_message(context.bot, chat_id=t_uid, text=f"🎁 আপনার অ্যাকাউন্টে ৳{amt:.2f} BDT যোগ করা হয়েছে!")
        elif op == "cut":
            update_balance(t_uid, -amt)
            new_b = get_balance(t_uid)
            await update.message.reply_text(f"✂️ ব্যালেন্স কেটে নেওয়া হয়েছে!\nইউজার: `{t_uid}`\nকর্তনকৃত: -৳{amt:.2f}\nনতুন ব্যালেন্স: ৳{new_b:.2f}", parse_mode="Markdown")
            await safe_send_message(context.bot, chat_id=t_uid, text=f"⚠️ আপনার অ্যাকাউন্ট থেকে ৳{amt:.2f} BDT কেটে নেওয়া হয়েছে।")

        context.user_data.clear()
        return

    # বাটন অফ নোটিশ
    if is_admin(user.id) and state == "waiting_btn_notice_text":
        b_key = context.user_data.get("editing_notice_bkey")
        if text:
            set_config(f"notice_{b_key}", text)
            clean_n, n_kb = parse_text_and_buttons(text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ **'{b_key.upper()}'** বাটনের অফ নোটিশ সফলভাবে সেভ হয়েছে!\n\n📝 প্রিভিউ:", parse_mode="Markdown")
            await safe_send_message(context.bot, chat_id=user.id, text=clean_n, reply_markup=n_kb)
        return

    # বাটন সুইচ গার্ড
    if not is_admin(user.id):
        button_switch_pairs = [
            (btn_sell, "sell"),
            (btn_bal, "balance"),
            (btn_wd, "withdraw"),
            (btn_hist, "history"),
            (btn_ref, "referral"),
            (btn_rules, "rules"),
            (btn_tut, "tutorial"),
            (btn_sup, "support")
        ]
        for b_title, b_key in button_switch_pairs:
            if text == b_title or (b_key == "sell" and "SELL FRESH GMAIL ACCOUNT" in text) or (b_key == "tutorial" and "TUTORIAL" in text.upper()):
                if not is_button_active(b_key):
                    notice = get_button_closed_notice(b_key, lang)
                    clean_n, n_kb = parse_text_and_buttons(notice)
                    dismiss_btn = InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")
                    if n_kb:
                        inline_rows = [list(r) for r in n_kb.inline_keyboard]
                        inline_rows.append([dismiss_btn])
                        final_kb = InlineKeyboardMarkup(inline_rows)
                    else:
                        final_kb = InlineKeyboardMarkup([[dismiss_btn]])
                    await safe_send_message(context.bot, chat_id=user.id, text=clean_n, reply_markup=final_kb)
                    return

    # ইনবক্স ফাইল আপলোড উইজার্ড: ধাপ ১
    if is_admin(user.id) and state == "wizard_waiting_uid":
        if not text.isdigit():
            await update.message.reply_text("❌ সঠিক সংখ্যায় টেলিগ্রাম আইডি দিন। (যেমন: `123456789`)")
            return
        target_uid = int(text)
        t_user = get_user(target_uid)
        if not t_user:
            await update.message.reply_text("❌ এই ইউজার আইডি বট ডাটাবেজে পাওয়া যায়নি! ইউজারকে বটে অন্তত একবার /start দিতে হবে।")
            return

        context.user_data["wizard_target_uid"] = target_uid
        context.user_data["state"] = "wizard_waiting_file"
        await update.message.reply_text(
            f"✅ ইউজার শনাক্ত হয়েছে: `{target_uid}`\n\n"
            "এবার ইউজারের পাঠানো এক্সেল বা সিএসভি ফাইলটি (.xlsx, .csv) এখানে পাঠিয়ে দিন:\n(বাতিল করতে /cancel পাঠান)",
            parse_mode="Markdown"
        )
        return

    # ইনবক্স ফাইল আপলোড উইজার্ড: ধাপ ২
    if is_admin(user.id) and state == "wizard_waiting_file":
        doc = update.message.document
        if not doc:
            await update.message.reply_text("❌ অনুগ্রহ করে ইউজারের এক্সেল বা সিএসভি ফাইল (.xlsx, .csv) পাঠান।")
            return

        fn = doc.file_name.lower()
        if not (fn.endswith(".xlsx") or fn.endswith(".xls") or fn.endswith(".csv")):
            await update.message.reply_text("❌ শুধুমাত্র Excel বা CSV ফাইল সাপোর্টেড।")
            return

        target_uid = context.user_data.get("wizard_target_uid")
        target_user = get_user(target_uid)

        os.makedirs("downloads", exist_ok=True)
        local_path = os.path.join("downloads", f"wiz_{target_uid}_{doc.file_name}")
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(local_path)

        try:
            valid_emails = validate_gmail_file(local_path)
        except Exception as e:
            if os.path.exists(local_path):
                os.remove(local_path)
            await update.message.reply_text(f"❌ ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        if os.path.exists(local_path):
            os.remove(local_path)

        if not valid_emails:
            await update.message.reply_text("❌ ফাইলে কোনো ভ্যালিড @gmail.com পাওয়া যায়নি!")
            return

        save_submitted_emails(valid_emails)
        emails_json = json.dumps(valid_emails)

        uploader_name = get_admin_name(user.id)

        db_execute("""
            INSERT INTO submissions (user_id, file_id, emails_json, status, handled_by, created_at)
            VALUES (%s, %s, %s, 'pending', %s, %s)
        """, (target_uid, doc.file_id, emails_json, uploader_name, datetime.now().isoformat()))

        last_sub = db_execute("SELECT id FROM submissions WHERE user_id=%s ORDER BY id DESC LIMIT 1", (target_uid,), fetchone=True)
        sub_id = last_sub["id"] if last_sub else 1
        increase_file_count(target_uid)
        context.user_data.clear()

        chat_url = get_user_chat_url(target_uid, target_user["username"])
        mail_preview = "\n".join([f"{i+1}. {m}" for i, m in enumerate(valid_emails)])
        admin_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("💬 ইউজারের সাথে চ্যাট", url=chat_url)],
            [InlineKeyboardButton("📥 ধাপ ১: প্রাথমিক বাছাই ও রিসিভ", callback_data=f"sub_s1_open:{sub_id}")],
            [InlineKeyboardButton("❌ পুরো ফাইল বাতিল", callback_data=f"sub_rj_all_prompt:1:{sub_id}")]
        ])

        await update.message.reply_text(
            f"🎉 **ইনবক্স ফাইল সফলভাবে আপনার ব্যক্তিগত রিভিউ তালিকায় যুক্ত হয়েছে!**\n\n"
            f"📁 ফাইল আইডি: `#{sub_id}`\n"
            f"👤 ইউজার: `{target_uid}`\n"
            f"✉️ মোট জিমেইল: {len(valid_emails)} টি\n\n"
            f"📋 তালিকা:\n{mail_preview}\n\n"
            f"🔒 দায়িত্বপ্রাপ্ত: **{uploader_name}** (অন্য অ্যাডমিনদের কাছে এই ফাইল পাঠানো হয়নি)",
            reply_markup=admin_kb,
            parse_mode="Markdown"
        )

        await safe_send_message(
            context.bot,
            chat_id=target_uid,
            text=(
                f"📥 **আপনার ফাইল অ্যাডমিন কর্তৃক সিস্টেমে যুক্ত করা হয়েছে!**\n\n"
                f"📁 ফাইল আইডি: `#{sub_id}`\n"
                f"✉️ মোট ভ্যালিড Gmail: {len(valid_emails)} টি\n"
                f"⚡ স্ট্যাটাস: 🟡 নতুন জমা (পর্যালোচনার অপেক্ষায়)\n\n"
                f"আপনি বটের '📜 হিস্ট্রি' বাটনে চাপ দিয়ে লাইভ অগ্রগতি দেখতে পারবেন।"
            )
        )
        return

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
                    f"✅ কারণ সেভ হয়েছে!\n\n"
                    f"⚠️ **পরবর্তী জিমেইল ({next_idx+1}/{len(flow['rejected_list'])})**\n"
                    f"📧 জিমেইল: `{next_mail}`\n\n"
                    f"কারণ নির্বাচন করুন:",
                    reply_markup=kb,
                    parse_mode="Markdown"
                )
            else:
                rejected_items = [(em, flow["reasons"].get(em, "Invalid")) for em in flow["rejected_list"]]
                class DummyQuery:
                    def __init__(self, c_id, m_id):
                        self.chat_id = c_id
                        self.message_id = m_id
                    async def edit_message_caption(self, caption, reply_markup=None, parse_mode=None):
                        try:
                            await context.bot.edit_message_caption(chat_id=self.chat_id, message_id=self.message_id, caption=caption, reply_markup=reply_markup, parse_mode=parse_mode)
                        except Exception:
                            await update.message.reply_text(caption, reply_markup=reply_markup)

                dq = DummyQuery(flow.get("chat_id", user.id), flow.get("message_id"))

                if stage_num == 1:
                    await finalize_stage1(context, dq, sub_id, flow["in_review_list"], rejected_items, user.id)
                elif stage_num == 3:
                    await finalize_stage3(context, dq, sub_id, flow["accepted_list"], rejected_items, user.id)
                elif stage_num == 99:
                    await finalize_reject_all(context, dq, sub_id, rejected_items, user.id)
                
                context.user_data.pop("pending_rejection_flow", None)
        return

    if is_admin(user.id) and state == "waiting_cfg_val":
        cfg_key = context.user_data.get("editing_cfg_key")
        try:
            val = float(text)
            if val < 0:
                await update.message.reply_text("❌ ঋণাত্মক মান গ্রহণযোগ্য নয়।")
                return
            set_config(cfg_key, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! **{cfg_key}** এর নতুন মান **{text}** সেভ হয়েছে।", parse_mode="Markdown")
        except ValueError:
            await update.message.reply_text("❌ সঠিক সংখ্যায় মান লিখে পাঠান।")
        return

    if is_admin(user.id) and state == "waiting_add_ref_val":
        context.user_data.clear()
        parts = text.split()
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            t_user, cnt = int(parts[0]), int(parts[1])
            if get_user(t_user):
                add_manual_referral(t_user, cnt)
                await update.message.reply_text(f"✅ সফল হয়েছে! ইউজার `{t_user}`-এর অ্যাকাউন্টে {cnt} রেফার যোগ করা হয়েছে।")
            else:
                await update.message.reply_text("❌ ইউজার ডাটাবেজে পাওয়া যায়নি!")
        else:
            await update.message.reply_text("❌ ফরম্যাট ভুল! `<user_id> <সংখ্যা>` এভাবে দিন।")
        return

    if is_admin(user.id) and state == "dyn_waiting_title":
        context.user_data["dyn_title"] = text
        context.user_data["state"] = "dyn_waiting_content"
        await update.message.reply_text(
            f"✅ বাটনের নাম: **{text}**\n\n"
            "এবার এই বাটনের তথ্য পাঠিয়ে দিন:\n\n"
            "🔹 **যদি লিংক দিতে চান:** সরাসরি URL লিখে পাঠান (যেমন: `https://t.me/...`)\n"
            "🔹 **যদি মেসেজ দিতে চান:** মেসেজটি লিখে দিন।\n"
            "💡 **টিপস:** একাধিক মেসেজ পাঠাতে মাঝখানে `---SPLIT---` লিখুন।",
            parse_mode="Markdown"
        )
        return

    if is_admin(user.id) and state == "dyn_waiting_content":
        parent_id = context.user_data.get("dyn_parent_id", 0)
        title = context.user_data.get("dyn_title")

        b_type = "url" if text.startswith(("http://", "https://", "tg://")) else "message"

        db_execute(
            "INSERT INTO dynamic_buttons (parent_id, title, btn_type, content, show_in_reply) VALUES (%s, %s, %s, %s, 1)",
            (parent_id, title, b_type, text)
        )
        context.user_data.clear()
        await update.message.reply_text(f"🎉 **বাটন সফলভাবে তৈরি হয়েছে!**\n\nনাম: {title}\nধরন: {b_type.upper()}", reply_markup=get_bottom_keyboard(user.id, lang))
        return

    dyn_btn_row = db_execute("SELECT * FROM dynamic_buttons WHERE parent_id=0 AND title=%s", (text,), fetchone=True)
    if dyn_btn_row:
        if dyn_btn_row["btn_type"] == "url":
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔗 ওপেন করুন", url=dyn_btn_row["content"])]])
            await update.message.reply_text(f"👉 **{dyn_btn_row['title']}** লিংকে যেতে নিচের বাটনে চাপ দিন:", reply_markup=kb, parse_mode="Markdown")
        else:
            messages = dyn_btn_row["content"].split("---SPLIT---")
            for m in messages:
                cleaned = m.strip()
                if cleaned:
                    cl, kb = parse_text_and_buttons(cleaned)
                    await safe_send_message(context.bot, chat_id=user.id, text=cl, reply_markup=kb)
            
            sub_markup = build_dynamic_sub_markup(dyn_btn_row["id"])
            if sub_markup:
                await safe_send_message(context.bot, chat_id=user.id, text=f"📂 **{dyn_btn_row['title']}** এর সাব-মেনু:", reply_markup=sub_markup)
        return

    # সেফ ব্রডকাস্ট ইঞ্জিন
    if is_admin(user.id) and state == "admin_broadcast":
        clean_bc, bc_kb = parse_text_and_buttons(text)
        users = db_execute("SELECT user_id FROM users", fetchall=True)
        user_count = 0
        if users:
            for u in users:
                res = await safe_send_message(context.bot, chat_id=u["user_id"], text=clean_bc, reply_markup=bc_kb)
                if res:
                    user_count += 1

        group_res = await safe_send_message(context.bot, chat_id=FORCE_GROUP_CHAT_ID, text=clean_bc, reply_markup=bc_kb)
        group_posted = group_res is not None

        context.user_data.clear()
        
        report_msg = (
            f"📢 **ব্রডকাস্ট সম্পন্ন হয়েছে!**\n\n"
            f"👤 ইউজার ইনবক্সে পৌঁছেছে: {user_count}/{len(users) if users else 0} জনের কাছে\n"
            f"👥 অফিসিয়াল গ্রুপে পোস্ট: {'✅ সফল' if group_posted else '❌ ব্যর্থ (বটকে গ্রুপে অ্যাড করে অ্যাডমিন দিন)'}\n"
            f"📢 পেমেন্ট প্রুফ চ্যানেল: 🚫 বাদ রাখা হয়েছে"
        )
        await update.message.reply_text(report_msg, parse_mode="Markdown")
        return

    if is_button_visible("support") and text == btn_sup:
        sup_text = get_custom_msg("support", lang)
        clean_sup, sup_kb = parse_text_and_buttons(sup_text)
        def_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("💬 সরাসরি মেসেজ পাঠান / Support", url=SUPPORT_URL)],
            [InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]
        ])
        await safe_send_message(context.bot, chat_id=user.id, text=clean_sup, reply_markup=(sup_kb or def_kb))
        return

    if is_button_visible("tutorial") and (text == btn_tut or "TUTORIAL" in text.upper()):
        tut_text = get_custom_msg("tutorial", lang)
        clean_tut, tut_kb = parse_text_and_buttons(tut_text)
        def_tut_kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📺 Watch Video Tutorial", url="https://youtube.com")],
            [InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]
        ])
        await safe_send_message(context.bot, chat_id=user.id, text=clean_tut, reply_markup=(tut_kb or def_tut_kb))
        return

    if not is_admin(user.id):
        joined = await is_user_joined_all(context.bot, user.id)
        if not joined:
            rejoin_txt = (
                "⚠️ আপনি আমাদের চ্যানেল বা গ্রুপে যুক্ত নেই!\n\n👇 দয়া করে জয়েন হয়ে ভেরিফাই বাটনে চাপ দিন:"
                if lang == 'bn' else
                "⚠️ You are not joined in our channel or group!\n\n👇 Please join and click Verify:"
            )
            await update.message.reply_text(
                rejoin_txt,
                reply_markup=get_rejoin_markup(lang)
            )
            return

    owner = get_owner_id()
    if owner and int(user.id) == int(owner) and state == "waiting_helper_new_name":
        target_id = context.user_data.get("target_helper_id")
        if text:
            set_helper_name(target_id, text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! অ্যাডমিন `{target_id}` এর নাম পরিবর্তন করে **{text}** রাখা হয়েছে।", parse_mode="Markdown")
        return

    if is_admin(user.id) and state == "waiting_new_msg_text":
        key = context.user_data.get("editing_msg_key")
        if text:
            set_custom_msg(key, text, lang)
            clean_m, kb_m = parse_text_and_buttons(text)
            context.user_data.clear()
            await update.message.reply_text(f"✅ সফল হয়েছে! '{MESSAGE_NAMES.get(key)}' মেসেজটি সেভ হয়েছে। প্রিভিউ নিচে দেখুন:")
            await safe_send_message(context.bot, chat_id=user.id, text=clean_m, reply_markup=kb_m)
        return

    if is_admin(user.id) and state == "waiting_new_button_title":
        btn_key = context.user_data.get("editing_btn_key")
        if text:
            set_button_title(btn_key, text, lang)
            context.user_data.clear()
            await update.message.reply_text(f"✅ বাটন নাম আপডেট হয়েছে:\n{text}", reply_markup=get_bottom_keyboard(user.id, lang))
        return

    if is_button_visible("history") and text == btn_hist:
        if lang == 'en':
            hist_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⏳ Live Ongoing Submissions (Review)", callback_data="user_hist_live")],
                [InlineKeyboardButton("📜 All Past Submission History", callback_data="user_hist_all")],
                [InlineKeyboardButton("💸 Withdrawal History (Live & Past)", callback_data="user_hist_withdrawals")],
                [InlineKeyboardButton("❌ Close", callback_data="dismiss_msg")]
            ])
            prompt_text = (
                "📜 **Select your history category:**\n\n"
                "🔹 **Live Ongoing Submissions:** Files currently under review or 24-48h observation.\n"
                "🔹 **All Submission History:** Details of all your previous submissions.\n"
                "🔹 **Withdrawal History:** Live info of all pending and successful withdrawals."
            )
        else:
            hist_kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("⏳ লাইভ চলমান সাবমিশন (Review)", callback_data="user_hist_live")],
                [InlineKeyboardButton("📜 অতীতের সকল সাবমিশন হিস্ট্রি", callback_data="user_hist_all")],
                [InlineKeyboardButton("💸 উইথড্র হিস্ট্রি (লাইভ ও অতীত)", callback_data="user_hist_withdrawals")],
                [InlineKeyboardButton("❌ বন্ধ করুন", callback_data="dismiss_msg")]
            ])
            prompt_text = (
                "📜 **আপনার হিস্ট্রি ক্যাটাগরি নির্বাচন করুন:**\n\n"
                "🔹 **লাইভ চলমান সাবমিশন:** বর্তমানে রিভিউ বা ২৪-৪৮ ঘণ্টা পর্যালোচনায় থাকা ফাইল।\n"
                "🔹 **সকল সাবমিশন হিস্ট্রি:** আপনার জমা দেওয়া পূর্বের সকল ফাইলের বিস্তারিত।\n"
                "🔹 **উইথড্র হিস্ট্রি:** আপনার পেন্ডিং ও সফল হওয়া সকল উইথড্রর লাইভ তথ্য।"
            )

        await update.message.reply_text(
            prompt_text,
            reply_markup=hist_kb,
            parse_mode="Markdown"
        )
        return

    if is_button_visible("sell") and (text == btn_sell or "SELL FRESH GMAIL ACCOUNT" in text):
        count = get_today_file_count(user.id)
        d_lim = get_config("daily_file_limit", int)
        e_lim = get_config("max_emails_per_file", int)
        if count >= d_lim:
            limit_msg = f"❌ Daily limit reached ({d_lim}/{d_lim})." if lang == 'en' else f"❌ দৈনিক লিমিট শেষ ({d_lim}/{d_lim})।"
            await update.message.reply_text(limit_msg)
            return

        context.user_data["state"] = "waiting_file"
        sell_msg = get_custom_msg("sell", lang)
        clean_s, s_kb = parse_text_and_buttons(sell_msg)
        cancel_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বাতিল করুন" if lang == 'bn' else "❌ Cancel", callback_data="dismiss_msg")]])
        
        info_text = (
            f"{clean_s}\n\n"
            f"📊 Today's Submissions: {count}/{d_lim}\n"
            f"💰 Rate per valid Gmail: ৳{rate:.2f} BDT\n"
            f"⚠️ Max {e_lim} Gmails allowed per file."
            if lang == 'en' else
            f"{clean_s}\n\n"
            f"📊 আজকের সাবমিশন: {count}/{d_lim}\n"
            f"💰 প্রতি ভ্যালিড জিমেইল রেট: ৳{rate:.2f} BDT\n"
            f"⚠️ ফাইলে সর্বোচ্চ {e_lim}টি জিমেইল থাকতে পারবে।"
        )

        await safe_send_message(
            context.bot,
            chat_id=user.id,
            text=info_text,
            reply_markup=(s_kb or cancel_kb)
        )
        return

    elif is_button_visible("balance") and text == btn_bal:
        b_bdt = get_balance(user.id)
        u_rate = get_config("usdt_rate", float)
        b_usdt = b_bdt / u_rate if u_rate > 0 else 0
        refs = get_referral_count(user.id)
        bal_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
        await update.message.reply_text(
            f"💰 **YOUR ACCOUNT BALANCE**\n\n"
            f"🔹 BDT Balance: ৳{b_bdt:.2f} BDT\n"
            f"🔹 USDT Balance: ${b_usdt:.2f} USDT\n"
            f"👥 Total Referrals (Verified): {refs} users",
            reply_markup=bal_kb,
            parse_mode="Markdown"
        )
        return

    elif is_button_visible("withdraw") and text == btn_wd:
        bal = get_balance(user.id)
        u_rate = get_config("usdt_rate", float)
        min_w = get_config("min_withdraw", float)
        w_fee = get_config("withdraw_fee_percent", float)
        bal_usdt = bal / u_rate if u_rate > 0 else 0
        if bal < min_w:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
            low_msg = f"❌ Minimum withdrawal is ৳{min_w:.2f} BDT. You have: ৳{bal:.2f} BDT" if lang == 'en' else f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.2f} BDT। আপনার আছে: ৳{bal:.2f} BDT"
            await update.message.reply_text(low_msg, reply_markup=kb)
            return

        context.user_data["state"] = "withdraw_amount"
        cancel_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
        
        wd_prompt = (
            f"💸 **WITHDRAWAL SYSTEM**\n\n"
            f"💰 Balance: ৳{bal:.2f} BDT (${bal_usdt:.2f} USDT)\n"
            f"🔹 Minimum Withdraw: ৳{min_w:.2f} BDT\n"
            f"🔹 Fee: {w_fee:.1f}%\n"
            f"🔹 Dollar Rate: $1 = ৳{u_rate:.2f}\n\n"
            f"How much do you want to withdraw? Send in numbers:\n(Send /cancel to abort)"
            if lang == 'en' else
            f"💸 **WITHDRAWAL SYSTEM**\n\n"
            f"💰 ব্যালেন্স: ৳{bal:.2f} BDT (${bal_usdt:.2f} USDT)\n"
            f"🔹 সর্বনিম্ন উইথড্র: ৳{min_w:.2f} BDT\n"
            f"🔹 ফি: {w_fee:.1f}%\n"
            f"🔹 ডলার রেট: $1 = ৳{u_rate:.2f}\n\n"
            f"কত টাকা উইথড্র করতে চান? সংখ্যায় লিখে পাঠান:\n(বাতিল করতে /cancel পাঠান)"
        )

        await update.message.reply_text(
            wd_prompt,
            reply_markup=cancel_kb,
            parse_mode="Markdown"
        )
        return

    elif is_button_visible("referral") and text == btn_ref:
        b_info = await context.bot.get_me()
        ref_link = f"https://t.me/{b_info.username}?start={user.id}"
        refs = get_referral_count(user.id)
        ref_bonus = get_config("referral_bonus", float)
        earned = refs * ref_bonus

        ref_template = get_custom_msg("referral", lang)
        custom_ref = ref_template.replace("{link}", ref_link).replace("{bonus}", str(ref_bonus))
        clean_r, r_kb = parse_text_and_buttons(custom_ref)

        msg = (
            f"{clean_r}\n\n"
            f"📊 Your Referral Stats:\n"
            f"🔹 Verified Successful Referrals: {refs}\n"
            f"🔹 Total Earnings: ৳{earned:.2f} BDT"
            if lang == 'en' else
            f"{clean_r}\n\n"
            f"📊 আপনার রেফারেল পরিসংখ্যান:\n"
            f"🔹 ভেরিফায়েড সফল রেফার: {refs} জন\n"
            f"🔹 মোট অর্জিত আয়: ৳{earned:.2f} BDT"
        )
        def_r_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
        await safe_send_message(context.bot, chat_id=user.id, text=msg, reply_markup=(r_kb or def_r_kb))
        return

    elif is_button_visible("rules") and text == btn_rules:
        name = user.first_name or "User"
        rules = get_custom_msg("rules", lang).replace("{name}", name)
        clean_rul, rul_kb = parse_text_and_buttons(rules)
        def_rul_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
        await safe_send_message(context.bot, chat_id=user.id, text=clean_rul, reply_markup=(rul_kb or def_rul_kb))
        return

    if is_admin(user.id) and state == "admin_single_user":
        try:
            context.user_data["target_user"] = int(text)
            context.user_data["state"] = "admin_single_text"
            await update.message.reply_text("মেসেজটি লিখে পাঠান:\n(বাটন দিতে চাইলে শেষে `[বাটন | লিংক]` দিন)")
        except ValueError:
            await update.message.reply_text("❌ সঠিক নিউমেরিক আইডি দিন।")
        return

    if is_admin(user.id) and state == "admin_single_text":
        t_id = context.user_data.get("target_user")
        clean_s_msg, s_markup = parse_text_and_buttons(text)
        res = await safe_send_message(context.bot, chat_id=t_id, text=clean_s_msg, reply_markup=s_markup)
        if res:
            await update.message.reply_text("✅ মেসেজ পাঠানো হয়েছে।")
        else:
            await update.message.reply_text("❌ মেসেজ পাঠানো সম্ভব হয়নি (ইউজার ব্লক করে থাকতে পারে)।")
        context.user_data.clear()
        return

    if state == "withdraw_amount":
        try:
            amount = float(text)
        except ValueError:
            err_num = "❌ Please enter a valid number." if lang == 'en' else "❌ সঠিক সংখ্যা লিখুন।"
            await update.message.reply_text(err_num)
            return

        min_w = get_config("min_withdraw", float)
        w_fee = get_config("withdraw_fee_percent", float)
        u_rate = get_config("usdt_rate", float)

        if amount < min_w:
            low_w = f"❌ Minimum withdrawal is ৳{min_w:.2f} BDT." if lang == 'en' else f"❌ সর্বনিম্ন উইথড্র ৳{min_w:.2f} BDT।"
            await update.message.reply_text(low_w)
            return

        if amount > get_balance(user.id):
            no_bal = "❌ Insufficient balance." if lang == 'en' else "❌ ব্যালেন্স পর্যাপ্ত নেই।"
            await update.message.reply_text(no_bal)
            return

        fee = amount * w_fee / 100
        rec = amount - fee
        rec_usdt = rec / u_rate if u_rate > 0 else 0

        context.user_data.update({"withdraw_amount": amount, "withdraw_fee": fee, "withdraw_receive": rec})
        context.user_data["state"] = "withdraw_method"

        keyboard = [
            [InlineKeyboardButton("💳 bKash", callback_data="method_bkash")],
            [InlineKeyboardButton("💳 Nagad", callback_data="method_nagad")],
            [InlineKeyboardButton("💰 Binance (UID)", callback_data="method_binance")],
            [InlineKeyboardButton("❌ বাতিল করুন" if lang == 'bn' else "❌ Cancel", callback_data="dismiss_msg")]
        ]
        summary_text = (
            f"💸 Withdrawal Summary:\nWithdraw: ৳{amount:.2f} BDT\nFee: ৳{fee:.2f} BDT\nYou will get: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\nSelect method:"
            if lang == 'en' else
            f"💸 উইথড্র সামারি:\nউত্তোলন: ৳{amount:.2f} BDT\nফি: ৳{fee:.2f} BDT\nপাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\nমেথড নির্বাচন করুন:"
        )
        await update.message.reply_text(
            summary_text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    if state == "withdraw_account":
        method = context.user_data.get("withdraw_method")

        if method in ["Bkash", "Nagad"] and not BD_PHONE_REGEX.match(text):
            err_phone = f"❌ Invalid {method} number! Provide 11-digit Bangladeshi number." if lang == 'en' else f"❌ ভুল {method} নম্বর! ১১ ডিজিটের বাংলাদেশি নম্বর দিন।"
            await update.message.reply_text(err_phone)
            return
        elif method == "Binance" and not BINANCE_UID_REGEX.match(text):
            err_uid = "❌ Invalid Binance UID! Enter 7-10 digit numbers only." if lang == 'en' else "❌ ভুল Binance UID! শুধুমাত্র ৭-১০ ডিজিটের সংখ্যা দিন।"
            await update.message.reply_text(err_uid)
            return

        amt = context.user_data["withdraw_amount"]
        fee = context.user_data["withdraw_fee"]
        rec = context.user_data["withdraw_receive"]
        u_rate = get_config("usdt_rate", float)
        rec_usdt = rec / u_rate if u_rate > 0 else 0

        db_execute("""
            INSERT INTO withdrawals (user_id, amount, fee, receive_amount, method, account, status, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, 'pending', %s)
        """, (user.id, amt, fee, rec, method, text, datetime.now().isoformat()))

        last_w = db_execute("SELECT id FROM withdrawals WHERE user_id=%s ORDER BY id DESC LIMIT 1", (user.id,), fetchone=True)
        w_id = last_w["id"] if last_w else 1
        context.user_data.clear()

        confirm_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বন্ধ করুন" if lang == 'bn' else "❌ Close", callback_data="dismiss_msg")]])
        success_wd = (
            f"✅ Withdrawal request successful!\nMethod: {method}\nAccount: `{text}`\nYou will get: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\n⚡ You can check status in '📜 History' > '💸 Withdrawal History'."
            if lang == 'en' else
            f"✅ উইথড্র রিকোয়েস্ট সফল হয়েছে!\nমেথড: {method}\nঅ্যাকাউন্ট: `{text}`\nপাবেন: ৳{rec:.2f} BDT (${rec_usdt:.2f} USDT)\n\n⚡ আপনার উইথড্র স্ট্যাটাস '📜 হিস্ট্রি' > '💸 উইথড্র হিস্ট্রি'-তে দেখতে পাবেন।"
        )
        await update.message.reply_text(
            success_wd,
            reply_markup=confirm_kb,
            parse_mode="Markdown"
        )

        owner_id = get_owner_id()
        if owner_id:
            action_buttons = [
                [
                    InlineKeyboardButton("✅ Approve", callback_data=f"approve_withdraw:{w_id}"),
                    InlineKeyboardButton("❌ Reject", callback_data=f"reject_withdraw:{w_id}")
                ]
            ]
            keyboard = InlineKeyboardMarkup(action_buttons)
            await safe_send_message(
                context.bot,
                chat_id=int(owner_id),
                text=(
                    f"🔔 **নতুন উইথড্র রিকোয়েস্ট (Owner Only)** (#{w_id})\n\n"
                    f"🆔 ইউজার: `{user.id}`\n"
                    f"💰 তোলার পরিমাণ: ৳{amt:.2f} BDT\n"
                    f"💸 পেমেন্ট দিতে হবে: ৳{rec:.2f} (${rec_usdt:.2f})\n"
                    f"💳 মেথড: **{method}**\n"
                    f"📌 অ্যাকাউন্ট/নম্বর: `{text}`\n\n"
                    f"💡 নম্বরের ওপর আঙুল ছোঁয়ালেই (ট্যাপ করলেই) কপি হয়ে যাবে। টাকা পাঠিয়ে অ্যাপ্রুভ বাটনে চাপ দিন।"
                ),
                reply_markup=keyboard
            )
        return

    # সাধারণ ইউজারদের ফাইল সাবমিশন
    if state == "waiting_file":
        doc = update.message.document
        if not doc:
            err_doc = "❌ Please send an Excel or CSV file (.xlsx, .csv)." if lang == 'en' else "❌ এক্সেল বা সিএসভি ফাইল (.xlsx, .csv) পাঠান।"
            await update.message.reply_text(err_doc)
            return

        fn = doc.file_name.lower()
        if not (fn.endswith(".xlsx") or fn.endswith(".xls") or fn.endswith(".csv")):
            err_ext = "❌ Only Excel or CSV files are supported." if lang == 'en' else "❌ শুধুমাত্র Excel বা CSV ফাইল সাপোর্টেড।"
            await update.message.reply_text(err_ext)
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
            err_read = f"❌ Error reading file: {e}" if lang == 'en' else f"❌ ফাইল পড়তে সমস্যা হয়েছে: {e}"
            await update.message.reply_text(err_read)
            return

        if os.path.exists(local_path):
            os.remove(local_path)

        if not valid_emails:
            err_no = "❌ No valid @gmail.com found in the file!" if lang == 'en' else "❌ ফাইলে কোনো ভ্যালিড @gmail.com পাওয়া যায়নি!"
            await update.message.reply_text(err_no)
            return

        m_lim = get_config("max_emails_per_file", int)
        if len(valid_emails) > m_lim:
            err_lim = (
                f"❌ File has {len(valid_emails)} Gmails.\nMax allowed per file is {m_lim}."
                if lang == 'en' else
                f"❌ ফাইলে {len(valid_emails)}টি জিমেইল রয়েছে。\nএকটি ফাইলে সর্বোচ্চ {m_lim}টি জিমেইল দেওয়া অনুমোদিত।"
            )
            await update.message.reply_text(err_lim)
            return

        duplicates = check_recent_duplicate_emails(valid_emails)
        if duplicates:
            cd_days = get_config("duplicate_check_days", int)
            dup_list = "\n".join([f"• {d}" for d in duplicates])
            err_dup = (
                f"❌ **File not accepted!**\n\nThe following Gmails were submitted within the last {cd_days} days:\n{dup_list}"
                if lang == 'en' else
                f"❌ **ফাইল গ্রহণ করা হয়নি!**\n\nনিচের জিমেইলগুলো বিগত {cd_days} দিনের মধ্যে জমা দেওয়া হয়েছিল:\n{dup_list}"
            )
            await update.message.reply_text(err_dup)
            return

        save_submitted_emails(valid_emails)

        emails_json = json.dumps(valid_emails)
        db_execute("""
            INSERT INTO submissions (user_id, file_id, emails_json, status, created_at)
            VALUES (%s, %s, %s, 'pending', %s)
        """, (user.id, doc.file_id, emails_json, datetime.now().isoformat()))

        last_sub = db_execute("SELECT id FROM submissions WHERE user_id=%s ORDER BY id DESC LIMIT 1", (user.id,), fetchone=True)
        sub_id = last_sub["id"] if last_sub else 1
        increase_file_count(user.id)
        context.user_data.clear()

        success_sub = (
            f"✅ File submitted successfully!\n\n🔍 Total valid Gmails: {len(valid_emails)}.\n⏳ Admin will observe Gmails for 24-48 hours after initial screening."
            if lang == 'en' else
            f"✅ ফাইল সফলভাবে জমা হয়েছে!\n\n🔍 মোট ভ্যালিড Gmail: {len(valid_emails)} টি。\n⏳ অ্যাডমিন প্রাথমিক বাছাইয়ের পর জিমেইলগুলো ২৪ থেকে ৪৮ ঘণ্টা পর্যবেক্ষণ করবেন।"
        )
        await update.message.reply_text(success_sub)

        await dispatch_submission_to_admins(context, sub_id, user.id, user.username, doc.file_id, valid_emails)
        return

    await update.message.reply_text("নিচের বাটনগুলো ব্যবহার করুন:", reply_markup=get_bottom_keyboard(user.id, lang))


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
        cancel_kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ বাতিল করুন", callback_data="dismiss_msg")]])
        await query.edit_message_text(f"💳 নির্বাচিত: {method}\n\n👉 {hint}", reply_markup=cancel_kb)


async def refinfo_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_chat or update.effective_chat.type != "private":
        return
    if not is_admin(update.effective_user.id):
        return

    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text("ব্যবহার নিয়ম: `/refinfo <user_id>`")
        return

    target_id = int(context.args[0])
    u = get_user(target_id)
    if not u:
        await update.message.reply_text("❌ এই ইউজার ডাটাবেজে পাওয়া যায়নি।")
        return

    refs = get_referral_details(target_id)
    total_c = get_referral_count(target_id)

    res_text = f"📊 **রেফারেল হিস্ট্রি:** `{target_id}`\n"
    res_text += f"মোট ভেরিফায়েড রেফার: **{total_c}** (অর্গানিক ভেরিফায়েড: {len([r for r in refs if r['ref_bonus_rewarded']==1]) if refs else 0}, ম্যানুয়াল: {u.get('manual_ref_count', 0)})\n\n"

    if not refs:
        res_text += "ℹ️ এই ইউজারের লিংকে এখনো কেউ জয়েন করেনি।"
    else:
        res_text += "📋 **রেফার করা ইউজারদের তালিকা:**\n"
        for r in refs[:30]:
            un = f"@{r['username']}" if r['username'] else "No username"
            st = "🟢 ভেরিফায়েড (কাজ সম্পন্ন)" if r.get('ref_bonus_rewarded') == 1 else "🟡 পেন্ডিং (এখনো ফাইল সম্পন্ন করেনি)"
            dt = r['created_at'][:10] if r['created_at'] else "N/A"
            res_text += f"• `{r['user_id']}` ({un}) - {st} ({dt})\n"

    await update.message.reply_text(res_text, parse_mode="Markdown")


# =========================================================
# MAIN (WEBHOOK MODE WITH HEALTH CHECK ENDPOINT)
# =========================================================

def main():
    init_db()

    application = Application.builder().token(BOT_TOKEN).concurrent_updates(True).build()

    application.add_handler(CommandHandler("start", start, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("admin", admin, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("setadmin", set_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("addadmin", add_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("removeadmin", remove_admin_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("adminlist", admin_list_cmd, filters=filters.ChatType.PRIVATE))
    application.add_handler(CommandHandler("refinfo", refinfo_cmd, filters=filters.ChatType.PRIVATE))

    application.add_handler(CallbackQueryHandler(withdrawal_method_handler, pattern=r"^method_"))
    application.add_handler(CallbackQueryHandler(button_handler))
    
    application.add_handler(MessageHandler(filters.ChatType.PRIVATE & ~filters.COMMAND, message_handler))

    if RENDER_EXTERNAL_URL:
        clean_url = RENDER_EXTERNAL_URL.rstrip("/")
        webhook_target = f"{clean_url}/{BOT_TOKEN}"
        print(f"Starting Superfast Webhook on port {PORT}: {webhook_target}")

        application.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            url_path=BOT_TOKEN,
            webhook_url=webhook_target,
            allowed_updates=Update.ALL_TYPES,
        )
    else:
        print("Starting Long Polling fallback...")
        application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
