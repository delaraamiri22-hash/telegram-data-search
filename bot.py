#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Telegram file-search bot for datasets the operator owns or is authorized to use.
Supports CSV, TXT, XLSX. SQLite FTS5 is used for indexed text search.
"""
import os
import re
import csv
import json
import sqlite3
import logging
from pathlib import Path
from datetime import datetime

from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes, filters
)

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = BASE / "search_index.sqlite3"
TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_IDS = {int(x.strip()) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "50"))
MAX_ROWS_PER_FILE = int(os.getenv("MAX_ROWS_PER_FILE", "300000"))
RESULTS_PER_PAGE = 8

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger("data-search-bot")

# Avoid indexing common highly sensitive personal/financial identifiers.
BLOCKED_COLUMN_WORDS = (
    "password", "passwd", "secret", "cvv", "cvc", "card_number",
    "credit_card", "bank_account", "iban", "ssn", "national_id",
    "رمزعبور", "رمز عبور", "شماره کارت", "کد امنیتی", "cvv",
    "شماره حساب", "شماره شبا", "کد ملی"
)

def is_admin(user_id: int) -> bool:
    return not ADMIN_IDS or user_id in ADMIN_IDS

def clean_text(value) -> str:
    if value is None:
        return ""
    s = str(value).replace("\x00", " ").strip()
    return re.sub(r"\s+", " ", s)

def blocked_header(header: str) -> bool:
    h = clean_text(header).lower().replace(" ", "_")
    return any(word in h or word in clean_text(header).lower() for word in BLOCKED_COLUMN_WORDS)

def db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            row_no INTEGER NOT NULL,
            body TEXT NOT NULL
        )
    """)
    con.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS records_fts
        USING fts5(body, source UNINDEXED, row_no UNINDEXED, tokenize='unicode61')
    """)
    con.commit()
    return con

def add_record(con, source, row_no, body):
    body = clean_text(body)
    if not body:
        return
    cur = con.execute(
        "INSERT INTO records(source,row_no,body) VALUES(?,?,?)",
        (source, row_no, body)
    )
    con.execute(
        "INSERT INTO records_fts(rowid,body,source,row_no) VALUES(?,?,?,?)",
        (cur.lastrowid, body, source, row_no)
    )

def index_csv(path: Path, con):
    count = 0
    # Try UTF-8 BOM first, then common fallback.
    try:
        f = path.open("r", encoding="utf-8-sig", newline="", errors="replace")
    except Exception:
        f = path.open("r", encoding="cp1256", newline="", errors="replace")
    with f:
        sample = f.read(8192)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except Exception:
            dialect = csv.excel
        reader = csv.reader(f, dialect)
        headers = next(reader, [])
        safe_indices = [i for i, h in enumerate(headers) if not blocked_header(h)]
        if not safe_indices:
            raise ValueError("ستون قابل‌جست‌وجوی مجاز پیدا نشد.")
        for row_no, row in enumerate(reader, start=2):
            if count >= MAX_ROWS_PER_FILE:
                break
            parts = []
            for i in safe_indices:
                if i < len(row):
                    val = clean_text(row[i])
                    if val:
                        parts.append(f"{clean_text(headers[i])}: {val}")
            if parts:
                add_record(con, path.name, row_no, " | ".join(parts))
                count += 1
    return count

def index_txt(path: Path, con):
    count = 0
    with path.open("r", encoding="utf-8-sig", errors="replace") as f:
        for row_no, line in enumerate(f, start=1):
            if count >= MAX_ROWS_PER_FILE:
                break
            line = clean_text(line)
            if line:
                add_record(con, path.name, row_no, line)
                count += 1
    return count

def index_xlsx(path: Path, con):
    if load_workbook is None:
        raise RuntimeError("کتابخانه openpyxl نصب نیست.")
    count = 0
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            iterator = ws.iter_rows(values_only=True)
            headers = next(iterator, None)
            if not headers:
                continue
            safe_indices = [i for i, h in enumerate(headers) if not blocked_header(h)]
            if not safe_indices:
                continue
            for row_no, row in enumerate(iterator, start=2):
                if count >= MAX_ROWS_PER_FILE:
                    return count
                parts = []
                for i in safe_indices:
                    if i < len(row):
                        val = clean_text(row[i])
                        if val:
                            parts.append(f"{clean_text(headers[i])}: {val}")
                if parts:
                    add_record(con, f"{path.name}/{ws.title}", row_no, " | ".join(parts))
                    count += 1
    finally:
        wb.close()
    return count

def index_file(path: Path):
    suffix = path.suffix.lower()
    con = db()
    try:
        # Replace prior index for this exact uploaded filename.
        con.execute("DELETE FROM records_fts WHERE source LIKE ?", (path.name + "%",))
        con.execute("DELETE FROM records WHERE source LIKE ?", (path.name + "%",))
        if suffix == ".csv":
            n = index_csv(path, con)
        elif suffix == ".txt":
            n = index_txt(path, con)
        elif suffix == ".xlsx":
            n = index_xlsx(path, con)
        else:
            raise ValueError("فقط CSV، TXT و XLSX پشتیبانی می‌شوند.")
        con.commit()
        return n
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()

def search_records(query: str, limit=RESULTS_PER_PAGE):
    # Quote each token so punctuation cannot become FTS operators.
    tokens = [t for t in re.findall(r"\w+", query, flags=re.UNICODE) if len(t) > 1]
    if not tokens:
        return []
    match_expr = " AND ".join('"' + t.replace('"', '""') + '"' for t in tokens)
    con = db()
    try:
        rows = con.execute("""
            SELECT source, row_no, snippet(records_fts, 0, '【', '】', ' … ', 18)
            FROM records_fts
            WHERE records_fts MATCH ?
            LIMIT ?
        """, (match_expr, limit)).fetchall()
        return rows
    finally:
        con.close()

def stats():
    con = db()
    try:
        total = con.execute("SELECT count(*) FROM records").fetchone()[0]
        files = con.execute("SELECT count(DISTINCT source) FROM records").fetchone()[0]
        return total, files
    finally:
        con.close()

def keyboard():
    return ReplyKeyboardMarkup(
        [[KeyboardButton("🔎 جست‌وجو"), KeyboardButton("📁 راهنما")],
         [KeyboardButton("📊 آمار")]],
        resize_keyboard=True
    )

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("این ربات خصوصی است و دسترسی شما فعال نیست.")
        return
    await update.message.reply_text(
        "سلام! 👋\n"
        "فایل‌های مجاز CSV، TXT یا XLSX خودت را همین‌جا بفرست.\n"
        "ربات فایل را یک‌بار فهرست‌سازی می‌کند؛ بعد جست‌وجو سریع‌تر انجام می‌شود.\n\n"
        "برای شروع، فایل را ارسال کن یا «🔎 جست‌وجو» را بزن.",
        reply_markup=keyboard()
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "۱) فایل CSV، TXT یا XLSX را ارسال کن.\n"
        "۲) منتظر پیام پایان فهرست‌سازی بمان.\n"
        "۳) عبارت جست‌وجو را بفرست.\n"
        "۴) برای نتیجه دقیق‌تر چند واژه بنویس.\n\n"
        "فقط داده‌هایی را بارگذاری کن که مالکشان هستی یا مجوز استفاده از آن‌ها را داری. "
        "ستون‌های شناخته‌شده مربوط به رمزها و شناسه‌ها و اطلاعات مالی حساس فهرست نمی‌شوند."
    )

async def stats_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    total, files = stats()
    await update.message.reply_text(f"📊 آمار فهرست\nتعداد ردیف‌های فهرست‌شده: {total:,}\nتعداد منابع: {files:,}")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("این ربات خصوصی است.")
        return
    doc = update.message.document
    filename = Path(doc.file_name or "upload").name
    suffix = Path(filename).suffix.lower()
    if suffix not in (".csv", ".txt", ".xlsx"):
        await update.message.reply_text("این نوع فایل پشتیبانی نمی‌شود. فقط CSV، TXT و XLSX بفرست.")
        return
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text(f"حجم فایل بیشتر از سقف فعلی ({MAX_FILE_MB} مگابایت) است.")
        return
    safe_name = re.sub(r"[^A-Za-z0-9._\-\u0600-\u06FF]", "_", filename)
    dest = DATA_DIR / safe_name
    await update.message.reply_text("فایل دریافت شد؛ دارم آن را فهرست‌سازی می‌کنم. برای فایل بزرگ ممکن است کمی زمان ببرد.")
    try:
        tg_file = await context.bot.get_file(doc.file_id)
        await tg_file.download_to_drive(custom_path=str(dest))
        n = index_file(dest)
        await update.message.reply_text(f"✅ تمام شد!\nردیف‌های فهرست‌شده: {n:,}\nحالا عبارت موردنظرت را بفرست.")
    except Exception as e:
        log.exception("Indexing failed")
        await update.message.reply_text(f"❌ پردازش فایل انجام نشد: {str(e)[:300]}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("این ربات خصوصی است.")
        return
    text = clean_text(update.message.text)
    if text == "📁 راهنما":
        await help_cmd(update, context)
        return
    if text == "📊 آمار":
        await stats_cmd(update, context)
        return
    if text == "🔎 جست‌وجو":
        await update.message.reply_text("عبارت موردنظرت را بفرست؛ مثلاً نام یک محصول یا کد یک پرونده غیرحساس.")
        return
    results = search_records(text)
    if not results:
        await update.message.reply_text("نتیجه‌ای پیدا نشد. املای عبارت یا تعداد واژه‌ها را تغییر بده.")
        return
    chunks = []
    for source, row_no, snippet in results:
        chunks.append(f"📄 {source} | ردیف {row_no}\n{snippet}")
    await update.message.reply_text("🔎 نتایج:\n\n" + "\n\n".join(chunks))

async def error_handler(update, context):
    log.exception("Unhandled error", exc_info=context.error)

def main():
    if not TOKEN:
        raise SystemExit("BOT_TOKEN تنظیم نشده است. فایل .env.example را ببین.")
    if not ADMIN_IDS:
        raise SystemExit("ADMIN_IDS تنظیم نشده است؛ شناسه عددی تلگرام مدیر را وارد کن.")
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("stats", stats_cmd))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)
    log.info("Bot is running")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
