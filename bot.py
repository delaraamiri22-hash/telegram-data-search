import os,re,csv,sqlite3,logging,asyncio
from pathlib import Path
from telegram import Update,ReplyKeyboardMarkup,KeyboardButton
from telegram.ext import Application,CommandHandler,MessageHandler,ContextTypes,filters
from openpyxl import load_workbook

BASE=Path(__file__).resolve().parent
DATA=BASE/"data"; DATA.mkdir(exist_ok=True)
DB=Path(os.getenv("DB_PATH",str(BASE/"search.sqlite3")))
TOKEN=os.getenv("BOT_TOKEN","").strip()
ADMINS={int(x) for x in os.getenv("ADMIN_IDS","").split(",") if x.strip()}
MAX_MB=int(os.getenv("MAX_FILE_MB","20"))
MAX_ROWS=int(os.getenv("MAX_ROWS_PER_FILE","200000"))
logging.basicConfig(level=logging.INFO)
log=logging.getLogger("bot")
SENSITIVE=("password","passwd","secret","cvv","cvc","card_number","credit_card","bank_account","iban","ssn","national_id","رمز عبور","شماره کارت","شماره حساب","شماره شبا","کد ملی","کد امنیتی")
def clean(x): return re.sub(r"\\s+"," ",str(x or "").replace("\\x00"," ").strip())
def blocked(x):
 s=clean(x).lower()
 return any(w in s or w in s.replace(" ","_") for w in SENSITIVE)
def conn():
 c=sqlite3.connect(DB)
 c.execute("CREATE TABLE IF NOT EXISTS rec(id INTEGER PRIMARY KEY,src TEXT,rowno INTEGER,body TEXT)")
 c.execute("CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(body,src UNINDEXED,rowno UNINDEXED,tokenize='unicode61')")
 c.commit(); return c
def put(c,src,r,b):
 if clean(b):
  cur=c.execute("INSERT INTO rec(src,rowno,body) VALUES(?,?,?)",(src,r,clean(b)))
  c.execute("INSERT INTO fts(rowid,body,src,rowno) VALUES(?,?,?,?)",(cur.lastrowid,clean(b),src,r))
def index_file(p):
 c=conn()
 try:
  c.execute("DELETE FROM fts WHERE src LIKE ?",(p.name+"%",)); c.execute("DELETE FROM rec WHERE src LIKE ?",(p.name+"%",))
  n=0; ext=p.suffix.lower()
  if ext==".txt":
   with p.open(encoding="utf-8-sig",errors="replace") as f:
    for r,line in enumerate(f,1):
     if n>=MAX_ROWS: break
     if clean(line): put(c,p.name,r,line); n+=1
  elif ext==".csv":
   with p.open(encoding="utf-8-sig",errors="replace",newline="") as f:
    sample=f.read(8192); f.seek(0)
    try: d=csv.Sniffer().sniff(sample,delimiters=",;\\t|")
    except Exception: d=csv.excel
    rd=csv.reader(f,d); heads=next(rd,[]); ix=[i for i,h in enumerate(heads) if not blocked(h)]
    if not ix: raise ValueError("ستون قابل جست‌وجوی مجاز پیدا نشد")
    for r,row in enumerate(rd,2):
     if n>=MAX_ROWS: break
     vals=[f"{clean(heads[i])}: {clean(row[i])}" for i in ix if i<len(row) and clean(row[i])]
     if vals: put(c,p.name,r," | ".join(vals)); n+=1
  elif ext==".xlsx":
   wb=load_workbook(p,read_only=True,data_only=True)
   try:
    for ws in wb.worksheets:
     it=ws.iter_rows(values_only=True); heads=next(it,None)
     if not heads: continue
     ix=[i for i,h in enumerate(heads) if not blocked(h)]
     for r,row in enumerate(it,2):
      if n>=MAX_ROWS: break
      vals=[f"{clean(heads[i])}: {clean(row[i])}" for i in ix if i<len(row) and clean(row[i])]
      if vals: put(c,f"{p.name}/{ws.title}",r," | ".join(vals)); n+=1
   finally: wb.close()
  else: raise ValueError("فقط CSV، TXT و XLSX پشتیبانی می‌شوند")
  c.commit(); return n
 except: c.rollback(); raise
 finally: c.close()
def search(q):
 toks=[x for x in re.findall(r"\\w+",q) if len(x)>1]
 if not toks:return []
 expr=" AND ".join('"'+x.replace('"','""')+'"' for x in toks)
 c=conn()
 try:return c.execute("SELECT src,rowno,snippet(fts,0,'【','】',' … ',18) FROM fts WHERE fts MATCH ? LIMIT 8",(expr,)).fetchall()
 finally:c.close()
def kb():return ReplyKeyboardMarkup([[KeyboardButton("🔎 جست‌وجو"),KeyboardButton("📁 راهنما")],[KeyboardButton("📊 آمار")]],resize_keyboard=True)
def admin(u):return bool(ADMINS) and u in ADMINS
async def start(u,c):
 if not admin(u.effective_user.id): await u.message.reply_text("دسترسی فعال نیست."); return
 await u.message.reply_text("سلام! فایل CSV، TXT یا XLSX مجاز خودت را بفرست؛ پس از فهرست‌سازی جست‌وجو کن.",reply_markup=kb())
async def helpcmd(u,c):
 await u.message.reply_text("فایل CSV/TXT/XLSX را بفرست، منتظر پایان فهرست‌سازی بمان و سپس عبارت جست‌وجو را ارسال کن. داده‌های حساس یا خصوصی اشخاص را بارگذاری نکن.")
async def stats(u,c):
 if not admin(u.effective_user.id): return
 db=conn()
 try:n=db.execute("SELECT count(*) FROM rec").fetchone()[0]; f=db.execute("SELECT count(DISTINCT src) FROM rec").fetchone()[0]
 finally:db.close()
 await u.message.reply_text(f"ردیف‌ها: {n:,}\\nمنابع: {f:,}")
async def doc(u,c):
 if not admin(u.effective_user.id): await u.message.reply_text("دسترسی فعال نیست."); return
 d=u.message.document; name=Path(d.file_name or "upload").name
 if Path(name).suffix.lower() not in (".csv",".txt",".xlsx"): await u.message.reply_text("فقط CSV، TXT و XLSX."); return
 if d.file_size and d.file_size>MAX_MB*1024*1024: await u.message.reply_text(f"حداکثر حجم فعلی {MAX_MB} مگابایت است."); return
 p=DATA/re.sub(r"[^A-Za-z0-9._\\-\\u0600-\\u06FF]","_",name)
 await u.message.reply_text("فایل دریافت شد؛ فهرست‌سازی شروع شد.")
 try:
  tf=await c.bot.get_file(d.file_id); await tf.download_to_drive(custom_path=str(p)); n=index_file(p)
  await u.message.reply_text(f"تمام شد؛ {n:,} ردیف فهرست شد. عبارت جست‌وجو را بفرست.")
 except Exception as e: log.exception("index"); await u.message.reply_text("خطا در پردازش: "+str(e)[:200])
async def txt(u,c):
 if not admin(u.effective_user.id): return
 q=clean(u.message.text)
 if q=="📁 راهنما":return await helpcmd(u,c)
 if q=="📊 آمار":return await stats(u,c)
 if q=="🔎 جست‌وجو":return await u.message.reply_text("عبارت جست‌وجو را بفرست.")
 rows=search(q)
 if not rows:return await u.message.reply_text("نتیجه‌ای پیدا نشد.")
 await u.message.reply_text("🔎 نتایج:\\n\\n"+"\\n\\n".join(f"📄 {s} | ردیف {r}\\n{v}" for s,r,v in rows))
async def main():
 if not TOKEN or not ADMINS: raise RuntimeError("BOT_TOKEN و ADMIN_IDS را در Environment تنظیم کن")
 port=int(os.getenv("PORT","10000")); app=Application.builder().token(TOKEN).build()
 app.add_handler(CommandHandler("start",start)); app.add_handler(CommandHandler("help",helpcmd)); app.add_handler(CommandHandler("stats",stats))
 app.add_handler(MessageHandler(filters.Document.ALL,doc)); app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,txt))
 await app.initialize(); await app.start()
 await app.updater.start_webhook(listen="0.0.0.0",port=port,url_path=TOKEN,webhook_url=f"https://{os.environ['RENDER_EXTERNAL_HOSTNAME']}/{TOKEN}",allowed_updates=Update.ALL_TYPES)
 log.info("Webhook started")
 await asyncio.Event().wait()
if __name__=="__main__":asyncio.run(main())
