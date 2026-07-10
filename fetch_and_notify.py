import os, time, feedparser, requests, sqlite3, hashlib, logging
from datetime import datetime, timedelta, timezone
from dateutil import parser

# ===== TIMEZONE =====
IST = timezone(timedelta(hours=5, minutes=30))

FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
DB_FILE = "announcements.db"
LOG_FILE = "run_log.txt"
EXCEL_FILE = "announcements.xlsx"
LOCK_FILE = ".lock"

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
RUN_SOURCE = os.getenv("RUN_SOURCE", "cronjobs.org")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0 Safari/537.36",
    "Referer": "https://www.nseindia.com/",
    "Accept": "application/xml,text/xml,*/*",
}

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler('bot.log'), logging.StreamHandler()])
logger = logging.getLogger(__name__)

# ===== LOCK =====
class FileLock:
    def __init__(self, f=LOCK_FILE): self.f = f
    def acquire(self):
        if os.path.exists(self.f):
            if time.time() - os.path.getmtime(self.f) > 300:
                os.remove(self.f)
                return self.acquire()
            return False
        with open(self.f, 'w') as fp: fp.write(str(os.getpid()))
        return True
    def release(self):
        if os.path.exists(self.f): os.remove(self.f)

# ===== DATABASE =====
class DB:
    def __init__(self, f=DB_FILE):
        self.f = f
        self.init()
    def init(self):
        conn = sqlite3.connect(self.f)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS announcements (
            guid TEXT PRIMARY KEY,
            link TEXT UNIQUE,
            title TEXT, subject TEXT, description TEXT,
            date TEXT, time TEXT, sentiment TEXT,
            pub_date TIMESTAMP, hash TEXT UNIQUE,
            sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        c.execute("PRAGMA table_info(announcements)")
        cols = [r[1] for r in c.fetchall()]
        if 'pub_date' not in cols:
            c.execute('ALTER TABLE announcements ADD COLUMN pub_date TIMESTAMP')
            c.execute('UPDATE announcements SET pub_date = sent_at WHERE pub_date IS NULL')
        c.execute('CREATE INDEX IF NOT EXISTS idx_pub_date ON announcements(pub_date)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_hash ON announcements(hash)')
        conn.commit()
        conn.close()
    def exists(self, guid, link, h):
        conn = sqlite3.connect(self.f)
        c = conn.cursor()
        c.execute('SELECT 1 FROM announcements WHERE guid=? OR link=? OR hash=?', (guid, link, h))
        r = c.fetchone() is not None
        conn.close()
        return r
    def add(self, data):
        conn = sqlite3.connect(self.f)
        c = conn.cursor()
        try:
            c.execute('''INSERT OR IGNORE INTO announcements
                (guid, link, title, subject, description, date, time, sentiment, pub_date, hash)
                VALUES (?,?,?,?,?,?,?,?,?,?)''',
                (data['guid'], data['link'], data['title'], data['subject'],
                 data['description'], data['date'], data['time'], data['sentiment'],
                 data['pub_date'], data['hash']))
            conn.commit()
            return c.rowcount > 0
        except: return False
        finally: conn.close()
    def cleanup(self, days=30):
        conn = sqlite3.connect(self.f)
        c = conn.cursor()
        c.execute('DELETE FROM announcements WHERE sent_at < datetime("now", ?)', (f'-{days} days',))
        d = c.rowcount
        conn.commit()
        conn.close()
        if d: logger.info(f"🧹 Cleaned {d} old")
        return d
    def get_total_count(self):
        conn = sqlite3.connect(self.f)
        c = conn.cursor()
        c.execute('SELECT COUNT(*) FROM announcements')
        return c.fetchone()[0]

# ===== SENTIMENT (now checks title, subject, description) =====
def sentiment_emoji(title, subject, description):
    text = f"{title} {subject} {description}".lower()
    
    BAD = [
        "resignation", "cessation", "insolvency", "litigation", "dispute",
        "default", "delay", "penalt", "fine", "action initiated",
        "action taken", "orders passed", "takeover", "corporate insolvency",
        "winding up", "reduction in capital", "downgrade"
    ]
    GOOD = [
        "dividend", "bonus", "buyback", "acquisition", "awarding of order",
        "bagging", "receiving of order", "credit rating- new", "capacity addition",
        "commencement of commercial production", "investor presentation",
        "allotment of securities", "amalgamation", "merger", "upgrade",
        "record date", "scheme of arrangement"
    ]
    WARNING = [
        "caution", "warning", "update", "clarification", "announcement",
        "postponed", "adjourned", "suspended", "cancelled"
    ]
    
    if any(k in text for k in BAD):
        return "🔴"
    if any(k in text for k in WARNING):
        return "⚠️"
    if any(k in text for k in GOOD):
        return "🟢"
    return "🟡"

# ===== HELPERS =====
def skip(title, subject, link):
    if not link: return True
    skip_words = ["declaration of nav", "net asset value", "mutual fund", "etf"]
    return any(k in f"{title} {subject}".lower() for k in skip_words)

def truncate(t, n=220):
    t = " ".join(t.split())
    return t if len(t)<=n else t[:n].rsplit(' ',1)[0]+' ...'

def split_subject(d):
    d = " ".join(d.split())
    if "|SUBJECT:" in d:
        a,b = d.split("|SUBJECT:",1)
        return a.strip(), b.strip()
    return d.strip(), ""

def split_datetime(p):
    p = p.strip()
    if ' ' in p:
        d,t = p.split(' ',1)
        return d,t
    return p, ""

def escape(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def hash_content(title, subject, desc):
    return hashlib.md5(f"{title}|{subject}|{desc}".encode()).hexdigest()

def parse_pub_ist(pub_str):
    try:
        dt = parser.parse(pub_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=IST)
        return dt.astimezone(IST).replace(tzinfo=None)
    except:
        return None

def send_telegram(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False
    try:
        r = requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
                          data={"chat_id": TELEGRAM_CHAT_ID, "text": text,
                                "parse_mode": "HTML", "disable_web_page_preview": False}, timeout=15)
        return r.ok
    except Exception as e:
        logger.error(f"Telegram error: {e}")
        return False

def fetch_feed():
    for a in range(1,4):
        try:
            r = requests.get(FEED_URL, headers=HEADERS, timeout=30)
            r.raise_for_status()
            return feedparser.parse(r.content)
        except Exception as e:
            logger.warning(f"Attempt {a} failed: {e}")
            time.sleep(5*a)
    return None

def write_excel(db):
    try:
        from openpyxl import Workbook
    except:
        return 0
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Today")
    ws.append(["Sentiment","Company","Subject","Description","Date","Time","Link","Sent At"])
    conn = sqlite3.connect(db.f)
    c = conn.cursor()
    c.execute('''SELECT sentiment,title,subject,description,date,time,link,sent_at
                 FROM announcements WHERE DATE(sent_at)=DATE('now','localtime')
                 ORDER BY sent_at DESC LIMIT 500''')
    cnt=0
    for row in c:
        ws.append([str(x) if x else "" for x in row])
        cnt+=1
    conn.close()
    wb.save(EXCEL_FILE)
    logger.info(f"📊 Excel updated with {cnt} today's entries")
    return cnt

# ===== MAIN =====
def main():
    logger.info(f"🚀 Starting at {datetime.now()}")
    lock = FileLock()
    if not lock.acquire():
        logger.warning("Another instance running, exit")
        return
    try:
        db = DB()
        db.cleanup(30)

        feed = fetch_feed()
        if not feed:
            logger.error("No feed")
            return

        now_ist = datetime.now(IST).replace(tzinfo=None)
        six_ago_ist = now_ist - timedelta(minutes=6)
        logger.info(f"⏰ Window (IST): {six_ago_ist.strftime('%H:%M:%S')} – {now_ist.strftime('%H:%M:%S')}")

        new_items = []
        dups = 0
        skips = 0

        for entry in feed.entries:
            pub = getattr(entry, "published", "")
            if not pub:
                continue

            entry_time = parse_pub_ist(pub)
            if entry_time is None:
                continue

            if not (six_ago_ist <= entry_time <= now_ist):
                continue

            guid = entry.get("id", entry.get("link", ""))
            link = entry.get("link", "")
            title = getattr(entry, "title", "NSE Announcement")
            raw_desc = getattr(entry, "summary", "")
            if not link:
                skips += 1
                continue

            desc, subject = split_subject(raw_desc)
            date_str, time_str = split_datetime(pub)

            if skip(title, subject, link):
                skips += 1
                continue

            h = hash_content(title, subject, desc)
            if db.exists(guid, link, h):
                dups += 1
                continue

            pub_iso = entry_time.isoformat()
            data = {
                'guid': guid, 'link': link, 'title': title.strip(),
                'subject': subject, 'description': desc,
                'date': date_str, 'time': time_str,
                'sentiment': sentiment_emoji(title, subject, desc),   # ← NOW WITH 3 ARGS
                'pub_date': pub_iso,
                'hash': h,
                'pub': pub
            }

            if db.add(data):
                new_items.append(data)

        # Oldest first
        new_items.sort(key=lambda x: x['pub_date'])

        if new_items:
            logger.info(f"📤 Sending {len(new_items)} alerts (OLDEST FIRST)...")
            for i, item in enumerate(new_items[:3]):
                logger.info(f"  #{i+1} pub_date: {item['pub_date']} – {item['title'][:40]}")

            # Send in batches of 10
            for i in range(0, len(new_items), 10):
                batch = new_items[i:i+10]
                for data in batch:
                    # Message with document link
                    msg = (
                        f"{data['sentiment']} <a href=\"{data['link']}\">{escape(data['title'])}</a>\n"
                        f"📄 {escape(data['subject'])}\n"
                        f"{escape(truncate(data['description'], 220))}\n"
                        f"📎 <a href=\"{data['link']}\">View Document</a>\n"
                        f"🕐 {data['pub']}"
                    )
                    send_telegram(msg)
                    time.sleep(0.3)
                time.sleep(1)

        if new_items:
            write_excel(db)

        total = db.get_total_count()
        logger.info(f"📊 New={len(new_items)}, Dups={dups}, Skipped={skips}, DB total={total}")
        with open(LOG_FILE, "a") as f:
            f.write(f"{datetime.now()}: New={len(new_items)}, Dups={dups}, Skipped={skips}\n")

    except Exception as e:
        logger.error(f"Fatal: {e}", exc_info=True)
    finally:
        lock.release()
        logger.info("✅ Done")

if __name__ == "__main__":
    main()