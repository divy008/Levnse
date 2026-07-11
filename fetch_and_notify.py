import os
import time
import json
import feedparser
import requests
import sqlite3
import hashlib
import logging
from datetime import datetime, timedelta, timezone
from dateutil import parser

# ===== ENVIRONMENT VARIABLES =====
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
GOOGLE_SHEETS_WEB_APP_URL = os.getenv("GOOGLE_SHEETS_WEB_APP_URL")  # Optional
RUN_SOURCE = os.getenv("RUN_SOURCE", "cronjobs.org")

if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
    print("⚠️ Warning: Telegram credentials missing. Alerts will not be sent.")

# ===== CONFIGURATION =====
FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
DB_FILE = "announcements.db"
LOG_FILE = "run_log.txt"
EXCEL_FILE = "announcements.xlsx"
LOCK_FILE = ".lock"

IST = timezone(timedelta(hours=5, minutes=30))   # UTC+5:30

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0 Safari/537.36",
    "Referer": "https://www.nseindia.com/",
    "Accept": "application/xml,text/xml,*/*",
}

# ===== LOGGING =====
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('bot.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ===== FILE LOCK =====
class FileLock:
    def __init__(self, lock_file=LOCK_FILE):
        self.lock_file = lock_file

    def acquire(self):
        if os.path.exists(self.lock_file):
            # Stale lock older than 5 minutes
            if time.time() - os.path.getmtime(self.lock_file) > 300:
                os.remove(self.lock_file)
                return self.acquire()
            return False
        with open(self.lock_file, 'w') as f:
            f.write(str(os.getpid()))
        return True

    def release(self):
        if os.path.exists(self.lock_file):
            os.remove(self.lock_file)

# ===== DATABASE =====
class AnnouncementDB:
    def __init__(self, db_file=DB_FILE):
        self.db_file = db_file
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS announcements (
                guid TEXT PRIMARY KEY,
                link TEXT UNIQUE,
                title TEXT,
                subject TEXT,
                description TEXT,
                date TEXT,
                time TEXT,
                sentiment TEXT,
                pub_date TIMESTAMP,
                hash TEXT UNIQUE,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        # Add pub_date if missing (migration)
        c.execute("PRAGMA table_info(announcements)")
        columns = [row[1] for row in c.fetchall()]
        if 'pub_date' not in columns:
            c.execute('ALTER TABLE announcements ADD COLUMN pub_date TIMESTAMP')
            c.execute('UPDATE announcements SET pub_date = sent_at WHERE pub_date IS NULL')
        c.execute('CREATE INDEX IF NOT EXISTS idx_pub_date ON announcements(pub_date)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_link ON announcements(link)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_hash ON announcements(hash)')
        conn.commit()
        conn.close()
        logger.info("✅ Database initialized")

    def is_duplicate(self, guid, link, content_hash):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT 1 FROM announcements WHERE guid=? OR link=? OR hash=?', (guid, link, content_hash))
        result = c.fetchone() is not None
        conn.close()
        return result

    def add_announcement(self, data):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        try:
            c.execute('''
                INSERT OR IGNORE INTO announcements
                (guid, link, title, subject, description, date, time, sentiment, pub_date, hash)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['guid'], data['link'], data['title'], data['subject'],
                data['description'], data['date'], data['time'], data['sentiment'],
                data['pub_date'], data['hash']
            ))
            conn.commit()
            return c.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def cleanup_old(self, days=30):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('DELETE FROM announcements WHERE sent_at < datetime("now", ?)', (f'-{days} days',))
        deleted = c.rowcount
        conn.commit()
        conn.close()
        if deleted:
            logger.info(f"🧹 Cleaned up {deleted} old announcements")
        return deleted

    def get_total_count(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT COUNT(*) FROM announcements')
        count = c.fetchone()[0]
        conn.close()
        return count

# ===== SENTIMENT EMOJI (checks title + subject + description) =====
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

# ===== HELPER FUNCTIONS =====
def should_skip(title, subject, link):
    if not link:
        return True
    skip_words = ["declaration of nav", "net asset value", "mutual fund", "etf"]
    text = f"{title} {subject}".lower()
    return any(k in text for k in skip_words)

def truncate(text, max_chars=220):
    text = " ".join(text.split())
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(' ', 1)[0] + ' ...'

def split_subject(description):
    description = " ".join(description.split())
    if "|SUBJECT:" in description:
        desc_part, subject_part = description.split("|SUBJECT:", 1)
        return desc_part.strip(), subject_part.strip()
    return description.strip(), ""

def split_datetime(pub):
    pub = pub.strip()
    if ' ' in pub:
        date_part, time_part = pub.split(' ', 1)
        return date_part, time_part
    return pub, ""

def escape_html(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def create_hash(title, subject, description):
    return hashlib.md5(f"{title}|{subject}|{description}".encode()).hexdigest()

def parse_pub_ist(pub_str):
    """Parse RSS pub date, assume IST if naive, return naive IST datetime."""
    try:
        dt = parser.parse(pub_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=IST)
        return dt.astimezone(IST).replace(tzinfo=None)  # naive IST
    except Exception as e:
        logger.warning(f"Could not parse pub date: {pub_str} – {e}")
        return None

# ===== SEND TELEGRAM =====
def send_telegram(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        r = requests.post(url, data={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": text,
            "parse_mode": "HTML",
            "disable_web_page_preview": False
        }, timeout=15)
        return r.ok
    except Exception as e:
        logger.error(f"❌ Telegram error: {e}")
        return False

# ===== FETCH RSS FEED =====
def fetch_feed():
    for attempt in range(1, 4):
        try:
            resp = requests.get(FEED_URL, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            return feedparser.parse(resp.content)
        except Exception as e:
            logger.warning(f"⚠️ Attempt {attempt} failed: {e}")
            time.sleep(5 * attempt)
    return None

# ===== SEND TO GOOGLE SHEETS (via Web App POST) =====
def send_to_google_sheets(new_items, web_app_url):
    """Send new announcements directly to Google Sheets Master Sheet."""
    if not new_items or not web_app_url:
        return False

    data_to_send = []
    for item in new_items:
        data_to_send.append({
            "company": item['title'],
            "subject": item['subject'],
            "description": item['description'],
            "date": item['date'],
            "time": item['time'],
            "link": item['link'],
            "sentiment": item['sentiment']
        })

    payload = {"action": "add_records", "data": data_to_send}

    try:
        response = requests.post(web_app_url, json=payload, timeout=30)
        if response.status_code == 200:
            logger.info(f"✅ Sent {len(data_to_send)} records to Google Sheets")
            return True
        else:
            logger.error(f"❌ Google Sheets POST failed: {response.text}")
            return False
    except Exception as e:
        logger.error(f"❌ Error sending to Google Sheets: {e}")
        return False

# ===== WRITE EXCEL (optional) =====
def write_excel(db):
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("⚠️ openpyxl not installed. Excel export skipped.")
        return 0

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Today's Announcements")
    ws.append(["Sentiment", "Company", "Subject", "Description", "Date", "Time", "Link", "Sent At"])

    conn = sqlite3.connect(db.db_file)
    c = conn.cursor()
    c.execute('''
        SELECT sentiment, title, subject, description, date, time, link, sent_at
        FROM announcements
        WHERE DATE(sent_at) = DATE('now', 'localtime')
        ORDER BY sent_at DESC
        LIMIT 500
    ''')
    count = 0
    for row in c:
        ws.append([str(x) if x else "" for x in row])
        count += 1
    conn.close()
    wb.save(EXCEL_FILE)
    logger.info(f"📊 Excel updated with {count} today's announcements")
    return count

# ===== MAIN =====
def main():
    logger.info(f"🚀 NSE Watch starting at {datetime.now()}")
    logger.info(f"📡 Source: {RUN_SOURCE}")

    lock = FileLock()
    if not lock.acquire():
        logger.warning("⚠️ Another instance is running. Exiting.")
        return

    try:
        db = AnnouncementDB()
        db.cleanup_old(30)

        feed = fetch_feed()
        if feed is None:
            logger.error("❌ Failed to fetch feed")
            return

        # Current time in IST (naive)
        now_ist = datetime.now(IST).replace(tzinfo=None)
        six_ago_ist = now_ist - timedelta(minutes=6)
        logger.info(f"⏰ Window (IST): {six_ago_ist.strftime('%H:%M:%S')} – {now_ist.strftime('%H:%M:%S')}")

        new_items = []
        duplicates = 0
        skipped = 0

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
                skipped += 1
                continue

            desc, subject = split_subject(raw_desc)
            date_str, time_str = split_datetime(pub)

            if should_skip(title, subject, link):
                skipped += 1
                continue

            content_hash = create_hash(title, subject, desc)
            if db.is_duplicate(guid, link, content_hash):
                duplicates += 1
                continue

            pub_iso = entry_time.isoformat()
            data = {
                'guid': guid,
                'link': link,
                'title': title.strip(),
                'subject': subject,
                'description': desc,
                'date': date_str,
                'time': time_str,
                'sentiment': sentiment_emoji(title, subject, desc),
                'pub_date': pub_iso,
                'hash': content_hash,
                'pub': pub   # original published string for Telegram
            }

            if db.add_announcement(data):
                new_items.append(data)

        # Sort oldest first (ascending pub_date)
        new_items.sort(key=lambda x: x['pub_date'])

        # Send Telegram alerts (oldest first)
        if new_items:
            logger.info(f"📤 Sending {len(new_items)} alerts (OLDEST FIRST)...")
            for i, item in enumerate(new_items[:3]):
                logger.info(f"  #{i+1} pub_date: {item['pub_date']} – {item['title'][:40]}")

            # Batch in chunks of 10
            batch_size = 10
            for i in range(0, len(new_items), batch_size):
                batch = new_items[i:i+batch_size]
                for data in batch:
                    title_html = escape_html(data['title'])
                    desc_html = escape_html(truncate(data['description'], 220))
                    subject_html = escape_html(data['subject'])

                    msg = (
                        f"{data['sentiment']} <a href=\"{data['link']}\">{title_html}</a>\n"
                        f"📄 {subject_html}\n"
                        f"{desc_html}\n"
                        f"📎 <a href=\"{data['link']}\">View Document</a>\n"
                        f"🕐 {data['pub']}"
                    )
                    send_telegram(msg)
                    time.sleep(0.3)
                if i + batch_size < len(new_items):
                    time.sleep(1)  # gap between batches

        # Send to Google Sheets (if URL provided)
        if new_items and GOOGLE_SHEETS_WEB_APP_URL:
            send_to_google_sheets(new_items, GOOGLE_SHEETS_WEB_APP_URL)
        elif new_items and not GOOGLE_SHEETS_WEB_APP_URL:
            logger.info("ℹ️ GOOGLE_SHEETS_WEB_APP_URL not set, skipping Google Sheets export.")

        # Excel export
        if new_items:
            write_excel(db)

        # Summary
        total = db.get_total_count()
        summary = f"""
📊 SUMMARY (Last 6 minutes IST):
   📨 New alerts: {len(new_items)}
   🔄 Duplicates skipped: {duplicates}
   ⏭️ Filtered skipped: {skipped}
   📅 Database total: {total}
   ⏰ Window: {six_ago_ist.strftime('%H:%M:%S')} – {now_ist.strftime('%H:%M:%S')}
        """
        logger.info(summary)
        with open(LOG_FILE, "a") as f:
            f.write(f"{datetime.now()}: New={len(new_items)}, Dups={duplicates}, Skipped={skipped}\n")

    except Exception as e:
        logger.error(f"❌ Fatal error: {e}", exc_info=True)
        raise
    finally:
        lock.release()
        logger.info("✅ Done")

if __name__ == "__main__":
    main()