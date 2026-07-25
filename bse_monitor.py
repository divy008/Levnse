import os
import time
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
RUN_SOURCE = os.getenv("RUN_SOURCE", "cronjobs.org")

if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
    print("⚠️ Warning: Telegram credentials missing. Alerts will not be sent.")

# ===== CONFIGURATION =====
FEED_URL = "https://beta.bseindia.com/data/xml/announcements.xml"
FALLBACK_FEED_URL = "https://www.bseindia.com/corporates/ann.html"
DB_FILE = "bse_announcements.db"
LOG_FILE = "bse_run_log.txt"
EXCEL_FILE = "bse_announcements.xlsx"
LOCK_FILE = ".bse_lock"

IST = timezone(timedelta(hours=5, minutes=30))  # UTC+5:30

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Referer": "https://www.bseindia.com/",
    "Accept": "application/xml,text/xml,*/*",
}

# ===== LOGGING =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("bse_bot.log"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)


# ===== FILE LOCK =====
class FileLock:
    def __init__(self, lock_file=LOCK_FILE):
        self.lock_file = lock_file

    def acquire(self):
        if os.path.exists(self.lock_file):
            if time.time() - os.path.getmtime(self.lock_file) > 300:
                os.remove(self.lock_file)
                return self.acquire()
            return False
        with open(self.lock_file, "w") as f:
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
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS bse_announcements (
                guid TEXT PRIMARY KEY,
                link TEXT UNIQUE,
                company TEXT,
                scripcode TEXT,
                description TEXT,
                date TEXT,
                time TEXT,
                sentiment TEXT,
                pub_date TIMESTAMP,
                hash TEXT UNIQUE,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_pub_date ON bse_announcements(pub_date)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_link ON bse_announcements(link)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_hash ON bse_announcements(hash)")
        conn.commit()
        conn.close()
        logger.info("✅ BSE Database initialized")

    def is_duplicate(self, guid, link, content_hash):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute(
            "SELECT 1 FROM bse_announcements WHERE guid=? OR link=? OR hash=?",
            (guid, link, content_hash),
        )
        result = c.fetchone() is not None
        conn.close()
        return result

    def add_announcement(self, data):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        try:
            c.execute(
                """
                INSERT OR IGNORE INTO bse_announcements
                (guid, link, company, scripcode, description, date, time, sentiment, pub_date, hash)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    data["guid"],
                    data["link"],
                    data["company"],
                    data["scripcode"],
                    data["description"],
                    data["date"],
                    data["time"],
                    data["sentiment"],
                    data["pub_date"],
                    data["hash"],
                ),
            )
            conn.commit()
            return c.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def cleanup_old(self, days=30):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute(
            "DELETE FROM bse_announcements WHERE sent_at < datetime('now', ?)",
            (f"-{days} days",),
        )
        deleted = c.rowcount
        conn.commit()
        conn.close()
        if deleted:
            logger.info(f"🧹 Cleaned up {deleted} old BSE announcements")
        return deleted

    def get_total_count(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM bse_announcements")
        count = c.fetchone()[0]
        conn.close()
        return count


# ===== SENTIMENT EMOJI =====
def sentiment_emoji(company, description):
    text = f"{company} {description}".lower()

    BAD = [
        "resignation", "cessation", "insolvency", "litigation", "dispute",
        "default", "delay", "penalt", "fine", "action initiated",
        "action taken", "orders passed", "takeover", "corporate insolvency",
        "winding up", "reduction of share capital", "downgrade"
    ]
    GOOD = [
        "dividend", "bonus", "buyback", "acquisition", "awarding of order",
        "bagging", "receiving of order", "credit rating- new", "capacity addition",
        "commencement of commercial production", "investor presentation",
        "allotment of securities", "amalgamation", "merger", "upgrade",
        "record date", "scheme of arrangement", "financial results"
    ]
    WARNING = [
        "caution", "warning", "update", "clarification", "announcement",
        "postponed", "adjourned", "suspended", "cancelled", "trading window"
    ]

    if any(k in text for k in BAD):
        return "🔴"
    if any(k in text for k in WARNING):
        return "⚠️"
    if any(k in text for k in GOOD):
        return "🟢"
    return "🟡"


# ===== HELPER FUNCTIONS =====
def should_skip(company, description, link):
    if not link:
        return True
    skip_words = ["declaration of nav", "net asset value", "mutual fund", "etf"]
    text = f"{company} {description}".lower()
    return any(k in text for k in skip_words)

def truncate(text, max_chars=220):
    text = " ".join(text.split())
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + " ..."

def escape_html(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def create_hash(company, scripcode, description):
    return hashlib.md5(f"{company}|{scripcode}|{description}".encode()).hexdigest()

def parse_pub_ist(pub_str):
    try:
        dt = parser.parse(pub_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=IST)
        return dt.astimezone(IST).replace(tzinfo=None)
    except Exception as e:
        logger.warning(f"Could not parse pub date: {pub_str} – {e}")
        return None


# ===== SEND TELEGRAM =====
def send_telegram(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        r = requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": text,
                "parse_mode": "HTML",
                "disable_web_page_preview": False,
            },
            timeout=15,
        )
        return r.ok
    except Exception as e:
        logger.error(f"❌ Telegram error: {e}")
        return False


# ===== FETCH RSS FEED =====
def fetch_feed():
    urls = [FEED_URL, FALLBACK_FEED_URL]
    for url in urls:
        for attempt in range(1, 3):
            try:
                resp = requests.get(url, headers=HEADERS, timeout=25)
                resp.raise_for_status()
                parsed = feedparser.parse(resp.content)
                if parsed.entries:
                    return parsed
            except Exception as e:
                logger.warning(f"⚠️ Fetch attempt {attempt} failed for {url}: {e}")
                time.sleep(3)
    return None


# ===== EXCEL EXPORT =====
def write_excel(db):
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("⚠️ openpyxl not installed. Excel export skipped.")
        return 0

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Today's BSE Announcements")
    ws.append(["Sentiment", "Company", "ScripCode", "Description", "Date", "Time", "Link", "Sent At"])

    conn = sqlite3.connect(db.db_file)
    c = conn.cursor()
    c.execute(
        """
        SELECT sentiment, company, scripcode, description, date, time, link, sent_at
        FROM bse_announcements
        WHERE DATE(sent_at) = DATE('now', 'localtime')
        ORDER BY sent_at DESC
        LIMIT 500
    """
    )
    count = 0
    for row in c:
        ws.append([str(x) if x else "" for x in row])
        count += 1
    conn.close()
    wb.save(EXCEL_FILE)
    logger.info(f"📊 Excel updated with {count} announcements")
    return count


# ===== MAIN FUNCTION =====
def main():
    logger.info(f"🚀 BSE Watch starting at {datetime.now()}")
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
            logger.error("❌ Failed to fetch BSE feed")
            return

        now_ist = datetime.now(IST).replace(tzinfo=None)
        window_start_ist = now_ist - timedelta(minutes=15)
        logger.info(f"⏰ Window (IST): {window_start_ist.strftime('%H:%M:%S')} – {now_ist.strftime('%H:%M:%S')}")

        new_items = []
        duplicates = 0
        skipped = 0

        for entry in feed.entries:
            pub = getattr(entry, "published", getattr(entry, "pubdate", ""))
            if not pub:
                continue

            entry_time = parse_pub_ist(pub)
            if entry_time is None:
                continue

            if not (window_start_ist <= entry_time <= now_ist):
                continue

            link = entry.get("link", "")
            guid = entry.get("id", link)
            company = getattr(entry, "title", "BSE Company")
            description = getattr(entry, "summary", getattr(entry, "description", ""))
            scripcode = getattr(entry, "scripcode", "")

            if should_skip(company, description, link):
                skipped += 1
                continue

            date_str = entry_time.strftime("%d-%b-%Y")
            time_str = entry_time.strftime("%H:%M:%S")

            content_hash = create_hash(company, scripcode, description)
            if db.is_duplicate(guid, link, content_hash):
                duplicates += 1
                continue

            data = {
                "guid": guid,
                "link": link,
                "company": company.strip(),
                "scripcode": scripcode,
                "description": description.strip(),
                "date": date_str,
                "time": time_str,
                "sentiment": sentiment_emoji(company, description),
                "pub_date": entry_time.isoformat(),
                "hash": content_hash,
                "pub": pub,
            }

            if db.add_announcement(data):
                new_items.append(data)

        new_items.sort(key=lambda x: x["pub_date"])

        if new_items:
            logger.info(f"📤 Sending {len(new_items)} alerts to Telegram...")
            batch_size = 10
            for i in range(0, len(new_items), batch_size):
                batch = new_items[i : i + batch_size]
                for data in batch:
                    company_html = escape_html(data["company"])
                    desc_html = escape_html(truncate(data["description"], 250))
                    scrip_info = f" ({data['scripcode']})" if data['scripcode'] else ""

                    msg = (
                        f"{data['sentiment']} <b>{company_html}</b>{scrip_info}\n"
                        f"{desc_html}\n"
                        f"📎 <a href=\"{data['link']}\">View Document</a>\n"
                        f"🕐 {data['time']} | {data['date']}"
                    )
                    send_telegram(msg)
                    time.sleep(0.3)
                if i + batch_size < len(new_items):
                    time.sleep(1)

            write_excel(db)

        total = db.get_total_count()
        logger.info(
            f"📊 SUMMARY: New={len(new_items)}, Dups={duplicates}, Skipped={skipped}, Total DB={total}"
        )

    except Exception as e:
        logger.error(f"❌ Fatal error: {e}", exc_info=True)
        raise
    finally:
        lock.release()
        logger.info("✅ Done")


if __name__ == "__main__":
    main()
