import os
import time
import feedparser
import requests
from datetime import datetime, timedelta
import sqlite3
import hashlib
import logging
from dateutil import parser

# ===== CONFIGURATION =====
FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
DB_FILE = os.path.join(os.path.dirname(__file__), "announcements.db")
LOG_FILE = os.path.join(os.path.dirname(__file__), "run_log.txt")
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "announcements.xlsx")
LOCK_FILE = os.path.join(os.path.dirname(__file__), ".lock")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
RUN_SOURCE = os.getenv("RUN_SOURCE", "cronjobs.org")

if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
    print("⚠️ Warning: Telegram credentials missing!")

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

# ===== LOCK MECHANISM =====
class FileLock:
    def __init__(self, lock_file=LOCK_FILE):
        self.lock_file = lock_file
    
    def acquire(self):
        if os.path.exists(self.lock_file):
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
        
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='announcements'")
        table_exists = c.fetchone() is not None
        
        if table_exists:
            c.execute("PRAGMA table_info(announcements)")
            columns = [column[1] for column in c.fetchall()]
            
            if 'pub_date' not in columns:
                c.execute('ALTER TABLE announcements ADD COLUMN pub_date TIMESTAMP')
                c.execute('UPDATE announcements SET pub_date = sent_at WHERE pub_date IS NULL')
                logger.info("✅ Added pub_date column to existing database")
        
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
        
        c.execute('CREATE INDEX IF NOT EXISTS idx_pub_date ON announcements(pub_date)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_link ON announcements(link)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_hash ON announcements(hash)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_sent_at ON announcements(sent_at)')
        
        conn.commit()
        conn.close()
        logger.info("✅ Database initialized")
    
    def is_duplicate(self, guid, link, content_hash):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT 1 FROM announcements WHERE guid = ? OR link = ? OR hash = ?', 
                 (guid, link, content_hash))
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
                data['guid'], data['link'], data['title'],
                data['subject'], data['description'],
                data['date'], data['time'], data['sentiment'],
                data['pub_date'], data['hash']
            ))
            conn.commit()
            return c.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()
    
    def get_recent_count(self, minutes=5):
        """Get count of announcements from last N minutes"""
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''
            SELECT COUNT(*) FROM announcements 
            WHERE sent_at > datetime('now', ?)
        ''', (f'-{minutes} minutes',))
        count = c.fetchone()[0]
        conn.close()
        return count
    
    def get_today_count(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''
            SELECT COUNT(*) FROM announcements 
            WHERE DATE(sent_at) = DATE('now', 'localtime')
        ''')
        count = c.fetchone()[0]
        conn.close()
        return count
    
    def get_total_count(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT COUNT(*) FROM announcements')
        count = c.fetchone()[0]
        conn.close()
        return count
    
    def cleanup_old(self, days=30):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('DELETE FROM announcements WHERE sent_at < datetime("now", ?)', 
                 (f'-{days} days',))
        deleted = c.rowcount
        conn.commit()
        conn.close()
        if deleted:
            logger.info(f"🗑️ Cleaned up {deleted} old announcements")
        return deleted
    
    def count(self):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('SELECT COUNT(*) FROM announcements')
        count = c.fetchone()[0]
        conn.close()
        return count

# ===== HELPER FUNCTIONS =====
def sentiment_emoji(subject):
    BAD = ["resignation", "cessation", "insolvency", "litigation", "dispute", "default", "delay", "penalt", "fine", "downgrade"]
    GOOD = ["dividend", "bonus", "buyback", "acquisition", "awarding", "bagging", "merger", "upgrade", "record date"]
    
    s = subject.lower()
    if any(k in s for k in BAD):
        return "🔴"
    if any(k in s for k in GOOD):
        return "🟢"
    return "🟡"

def should_skip(title, subject, link):
    if not link or not link.strip():
        return True
    skip_keywords = ["declaration of nav", "net asset value", "mutual fund", "etf"]
    text = f"{title} {subject}".lower()
    if any(k in text for k in skip_keywords):
        return True
    return False

def truncate(text, max_chars=220):
    text = " ".join(text.split())
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + " ..."

def split_subject(description):
    description = " ".join(description.split())
    if "|SUBJECT:" in description:
        desc_part, subject_part = description.split("|SUBJECT:", 1)
        return desc_part.strip(), subject_part.strip()
    return description.strip(), ""

def split_datetime(pub):
    parts = pub.strip().split(" ", 1)
    return (parts[0], parts[1]) if len(parts) == 2 else (pub.strip(), "")

def escape_html(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def create_hash(title, subject, description):
    content = f"{title}|{subject}|{description}"
    return hashlib.md5(content.encode('utf-8')).hexdigest()

def parse_pub_date(pub_str):
    try:
        dt = parser.parse(pub_str)
        return dt.isoformat()
    except:
        return datetime.now().isoformat()

def send_telegram(text):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False
    
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        r = requests.post(url, data={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": text,
            "parse_mode": "HTML",
            "disable_web_page_preview": False,
        }, timeout=15)
        return r.ok
    except Exception as e:
        logger.error(f"❌ Telegram error: {e}")
        return False

def fetch_feed(max_attempts=3):
    for attempt in range(1, max_attempts + 1):
        try:
            resp = requests.get(FEED_URL, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            return feedparser.parse(resp.content)
        except Exception as e:
            logger.warning(f"⚠️ Attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                time.sleep(5 * attempt)
    return None

def write_excel(db):
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("⚠️ openpyxl not installed")
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

# ===== MAIN FUNCTION =====
def main():
    logger.info(f"🚀 NSE Watch starting at {datetime.now()}")
    logger.info(f"📡 Source: {RUN_SOURCE}")
    
    lock = FileLock()
    if not lock.acquire():
        logger.warning("⚠️ Another instance running. Exiting.")
        return
    
    try:
        db = AnnouncementDB()
        
        # Check if this is first run
        total_count = db.get_total_count()
        is_first_run = total_count == 0
        
        # Get current time for filtering
        now = datetime.now()
        five_minutes_ago = now - timedelta(minutes=5)
        
        if is_first_run:
            logger.info("🆕 FIRST RUN DETECTED! Creating baseline from last 5 minutes only...")
            logger.info(f"⏰ Baseline window: {five_minutes_ago.strftime('%H:%M:%S')} to {now.strftime('%H:%M:%S')}")
        
        # Cleanup old data (30 days)
        db.cleanup_old(30)
        
        # Fetch feed
        feed = fetch_feed()
        if feed is None:
            logger.error("❌ Failed to fetch feed")
            return
        
        # Get today's date
        today = now.date()
        today_str = today.strftime("%d-%b-%Y")
        
        logger.info(f"📅 Processing: {today_str}")
        
        new_count = 0
        duplicate_count = 0
        skipped_count = 0
        total_today = 0
        baseline_count = 0
        
        # Collect today's entries
        today_entries = []
        for entry in feed.entries:
            pub = getattr(entry, "published", "")
            if today_str in pub:
                today_entries.append(entry)
                total_today += 1
        
        logger.info(f"📨 Found {total_today} entries for today")
        
        # Process entries
        new_announcements = []
        
        for entry in today_entries:
            try:
                guid = entry.get("id", entry.get("link", ""))
                link = entry.get("link", "")
                title = getattr(entry, "title", "NSE Announcement")
                pub = getattr(entry, "published", "")
                raw_description = getattr(entry, "summary", "")
                
                if not link:
                    skipped_count += 1
                    continue
                
                description, subject = split_subject(raw_description)
                date_str, time_str = split_datetime(pub)
                
                if should_skip(title, subject, link):
                    skipped_count += 1
                    continue
                
                content_hash = create_hash(title, subject, description)
                
                if db.is_duplicate(guid, link, content_hash):
                    duplicate_count += 1
                    continue
                
                pub_date = parse_pub_date(pub)
                
                data = {
                    'guid': guid,
                    'link': link,
                    'title': title.strip(),
                    'subject': subject,
                    'description': description,
                    'date': date_str,
                    'time': time_str,
                    'sentiment': sentiment_emoji(subject),
                    'pub_date': pub_date,
                    'hash': content_hash,
                    'pub': pub
                }
                
                # Add to database
                if db.add_announcement(data):
                    new_count += 1
                    new_announcements.append(data)
                    
                    # Check if this entry is from last 5 minutes
                    try:
                        entry_time = parser.parse(pub)
                        if entry_time >= five_minutes_ago:
                            # This is a recent entry (within last 5 minutes)
                            pass  # Will be sent as alert
                        else:
                            # This is an old entry (baseline only)
                            baseline_count += 1
                    except:
                        pass
                
            except Exception as e:
                logger.error(f"❌ Error: {e}")
                continue
        
        # Send alerts ONLY for entries from last 5 minutes
        if new_announcements:
            # Filter only recent entries (last 5 minutes)
            recent_announcements = []
            baseline_announcements = []
            
            for data in new_announcements:
                try:
                    entry_time = parser.parse(data['pub'])
                    if entry_time >= five_minutes_ago:
                        recent_announcements.append(data)
                    else:
                        baseline_announcements.append(data)
                except:
                    # If can't parse, treat as recent
                    recent_announcements.append(data)
            
            # Sort recent announcements (newest first)
            recent_announcements.sort(key=lambda x: x['pub_date'], reverse=True)
            
            # Send alerts for recent announcements (newest first)
            if recent_announcements:
                # For first run, send a summary instead of individual alerts
                if is_first_run:
                    logger.info(f"📝 FIRST RUN: {len(baseline_announcements)} old entries saved as baseline")
                    logger.info(f"📝 FIRST RUN: {len(recent_announcements)} recent entries saved and alerted")
                    
                    # Send summary for first run
                    send_telegram(
                        f"🔄 NSE Watch initialized!\n"
                        f"📊 Baseline: {len(baseline_announcements)} old entries\n"
                        f"🆕 Recent alerts: {len(recent_announcements)} entries from last 5 minutes\n"
                        f"⏰ Time: {now.strftime('%H:%M:%S')}"
                    )
                    
                    # Still send recent alerts on first run (these are truly new)
                    for data in recent_announcements:
                        title_html = escape_html(data['title'])
                        desc_html = escape_html(truncate(data['description'], 220))
                        subject_html = escape_html(data['subject'])
                        
                        msg = (
                            f"{data['sentiment']} <a href=\"{data['link']}\">{title_html}</a>\n"
                            f"📄 {subject_html}\n"
                            f"{desc_html}\n"
                            f"🕐 {data['pub']}"
                        )
                        send_telegram(msg)
                        time.sleep(0.3)
                else:
                    # Normal run - send all recent alerts
                    logger.info(f"📤 Sending {len(recent_announcements)} alerts (newest first)...")
                    
                    if baseline_announcements:
                        logger.info(f"📝 {len(baseline_announcements)} older entries saved as baseline (no alerts)")
                    
                    for data in recent_announcements:
                        title_html = escape_html(data['title'])
                        desc_html = escape_html(truncate(data['description'], 220))
                        subject_html = escape_html(data['subject'])
                        
                        msg = (
                            f"{data['sentiment']} <a href=\"{data['link']}\">{title_html}</a>\n"
                            f"📄 {subject_html}\n"
                            f"{desc_html}\n"
                            f"🕐 {data['pub']}"
                        )
                        send_telegram(msg)
                        time.sleep(0.3)
        
        # Update Excel
        excel_count = write_excel(db) if new_count > 0 else 0
        
        # Summary
        summary = f"""
📊 TODAY'S SUMMARY ({today_str}):
   📨 Total in RSS: {total_today}
   ✅ New saved: {new_count}
   🔄 Duplicates: {duplicate_count}
   ⏭️ Filtered: {skipped_count}
   📊 Excel entries: {excel_count}
   📅 Database total: {db.count()}
   🆕 First run: {'Yes' if is_first_run else 'No'}
   ⏰ Last 5 min alerts: {len(recent_announcements) if 'recent_announcements' in locals() else 0}
   📝 Baseline entries: {baseline_count}
        """
        logger.info(summary)
        
        # Log to file
        with open(LOG_FILE, "a") as f:
            f.write(f"{datetime.now()}: Total={total_today}, New={new_count}, Dups={duplicate_count}, Skipped={skipped_count}, FirstRun={is_first_run}\n")
    
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        raise
    finally:
        lock.release()
        logger.info("✅ Done!")

if __name__ == "__main__":
    main()