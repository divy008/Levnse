import os
import json
import time
import feedparser
import requests
from datetime import datetime, timedelta
import sqlite3
from pathlib import Path
import hashlib
import sys
import logging

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

# ===== LOGGING SETUP =====
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
            # Check if lock is stale (older than 5 minutes)
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
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                hash TEXT UNIQUE
            )
        ''')
        c.execute('CREATE INDEX IF NOT EXISTS idx_link ON announcements(link)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_sent_at ON announcements(sent_at)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_hash ON announcements(hash)')
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
                (guid, link, title, subject, description, date, time, sentiment, hash)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['guid'], data['link'], data['title'],
                data['subject'], data['description'],
                data['date'], data['time'], data['sentiment'],
                data['hash']
            ))
            conn.commit()
            return c.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()
    
    def get_recent(self, days=30):
        conn = sqlite3.connect(self.db_file)
        c = conn.cursor()
        c.execute('''
            SELECT * FROM announcements 
            WHERE sent_at > datetime('now', ?)
            ORDER BY sent_at DESC
        ''', (f'-{days} days',))
        results = c.fetchall()
        conn.close()
        return results
    
    def cleanup_old(self, days=90):
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
    BAD = ["resignation", "cessation", "insolvency", "litigation", "dispute", "default", "delay", "penalt", "fine"]
    GOOD = ["dividend", "bonus", "buyback", "acquisition", "awarding", "bagging", "merger", "upgrade"]
    
    s = subject.lower()
    if any(k in s for k in BAD):
        return "🔴"
    if any(k in s for k in GOOD):
        return "🟢"
    return "🟡"

def should_skip(title, subject, link):
    if not link or not link.strip():
        return True
    skip_subjects = ["declaration of nav", "net asset value"]
    skip_titles = ["mutual fund", "etf"]
    
    if any(k in subject.lower() for k in skip_subjects):
        return True
    if any(k in title.lower() for k in skip_titles):
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
    """Create unique hash for content to detect duplicates"""
    content = f"{title}|{subject}|{description}"
    return hashlib.md5(content.encode('utf-8')).hexdigest()

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
    """Export database to Excel with batching"""
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("⚠️ openpyxl not installed, skipping Excel export")
        return
    
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Announcements")
    ws.append(["Sentiment", "Company", "Subject", "Description", "Date", "Time", "Link", "Sent At"])
    
    conn = sqlite3.connect(db.db_file)
    c = conn.cursor()
    c.execute('SELECT sentiment, title, subject, description, date, time, link, sent_at FROM announcements ORDER BY sent_at DESC LIMIT 5000')
    
    count = 0
    for row in c:
        ws.append([str(x) if x else "" for x in row])
        count += 1
    
    conn.close()
    wb.save(EXCEL_FILE)
    logger.info(f"📊 Excel file updated with {count} recent announcements")

# ===== MAIN FUNCTION =====
def main():
    logger.info(f"🚀 NSE Bot starting at {datetime.now()}")
    logger.info(f"📡 Source: {RUN_SOURCE}")
    
    # Acquire lock (prevent concurrent runs)
    lock = FileLock()
    if not lock.acquire():
        logger.warning("⚠️ Another instance is already running. Exiting.")
        return
    
    try:
        # Initialize database
        db = AnnouncementDB()
        logger.info(f"📊 Database has {db.count()} total announcements")
        
        # Cleanup old data (90 days)
        db.cleanup_old(90)
        
        # Fetch feed
        feed = fetch_feed()
        if feed is None:
            logger.error("❌ Failed to fetch feed. Exiting.")
            return
        
        # Process entries
        new_count = 0
        skipped_count = 0
        duplicate_count = 0
        error_count = 0
        
        # Process entries in reverse (oldest first) - limit to 50 per run
        entries_to_process = list(reversed(feed.entries))[:50]
        logger.info(f"📨 Processing {len(entries_to_process)} entries...")
        
        for entry in entries_to_process:
            try:
                guid = entry.get("id", entry.get("link", ""))
                link = entry.get("link", "")
                title = getattr(entry, "title", "New NSE Announcement")
                pub = getattr(entry, "published", "")
                raw_description = getattr(entry, "summary", "")
                
                # Skip if no GUID or link
                if not guid or not link:
                    skipped_count += 1
                    continue
                
                # Process description
                description, subject = split_subject(raw_description)
                date_str, time_str = split_datetime(pub)
                
                # Skip unwanted announcements
                if should_skip(title, subject, link):
                    skipped_count += 1
                    continue
                
                # Create content hash
                content_hash = create_hash(title, subject, description)
                
                # Check duplicate
                if db.is_duplicate(guid, link, content_hash):
                    duplicate_count += 1
                    continue
                
                # Prepare data
                data = {
                    'guid': guid,
                    'link': link,
                    'title': title.strip(),
                    'subject': subject,
                    'description': description,
                    'date': date_str,
                    'time': time_str,
                    'sentiment': sentiment_emoji(subject),
                    'hash': content_hash
                }
                
                # Save to database
                if db.add_announcement(data):
                    new_count += 1
                    
                    # Send Telegram
                    title_html = escape_html(title)
                    desc_html = escape_html(truncate(description, 220))
                    subject_html = escape_html(subject)
                    
                    msg = (
                        f"{data['sentiment']} <a href=\"{link}\">{title_html}</a>\n"
                        f"📄 {subject_html}\n"
                        f"{desc_html}\n"
                        f"🕐 {pub}"
                    )
                    send_telegram(msg)
                    time.sleep(0.5)  # Rate limiting
                
            except Exception as e:
                error_count += 1
                logger.error(f"❌ Error processing entry: {e}")
                continue
        
        # Update Excel file if there are new announcements
        if new_count > 0:
            write_excel(db)
        
        # Summary
        summary = f"""
📊 SUMMARY:
   ✅ New announcements: {new_count}
   🔄 Duplicates skipped: {duplicate_count}
   ⏭️ Filtered skipped: {skipped_count}
   ❌ Errors: {error_count}
   📅 Database size: {db.count()} total
   📁 Source: {RUN_SOURCE}
        """
        logger.info(summary)
        
        # Log to file
        with open(LOG_FILE, "a") as f:
            f.write(f"{datetime.now()}: New={new_count}, Dups={duplicate_count}, Skipped={skipped_count}, Source={RUN_SOURCE}\n")
        
        # Send summary to Telegram if there were new announcements
        if new_count > 0:
            send_telegram(f"✅ NSE Bot: {new_count} new announcements processed successfully!")
        
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        send_telegram(f"❌ NSE Bot crashed: {str(e)[:100]}")
        raise
    
    finally:
        lock.release()
        logger.info("✅ Done!")

if __name__ == "__main__":
    main()