import os
import json
import time
import feedparser
import requests

FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
STATE_FILE = os.path.join(os.path.dirname(__file__), "seen.json")
LOG_FILE = os.path.join(os.path.dirname(__file__), "log.json")
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "announcements.xlsx")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
    raise ValueError("Missing critical Telegram environment variables! Check TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID.")

# NSE blocks requests without browser-like headers
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Referer": "https://www.nseindia.com/",
    "Accept": "application/xml,text/xml,*/*",
}


def load_seen():
    """Loads history as a list to maintain chronological order."""
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE) as f:
            data = json.load(f)
            return data if isinstance(data, list) else list(data)
    return []


def save_seen(seen_list):
    """Saves only the most recent 3000 items sequentially."""
    with open(STATE_FILE, "w") as f:
        json.dump(seen_list[-3000:], f)


def escape_html(text):
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# Keyword-based sentiment classification for NSE announcement subjects.
BAD_KEYWORDS = [
    "resignation", "cessation", "insolvency", "litigation", "dispute",
    "default", "delay", "penalt", "fine", "action initiated", "action taken",
    "orders passed", "takeover", "corporate insolvency", "winding up",
    "reduction in capital", "downgrade",
]

GOOD_KEYWORDS = [
    "dividend", "bonus", "buyback", "acquisition", "awarding of order",
    "bagging", "receiving of order", "credit rating- new", "capacity addition",
    "commencement of commercial production", "investor presentation",
    "allotment of securities", "amalgamation", "merger", "upgrade",
    "record date", "scheme of arrangement",
]


def sentiment_emoji(subject):
    s = subject.lower()
    if any(k in s for k in BAD_KEYWORDS):
        return "\U0001F534"  # 🔴
    if any(k in s for k in GOOD_KEYWORDS):
        return "\U0001F7E2"  # 🟢
    return "\U0001F7E1"  # 🟡


def truncate(text, max_chars=220):
    text = " ".join(text.split())  # collapse newlines/extra whitespace
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + " ..."


def split_subject(description):
    """NSE descriptions often end with '|SUBJECT: xxx'. Split it out."""
    description = " ".join(description.split())
    if "|SUBJECT:" in description:
        desc_part, subject_part = description.split("|SUBJECT:", 1)
        return desc_part.strip(), subject_part.strip()
    return description.strip(), ""


def split_datetime(pub):
    """pubDate looks like '03-Jul-2026 12:09:34' -> ('03-Jul-2026', '12:09:34')"""
    parts = pub.strip().split(" ", 1)
    if len(parts) == 2:
        return parts[0], parts[1]
    return pub.strip(), ""


def load_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE) as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return []
    return []


def save_log(log):
    with open(LOG_FILE, "w") as f:
        json.dump(log, f, indent=2, ensure_ascii=False)


def write_excel(log):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Announcements"
    headers = ["Sentiment", "Company", "Subject", "Description", "Date", "Time", "Link"]
    ws.append(headers)
    for row in log:
        ws.append([
            row.get("sentiment", ""),
            row.get("company", ""),
            row.get("subject", ""),
            row.get("description", ""),
            row.get("date", ""),
            row.get("time", ""),
            row.get("link", ""),
        ])
    # basic column widths so it's readable
    widths = [10, 30, 30, 60, 14, 12, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    wb.save(EXCEL_FILE)


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
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
    if not r.ok:
        print("Telegram send failed:", r.text)


def fetch_feed(max_attempts=3):
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            resp = requests.get(FEED_URL, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            return feedparser.parse(resp.content)
        except requests.exceptions.RequestException as e:
            last_error = e
            print(f"Attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                time.sleep(5 * attempt)  # 5s, 10s backoff
    print(f"All {max_attempts} attempts failed. Last error: {last_error}")
    return None


def main():
    feed = fetch_feed()
    if feed is None:
        print("Skipping this run due to fetch failure. Will retry next scheduled run.")
        return

    seen_list = load_seen()
    seen_set = set(seen_list)  # Using a set lookup here is highly efficient
    is_first_run = len(seen_list) == 0
    new_items = []

    for entry in feed.entries:
        guid = entry.get("id", entry.link)
        if guid not in seen_set:
            new_items.append(entry)
            seen_list.append(guid)  # Keeps chronological sequence

    log = load_log()
    existing_links = {row.get("link", "") for row in log if row.get("link")}

    # Force write files even on empty/first runs so Git always finds them
    if not os.path.exists(LOG_FILE) or is_first_run:
        save_log(log)
    if not os.path.exists(EXCEL_FILE) or is_first_run:
        write_excel(log)

    if is_first_run:
        print(f"First run: recorded {len(seen_list)} existing items as baseline, no alerts sent.")
        save_seen(seen_list)
        return

    telegram_sent_count = 0

    for entry in reversed(new_items):  # oldest first
        title = getattr(entry, "title", "New NSE Announcement")
        link = getattr(entry, "link", "")
        pub = getattr(entry, "published", "")
        raw_description = getattr(entry, "summary", "")

        description, subject = split_subject(raw_description)
        date_str, time_str = split_datetime(pub)
        emoji = sentiment_emoji(subject)

        # Duplicate protection: skip if link already logged
        if link and link in existing_links:
            continue

        log.append({
            "company": title.strip(),
            "subject": subject,
            "description": description,
            "date": date_str,
            "time": time_str,
            "link": link,
            "sentiment": emoji,
        })
        if link:
            existing_links.add(link)

        title_html = escape_html(title)
        desc_html = escape_html(truncate(description, max_chars=220))
        subject_html = escape_html(subject)
        link_html = escape_html(link)

        msg = (
            f"{emoji} <a href=\"{link_html}\">{title_html}</a>\n"
            f"\U0001F4C4 {subject_html}\n"
            f"{desc_html}\n"
            f"\U0001F553 {pub}"
        )
        send_telegram(msg)
        telegram_sent_count += 1
        time.sleep(1.5)

    if telegram_sent_count > 0:
        save_log(log)
        write_excel(log)

    print(f"Sent {telegram_sent_count} new announcement(s).")
    save_seen(seen_list)


if __name__ == "__main__":
    main()
