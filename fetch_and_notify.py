import os
import json
import feedparser
import requests

FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
STATE_FILE = os.path.join(os.path.dirname(__file__), "seen.json")
LOG_FILE = os.path.join(os.path.dirname(__file__), "log.json")
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "announcements.xlsx")

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]

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
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE) as f:
            return set(json.load(f))
    return set()


def save_seen(seen):
    # keep file bounded so it doesn't grow forever
    with open(STATE_FILE, "w") as f:
        json.dump(list(seen)[-3000:], f)


def escape_html(text):
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


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
            return json.load(f)
    return []


def save_log(log):
    with open(LOG_FILE, "w") as f:
        json.dump(log, f, indent=2, ensure_ascii=False)


def write_excel(log):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Announcements"
    headers = ["Company", "Subject", "Description", "Date", "Time", "Link"]
    ws.append(headers)
    for row in log:
        ws.append([
            row.get("company", ""),
            row.get("subject", ""),
            row.get("description", ""),
            row.get("date", ""),
            row.get("time", ""),
            row.get("link", ""),
        ])
    # basic column widths so it's readable
    widths = [30, 30, 60, 14, 12, 45]
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


def main():
    resp = requests.get(FEED_URL, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    feed = feedparser.parse(resp.content)

    seen = load_seen()
    is_first_run = len(seen) == 0
    new_items = []

    for entry in feed.entries:
        guid = entry.get("id", entry.link)
        if guid not in seen:
            new_items.append(entry)
            seen.add(guid)

    if is_first_run:
        # don't spam on first ever run, just record baseline
        print(f"First run: recorded {len(seen)} existing items as baseline, no alerts sent.")
        save_seen(seen)
        return

    log = load_log()

    for entry in reversed(new_items):  # oldest first
        title = getattr(entry, "title", "New NSE Announcement")
        link = getattr(entry, "link", "")
        pub = getattr(entry, "published", "")
        raw_description = getattr(entry, "summary", "")

        description, subject = split_subject(raw_description)
        date_str, time_str = split_datetime(pub)

        # append structured record for JSON/Excel log
        log.append({
            "company": title.strip(),
            "subject": subject,
            "description": description,
            "date": date_str,
            "time": time_str,
            "link": link,
        })

        title_html = escape_html(title)
        desc_html = escape_html(truncate(description, max_chars=220))

        msg = (
            f"\U0001F4E2 <a href=\"{link}\">{title_html}</a>\n"
            f"{desc_html}\n"
            f"\U0001F553 {pub}"
        )
        send_telegram(msg)

    save_log(log)
    write_excel(log)

    print(f"Sent {len(new_items)} new announcement(s).")
    save_seen(seen)


if __name__ == "__main__":
    main()
