import os
import json
import feedparser
import requests

FEED_URL = "https://nsearchives.nseindia.com/content/RSS/Online_announcements.xml"
STATE_FILE = os.path.join(os.path.dirname(__file__), "seen.json")

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


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    r = requests.post(
        url,
        data={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": text,
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

    for entry in reversed(new_items):  # oldest first
        title = getattr(entry, "title", "New NSE Announcement")
        link = getattr(entry, "link", "")
        pub = getattr(entry, "published", "")
        msg = f"\U0001F4E2 {title}\n\U0001F553 {pub}\n\U0001F517 {link}"
        send_telegram(msg)

    print(f"Sent {len(new_items)} new announcement(s).")
    save_seen(seen)


if __name__ == "__main__":
    main()
  
