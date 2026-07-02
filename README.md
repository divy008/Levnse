# NSE Announcement Notifier — 100% Free Setup

No paid tier, no credit card, no bank details anywhere in this stack.

## What it uses
- **GitHub Actions** — free scheduled runs. Unlimited minutes on a **public** repo.
- **Telegram Bot API** — free, no signup limits, sends you a message the instant new items appear.

## Setup (10 minutes)

### 1. Create a Telegram bot
1. Open Telegram, search for **@BotFather**, start a chat.
2. Send `/newbot`, follow the prompts, give it any name.
3. BotFather gives you a **token** like `123456789:AAExxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`. Save it.
4. Send any message to your new bot (search its username and message it directly).
5. Visit this URL in your browser (replace TOKEN):
   `https://api.telegram.org/botTOKEN/getUpdates`
   Find `"chat":{"id":123456789,...}` — that number is your **chat ID**.

### 2. Create a GitHub repo
1. Go to github.com (free account, no card needed) → New repository → make it **Public**
   (public repos get unlimited free Actions minutes; private repos have a limited free quota).
2. Upload these three files, keeping the folder structure:
   - `fetch_and_notify.py`
   - `.github/workflows/nse-watch.yml`
   - `README.md`

### 3. Add your secrets
In the repo: **Settings → Secrets and variables → Actions → New repository secret**
- `TELEGRAM_BOT_TOKEN` → the token from step 1
- `TELEGRAM_CHAT_ID` → the chat ID from step 1

### 4. Run it
- Go to the **Actions** tab → select "NSE Announcements Watcher" → **Run workflow** (manual first run).
- The first run just records a baseline (no alert spam of old news).
- After that, it runs automatically every 10 minutes and messages you on Telegram
  the moment a new announcement appears.

## Notes
- You can lower the interval in `nse-watch.yml` (`cron: "*/5 * * * *"` for every 5 min),
  but GitHub's scheduler doesn't guarantee exact timing under load — treat it as "roughly every N minutes."
- If you ever see zero results even during market hours, NSE is likely rate-limiting the
  GitHub Actions IP. If that happens, ask me and I'll show you a fallback (e.g. routing
  the request through a free proxy or switching poll frequency).
- Everything here — GitHub, Telegram, the script — stays free indefinitely at this scale.
  No feature will suddenly ask for a card.
