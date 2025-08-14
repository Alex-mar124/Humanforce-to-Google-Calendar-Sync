
# Humanforce → Google Calendar Sync

A tool that automatically downloads your Humanforce roster and syncs it to your Google Calendar.


## 🚀 Features
- Downloads `.ics` roster files for **current** and **next month** from Humanforce.
- Parses and uploads events to Google Calendar.
- Prevents duplicate events by checking for existing entries within a 5-hour window.
- Works in the background with **Windows Task Scheduler**.

## 📥 Installation

### 1. [Download the Latest Release](https://github.com/Alex-mar124/Humanforce-to-Google-Calendar-Sync/releases/latest)

### 2. Install Python
- Install **Python 3.10+** from [python.org](https://www.python.org/downloads/).

### Run the One Click Setup
Run the included install.bat script by double-clicking it.

### 2. Enable Google Calendar API

1. Go to **Google Cloud Console** → [Enable Google Calendar API](https://developers.google.com/calendar/api/quickstart/python).
2. Click **Enable the API** and **Create Credentials** → OAuth 2.0 Client ID.
3. Download the `credentials.json` file and place it in the app’s data folder (or next to the `.py` file).
4. On first run, your browser will open to authenticate with Google.

### 3. Configure HumanForce Settings

* Edit the settings with your infomation:

  * `HUMANFORCE_URL`
  * `USERNAME`
  * `PASSWORD`
  * `CALENDAR_ID` (Google Calendar ID where shifts should be added)


---

## ▶ Usage

Run manually:

```bash
python sync.py
```

Or sync without opening the GUI:

```bash
python sync.py --sync
```

**Daily Automatic Sync**

* Use **Windows Task Scheduler** to run:

```bash
python sync.py --sync
```

every day at your preferred time.

---

## 🗂 Data Files

* `credentials.json` → Your Google OAuth client credentials.
* `token.json` → Created after authenticating with Google.
* `.env` → Stores Humanforce + calendar configuration.
* `downloads/` → Temporary folder for downloaded roster `.ics` files.

---

## 💡 Tips

* You can find your Google Calendar ID under **Calendar Settings → Integrate calendar**.
* If you change `credentials.json`, delete `token.json` so the app will re-authenticate.


