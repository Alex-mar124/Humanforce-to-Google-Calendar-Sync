# Humanforce to Google Calendar Sync

This Python script automates the process of logging into the Humanforce web portal, downloading your shift roster as an `.ics` file (for both the current and next month), parsing the shift data, and syncing it to your Google Calendar. It includes logic to update existing shifts and prevent duplicates even for overnight shifts!

---

## 🚀 Features

- Headless browser automation using Playwright to log in and download `.ics` files
- Supports both current and next month shifts
- Parses `.ics` files for shift data including notes (description)
- Syncs events to Google Calendar
- Overwrites existing events within a 5-hour match window
- Safe to run multiple times without duplicates

---

## 🧩 Requirements

- Python 3.9+
- Google Account (for Calendar access)
- Humanforce login credentials

---

## 🛠️ Installation & Setup

1. **Install dependencies**

```bash
pip install playwright icalendar google-api-python-client google-auth-httplib2 google-auth-oauthlib
python -m playwright install
```

2. **Download Google Calendar API credentials**

Go to [Google Developer Console](https://console.cloud.google.com/):
- Create a new project
- Enable the **Google Calendar API**
- Go to “Credentials” → “OAuth Client ID”
  - Choose: **Desktop App**
- Download the creds and save it as `credentials.json`
- Place it in the same directory as this script
- For Full instructions follow **[this tutorial](https://developers.google.com/workspace/calendar/api/quickstart/python)**

📝 On first run, you'll be prompted to log into your Google account (saved locally as `token.json`.)

3. **Set your Humanforce credentials and Google Calender**

❗ **Recommended to create a new calender so your existing calender dosnt interfere** ❗

Edit the script and update:
```python
USERNAME = "your_email_here"
PASSWORD = "your_password_here"
CALENDAR_ID = ""
```

---

## 📆 How It Works

- Downloads `.ics` files for both **this month** and **next month**
- Parses each shift, including description/notes
- For each shift:
  - Checks for existing Google Calendar events **within 5 hours of the start time**
  - If found, updates the event
  - If not found, inserts a new one

---

## 🔐 Permissions & Security

- Your Google login is handled through official OAuth 2.0
- A token is saved as `token.json` for repeat access
- You can safely delete `token.json` to log in again
- Coworkers can use the same `credentials.json` and authorize with their own Google account

---

## 🤝 Sharing with Others

You may share this script with your coworkers. They can:
- Use your `credentials.json` (OAuth client)
- Sign in with their own Google account
- Each person’s shifts are synced to their personal calendar

🔒 Google API usage will be attributed to your API project unless others create their own credentials.

---

## ⚠️ Limitations

- Only supports **one shift per day** (within 5hrs) if you have split or overlapping shifts, you’ll need to adjust the logic
- Humanforce must allow `.ics` downloads via clickable links on the roster page
- `UID` fields in the `.ics` are not stable, so deduplication is done via timestamp comparison (±5 hours)
- Google Calendar API quota applies (but this script is very lightweight)

---

## ✅ Recommended Automation

To run this script daily:

- On **Windows**:
  - Use Task Scheduler to run `python humanforce_sync.py`
- On **Linux/macOS**:
  - Use a cron job: `0 8 * * * /usr/bin/python3 /path/to/humanforce_sync.py`

---

## 🧼 Troubleshooting

- **Shifts are duplicated?**
  - Your `.ics` file might be regenerating unstable UIDs, this script avoids that by matching by time.
- **Nothing appears in your calendar?**
  - Check that `token.json` exists and you signed into the correct Google account.
- **Overlapping shifts?**
  - You may need to adjust the match logic in the script.

---

