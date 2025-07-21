import asyncio
from playwright.async_api import async_playwright
from datetime import datetime, timedelta
from urllib.parse import quote
import os
import re
import pytz

from icalendar import Calendar
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# ----------------------------
# CONFIGURATION
# ----------------------------
HUMANFORCE_URL = "https://maroondah.humanforce.com"
USERNAME = "your_email_here"
PASSWORD = "your_password_here"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
SCOPES = ['https://www.googleapis.com/auth/calendar.events']
CALENDAR_ID = "your_calendar_id_here"

# ----------------------------
# 1. Download .ics files for this + next month
# ----------------------------
async def download_icals():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    today = datetime.today()
    first_this_month = today.replace(day=1)
    first_next_month = (first_this_month.replace(day=28) + timedelta(days=4)).replace(day=1)

    dates = [
        ("this", first_this_month.strftime("%d/%m/%Y")),
        ("next", first_next_month.strftime("%d/%m/%Y"))
    ]

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        # Login
        await page.goto(HUMANFORCE_URL)
        await page.fill('input[name="UserName"]', USERNAME)
        await page.fill('input[name="Password"]', PASSWORD)
        await page.click('button[type="submit"]')
        await page.wait_for_url(re.compile(r"https://maroondah\.humanforce\.com/Home/?"), timeout=20000)

        files = []
        for label, date_str in dates:
            print(f"\n🔄 Getting roster for: {label} month")

            # Go to the Roster page
            await page.goto(f"{HUMANFORCE_URL}/Roster")
            await page.wait_for_selector("a.button-export")

            # Change the link's href to use current date
            await page.evaluate(f"""
                () => {{
                    const link = document.querySelector("a.button-export");
                    if (link) {{
                        link.href = "/Roster/ExportToICS?from={quote(date_str)}";
                    }}
                }}
            """)

            # Expect download and click
            async with page.expect_download() as download_info:
                await page.click("a.button-export")
            download = await download_info.value
            file_path = os.path.join(DOWNLOAD_DIR, f"roster_{label}.ics")
            await download.save_as(file_path)
            print(f"✅ Downloaded iCal for {label} month: {file_path}")
            files.append(file_path)

        await browser.close()
        return files

# ----------------------------
# 2. Parse .ics to extract events
# ----------------------------
def parse_ical(file_path):
    events = []
    with open(file_path, 'rb') as f:
        cal = Calendar.from_ical(f.read())

    for component in cal.walk():
        if component.name == "VEVENT":
            event = {
                "summary": str(component.get("summary")),
                "start": component.get("dtstart").dt,
                "end": component.get("dtend").dt,
                "uid": str(component.get("uid")).strip(),
                "description": str(component.get("description") or "").strip()
            }
            events.append(event)

    print(f"📆 Parsed {len(events)} events from {file_path}")
    return events

# ----------------------------
# 3. Connect to Google Calendar and upload events
# ----------------------------
def get_calendar_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)

def upload_events_to_gcal(events):
    service = get_calendar_service()
    timezone = pytz.timezone("Australia/Melbourne")

    for event in events:
        start = event['start']
        end = event['end']

        # Ensure timezone-aware
        if isinstance(start, datetime) and start.tzinfo is None:
            start = timezone.localize(start)
        if isinstance(end, datetime) and end.tzinfo is None:
            end = timezone.localize(end)

        g_event = {
            'summary': event['summary'],
            'start': {'dateTime': start.isoformat(), 'timeZone': str(start.tzinfo)},
            'end': {'dateTime': end.isoformat(), 'timeZone': str(end.tzinfo)},
            'description': event.get("description", "Imported from Humanforce"),
        }

        try:
            # Look for existing events in the window (5hr tolerance)
            results = service.events().list(
                calendarId=CALENDAR_ID,
                timeMin=start.isoformat(),
                timeMax=end.isoformat(),
                singleEvents=True,
                maxResults=5,
                orderBy="startTime"
            ).execute()

            existing = results.get('items', [])
            matched = None

            for e in existing:
                estart = e.get("start", {}).get("dateTime")
                eend = e.get("end", {}).get("dateTime")
                if estart and eend:
                    estart_dt = datetime.fromisoformat(estart)
                    eend_dt = datetime.fromisoformat(eend)
                    # Check if the event overlaps within a 5-hour window
                    if abs((start - estart_dt).total_seconds()) <= 18000:
                        matched = e
                        break

            if matched:
                event_id = matched["id"]
                service.events().update(calendarId=CALENDAR_ID, eventId=event_id, body=g_event).execute()
                print(f"🔁 Updated: {event['summary']} ({start} → {end})")
            else:
                service.events().insert(calendarId=CALENDAR_ID, body=g_event).execute()
                print(f"✅ Created: {event['summary']} ({start} → {end})")

        except Exception as e:
            print(f"⚠️ Error: {event['summary']} — {e}")


# ----------------------------
# 4. Main Runner
# ----------------------------
if __name__ == "__main__":
    ics_files = asyncio.run(download_icals())
    all_events = []
    for ics in ics_files:
        all_events.extend(parse_ical(ics))
    upload_events_to_gcal(all_events)
