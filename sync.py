#!/usr/bin/env python3
import os
import sys
import re
import json
import asyncio
import threading
import subprocess
from datetime import datetime, timedelta
from urllib.parse import quote
from argparse import ArgumentParser
from typing import List, Dict, Any
import tempfile

# Third‑party
from icalendar import Calendar
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import pytz

# Optional runtime imports (installed on demand by Quick Setup)
try:
    from playwright.async_api import async_playwright
except Exception:
    async_playwright = None

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None



APP_NAME = "HumanforceGCalSync"
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DOWNLOAD_DIR_DEFAULT = os.path.join(BASE_DIR, "downloads")
ENV_PATH = os.path.join(BASE_DIR, ".env")
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")
CRED_PATH = os.path.join(BASE_DIR, "credentials.json")
SCOPES = ['https://www.googleapis.com/auth/calendar.events']
LOG_PATH = os.path.join(BASE_DIR, "sync.log")
HISTORY_PATH = os.path.join(BASE_DIR, "history.json")
HISTORY_MAX = 100

# ----------------------------
# Config helpers
# ----------------------------
class Config:
    def __init__(self):
        self.HUMANFORCE_URL = os.getenv("HUMANFORCE_URL", "https://maroondah.humanforce.com")
        self.USERNAME = os.getenv("USERNAME", "")
        self.PASSWORD = os.getenv("PASSWORD", "")
        self.DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", DOWNLOAD_DIR_DEFAULT)
        self.CALENDAR_ID = os.getenv("CALENDAR_ID", "")
        self.TIMEZONE = os.getenv("TIMEZONE", "Australia/Melbourne")
        # scheduling
        self.DAILY_SYNC_ENABLED = os.getenv("DAILY_SYNC_ENABLED", "false").lower() == "true"
        self.DAILY_SYNC_HOUR = int(os.getenv("DAILY_SYNC_HOUR", "7"))
        # ui
        self.DARK_MODE = os.getenv("DARK_MODE", "false").lower() == "true"

    def to_env(self) -> str:
        lines = [
            f"HUMANFORCE_URL={self.HUMANFORCE_URL}",
            f"USERNAME={self.USERNAME}",
            f"PASSWORD={self.PASSWORD}",
            f"DOWNLOAD_DIR={self.DOWNLOAD_DIR}",
            f"CALENDAR_ID={self.CALENDAR_ID}",
            f"TIMEZONE={self.TIMEZONE}",
            f"DAILY_SYNC_ENABLED={'true' if self.DAILY_SYNC_ENABLED else 'false'}",
            f"DAILY_SYNC_HOUR={self.DAILY_SYNC_HOUR}",
            f"DARK_MODE={'true' if self.DARK_MODE else 'false'}",
        ]
        return "\n".join(lines) + "\n"

    def save(self):
        with open(ENV_PATH, "w", encoding="utf-8") as f:
            f.write(self.to_env())


def ensure_dotenv_loaded() -> None:
    global load_dotenv
    if load_dotenv is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-dotenv"])
        from dotenv import load_dotenv as _ld
        load_dotenv = _ld
    load_dotenv(ENV_PATH)


# ----------------------------
# Logging & history
# ----------------------------

def _is_valid_json_file(path: str) -> bool:
    """Return True if the file at path exists and contains valid JSON; False otherwise."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r", encoding="utf-8") as f:
            json.load(f)
        return True
    except Exception:
        return False



def _append_log(line: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {line}\n")


def _load_history() -> List[Dict[str, Any]]:
    if not os.path.exists(HISTORY_PATH):
        return []
    try:
        with open(HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_history(entries: List[Dict[str, Any]]) -> None:
    entries = entries[-HISTORY_MAX:]
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)


# ----------------------------
# Core logic (download + parse + upload) with progress callbacks
# ----------------------------
async def download_icals(cfg: Config, on_log=None):
    log = on_log or (lambda *a, **k: None)
    os.makedirs(cfg.DOWNLOAD_DIR, exist_ok=True)
    log(f"Ensured downloads folder: {cfg.DOWNLOAD_DIR}")

    today = datetime.today()
    first_this_month = today.replace(day=1)
    first_next_month = (first_this_month.replace(day=28) + timedelta(days=4)).replace(day=1)

    dates = [
        ("this", first_this_month.strftime("%d/%m/%Y")),
        ("next", first_next_month.strftime("%d/%m/%Y")),
    ]

    global async_playwright
    if async_playwright is None:
        log("Installing Playwright (first‑time setup)…")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "playwright>=1.38"])  # may take a moment
        from playwright.async_api import async_playwright as _apw
        async_playwright = _apw
        log("Installing Chromium runtime for Playwright…")
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])  # one‑off
        log("Playwright ready.")

    log("Launching headless Chromium…")
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        log("Opening Humanforce login page…")
        await page.goto(cfg.HUMANFORCE_URL)
        await page.fill('input[name="UserName"]', cfg.USERNAME)
        await page.fill('input[name="Password"]', cfg.PASSWORD)
        await page.click('button[type="submit"]')
        await page.wait_for_url(re.compile(r"https://.+/Home/?"), timeout=30000)
        log("Logged in successfully.")

        files = []
        for label, date_str in dates:
            log(f"Preparing export for {label} month (from {date_str})…")
            await page.goto(f"{cfg.HUMANFORCE_URL}/Roster")
            await page.wait_for_selector("a.button-export")

            await page.evaluate(
                f"""
                () => {{
                    const link = document.querySelector("a.button-export");
                    if (link) {{ link.href = "/Roster/ExportToICS?from={quote(date_str)}"; }}
                }}
                """
            )

            async with page.expect_download() as dl_info:
                await page.click("a.button-export")
            download = await dl_info.value
            file_path = os.path.join(cfg.DOWNLOAD_DIR, f"roster_{label}.ics")
            await download.save_as(file_path)
            log(f"Saved: {file_path}")
            files.append(file_path)

        await browser.close()
        log("Browser closed.")
        return files


def parse_ical(file_path, on_log=None):
    log = on_log or (lambda *a, **k: None)
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
                "description": str(component.get("description") or "").strip(),
            }
            events.append(event)
    log(f"Parsed {len(events)} events from {os.path.basename(file_path)}")
    return events


def get_calendar_service():
    """Build and return an authenticated Google Calendar service.

    Robust against a corrupted/empty token.json (common cause of JSONDecodeError: 'Expecting value').
    If token.json is invalid, it is removed and the OAuth flow is triggered again.
    Also validates credentials.json before attempting the OAuth flow to give a clearer error.
    """
    creds = None

    # Handle existing token.json safely
    if _is_valid_json_file(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        except Exception as e:
            _append_log(f"token.json could not be parsed: {e}. Will re-authenticate.")
            creds = None
    elif os.path.exists(TOKEN_PATH):
        # token.json exists but is invalid/empty
        try:
            _append_log("token.json is empty or invalid JSON. Deleting and re-authenticating…")
            os.remove(TOKEN_PATH)
        except Exception as e:
            _append_log(f"Failed to remove invalid token.json: {e}")

    # If we don't have valid creds, run the OAuth flow
    if not creds or not creds.valid:
        if not _is_valid_json_file(CRED_PATH):
            raise FileNotFoundError(
                "credentials.json is missing or invalid JSON — place your Google OAuth client file next to this script."
            )
        flow = InstalledAppFlow.from_client_secrets_file(CRED_PATH, SCOPES)
        creds = flow.run_local_server(port=0)
        try:
            with open(TOKEN_PATH, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())
        except Exception as e:
            _append_log(f"Warning: failed to write token.json: {e}")

    return build('calendar', 'v3', credentials=creds)


def upload_events_to_gcal(cfg: Config, events, on_log=None) -> Dict[str, int]:
    log = on_log or (lambda *a, **k: None)
    service = get_calendar_service()
    timezone = pytz.timezone(cfg.TIMEZONE)

    stats = {"created": 0, "updated": 0, "errors": 0}

    for event in events:
        start = event['start']
        end = event['end']

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
            time_min = (start - timedelta(hours=5)).isoformat()
            time_max = (end + timedelta(hours=5)).isoformat()
            results = service.events().list(
                calendarId=cfg.CALENDAR_ID,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                maxResults=10,
                orderBy="startTime",
            ).execute()

            existing = results.get('items', [])
            matched = None
            for e in existing:
                estart = e.get("start", {}).get("dateTime")
                if not estart:
                    continue
                estart_dt = datetime.fromisoformat(estart.replace('Z', '+00:00'))
                if abs((start - estart_dt).total_seconds()) <= 18000 and e.get('summary') == event['summary']:
                    matched = e
                    break

            if matched:
                service.events().update(calendarId=cfg.CALENDAR_ID, eventId=matched['id'], body=g_event).execute()
                stats["updated"] += 1
                log(f"🔁 Updated: {event['summary']} ({start} → {end})")
            else:
                service.events().insert(calendarId=cfg.CALENDAR_ID, body=g_event).execute()
                stats["created"] += 1
                log(f"✅ Created: {event['summary']} ({start} → {end})")
        except Exception as ex:
            stats["errors"] += 1
            log(f"⚠️ Error: {event['summary']} — {ex}")

    return stats


async def do_sync(cfg: Config, on_log=None) -> Dict[str, int]:
    log = on_log or (lambda *a, **k: None)
    log("Starting sync…")
    ics_files = await download_icals(cfg, on_log=log)
    all_events = []
    for path in ics_files:
        all_events.extend(parse_ical(path, on_log=log))
    log(f"Uploading {len(all_events)} events to Google Calendar…")
    stats = upload_events_to_gcal(cfg, all_events, on_log=log)
    log("Sync complete.")
    return stats


# ----------------------------
# Scheduling helpers
# ----------------------------

def is_windows() -> bool:
    return os.name == 'nt'

def create_windows_task(hour: int):
    task_name = APP_NAME
    python_exe = sys.executable
    script_path = os.path.abspath(__file__)

    cmd = f'"{python_exe}" "{script_path}" --sync'
    time_str = f"{hour:02d}:00"
    create_cmd = [
        "schtasks", "/Create", "/SC", "DAILY", "/TN", task_name, "/TR", cmd, "/ST", time_str, "/F",
    ]
    try:
        subprocess.check_call(create_cmd)
        return True, f"Scheduled daily sync at {time_str}."
    except subprocess.CalledProcessError as e:
        return False, f"Failed to create task: {e}"


def delete_windows_task():
    task_name = APP_NAME
    try:
        subprocess.check_call(["schtasks", "/Delete", "/TN", task_name, "/F"])
        return True, "Removed scheduled task."
    except subprocess.CalledProcessError as e:
        return False, f"Failed to delete task: {e}"


def suggest_unix_schedule(hour: int) -> str:
    script_path = os.path.abspath(__file__)
    return (
        "Cron suggestion:\n"
        f"0 {hour} * * * {sys.executable} {script_path} --sync >> $HOME/{APP_NAME}.log 2>&1\n"
        "macOS launchd (create ~/Library/LaunchAgents/com.user.hfgsync.plist) is also an option."
    )

# ----------------------------
# Autostart (Tray) registration helpers
# ----------------------------
def _windows_startup_folder() -> str:
    return os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs\Startup')

def _find_pythonw() -> str:
    # Prefer pythonw.exe to avoid console window on login
    exe = sys.executable
    #try pythonw.exe in the same dir
    if exe.lower().endswith("python.exe"):
        cand = exe[:-4] + "w.exe"
        if os.path.exists(cand):
            return cand
    cand = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.exists(cand):
        return cand
    return exe  # fallback

def _windows_startup_folder() -> str:
    return os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs\Startup')

def _find_pythonw() -> str:
    exe = sys.executable
    # Prefer pythonw.exe to avoid console window on login
    if exe.lower().endswith("python.exe"):
        cand = exe[:-4] + "w.exe"
        if os.path.exists(cand):
            return cand
    cand = os.path.join(os.path.dirname(exe), "pythonw.exe")
    if os.path.exists(cand):
        return cand
    return exe  # fallback to python.exe




# ----------------------------
# GUI (Tkinter)
# ----------------------------
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Humanforce → Google Calendar Sync")
        self.geometry("900x600")
        self.minsize(860, 560)

        ensure_dotenv_loaded()
        self.cfg = Config()

        self._init_style()

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.quick = ttk.Frame(nb)
        self.syncset = ttk.Frame(nb)
        self.history = ttk.Frame(nb)

        nb.add(self.quick, text="Quick Setup")
        nb.add(self.syncset, text="Sync & Settings")
        nb.add(self.history, text="History")

        self._build_quick()
        self._build_syncset()
        self._build_history()


    # ---------- style / dark mode ----------
    def _init_style(self):
        self.style = ttk.Style(self)
        bg_light = "#F5F7FB"
        bg_dark = "#1F2430"
        fg_dark = "#E7EAF0"

        # Try Azure theme if present
        try:
            self.tk.call("source", os.path.join(BASE_DIR, "azure.tcl"))
            self.style.theme_use("azure")
            if self.cfg.DARK_MODE:
                # Some Azure builds support azure-dark; fall back to manual tint if not.
                try:
                    self.style.theme_use("azure-dark")
                except Exception:
                    self.configure(bg=bg_dark)
        except Exception:
            # fall back built-in
            self.style.theme_use("clam")
            self.configure(bg=bg_dark if self.cfg.DARK_MODE else bg_light)

        # Basic tweaks
        self.style.configure("TButton", padding=8)
        self.style.configure("TEntry", padding=4)
        self.style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        self.style.configure("Subheader.TLabel", font=("Segoe UI", 11))

    # ---------- helpers ----------
    @staticmethod
    def _ts():
        return datetime.now().strftime("%H:%M:%S")

    def _log_file_and_gui(self, s: str):
        _append_log(s)
        self._log_out(s)

    def _log_out(self, msg: str):
        self.out.configure(state="normal")
        self.out.insert(tk.END, f"[{self._ts()}] {msg}\n")
        self.out.see(tk.END)
        self.out.configure(state="disabled")
        self.update_idletasks()

    def _log_quick(self, msg: str):
        self.quick_log.configure(state="normal")
        self.quick_log.insert(tk.END, f"[{self._ts()}] {msg}\n")
        self.quick_log.see(tk.END)
        self.quick_log.configure(state="disabled")
        self.update_idletasks()

    def _set_busy(self, busy: bool, text: str = ""):
        if busy:
            self.status_var.set(text or "Working…")
            self.pbar.start(10)
            for b in self.disable_while_busy:
                b.configure(state="disabled")
        else:
            self.status_var.set(text or "Idle")
            self.pbar.stop()
            for b in self.disable_while_busy:
                b.configure(state="normal")
        self.update_idletasks()

    # ---------- Quick Setup Tab ----------
    def _build_quick(self):
        frm = self.quick
        ttk.Label(frm, text="One‑click Install", style="Header.TLabel").pack(anchor="w", pady=(10,4))
        ttk.Label(frm, text=(
            "Prepare your environment: install deps, Playwright Chromium, and save .env."
        ), style="Subheader.TLabel").pack(anchor="w", pady=(0,8))

        row = ttk.Frame(frm)
        row.pack(anchor="w", pady=(0,10))
        self.btn_pick_creds = ttk.Button(row, text="Select credentials.json…", command=self._pick_creds)
        self.btn_install = ttk.Button(row, text="Run One‑click Install", command=self._one_click_install)
        self.btn_pick_creds.grid(row=0, column=0, padx=(0,10))
        self.btn_install.grid(row=0, column=1)

        # Optional deps toggles
        toggles = ttk.Frame(frm)
        toggles.pack(anchor="w", pady=(0,8))
        self.var_dark = tk.BooleanVar(value=self.cfg.DARK_MODE)
        ttk.Checkbutton(toggles, text="Enable Dark Mode", variable=self.var_dark, command=self._toggle_dark).grid(row=0, column=0, padx=(0,18))

        self.quick_log = tk.Text(frm, height=14, wrap="word", font=("Consolas", 10))
        self.quick_log.pack(fill=tk.BOTH, expand=True, pady=(8,0))
        self.quick_log.configure(state="disabled")

        bar = ttk.Frame(frm)
        bar.pack(fill=tk.X, pady=(6,4))
        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(bar, textvariable=self.status_var).pack(side=tk.LEFT)
        self.pbar = ttk.Progressbar(bar, mode="indeterminate", length=220)
        self.pbar.pack(side=tk.RIGHT)

    def _toggle_dark(self):
        self.cfg.DARK_MODE = self.var_dark.get()
        self.cfg.save()
        messagebox.showinfo("Theme", "Restart the app to fully apply the theme.")

    def _pick_creds(self):
        path = filedialog.askopenfilename(title="Pick Google OAuth credentials.json", filetypes=[("JSON","*.json"), ("All","*.*")])
        if path:
            try:
                with open(path, 'rb') as src, open(CRED_PATH, 'wb') as dst:
                    dst.write(src.read())
                messagebox.showinfo("Credentials", f"Copied to {CRED_PATH}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _one_click_install(self):
        def worker():
            try:
                self._set_busy(True, "Installing…")
                os.makedirs(self.cfg.DOWNLOAD_DIR, exist_ok=True)
                self._log_quick(f"Ensured downloads folder at {self.cfg.DOWNLOAD_DIR}")

                reqs = [
                    "icalendar>=5.0.0",
                    "google-api-python-client>=2.100.0",
                    "google-auth-oauthlib>=1.2.0",
                    "pytz",
                    "playwright>=1.38",
                    "python-dotenv>=1.0.0",
                ]

                self._log_quick("Installing Playwright Chromium…")
                subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"]) 
                self._log_quick("Playwright Chromium installed.")

                # Save UI toggles
                self.cfg.DARK_MODE = self.var_dark.get()
                self.cfg.save()
                self._log_quick("Saved settings to .env")


                messagebox.showinfo("Setup complete", "Environment prepared successfully.")
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Install error", f"A setup step failed: {e}")
            except Exception as ex:
                messagebox.showerror("Error", str(ex))
            finally:
                self._set_busy(False, "Idle")

        threading.Thread(target=worker, daemon=True).start()


    # ---------- Sync & Settings Tab ----------
    def _build_syncset(self):
        frm = self.syncset
        pad = {"padx": 6, "pady": 6}

        ttk.Label(frm, text="Sync Settings", style="Header.TLabel").grid(row=0, column=0, sticky="w", **pad)

        labels = [
            ("Humanforce URL", "HUMANFORCE_URL"),
            ("Username", "USERNAME"),
            ("Password", "PASSWORD"),
            ("Download folder", "DOWNLOAD_DIR"),
            ("Google Calendar ID", "CALENDAR_ID"),
            ("Timezone", "TIMEZONE"),
        ]
        self.vars = {}
        r = 1
        for label, key in labels:
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="e", **pad)
            var = tk.StringVar(value=getattr(self.cfg, key))
            entry = ttk.Entry(frm, textvariable=var, width=58, show="*" if key=="PASSWORD" else "")
            entry.grid(row=r, column=1, sticky="w", **pad)
            if key == "DOWNLOAD_DIR":
                ttk.Button(frm, text="…", command=self._pick_download_dir).grid(row=r, column=2, sticky="w")
            self.vars[key] = var
            r += 1

        # Scheduling row
        self.daily_enabled = tk.BooleanVar(value=self.cfg.DAILY_SYNC_ENABLED)
        self.daily_hour = tk.IntVar(value=self.cfg.DAILY_SYNC_HOUR)
        ttk.Checkbutton(frm, text="Run daily sync in background", variable=self.daily_enabled).grid(row=r, column=1, sticky="w", **pad)
        ttk.Label(frm, text="Hour (0-23):").grid(row=r, column=0, sticky="e", **pad)
        ttk.Spinbox(frm, from_=0, to=23, width=6, textvariable=self.daily_hour).grid(row=r, column=2, sticky="w", **pad)
        r += 1

        # Buttons
        btnrow = ttk.Frame(frm)
        btnrow.grid(row=r, column=0, columnspan=3, sticky="w", padx=6, pady=(10,6))
        self.btn_save = ttk.Button(btnrow, text="Save", command=self._save)
        self.btn_sync = ttk.Button(btnrow, text="Sync Now", command=self._sync_now)
        self.btn_test = ttk.Button(btnrow, text="Test Download", command=self._test_download)
        self.btn_install_task = ttk.Button(btnrow, text="Install/Update Background Task", command=self._install_task)
        self.btn_remove_task = ttk.Button(btnrow, text="Remove Background Task", command=self._remove_task)
        self.btn_open_log = ttk.Button(btnrow, text="Open Log", command=self._open_log)
        self.btn_clear = ttk.Button(btnrow, text="Clear Console", command=lambda: self._set_console(""))

        self.btn_save.grid(row=0, column=0, padx=(0,8))
        self.btn_sync.grid(row=0, column=1, padx=(0,8))
        self.btn_test.grid(row=0, column=2, padx=(0,8))
        self.btn_install_task.grid(row=0, column=3, padx=(0,8))
        self.btn_remove_task.grid(row=0, column=4, padx=(0,8))
        self.btn_open_log.grid(row=0, column=5, padx=(0,8))
        self.btn_clear.grid(row=0, column=6)

        self.disable_while_busy = [
            self.btn_save, self.btn_sync, self.btn_test, self.btn_install_task, self.btn_remove_task
        ]

        # Output console
        self.out = tk.Text(frm, height=14, wrap="word", font=("Consolas", 10))
        self.out.grid(row=r+1, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
        self.out.configure(state="disabled")
        frm.grid_rowconfigure(r+1, weight=1)
        frm.grid_columnconfigure(1, weight=1)

        # status + progress
        bar = ttk.Frame(frm)
        bar.grid(row=r+2, column=0, columnspan=3, sticky="ew", padx=6, pady=(0,6))
        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(bar, textvariable=self.status_var).pack(side=tk.LEFT)
        self.pbar = ttk.Progressbar(bar, mode="indeterminate", length=240)
        self.pbar.pack(side=tk.RIGHT)

    # ---------- History Tab ----------
    def _build_history(self):
        frm = self.history
        pad = {"padx": 6, "pady": 6}
        ttk.Label(frm, text="Recent Syncs", style="Header.TLabel").grid(row=0, column=0, sticky="w", **pad)
        self.tree = ttk.Treeview(frm, columns=("time","created","updated","errors","note"), show="headings", height=10)
        self.tree.heading("time", text="Time")
        self.tree.heading("created", text="Created")
        self.tree.heading("updated", text="Updated")
        self.tree.heading("errors", text="Errors")
        self.tree.heading("note", text="Note")
        self.tree.column("time", width=180)
        for c in ("created","updated","errors"):
            self.tree.column(c, width=90, anchor="center")
        self.tree.column("note", width=300)
        self.tree.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=6, pady=6)
        frm.grid_rowconfigure(1, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        # sparkline canvas
        self.spark = tk.Canvas(frm, height=60)
        self.spark.grid(row=2, column=0, columnspan=3, sticky="ew", padx=6, pady=(0,6))

        # controls
        row = ttk.Frame(frm)
        row.grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(0,6))
        ttk.Button(row, text="Refresh", command=self._refresh_history).grid(row=0, column=0, padx=(0,8))
        ttk.Button(row, text="Open Log", command=self._open_log).grid(row=0, column=1)

        self._refresh_history()

    def _refresh_history(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        hist = _load_history()
        for item in reversed(hist[-50:]):  # show last 50
            self.tree.insert('', 'end', values=(
                item.get('time',''), item.get('created',0), item.get('updated',0), item.get('errors',0), item.get('note','')
            ))
        self._draw_sparkline(hist)

    def _draw_sparkline(self, hist: List[Dict[str, Any]]):
        self.spark.delete("all")
        if not hist:
            return
        w = self.spark.winfo_width() or 600
        h = 60
        maxv = max((i.get('created',0)+i.get('updated',0)+i.get('errors',0)) for i in hist) or 1
        pts = []
        last = hist[-50:]
        for idx, i in enumerate(last):
            total = i.get('created',0)+i.get('updated',0)+i.get('errors',0)
            x = int(idx * (w/max(1, len(last)-1)))
            y = int(h - (total/maxv)* (h-10) - 5)
            pts.append((x,y))
        for a,b in zip(pts, pts[1:]):
            self.spark.create_line(a[0],a[1], b[0],b[1])

    # ---------- common actions ----------
    def _set_console(self, text: str):
        self.out.configure(state="normal")
        self.out.delete("1.0", tk.END)
        if text:
            self.out.insert(tk.END, text)
        self.out.configure(state="disabled")

    def _pick_download_dir(self):
        path = filedialog.askdirectory(title="Choose download folder")
        if path:
            self.vars["DOWNLOAD_DIR"].set(path)

    def _save(self, silent: bool = False):
        for key, var in self.vars.items():
            setattr(self.cfg, key, var.get())
        self.cfg.DAILY_SYNC_ENABLED = self.daily_enabled.get()
        self.cfg.DAILY_SYNC_HOUR = int(self.daily_hour.get())
        # ui toggles from quick tab (persist)
        self.cfg.DARK_MODE = self.var_dark.get()
        self.cfg.save()
        if not silent:
            messagebox.showinfo("Saved", "Settings saved to .env")

    def _open_log(self):
        try:
            if sys.platform.startswith('win'):
                os.startfile(LOG_PATH)  # type: ignore
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', LOG_PATH])
            else:
                subprocess.Popen(['xdg-open', LOG_PATH])
        except Exception as e:
            messagebox.showerror("Open Log", str(e))

    def _sync_now(self):
        self._save(silent=True)
        self._log_out("Sync requested…")
        def worker():
            try:
                self._set_busy(True, "Syncing…")
                stats = asyncio.run(do_sync(self.cfg, on_log=self._log_file_and_gui))
                note = "ok" if stats.get('errors',0)==0 else "with errors"
                self._record_history(stats, note)
                self._refresh_history()
            except Exception as e:
                self._record_history({"created":0,"updated":0,"errors":1}, f"exception: {e}")
                self._log_file_and_gui(f"Sync error: {e}")
                messagebox.showerror("Sync error", str(e))
            finally:
                self._set_busy(False, "Idle")
        threading.Thread(target=worker, daemon=True).start()

    def _test_download(self):
        self._save(silent=True)
        self._log_out("Test download requested…")
        def worker():
            try:
                self._set_busy(True, "Downloading…")
                files = asyncio.run(download_icals(self.cfg, on_log=self._log_file_and_gui))
                self._log_file_and_gui("Downloaded files: " + ", ".join(files))
            except Exception as e:
                self._log_file_and_gui(f"Download error: {e}")
                messagebox.showerror("Download error", str(e))
            finally:
                self._set_busy(False, "Idle")
        threading.Thread(target=worker, daemon=True).start()

    def _install_task(self):
        self._save(silent=True)
        if is_windows():
            self._log_file_and_gui("Installing/updating background task…")
            ok, msg = create_windows_task(self.cfg.DAILY_SYNC_HOUR)
            self._log_file_and_gui(msg)
            if ok and not self.cfg.DAILY_SYNC_ENABLED:
                delete_windows_task()
                self._log_file_and_gui("Disabled daily sync (task removed).")
            elif ok and self.cfg.DAILY_SYNC_ENABLED:
                messagebox.showinfo("Scheduled", msg)
        else:
            hint = suggest_unix_schedule(self.cfg.DAILY_SYNC_HOUR)
            messagebox.showinfo("How to schedule", hint)
            self._log_file_and_gui(hint)

    def _remove_task(self):
        if is_windows():
            self._log_file_and_gui("Removing background task…")
            ok, msg = delete_windows_task()
            self._log_file_and_gui(msg)
            if ok:
                messagebox.showinfo("Removed", msg)
        else:
            self._log_file_and_gui("On macOS/Linux, remove the corresponding cron/launchd entry.")

    def _record_history(self, stats: Dict[str,int], note: str = ""):
        hist = _load_history()
        hist.append({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created": int(stats.get("created",0)),
            "updated": int(stats.get("updated",0)),
            "errors": int(stats.get("errors",0)),
            "note": note,
        })
        _save_history(hist)

    # ---------- tray support ----------
   
# ----------------------------
# Tests (run with --selftest)
# ----------------------------
import unittest

class SelfTests(unittest.TestCase):
    def test_to_env_join_and_trailing_newline(self):
        os.environ["HUMANFORCE_URL"] = "https://example.com"
        os.environ["USERNAME"] = "u"
        os.environ["PASSWORD"] = "p"
        os.environ["DOWNLOAD_DIR"] = "/tmp/dl"
        os.environ["CALENDAR_ID"] = "cal"
        os.environ["TIMEZONE"] = "Australia/Melbourne"
        os.environ["DAILY_SYNC_ENABLED"] = "true"
        os.environ["DAILY_SYNC_HOUR"] = "6"
        os.environ["DARK_MODE"] = "true"
        os.environ["ENABLE_TRAY"] = "false"
        os.environ["START_MINIMIZED"] = "false"
        cfg = Config()
        env_str = cfg.to_env()
        self.assertTrue(env_str.endswith("\n"))
        for key in [
            "HUMANFORCE_URL","USERNAME","PASSWORD","DOWNLOAD_DIR","CALENDAR_ID","TIMEZONE",
            "DAILY_SYNC_ENABLED","DAILY_SYNC_HOUR","DARK_MODE","ENABLE_TRAY","START_MINIMIZED"
        ]:
            self.assertIn(key+"=", env_str)

    def test_unix_schedule_hint_contains_hour(self):
        hint = suggest_unix_schedule(5)
        self.assertIn("Cron suggestion:", hint)
        self.assertIn("0 5 * * *", hint)

    def test_history_roundtrip(self):
        items = [
            {"time":"2025-08-14 10:00:00","created":1,"updated":2,"errors":0,"note":"ok"},
            {"time":"2025-08-14 11:00:00","created":0,"updated":0,"errors":1,"note":"err"},
        ]
        _save_history(items)
        back = _load_history()
        self.assertGreaterEqual(len(back), 2)
        self.assertIn("time", back[-1])

    def test_is_valid_json_file_handles_empty_and_valid(self):
        # Create an empty temp file
        with tempfile.NamedTemporaryFile('w', delete=False) as tf:
            empty_path = tf.name
        try:
            self.assertFalse(_is_valid_json_file(empty_path))
        finally:
            os.remove(empty_path)
        # Create a valid JSON file
        with tempfile.NamedTemporaryFile('w', delete=False) as tf:
            tf.write('{}')
            valid_path = tf.name
        try:
            self.assertTrue(_is_valid_json_file(valid_path))
        finally:
            os.remove(valid_path)

# ----------------------------
# CLI entry
# ----------------------------

def main():
    ensure_dotenv_loaded()
    cfg = Config()

    parser = ArgumentParser()
    parser.add_argument("--sync", action="store_true", help="Run a headless sync and exit")
    parser.add_argument("--selftest", action="store_true", help="Run quick self tests and exit")
    args = parser.parse_args()

    if args.selftest:
        suite = unittest.defaultTestLoader.loadTestsFromTestCase(SelfTests)
        result = unittest.TextTestRunner(verbosity=2).run(suite)
        sys.exit(0 if result.wasSuccessful() else 1)

    if args.sync:
        def stdout_log(msg: str):
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            _append_log(msg)
        stats = asyncio.run(do_sync(cfg, on_log=stdout_log))
        # Record to history in headless mode too
        hist = _load_history()
        hist.append({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "created": int(stats.get("created",0)),
            "updated": int(stats.get("updated",0)),
            "errors": int(stats.get("errors",0)),
            "note": "headless",
        })
        _save_history(hist)
        return

    # Launch GUI
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
