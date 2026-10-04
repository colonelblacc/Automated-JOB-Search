#!/usr/bin/env python3
"""
Interactive Telegram Bot Controller for Ajith Shajan
Allows controlling job search, triggering scans, checking pipeline,
and marking applications directly from Telegram.

Pure Python standard library implementation (no heavy pip packages required).
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_FILE = os.path.join(WORKSPACE_DIR, ".env")

def load_env_file():
    if os.path.exists(ENV_FILE):
        try:
            with open(ENV_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
        except Exception:
            pass

load_env_file()

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
ALLOWED_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

if not BOT_TOKEN or not ALLOWED_CHAT_ID:
    print("❌ Error: TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID must be set.")
    sys.exit(1)

BASE_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"

def api_call(endpoint, payload=None):
    url = f"{BASE_URL}/{endpoint}"
    data = None
    headers = {}
    if payload:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"API Error ({endpoint}): {e}")
        return None

def send_message(chat_id, text, parse_mode="HTML"):
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": parse_mode,
        "disable_web_page_preview": True
    }
    return api_call("sendMessage", payload)

def handle_jobs(chat_id):
    from telegram_notifier import get_job_and_pipeline_data, build_telegram_html
    jobs, stats, active = get_job_and_pipeline_data()
    msg = build_telegram_html(jobs, stats, max_jobs=15, active_interviews=active)
    send_message(chat_id, msg)

def handle_status(chat_id):
    from telegram_notifier import get_job_and_pipeline_data, get_google_sheet_url
    _, stats, active = get_job_and_pipeline_data()
    sheet_url = get_google_sheet_url()
    text = (
        "📊 <b>APPLICATION PIPELINE STATUS:</b>\n"
        f"• <b>Total Applied:</b> {stats.get('total', 0)}\n"
        f"• <b>Under Review:</b> {stats.get('under_review', 0)}\n"
        f"• <b>Active Interviews:</b> {stats.get('interview', 0)}\n"
        f"• <b>Past Rejections:</b> {stats.get('rejected', 0)}\n\n"
        f"📑 <a href=\"{sheet_url}\"><b>Open Google Sheets Tracker</b></a>"
    )
    send_message(chat_id, text)

def handle_scan(chat_id):
    send_message(chat_id, "⏳ <i>Starting scan, matching 17 role families & updating Google Sheets...</i>")
    try:
        import subprocess
        # 1. Relevance Engine
        subprocess.run([sys.executable, os.path.join(WORKSPACE_DIR, "scripts", "relevance_engine.py")], check=True)
        # 2. Excel Tracker
        subprocess.run([sys.executable, os.path.join(WORKSPACE_DIR, "scripts", "generate_excel_tracker.py")], check=True)
        # 3. Google Sheets Sync
        subprocess.run([sys.executable, os.path.join(WORKSPACE_DIR, "scripts", "sync_to_google_sheets.py")], check=True)
        
        from telegram_notifier import get_job_and_pipeline_data, build_telegram_html
        jobs, stats = get_job_and_pipeline_data()
        msg = "✅ <b>SCAN & SYNC COMPLETE!</b>\n\n" + build_telegram_html(jobs, stats, max_jobs=5)
        send_message(chat_id, msg)
    except Exception as e:
        send_message(chat_id, f"❌ Scan failed: {e}")

def handle_help(chat_id):
    text = (
        "🤖 <b>HARDWARE JOB SEARCH CONTROLLER</b>\n\n"
        "Available Commands:\n"
        "• <b>/jobs</b> or <b>/radar</b> — Show top 90%+ match 0-2y hardware openings with direct links\n"
        "• <b>/status</b> — Pipeline status (applied, under review, interview)\n"
        "• <b>/scan</b> — Trigger full scan & sync directly to Google Sheets\n"
        "• <b>/sheet</b> — Get link to live Google Sheets Command Center\n"
        "• <b>/help</b> — Show this guide"
    )
    send_message(chat_id, text)

def process_update(update):
    msg = update.get("message") or update.get("channel_post")
    if not msg:
        return

    chat = msg.get("chat", {})
    chat_id = str(chat.get("id"))
    text = (msg.get("text") or "").strip()

    # Security check: only allow authorized chat or group
    if chat_id != str(ALLOWED_CHAT_ID) and str(ALLOWED_CHAT_ID).lstrip("-100") not in chat_id:
        return

    cmd = text.split()[0].lower() if text else ""
    # Strip bot handle if sent in group e.g. /jobs@Ajith_Job_Radar_Bot
    if "@" in cmd:
        cmd = cmd.split("@")[0]

    if cmd in ["/start", "/help"]:
        handle_help(chat_id)
    elif cmd in ["/jobs", "/radar", "/new"]:
        handle_jobs(chat_id)
    elif cmd in ["/status", "/pipeline"]:
        handle_status(chat_id)
    elif cmd in ["/scan", "/refresh", "/sync"]:
        handle_scan(chat_id)
    elif cmd in ["/sheet", "/sheets", "/url"]:
        from telegram_notifier import get_google_sheet_url
        send_message(chat_id, f"📑 <a href=\"{get_google_sheet_url()}\"><b>Open Google Sheets Command Center</b></a>")

def run_bot_polling():
    print(f"🤖 Telegram Bot Controller is running (polling for chat ID {ALLOWED_CHAT_ID})...")
    offset = 0
    while True:
        try:
            updates = api_call("getUpdates", {"offset": offset, "timeout": 20})
            if updates and updates.get("ok"):
                for u in updates.get("result", []):
                    offset = u["update_id"] + 1
                    process_update(u)
        except KeyboardInterrupt:
            print("\nShutting down bot...")
            break
        except Exception as e:
            print(f"Polling loop error: {e}")
            time.sleep(3)
        time.sleep(1)

import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        response = json.dumps({
            "status": "healthy",
            "service": "telegram-job-bot",
            "timestamp": datetime.now().isoformat()
        }).encode("utf-8")
        self.wfile.write(response)

    def log_message(self, format, *args):
        # Suppress noisy health-check access logs
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    print(f"🌐 Health check HTTP server active on port {port} (Render Web Service ready)")
    server.serve_forever()

if __name__ == "__main__":
    # Start HTTP server on background thread if running as a web service
    if "PORT" in os.environ:
        t = threading.Thread(target=start_health_server, daemon=True)
        t.start()

    if len(sys.argv) > 1 and sys.argv[1] == "--single-poll":
        # Process pending updates once and exit
        updates = api_call("getUpdates", {"timeout": 5})
        if updates and updates.get("ok"):
            for u in updates.get("result", []):
                process_update(u)
    else:
        run_bot_polling()

