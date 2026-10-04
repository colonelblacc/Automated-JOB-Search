#!/usr/bin/env python3
"""
Interactive Telegram Bot Controller & Mobile-to-PC Bridge for Ajith Shajan
Allows controlling job search, triggering scans, checking pipeline,
and saving mobile messages/links directly to PC via web dashboard or local file.

Pure Python standard library implementation (no heavy pip packages required).
"""

import os
import sys
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_FILE = os.path.join(WORKSPACE_DIR, ".env")
NOTES_FILE = os.path.join(WORKSPACE_DIR, "data", "mobile_notes.json")
NOTES_MD_FILE = os.path.join(WORKSPACE_DIR, "data", "MOBILE_NOTES.md")

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

if not BOT_TOKEN:
    print("❌ Error: TELEGRAM_BOT_TOKEN must be set.")
    sys.exit(1)

BASE_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"

# ==============================================================================
# NOTES STORAGE & SYNC ENGINE
# ==============================================================================

def load_notes():
    if os.path.exists(NOTES_FILE):
        try:
            with open(NOTES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_notes(notes):
    os.makedirs(os.path.dirname(NOTES_FILE), exist_ok=True)
    with open(NOTES_FILE, "w", encoding="utf-8") as f:
        json.dump(notes, f, indent=2, ensure_ascii=False)
    
    # Also generate human-readable markdown for immediate IDE / local reading
    try:
        lines = [
            "# 📱 Mobile-to-PC Saved Messages & Notes",
            f"*Last Updated: {datetime.now().strftime('%Y-%m-%d %I:%M %p IST')}*",
            "",
            f"Total Notes: **{len(notes)}**",
            "",
            "---",
            ""
        ]
        for n in notes:
            lines.append(f"### Note #{n.get('id', 0)} — {n.get('timestamp', '')}")
            lines.append(f"**From:** {n.get('sender', 'Telegram')}")
            lines.append("")
            lines.append(f"> {n.get('text', '').replace('\n', '\n> ')}")
            lines.append("")
            lines.append("---")
            lines.append("")
        with open(NOTES_MD_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception as e:
        print(f"Error writing markdown notes: {e}")

def add_note(text, sender="Telegram"):
    notes = load_notes()
    new_id = (max([n.get("id", 0) for n in notes]) + 1) if notes else 1
    now_str = datetime.now().strftime("%d %b %Y, %I:%M %p IST")
    note = {
        "id": new_id,
        "text": text.strip(),
        "sender": sender,
        "timestamp": now_str,
        "has_link": ("http://" in text or "https://" in text)
    }
    notes.insert(0, note) # Newest on top
    save_notes(notes)
    return note, len(notes)

def delete_note(note_id):
    notes = load_notes()
    orig_len = len(notes)
    notes = [n for n in notes if str(n.get("id")) != str(note_id)]
    save_notes(notes)
    return len(notes) < orig_len

def clear_all_notes():
    save_notes([])

# ==============================================================================
# TELEGRAM BOT API CALLS
# ==============================================================================

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

def get_web_url():
    # If custom domain or render URL
    render_service = os.environ.get("RENDER_EXTERNAL_HOSTNAME")
    if render_service:
        return f"https://{render_service}"
    return "https://automated-job-search.onrender.com"

# ==============================================================================
# TELEGRAM COMMAND HANDLERS
# ==============================================================================

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

def handle_save_note(chat_id, content, sender="Telegram"):
    if not content.strip():
        send_message(chat_id, "⚠️ <i>Please provide a message after /note. Example:</i>\n<code>/note Found opening at TI: https://ti.com/job/123</code>")
        return
    note, total = add_note(content, sender=sender)
    web_url = get_web_url()
    reply = (
        f"📝 <b>Note #{note['id']} Saved to PC Hub!</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n"
        f"💬 <i>\"{note['text'][:300]}{'...' if len(note['text']) > 300 else ''}\"</i>\n\n"
        f"💻 <b>Access on Computer:</b>\n"
        f"👉 <a href=\"{web_url}/notes\"><b>Open Live Notes Hub ({web_url}/notes)</b></a>\n"
        f"📊 Total saved: <b>{total}</b> notes"
    )
    send_message(chat_id, reply)

def handle_list_notes(chat_id):
    notes = load_notes()
    if not notes:
        send_message(chat_id, "📭 <i>No notes saved yet. Use /note &lt;text&gt; to save messages to your PC!</i>")
        return
    
    web_url = get_web_url()
    lines = [
        f"📋 <b>MOBILE-TO-PC NOTES ({len(notes)} Total):</b>",
        "─────────────────────────"
    ]
    for n in notes[:10]:
        t = n.get("text", "")
        preview = t[:90] + ("..." if len(t) > 90 else "")
        lines.append(f"• <b>#{n.get('id')}</b> [{n.get('timestamp')}]:\n   {preview}")
        lines.append("")
    
    lines.append(f"🌐 <a href=\"{web_url}/notes\"><b>View & Copy all on Computer ({web_url}/notes)</b></a>")
    lines.append("🗑 <i>Use /delnote &lt;id&gt; to delete a note</i>")
    send_message(chat_id, "\n".join(lines))

def handle_delete_note(chat_id, args):
    if not args.strip():
        send_message(chat_id, "⚠️ <i>Specify note ID to delete. Example:</i> <code>/delnote 2</code>")
        return
    note_id = args.strip()
    if delete_note(note_id):
        send_message(chat_id, f"🗑 <b>Note #{note_id} deleted successfully.</b>")
    else:
        send_message(chat_id, f"⚠️ Note #{note_id} not found.")

def handle_help(chat_id):
    web_url = get_web_url()
    text = (
        "🤖 <b>HARDWARE JOB SEARCH & PC BRIDGE CONTROLLER</b>\n\n"
        "<b>Job Radar Commands:</b>\n"
        "• <b>/jobs</b> or <b>/radar</b> — Top 90%+ match 0-2y hardware openings\n"
        "• <b>/status</b> — Application pipeline stats\n"
        "• <b>/scan</b> — Trigger full scan & Google Sheets sync\n"
        "• <b>/sheet</b> — Link to live Google Sheets tracker\n\n"
        "<b>📱 Mobile-to-PC Notes Commands:</b>\n"
        "• <b>/note &lt;text or link&gt;</b> — Save message/link to your computer\n"
        "• <b>/notes</b> — List your saved messages\n"
        "• <b>/delnote &lt;id&gt;</b> — Delete a saved note\n"
        "• <i>(In direct DM, any plain text or link sent is automatically saved!)</i>\n\n"
        f"💻 <b>Live PC Hub:</b> <a href=\"{web_url}/notes\">{web_url}/notes</a>"
    )
    send_message(chat_id, text)

# ==============================================================================
# INCOMING MESSAGE DISPATCHER
# ==============================================================================

def process_update(update):
    msg = update.get("message") or update.get("channel_post")
    if not msg:
        return

    chat = msg.get("chat", {})
    chat_id = str(chat.get("id"))
    chat_type = chat.get("type", "unknown")
    text = (msg.get("text") or "").strip()
    from_user = msg.get("from", {})
    sender_name = from_user.get("first_name") or from_user.get("username") or "Mobile"

    print(f"📥 Received message: {repr(text)} from chat_id={chat_id} (type={chat_type})")

    # Security check: allow if sent in configured group/channel OR in a private direct message
    is_group_match = (
        ALLOWED_CHAT_ID and (
            chat_id == str(ALLOWED_CHAT_ID) or 
            str(ALLOWED_CHAT_ID).lstrip("-100") in chat_id or
            chat_id.lstrip("-100") == str(ALLOWED_CHAT_ID).lstrip("-100")
        )
    )
    is_private = (chat_type == "private")

    if not (is_group_match or is_private):
        print(f"⛔ Ignoring command from unauthorized chat_id={chat_id} (configured={ALLOWED_CHAT_ID})")
        return

    if not text:
        return

    tokens = text.split(maxsplit=1)
    raw_cmd = tokens[0].lower()
    cmd_arg = tokens[1] if len(tokens) > 1 else ""

    # Strip bot handle if sent in group e.g. /jobs@Ajith_Job_Radar_Bot
    if "@" in raw_cmd:
        cmd = raw_cmd.split("@")[0]
    else:
        cmd = raw_cmd

    print(f"⚡ Processing command: '{cmd}' for chat {chat_id}")

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
    elif cmd in ["/note", "/save", "/memo", "/add", "/link"]:
        handle_save_note(chat_id, cmd_arg, sender=sender_name)
    elif cmd in ["/notes", "/list"]:
        handle_list_notes(chat_id)
    elif cmd in ["/delnote", "/rmnote", "/delete"]:
        handle_delete_note(chat_id, cmd_arg)
    elif cmd in ["/clearnotes"]:
        clear_all_notes()
        send_message(chat_id, "🗑 <i>All saved notes cleared.</i>")
    elif is_private and not text.startswith("/"):
        # Convenience feature: In 1-on-1 private chat with the bot,
        # any message or link sent directly is automatically saved as a note!
        handle_save_note(chat_id, text, sender=sender_name)
    else:
        # Unknown group command
        pass

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

# ==============================================================================
# WEB SERVER & DESKTOP DASHBOARD (RENDER + PC ACCESS)
# ==============================================================================

def render_notes_html():
    notes = load_notes()
    cards_html = ""
    if not notes:
        cards_html = """
        <div class="empty-state">
            <div class="empty-icon">📭</div>
            <h3>No Messages or Notes Yet</h3>
            <p>Send <code>/note your message</code> to your Telegram bot, or simply text any link or note in your private bot chat from your phone.</p>
        </div>
        """
    else:
        import html, re
        for n in notes:
            nid = n.get("id", 0)
            raw_text = n.get("text", "")
            escaped_text = html.escape(raw_text)
            # Make URLs clickable
            linked_text = re.sub(
                r'(https?://[^\s<>"]+)',
                r'<a href="\1" target="_blank" class="text-link">\1 ↗</a>',
                escaped_text
            )
            # Preserve newlines
            formatted_text = linked_text.replace("\n", "<br>")
            ts = n.get("timestamp", "")
            sender = html.escape(n.get("sender", "Mobile"))

            cards_html += f"""
            <div class="note-card" id="card-{nid}">
                <div class="card-header">
                    <span class="badge">Note #{nid}</span>
                    <span class="timestamp">{ts} • {sender}</span>
                    <div class="card-actions">
                        <button class="btn-copy" onclick="copyNote('{nid}')">📋 Copy</button>
                        <a href="/delete?id={nid}" class="btn-del" onclick="return confirm('Delete Note #{nid}?')">🗑 Delete</a>
                    </div>
                </div>
                <div class="note-body" id="text-{nid}">{formatted_text}</div>
            </div>
            """

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ajith's Mobile-to-PC Notes Hub</title>
    <style>
        :root {{
            --bg: #090D16;
            --surface: #131B2E;
            --surface-hover: #1E293B;
            --border: #263352;
            --primary: #38BDF8;
            --primary-hover: #0284C7;
            --accent: #10B981;
            --text-main: #F1F5F9;
            --text-muted: #94A3B8;
            --danger: #EF4444;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }}
        body {{
            background: var(--bg);
            color: var(--text-main);
            padding: 24px 16px;
            display: flex;
            justify-content: center;
        }}
        .container {{
            width: 100%;
            max-width: 860px;
        }}
        .header {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }}
        .title h1 {{
            font-size: 22px;
            color: #FFFFFF;
            font-weight: 700;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .title p {{
            font-size: 13px;
            color: var(--text-muted);
        }}
        .header-actions {{
            display: flex;
            gap: 10px;
            align-items: center;
        }}
        .btn {{
            padding: 8px 16px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            text-decoration: none;
            cursor: pointer;
            border: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
        }}
        .btn-primary {{
            background: var(--primary);
            color: #0F172A;
        }}
        .btn-primary:hover {{
            background: var(--primary-hover);
            color: #FFFFFF;
        }}
        .btn-outline {{
            background: transparent;
            border: 1px solid var(--border);
            color: var(--text-main);
        }}
        .btn-outline:hover {{
            background: var(--surface-hover);
        }}
        .add-box {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 24px;
        }}
        .add-form {{
            display: flex;
            gap: 12px;
        }}
        .add-input {{
            flex: 1;
            background: #0B1120;
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 10px 14px;
            color: #FFFFFF;
            font-size: 14px;
        }}
        .add-input:focus {{
            outline: none;
            border-color: var(--primary);
        }}
        .note-card {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 18px 20px;
            margin-bottom: 14px;
            transition: transform 0.15s ease, border-color 0.15s ease;
        }}
        .note-card:hover {{
            border-color: #38BDF8;
            transform: translateY(-1px);
        }}
        .card-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
            gap: 8px;
        }}
        .badge {{
            background: #0284C7;
            color: #FFFFFF;
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 6px;
            text-transform: uppercase;
        }}
        .timestamp {{
            font-size: 12px;
            color: var(--text-muted);
            flex: 1;
            margin-left: 10px;
        }}
        .card-actions {{
            display: flex;
            gap: 8px;
        }}
        .btn-copy {{
            background: #1E293B;
            border: 1px solid var(--border);
            color: var(--primary);
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
        }}
        .btn-copy:hover {{
            background: #38BDF8;
            color: #090D16;
        }}
        .btn-del {{
            background: transparent;
            color: var(--danger);
            text-decoration: none;
            font-size: 12px;
            padding: 4px 8px;
            border-radius: 6px;
            border: 1px solid transparent;
        }}
        .btn-del:hover {{
            background: #EF444422;
            border-color: var(--danger);
        }}
        .note-body {{
            font-size: 14px;
            line-height: 1.6;
            color: #E2E8F0;
            word-break: break-word;
        }}
        .text-link {{
            color: var(--primary);
            text-decoration: underline;
            word-break: break-all;
        }}
        .empty-state {{
            background: var(--surface);
            border: 1px dashed var(--border);
            border-radius: 12px;
            padding: 48px 24px;
            text-align: center;
        }}
        .empty-icon {{ font-size: 40px; margin-bottom: 12px; }}
        .empty-state h3 {{ font-size: 17px; margin-bottom: 8px; }}
        .empty-state p {{ color: var(--text-muted); font-size: 13px; max-width: 480px; margin: 0 auto; line-height: 1.5; }}
        .footer {{
            text-align: center;
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 32px;
        }}
    </style>
    <script>
        function copyNote(nid) {{
            const el = document.getElementById('text-' + nid);
            const text = el.innerText || el.textContent;
            navigator.clipboard.writeText(text).then(() => {{
                const btn = document.querySelector('#card-' + nid + ' .btn-copy');
                const orig = btn.innerText;
                btn.innerText = '✓ Copied!';
                btn.style.background = '#10B981';
                btn.style.color = '#FFFFFF';
                setTimeout(() => {{
                    btn.innerText = orig;
                    btn.style.background = '#1E293B';
                    btn.style.color = 'var(--primary)';
                }}, 2000);
            }});
        }}
    </script>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="title">
                <h1>📱 Ajith's Mobile-to-PC Hub</h1>
                <p>Synced directly from your Telegram Bot • Real-time cloud access</p>
            </div>
            <div class="header-actions">
                <a href="/notes" class="btn btn-outline">🔄 Refresh</a>
                <a href="https://docs.google.com/spreadsheets/d/1MfgEtHPuEx3DvITdxzjfbl2GvmmipOrc5oKBzjZRECY" target="_blank" class="btn btn-primary">📊 Open Google Sheet</a>
            </div>
        </div>

        <div class="add-box">
            <form action="/add" method="GET" class="add-form">
                <input type="text" name="text" class="add-input" placeholder="Quick paste a note, link, or job lead from PC..." required autofocus>
                <button type="submit" class="btn btn-primary">➕ Save</button>
            </form>
        </div>

        <div class="notes-feed">
            {cards_html}
        </div>

        <div class="footer">
            Automated Job Search & Telemetry Bot • Connected to Telegram & Render
        </div>
    </div>
</body>
</html>
"""

class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path in ["/notes", "/view", "/hub"]:
            html_content = render_notes_html()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(html_content.encode("utf-8"))

        elif path in ["/api/notes", "/notes.json"]:
            notes = load_notes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"count": len(notes), "notes": notes}, indent=2).encode("utf-8"))

        elif path == "/add":
            text = query.get("text", [""])[0]
            if text:
                add_note(text, sender="Web/PC")
            self.send_response(302)
            self.send_header("Location", "/notes")
            self.end_headers()

        elif path == "/delete":
            note_id = query.get("id", [""])[0]
            if note_id:
                delete_note(note_id)
            self.send_response(302)
            self.send_header("Location", "/notes")
            self.end_headers()

        else:
            # Standard health check endpoint
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            notes = load_notes()
            response = json.dumps({
                "status": "healthy",
                "service": "telegram-job-bot",
                "total_notes": len(notes),
                "notes_hub": f"{get_web_url()}/notes",
                "timestamp": datetime.now().isoformat()
            }).encode("utf-8")
            self.wfile.write(response)

    def log_message(self, format, *args):
        # Suppress routine health-check logs
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    print(f"🌐 Notes Hub & Health Server active on port {port} (Access at /notes)")
    server.serve_forever()

if __name__ == "__main__":
    if "PORT" in os.environ:
        t = threading.Thread(target=start_health_server, daemon=True)
        t.start()

    if len(sys.argv) > 1 and sys.argv[1] == "--single-poll":
        updates = api_call("getUpdates", {"timeout": 5})
        if updates and updates.get("ok"):
            for u in updates.get("result", []):
                process_update(u)
    else:
        run_bot_polling()
