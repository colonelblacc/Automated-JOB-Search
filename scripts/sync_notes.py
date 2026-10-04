#!/usr/bin/env python3
"""
Sync Mobile Notes from Render Cloud to Local PC
Usage:
    python scripts/sync_notes.py
"""

import os
import sys
import json
import urllib.request
from datetime import datetime

# Ensure utf-8 output on Windows console
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

REMOTE_API = os.environ.get("RENDER_EXTERNAL_URL", "https://automated-job-search.onrender.com") + "/api/notes"
LOCAL_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
LOCAL_MD = os.path.join(LOCAL_DATA_DIR, "MOBILE_NOTES.md")
LOCAL_JSON = os.path.join(LOCAL_DATA_DIR, "mobile_notes.json")

def sync():
    os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
    print(f"📥 Connecting to Cloud Notes Hub: {REMOTE_API} ...")

    req = urllib.request.Request(
        REMOTE_API,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) MobileNotesSync/1.0"}
    )

    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            if res.status != 200:
                print(f"❌ Server returned HTTP {res.status}")
                return
            data = json.loads(res.read().decode("utf-8"))
            notes = data.get("notes", [])
    except Exception as e:
        print(f"⚠️ Could not connect to cloud notes ({e}). Checking local notes file...")
        if os.path.exists(LOCAL_JSON):
            with open(LOCAL_JSON, "r", encoding="utf-8") as f:
                notes = json.load(f)
        else:
            notes = []

    print(f"✅ Found {len(notes)} note(s). Updating local files...")

    # Write local JSON
    with open(LOCAL_JSON, "w", encoding="utf-8") as f:
        json.dump(notes, f, indent=2, ensure_ascii=False)

    # Format Markdown
    md_lines = [
        "# 📱 Mobile Notes Hub (Telegram -> PC)",
        f"> Last Synced: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} • Source: [Online Hub](https://automated-job-search.onrender.com/notes)",
        "",
        "---",
        ""
    ]

    if not notes:
        md_lines.append("_No notes saved yet. Send any text or link to your Telegram bot on your phone!_")
    else:
        for idx, item in enumerate(notes, 1):
            ts = item.get("timestamp", "Unknown time")
            sender = item.get("sender", "Mobile")
            text = item.get("text", "")
            nid = item.get("id", str(idx))
            
            md_lines.append(f"### Note #{nid} · `{sender}` ({ts})")
            md_lines.append(f"{text}")
            md_lines.append("")

    with open(LOCAL_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    print(f"📄 Local Markdown updated: {LOCAL_MD}")
    print("\n--- RECENT NOTES ---")
    if not notes:
        print("No notes found.")
    else:
        for n in notes[:5]:
            print(f"[{n.get('timestamp')}] {n.get('text')[:80]}")

if __name__ == "__main__":
    sync()
