#!/usr/bin/env python3
"""
Telegram Job Radar & Pipeline Notifier for Ajith Shajan
Pushes daily morning alerts with newly discovered 0-2y embedded/hardware jobs
and pipeline tracking stats directly to Telegram.
"""

import os
import sys
import json
import urllib.request
import urllib.parse
from datetime import datetime
import openpyxl

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_FILE = os.path.join(WORKSPACE_DIR, "AI_Embedded_Job_Hunt.xlsx")
URL_CACHE_FILE = os.path.join(WORKSPACE_DIR, "google_sheet_url.txt")
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

def get_google_sheet_url():
    env_url = os.environ.get("GOOGLE_SHEET_URL")
    if env_url:
        return env_url
    if os.path.exists(URL_CACHE_FILE):
        with open(URL_CACHE_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "https://docs.google.com/spreadsheets/d/1MfgEtHPuEx3DvITdxzjfbl2GvmmipOrc5oKBzjZRECY"

def get_job_and_pipeline_data():
    if not os.path.exists(XLSX_FILE):
        return [], {}

    wb = openpyxl.load_workbook(XLSX_FILE, data_only=True)
    
    # 1. Parse DASHBOARD & NEW JOBS
    new_jobs = []
    if "DASHBOARD & NEW JOBS" in wb.sheetnames:
        ws = wb["DASHBOARD & NEW JOBS"]
        # Find headers on row 5
        headers = [cell.value for cell in ws[5]]
        
        def col_idx(name):
            for i, h in enumerate(headers):
                if h and name.lower() in str(h).lower():
                    return i
            return None

        c_company = col_idx("company")
        c_title = col_idx("job title")
        c_score = col_idx("relevance score") or col_idx("score")
        c_fit = col_idx("why fit") or col_idx("why it matches")
        c_loc = col_idx("location")
        c_mode = col_idx("work mode")
        c_status = col_idx("job status")
        c_exp = col_idx("experience") or col_idx("exp")
        c_url = col_idx("authoritative job url") or col_idx("apply link") or col_idx("job url")
        c_company_url = col_idx("company url")

        for row in ws.iter_rows(min_row=6, values_only=True):
            if not row or not any(row):
                continue
            company = row[c_company] if c_company is not None and c_company < len(row) else None
            title = row[c_title] if c_title is not None and c_title < len(row) else None
            if not company or not title:
                continue

            status = str(row[c_status] or "").strip().upper()
            if status in ["APPLIED", "CLOSED", "REJECTED", "DUPLICATE", "NOT ELIGIBLE"]:
                continue

            score = row[c_score] if c_score is not None and c_score < len(row) else 0
            try:
                score_num = float(str(score).replace("%", "").strip())
            except Exception:
                score_num = 0

            url = row[c_url] if c_url is not None and c_url < len(row) else ""
            if not url and c_company_url is not None and c_company_url < len(row):
                url = row[c_company_url]
            url_str = str(url or "").strip()
            if "HYPERLINK(" in url_str:
                import re
                match = re.search(r'https?://[^",\s)]+', url_str)
                if match:
                    url_str = match.group(0)

            loc = row[c_loc] if c_loc is not None and c_loc < len(row) else "India"
            mode = row[c_mode] if c_mode is not None and c_mode < len(row) else "On-site"
            fit = row[c_fit] if c_fit is not None and c_fit < len(row) else ""

            new_jobs.append({
                "company": str(company).strip(),
                "title": str(title).strip(),
                "score": score_num,
                "location": str(loc).strip(),
                "mode": str(mode).strip(),
                "fit": str(fit).strip(),
                "url": url_str
            })

    # Sort descending by match score
    new_jobs.sort(key=lambda x: x["score"], reverse=True)

    # 2. Parse APPLICATIONS
    app_stats = {"total": 0, "under_review": 0, "interview": 0, "rejected": 0, "offer": 0}
    active_interviews = []
    if "APPLICATIONS" in wb.sheetnames:
        ws_app = wb["APPLICATIONS"]
        headers_app = [str(cell.value or "").strip().lower() for cell in ws_app[1]]
        def app_col(n):
            for i, h in enumerate(headers_app):
                if n in h:
                    return i
            return None

        c_app_comp = app_col("company")
        c_app_role = app_col("role") or app_col("job title")
        c_app_stat = app_col("status")
        c_app_stage = app_col("detail") or app_col("stage")
        c_app_next = app_col("next") or app_col("follow")
        c_app_notes = app_col("notes")

        for row in ws_app.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            st = str(row[c_app_stat] or "").strip().upper() if c_app_stat is not None else ""
            if not st:
                continue
            app_stats["total"] += 1
            if "INTERVIEW" in st:
                app_stats["interview"] += 1
                comp = str(row[c_app_comp] or "Company") if c_app_comp is not None else "Company"
                role = str(row[c_app_role] or "Role") if c_app_role is not None else "Role"
                stage = str(row[c_app_stage] or "Interview") if c_app_stage is not None else "Interview"
                next_act = str(row[c_app_next] or "") if c_app_next is not None else ""
                notes = str(row[c_app_notes] or "") if c_app_notes is not None else ""
                active_interviews.append({
                    "company": comp,
                    "role": role,
                    "stage": stage,
                    "next_action": next_act,
                    "notes": notes
                })
            elif "REJECT" in st:
                app_stats["rejected"] += 1
            elif "OFFER" in st:
                app_stats["offer"] += 1
            else:
                app_stats["under_review"] += 1

    return new_jobs, app_stats, active_interviews

def build_telegram_html(new_jobs, app_stats, max_jobs=15, active_interviews=None):
    now_str = datetime.now().strftime("%d %b %Y | %I:%M %p IST")
    sheet_url = get_google_sheet_url()

    lines = [
        "🌅 <b>MORNING EMBEDDED JOB RADAR</b>",
        f"📅 <i>{now_str}</i>",
        "─────────────────────────",
        f"⚡ <b>VERIFIED 0–2Y OPENINGS ({len(new_jobs)} Active):</b>\n"
    ]

    top_jobs = new_jobs[:max_jobs] if max_jobs else new_jobs
    if not top_jobs:
        lines.append("<i>No new unapplied openings today. Portals scanned & healthy!</i>\n")
    else:
        for idx, job in enumerate(top_jobs, 1):
            badge = "🔥" if job["score"] >= 95 else "⚡"
            lines.append(f"<b>{idx}. {job['company']}</b> — {job['title']}")
            lines.append(f"   {badge} <b>{int(job['score'])}% Match</b> | 📍 {job['location']} ({job['mode']})")
            if job['fit']:
                lines.append(f"   🎯 <i>{job['fit']}</i>")
            
            url = job['url']
            if url and url.startswith("http"):
                lines.append(f"   👉 <a href=\"{url}\"><b>View & Apply Directly</b></a>\n")
            else:
                lines.append(f"   👉 <i>Check Company Career Portal</i>\n")

    if active_interviews:
        lines.append("─────────────────────────")
        lines.append(f"🎯 <b>ACTIVE INTERVIEW PIPELINE ({len(active_interviews)} In Progress):</b>")
        for act in active_interviews:
            lines.append(f"• <b>{act['company']}</b> — {act['role']}")
            lines.append(f"   📌 <i>Stage:</i> {act['stage']}")
            if act['next_action'] and act['next_action'] != "-":
                lines.append(f"   ⏰ <b>Reminder:</b> Follow up / mail on <b>Tuesday ({act['next_action']})</b> if no response received.")
            lines.append("")

    lines.append("─────────────────────────")
    lines.append("📊 <b>APPLICATION PIPELINE:</b>")
    lines.append(f"• <b>Total Applied:</b> {app_stats.get('total', 0)}")
    lines.append(f"• <b>Under Review:</b> {app_stats.get('under_review', 0)}")
    lines.append(f"• <b>Active Interviews:</b> {app_stats.get('interview', 0)}")
    lines.append(f"• <b>Past Rejections:</b> {app_stats.get('rejected', 0)}")
    lines.append("")
    lines.append(f"📑 <a href=\"{sheet_url}\"><b>Open Google Sheets Command Center</b></a>")

    return "\n".join(lines)

def send_telegram_message(html_text: str):
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if not bot_token or not chat_id:
        print("\n" + "!" * 70)
        print("TELEGRAM CREDENTIALS NOT FOUND")
        print("!" * 70)
        print("Set TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID as environment variables,")
        print("or add them to your local .env file.")
        print("-" * 70)
        print("MESSAGE PREVIEW (HTML):")
        print("-" * 70)
        print(html_text)
        print("!" * 70 + "\n")
        return False

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"

    # Split message into chunks <= 4000 characters if needed
    chunks = []
    if len(html_text) <= 4000:
        chunks = [html_text]
    else:
        # Split by separator or double newlines
        parts = html_text.split("\n\n")
        cur_chunk = ""
        for p in parts:
            if len(cur_chunk) + len(p) + 2 < 3900:
                cur_chunk += ("\n\n" if cur_chunk else "") + p
            else:
                if cur_chunk:
                    chunks.append(cur_chunk)
                cur_chunk = p
        if cur_chunk:
            chunks.append(cur_chunk)

    all_ok = True
    for chunk in chunks:
        payload = {
            "chat_id": chat_id,
            "text": chunk,
            "parse_mode": "HTML",
            "disable_web_page_preview": True
        }
        try:
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                url, 
                data=data, 
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                res_json = json.loads(resp.read().decode("utf-8"))
                if not res_json.get("ok"):
                    print(f"❌ Telegram API Error: {res_json}")
                    all_ok = False
        except Exception as e:
            print(f"❌ Failed to send Telegram alert: {e}")
            all_ok = False

    if all_ok:
        print("✓ Successfully sent all active openings to Telegram!")
    return all_ok

if __name__ == "__main__":
    jobs, stats, active = get_job_and_pipeline_data()
    msg = build_telegram_html(jobs, stats, max_jobs=15, active_interviews=active)
    
    if "--preview" in sys.argv or "--dry-run" in sys.argv:
        print("=" * 60)
        print("TELEGRAM MESSAGE PREVIEW")
        print("=" * 60)
        print(msg)
        print("=" * 60)
    else:
        send_telegram_message(msg)
