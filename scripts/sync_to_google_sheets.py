#!/usr/bin/env python3
"""
Google Sheets Synchronizer for AI_Embedded_Job_Hunt
Syncs local AI_Embedded_Job_Hunt.xlsx (7 Sheets) with a live Google Spreadsheet via gspread.
Supports automatic discovery of Service Account JSON keys.
"""

import os
import sys
import glob
import json
import openpyxl
import gspread

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_FILE = os.path.join(WORKSPACE_DIR, "AI_Embedded_Job_Hunt.xlsx")
URL_CACHE_FILE = os.path.join(WORKSPACE_DIR, "google_sheet_url.txt")

def find_service_account_file():
    # 1. Check GCP_SA_KEY environment variable (ideal for GitHub Actions / CI)
    sa_env = os.environ.get("GCP_SA_KEY") or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_env:
        try:
            data = json.loads(sa_env)
            if data.get("type") == "service_account" and "client_email" in data:
                return "ENV_VAR", data.get("client_email"), data
        except Exception:
            pass

    # 2. Check default filenames in workspace
    candidates = [
        os.path.join(WORKSPACE_DIR, "service_account.json"),
        os.path.join(WORKSPACE_DIR, "credentials.json")
    ]
    # Check any .json in workspace
    for json_path in glob.glob(os.path.join(WORKSPACE_DIR, "*.json")):
        if "package" not in json_path and "release" not in json_path:
            candidates.append(json_path)

    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if data.get("type") == "service_account" and "client_email" in data:
                        return path, data.get("client_email"), data
            except Exception:
                continue
    return None, None, None

def sync_to_google_sheets(target_url_or_id: str = None):
    key_file, client_email, key_dict = find_service_account_file()
    
    if not key_file:
        print("\n" + "!" * 70)
        print("GOOGLE SHEETS SERVICE ACCOUNT KEY NOT FOUND")
        print("!" * 70)
        print("Please place your Service Account JSON key in this workspace.")
        print("!" * 70 + "\n")
        return False

    print("=" * 70)
    print("GOOGLE SHEETS SYNCHRONIZER — AI EMBEDDED JOB HUNT")
    print("=" * 70)
    if key_dict:
        client = gspread.service_account_from_dict(key_dict)
    else:
        client = gspread.service_account(filename=key_file)

    # Determine target URL
    sheet_url = target_url_or_id
    if not sheet_url and os.path.exists(URL_CACHE_FILE):
        with open(URL_CACHE_FILE, "r", encoding="utf-8") as f:
            sheet_url = f.read().strip()

    if not sheet_url:
        print("📋 ACTION REQUIRED TO CONNECT YOUR GOOGLE SHEET:")
        print("1. Open Google Sheets and create a blank sheet: https://sheets.new")
        print("2. Click the 'Share' button in the top right.")
        print(f"3. Add your service account as Editor:")
        print(f"   👉 {client_email}")
        print("4. Copy your Google Sheet link and run:")
        print("   python scripts/sync_to_google_sheets.py \"<PASTE_YOUR_GOOGLE_SHEET_URL>\"\n")
        print("Once linked once, the system remembers the URL automatically!")
        print("=" * 70)
        return False

    print(f"Connecting to Google Sheet: {sheet_url}...")
    try:
        if "docs.google.com" in sheet_url:
            sh = client.open_by_url(sheet_url)
        else:
            sh = client.open_by_key(sheet_url)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"\n❌ Error: Spreadsheet not found.")
        print(f"Did you share the sheet with '{client_email}' as Editor?")
        return False
    except Exception as e:
        print(f"\n❌ Error opening spreadsheet: {e}")
        print(f"Ensure '{client_email}' has Editor access to the sheet.")
        return False

    # Save target URL for future automated runs
    with open(URL_CACHE_FILE, "w", encoding="utf-8") as f:
        f.write(sh.url)

    print(f"✓ Connected to Google Sheet: \"{sh.title}\"")
    print(f"📂 Reading local 8-sheet workbook: {XLSX_FILE}")
    
    wb = openpyxl.load_workbook(XLSX_FILE, data_only=False)

    for sheet_name in wb.sheetnames:
        xl_ws = wb[sheet_name]
        data = []
        for row in xl_ws.iter_rows(values_only=True):
            cleaned_row = ["" if val is None else str(val) for val in row]
            if any(cleaned_row):
                data.append(cleaned_row)
        
        if not data:
            continue

        max_cols = max(len(r) for r in data)
        max_rows = len(data)

        try:
            gs_ws = sh.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            gs_ws = sh.add_worksheet(title=sheet_name, rows=max(max_rows + 20, 100), cols=max(max_cols + 5, 26))

        # Ensure sheet dimensions can fit the data
        if gs_ws.row_count < max_rows + 5:
            gs_ws.add_rows(max_rows + 10 - gs_ws.row_count)
        if gs_ws.col_count < max_cols + 2:
            gs_ws.add_cols(max_cols + 5 - gs_ws.col_count)
            
        gs_ws.clear()
        gs_ws.update(range_name='A1', values=data, value_input_option='USER_ENTERED')
        print(f"   ✓ Synced tab '{sheet_name}' ({len(data)} rows, {max_cols} cols, live formulas preserved)")

    # Clean default 'Sheet1' if unused
    try:
        existing_sheets = [s.title for s in sh.worksheets()]
        if "Sheet1" in existing_sheets and "Sheet1" not in wb.sheetnames:
            sh.del_worksheet(sh.worksheet("Sheet1"))
    except Exception:
        pass

    print("\n" + "=" * 70)
    print("🎉 ALL 7 SHEETS SUCCESSFULLY HOSTED & SYNCED TO GOOGLE SHEETS!")
    print(f"👉 Live URL: {sh.url}")
    print("=" * 70 + "\n")
    return True

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else None
    sync_to_google_sheets(target)
