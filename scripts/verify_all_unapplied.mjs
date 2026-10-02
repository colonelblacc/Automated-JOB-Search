#!/usr/bin/env node
/**
 * verify_all_unapplied.mjs
 * 
 * Automatically iterates through all unapplied job openings in the tracker.
 * Takes a visual screenshot of every job link via Playwright.
 * If a job is closed, 404, or expired, automatically flags/archives it.
 * Only verified live jobs remain in DASHBOARD & NEW JOBS and are alerted.
 */

import { chromium } from 'playwright';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';
import { verifyAndScreenshotUrl } from './verify_and_screenshot.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const WORKSPACE_DIR = path.dirname(__dirname);

// 1. Run python helper to get list of unapplied URLs
const pyHelper = `
import openpyxl, json, re

wb = openpyxl.load_workbook("AI_Embedded_Job_Hunt.xlsx", data_only=False)
ws = wb["DASHBOARD & NEW JOBS"]
headers = [c.value for c in ws[5]]

def col_idx(name):
    for i, h in enumerate(headers):
        if h and name.lower() in str(h).lower():
            return i
    return None

c_company = col_idx("company")
c_title = col_idx("job title")
c_id = col_idx("canonical job id")
c_url = col_idx("authoritative job url") or col_idx("apply link")
c_company_url = col_idx("company url")
c_status = col_idx("job status")

jobs = []
for row_idx, row in enumerate(ws.iter_rows(min_row=6, values_only=True), start=6):
    if not row or not any(row):
        continue
    st = str(row[c_status] or "").strip().upper()
    if st in ["APPLIED", "CLOSED", "REJECTED", "EXPIRED", "DEAD_LINK"]:
        continue
    
    company = str(row[c_company] or "").strip()
    title = str(row[c_title] or "").strip()
    job_id = str(row[c_id] or f"{company}_{row_idx}").strip()
    
    raw_url = str(row[c_url] or row[c_company_url] or "").strip()
    m = re.search(r'https?://[^",\\s)]+', raw_url)
    clean_url = m.group(0) if m else raw_url
    
    if clean_url and clean_url.startswith("http"):
        jobs.append({"row": row_idx, "company": company, "title": title, "id": job_id, "url": clean_url})

print(json.dumps(jobs))
`;

async function main() {
  console.log("============================================================");
  console.log("PLAYWRIGHT JOB LIVENESS & SCREENSHOT VERIFICATION ENGINE");
  console.log("============================================================");

  const res = spawnSync("python", ["-c", pyHelper], { cwd: WORKSPACE_DIR, encoding: "utf-8" });
  if (res.error || res.status !== 0) {
    console.error("Failed to read tracker:", res.stderr);
    process.exit(1);
  }

  let jobs = [];
  try {
    jobs = JSON.parse(res.stdout);
  } catch (e) {
    console.error("Failed to parse jobs list:", e);
    process.exit(1);
  }

  console.log(`Found ${jobs.length} unapplied jobs to verify with Playwright...\n`);

  const browser = await chromium.launch({ headless: true });
  const deadOrExpired = [];

  for (const job of jobs) {
    const outcome = await verifyAndScreenshotUrl(job.url, job.company, job.id, browser);
    console.log(`[${outcome.status}] ${job.company} — ${job.title}`);
    console.log(`   📸 Screenshot: ${path.basename(outcome.screenshotPath || '')}`);
    if (outcome.status !== 'VERIFIED_LIVE') {
      console.log(`   ⚠️ Reason: ${outcome.reason}`);
      deadOrExpired.push({ ...job, status: outcome.status, reason: outcome.reason });
    }
    console.log("");
  }

  await browser.close();

  // If any dead or expired jobs found, update the Excel tracker automatically
  if (deadOrExpired.length > 0) {
    console.log(`\nArchiving ${deadOrExpired.length} dead/expired jobs from NEW JOBS...`);
    const updateScript = `
import openpyxl, json

dead = json.loads('''${JSON.stringify(deadOrExpired)}''')
wb = openpyxl.load_workbook("AI_Embedded_Job_Hunt.xlsx")
ws = wb["DASHBOARD & NEW JOBS"]

dead_rows = {d["row"]: d["status"] for d in dead}
for r, st in dead_rows.items():
    ws.cell(row=r, column=12, value=st) # Column 12 is JOB STATUS

wb.save("AI_Embedded_Job_Hunt.xlsx")
print("Tracker updated successfully.")
`;
    spawnSync("python", ["-c", updateScript], { cwd: WORKSPACE_DIR });
    
    // Regenerate tracker to move expired jobs into ARCHIVED JOBS
    spawnSync("python", ["scripts/generate_excel_tracker.py"], { cwd: WORKSPACE_DIR });
    spawnSync("python", ["scripts/sync_to_google_sheets.py"], { cwd: WORKSPACE_DIR });
    console.log("✓ Google Sheets & local tracker updated with verified live openings.");
  } else {
    console.log("\n🎉 All scanned job links are VERIFIED LIVE and active!");
  }
}

main().catch(err => {
  console.error("Verification crashed:", err);
  process.exit(1);
});
