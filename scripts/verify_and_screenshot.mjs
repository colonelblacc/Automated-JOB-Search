#!/usr/bin/env node
/**
 * verify_and_screenshot.mjs
 * 
 * Verifies if job postings are still live and active using Playwright.
 * Captures visual screenshot proof into screenshots/
 * Filters out expired/closed/404 postings automatically.
 */

import { chromium } from 'playwright';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const WORKSPACE_DIR = path.dirname(__dirname);
const SCREENSHOTS_DIR = path.join(WORKSPACE_DIR, 'screenshots');

if (!fs.existsSync(SCREENSHOTS_DIR)) {
  fs.mkdirSync(SCREENSHOTS_DIR, { recursive: true });
}

// Expired and closure markers
const HARD_EXPIRED_PATTERNS = [
  /job (is )?no longer available/i,
  /job.*no longer open/i,
  /this job has expired/i,
  /job posting has expired/i,
  /no longer accepting applications/i,
  /this (position|role|job) (is )?no longer/i,
  /this (?:job|role|position)(?: listing)? is closed/i,
  /job (listing )?not found/i,
  /the page you are looking for doesn't exist/i,
  /applications?\s+(?:(?:have|are|is)\s+)?closed/i,
  /closed on \d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i,
  /we couldn't find that job/i,
  /requisition (is )?closed/i,
  /page not found/i,
  /404 not found/i
];

function sanitizeFilename(str) {
  return str.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 50);
}

export async function verifyAndScreenshotUrl(url, company = 'Unknown', jobId = 'job', browserInstance = null) {
  const ownBrowser = !browserInstance;
  const browser = browserInstance || await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
  });

  const page = await context.newPage();
  const dateStr = new Date().toISOString().split('T')[0];
  const filename = `${sanitizeFilename(company)}_${sanitizeFilename(jobId)}_${dateStr}.png`;
  const screenshotPath = path.join(SCREENSHOTS_DIR, filename);

  let result = {
    url,
    company,
    jobId,
    status: 'UNKNOWN',
    httpStatus: null,
    reason: '',
    screenshotPath: null
  };

  try {
    console.log(`🔍 Checking [${company}] ${url}...`);
    const resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 20000 });
    result.httpStatus = resp ? resp.status() : null;

    if (result.httpStatus === 404 || result.httpStatus === 410) {
      result.status = 'DEAD_LINK';
      result.reason = `HTTP status ${result.httpStatus}`;
      await page.screenshot({ path: screenshotPath });
      result.screenshotPath = screenshotPath;
      return result;
    }

    // Wait 2s for client-side hydration (e.g. Workday, Lever, Greenhouse)
    await page.waitForTimeout(2000);

    const bodyText = await page.evaluate(() => document.body ? document.body.innerText : '');
    const title = await page.title();

    // Check for expired banners
    let matchedPattern = null;
    for (const pattern of HARD_EXPIRED_PATTERNS) {
      if (pattern.test(bodyText) || pattern.test(title)) {
        matchedPattern = pattern;
        break;
      }
    }

    // Take screenshot as visual proof
    await page.screenshot({ path: screenshotPath, fullPage: false });
    result.screenshotPath = screenshotPath;

    if (matchedPattern) {
      result.status = 'EXPIRED';
      result.reason = `Matched expired pattern: ${matchedPattern}`;
    } else {
      result.status = 'VERIFIED_LIVE';
      result.reason = 'Page loaded with active job content';
    }

  } catch (err) {
    result.status = 'ERROR';
    result.reason = err.message;
    try {
      await page.screenshot({ path: screenshotPath });
      result.screenshotPath = screenshotPath;
    } catch (_) {}
  } finally {
    await context.close();
    if (ownBrowser) {
      await browser.close();
    }
  }

  return result;
}

// CLI Runner
async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.log('Usage: node verify_and_screenshot.mjs <url> [company] [jobId]');
    process.exit(1);
  }

  const url = args[0];
  const company = args[1] || 'Company';
  const jobId = args[2] || 'job';

  const res = await verifyAndScreenshotUrl(url, company, jobId);
  console.log('\n--- VERIFICATION RESULT ---');
  console.log(`Status:     ${res.status}`);
  console.log(`Reason:     ${res.reason}`);
  console.log(`Screenshot: ${res.screenshotPath}`);
  console.log('---------------------------\n');
}

if (process.argv[1] && process.argv[1].endsWith('verify_and_screenshot.mjs')) {
  main();
}
