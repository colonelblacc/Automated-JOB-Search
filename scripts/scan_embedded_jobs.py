#!/usr/bin/env python3
"""
Embedded.jobs Radar Scanner
Automated crawler and ingestion pipeline for https://embedded.jobs.
Probes open search endpoints for India + Internships, Entry Level, and Remote embedded roles.
Evaluates candidates using relevance_engine.py and outputs to cache for tracker generation.
"""

import os
import sys
import re
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime
from typing import List, Dict, Any
from bs4 import BeautifulSoup

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(WORKSPACE_DIR, "scripts"))

from relevance_engine import evaluate_job

CACHE_FILE = os.path.join(WORKSPACE_DIR, "data", "embedded_jobs_cache.json")
COMPANIES_FILE = os.path.join(WORKSPACE_DIR, "config", "companies.yml")

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

# Strict applied-only set to prevent polluting radar with companies already applied to
APPLIED_COMPANIES = {
    "kalkitech", "infineon technologies", "infineon", "amaya energy", "ge healthcare", 
    "daloft aerospace", "daloft", "nerolab", "riod energy", "ge vernova", 
    "bytebeam", "galaxeye space", "galaxeye", "park controls & communications", 
    "park controls", "tessolve", "blackfig technologies", "blackfig tech", 
    "statiq", "artpark (iisc bangalore)", "artpark", "arys garage", "breakout", 
    "south indian bank", "larsen & toubro (l&t)", "l&t",
    "tcs", "tata consultancy services", "emsyne", "emsyne technologies",
    "tosil systems", "tosil", "mistral solutions", "mistral"
}

QUERIES = [
    {
        "name": "India + Internship (Priority 1)",
        "params": {"country": "India", "employment": "internship", "time": "all"}
    },
    {
        "name": "India + Entry Level (Priority 2 & 3)",
        "params": {"country": "India", "experience": "entry level", "time": "all"}
    },
    {
        "name": "India + Recent 1 Month",
        "params": {"country": "India", "time": "1month"}
    },
    {
        "name": "Remote + Entry Level",
        "params": {"remote_work": "remote", "experience": "entry level", "time": "all"}
    }
]

def fetch_search_page(params: Dict[str, str], max_retries: int = 2) -> str:
    url = "https://embedded.jobs/search?" + urllib.parse.urlencode(params)
    for attempt in range(max_retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=12) as resp:
                return resp.read().decode('utf-8', errors='ignore')
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"   ⚠️ Request failed for {url}: {e}")
            time.sleep(1)
    return ""

def fetch_job_details(job_url: str) -> Dict[str, Any]:
    """Fetches single job page to parse Schema.org JobPosting JSON-LD."""
    details = {
        "description": "",
        "direct_apply_url": job_url,
        "date_posted": datetime.now().strftime("%Y-%m-%d"),
        "employment_type": "FULL_TIME",
        "hiring_org": ""
    }
    try:
        req = urllib.request.Request(job_url, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8', errors='ignore')

        soup = BeautifulSoup(html, 'html.parser')
        
        # Look for JSON-LD JobPosting
        for s in soup.find_all('script', type='application/ld+json'):
            try:
                data = json.loads(s.text.strip())
                if isinstance(data, dict) and data.get('@type') == 'JobPosting':
                    details["description"] = data.get("description", "")
                    details["date_posted"] = data.get("datePosted", "")[:10] if data.get("datePosted") else details["date_posted"]
                    details["employment_type"] = data.get("employmentType", "")
                    if isinstance(data.get("hiringOrganization"), dict):
                        details["hiring_org"] = data.get("hiringOrganization", {}).get("name", "")
                    break
            except Exception:
                continue

        # Look for Apply link
        apply_btn = soup.find('a', string=re.compile(r'Apply', re.I))
        if apply_btn and apply_btn.get('href'):
            href = apply_btn['href']
            if href.startswith('http') and 'embedded.jobs' not in href:
                details["direct_apply_url"] = href
    except Exception:
        pass
    return details

def extract_jobs_from_html(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, 'html.parser')
    extracted = []
    seen_urls = set()

    for a in soup.find_all('a', href=re.compile(r'/job/')):
        href = a['href']
        if href.startswith('/job/'):
            full_url = "https://embedded.jobs" + href
        else:
            full_url = href

        if full_url in seen_urls:
            continue
        seen_urls.add(full_url)

        title = a.get_text(strip=True).replace("View ", "").strip()
        card = a.find_parent('div', class_=re.compile(r'col|row|card', re.I)) or a.find_parent('div')
        card_text = card.get_text(separator=' · ', strip=True) if card else ""

        # Extract company from card text: typically `@ · CompanyName ·`
        company = "Unknown"
        comp_match = re.search(r'@\s*·\s*([^·]+?)\s*·', card_text)
        if comp_match:
            company = comp_match.group(1).strip()
        elif "with-" in full_url:
            # Fallback slug extraction: /job/Title-with-CompanyName-hash
            slug_match = re.search(r'-with-([a-zA-Z0-9\-]+)-[a-f0-9]+$', full_url)
            if slug_match:
                company = slug_match.group(1).replace('-', ' ').title()

        # Extract location: `📍Location ·`
        location = "India"
        loc_match = re.search(r'📍([^·]+)', card_text)
        if loc_match:
            location = loc_match.group(1).strip()

        # Job type from card
        job_type = "Full-time"
        if "intern" in title.lower() or "intern" in card_text.lower():
            job_type = "Internship"
        elif "trainee" in title.lower() or "graduate" in card_text.lower():
            job_type = "Graduate Program / Trainee"

        extracted.append({
            "title": title,
            "company": company,
            "location": location,
            "job_url": full_url,
            "job_type": job_type,
            "card_summary": card_text
        })

    return extracted

def run_embedded_jobs_scan() -> List[Dict[str, Any]]:
    print("=" * 70)
    print("EMBEDDED.JOBS PIPELINE SCANNER")
    print("=" * 70)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("Target: India / Remote | Priority 1 Internships & Entry-Level Firmware/HW")

    all_raw_jobs = []
    seen_urls = set()

    for q in QUERIES:
        print(f"\n🔍 Probing: {q['name']}...")
        html = fetch_search_page(q['params'])
        jobs = extract_jobs_from_html(html)
        print(f"   Found {len(jobs)} candidate listings")
        for j in jobs:
            if j["job_url"] not in seen_urls:
                seen_urls.add(j["job_url"])
                all_raw_jobs.append(j)
        time.sleep(0.5)

    print(f"\n📊 Total Unique Roles Harvested: {len(all_raw_jobs)}")

    qualified_jobs = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for raw in all_raw_jobs:
        comp_name = raw["company"].strip().lower()
        if comp_name in APPLIED_COMPANIES:
            continue

        # Location filter: Only allow India or Remote
        loc_lower = raw["location"].lower()
        is_allowed_loc = any(loc in loc_lower for loc in ["india", "bengaluru", "bangalore", "hyderabad", "kochi", "pune", "chennai", "delhi", "noida", "gurgaon", "remote"])
        has_blocked_country = any(fc in loc_lower for fc in ["france", "germany", "united states", "usa", "canada", "korea", "netherlands", "uk", "italy", "spain", "costa rica"])
        if has_blocked_country and not is_allowed_loc:
            continue

        # Fetch detailed description and JSON-LD
        details = fetch_job_details(raw["job_url"])
        full_desc = f"{raw['card_summary']} {details.get('description', '')}"

        eval_input = {
            "title": raw["title"],
            "description": full_desc,
            "location": raw["location"],
            "posted_date": details.get("date_posted") or today_str,
            "first_seen": today_str
        }

        eval_result = evaluate_job(eval_input)

        if eval_result.get("eligible") and eval_result.get("relevance_score", 0) >= 75:
            canon_id = f"emb_{re.sub(r'[^a-zA-Z0-9]', '_', raw['company'].lower())[:12]}_{raw['title'].lower()[:15]}"
            
            entry = {
                "posted": details.get("date_posted") or today_str,
                "seen": today_str,
                "freshness": eval_result.get("freshness_badge", "🟢 Fresh (1d)"),
                "days_open": eval_result.get("days_open", 1),
                "verified": today_str,
                "status": "OPEN",
                "id": f"JOB-EMB-{len(qualified_jobs)+1:03d}",
                "canonical_id": canon_id,
                "title": raw["title"],
                "role_family": eval_result.get("role_family", "Firmware"),
                "company": raw["company"],
                "category": "Embedded / Hardware Board",
                "location": raw["location"],
                "work_mode": eval_result.get("work_mode", "On-site"),
                "job_type": eval_result.get("job_type", raw["job_type"]),
                "exp": "0–2 yrs" if raw["job_type"] != "Internship" else "0 yrs (Student / Intern)",
                "salary": "Market Rate (₹4.5–10 LPA / ₹25k-45k/mo)",
                "eligibility": "ELIGIBLE",
                "eligibility_reason": f"P{eval_result.get('priority_tier', 3)}: {eval_result.get('eligibility_reason', 'Matches ECE & Embedded C background')}",
                "score": eval_result.get("relevance_score", 85),
                "priority_tier": eval_result.get("priority_tier", 3),
                "level": eval_result.get("match_level", "HIGH"),
                "action": eval_result.get("recommended_action", "APPLY"),
                "why_fit": "; ".join(eval_result.get("why_fit", ["Embedded systems alignment"])),
                "gaps": "; ".join(eval_result.get("gaps", [])) or "None identified",
                "source": "Embedded.jobs Portal",
                "source_type": "Specialized Board",
                "ats": "Embedded.jobs",
                "discovery_sources": "Embedded.jobs; Crawler",
                "job_url": raw["job_url"],
                "comp_url": "https://embedded.jobs"
            }
            qualified_jobs.append(entry)

    # Sort strictly: Priority 1 (Internships) -> Priority 2 (Trainees) -> Priority 3 (Full-time), then by score
    qualified_jobs.sort(key=lambda j: (j.get("priority_tier", 3), -j.get("score", 0)))

    print(f"⭐ Qualified High-Relevance Opportunities: {len(qualified_jobs)}")
    for i, qj in enumerate(qualified_jobs[:8]):
        print(f"   [{i+1}] {qj['title']} @ {qj['company']} ({qj['job_type']}) — Score: {qj['score']} | {qj['location']}")

    # Save to cache
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(qualified_jobs, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Cached qualified jobs to: {CACHE_FILE}")

    return qualified_jobs

if __name__ == "__main__":
    run_embedded_jobs_scan()
