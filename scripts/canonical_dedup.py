#!/usr/bin/env python3
"""
Canonical Job Normalization & Deduplication Engine
Implements the 3-Tier Multi-Source Observation Model:
  Level 1 (Authoritative): Company ATS (Greenhouse, Lever, Ashby, Workday, SmartRecruiters, Eightfold) & Career Portals
  Level 2 (Discovery): Naukri, LinkedIn Jobs, Indeed India, Wellfound, Glassdoor
  Level 3 (Supplemental): Google Jobs, Hacker News "Who is Hiring?", Campus Placement Portals

Transforms raw multi-source sightings into a single authoritative Canonical Job Record:
  - Extracts Requisition IDs (e.g. 200056513, HRC1765500, R5034526)
  - Generates stable CANONICAL_JOB_ID: {company_slug}_{req_id} or {company_slug}_{role_family}_{hash}
  - Resolves PRIMARY_SOURCE to authoritative ATS / career page
  - Aggregates DISCOVERY_SOURCES (e.g. "Workday; LinkedIn; Indeed; Naukri")
  - Tracks FIRST_SEEN, LAST_SEEN, POSTED_DATE, and JOB_STATUS
"""

import re
import hashlib
from typing import List, Dict, Any, Tuple
from datetime import datetime

# Company name standardizer (stripping legal suffixes)
COMPANY_ALIASES = {
    "qualcomm india": "Qualcomm",
    "qualcomm technologies inc": "Qualcomm",
    "infineon technologies ag": "Infineon Technologies",
    "infineon india": "Infineon Technologies",
    "texas instruments india": "Texas Instruments",
    "ti": "Texas Instruments",
    "microsoft corporation": "Microsoft",
    "microsoft india": "Microsoft",
    "ge vernova": "GE Vernova",
    "ge healthcare": "GE Healthcare",
    "general electric": "GE Healthcare",
    "larsen & toubro limited": "Larsen & Toubro (L&T)",
    "l&t": "Larsen & Toubro (L&T)",
    "larsen and toubro": "Larsen & Toubro (L&T)",
    "ideaforge technology": "ideaForge",
    "ather energy pvt ltd": "Ather Energy",
    "bytebeam iot": "Bytebeam",
    "riod energy pvt ltd": "RIOD Energy",
    "tosil systems private limited": "TOSIL Systems",
    "galaxeye space solutions": "GalaxEye Space",
    "park controls and communications": "Park Controls & Communications",
    "tessolve semiconductor": "Tessolve",
    "blackfig technologies": "Blackfig Technologies",
    "statiq ev": "Statiq"
}

LEVEL_1_PROVIDERS = [
    "workday", "greenhouse", "lever", "ashby", "smartrecruiters", 
    "eightfold", "direct_portal", "career page", "ats"
]

def normalize_company_name(name: str) -> str:
    cleaned = (name or "").strip()
    c_lower = cleaned.lower()
    for alias, canonical in COMPANY_ALIASES.items():
        if c_lower == alias or alias in c_lower:
            return canonical
            
    # Strip common suffixes
    cleaned = re.sub(r'(?i)\s+(pvt\.?\s*ltd\.?|limited|ltd\.?|inc\.?|corp\.?|llc|technologies|solutions)', '', cleaned).strip()
    return cleaned if cleaned else name

def extract_req_id(title: str, url: str) -> str:
    """Extracts job requisition / reference ID from title or URL."""
    # Look for patterns like (Job No: 200056513), (HRC1765500), (R5034526), jk=62efac7c85f66725, view/4460283803
    t_match = re.search(r'(?:job\s*(?:no|id|code)?[:#\s]*|req[:#\s]*|hrc|r\d+)([a-zA-Z0-9_\-]+)', title, re.IGNORECASE)
    if t_match:
        return t_match.group(1).upper()
        
    u_match = re.search(r'(?:job[s]?/|/jobs/view/|jk=|/job/|id=)(\d{6,12}|[a-zA-Z0-9_\-]{8,24})', url, re.IGNORECASE)
    if u_match:
        return u_match.group(1)
        
    return ""

def generate_canonical_id(company: str, title: str, role_family: str, location: str, req_id: str = "") -> str:
    comp_slug = re.sub(r'[^a-z0-9]', '', company.lower())[:8]
    
    if req_id and len(req_id) >= 4:
        clean_req = re.sub(r'[^a-zA-Z0-9]', '', req_id).lower()[:12]
        return f"{comp_slug}_{clean_req}"
        
    # Fallback to normalized semantic hash
    semantic = f"{comp_slug}_{role_family.lower()}_{title.lower()}_{location.lower()}"
    h = hashlib.md5(semantic.encode("utf-8")).hexdigest()[:6]
    return f"{comp_slug}_{h}"

def is_level_1_source(source_str: str) -> bool:
    s_lower = (source_str or "").lower()
    return any(p in s_lower for p in LEVEL_1_PROVIDERS)

def merge_raw_observations(observations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Merges multi-source sightings (ATS, LinkedIn, Indeed, Naukri, Wellfound)
    into single authoritative Canonical Job records.
    """
    grouped: Dict[str, List[Dict[str, Any]]] = {}

    for obs in observations:
        comp_norm = normalize_company_name(obs.get("company", ""))
        title = obs.get("title", "")
        url = obs.get("job_url") or obs.get("url") or ""
        role_family = obs.get("role_family") or "Firmware"
        loc = obs.get("location") or "India"
        
        req_id = obs.get("req_id") or extract_req_id(title, url)
        canonical_id = obs.get("canonical_id") or generate_canonical_id(comp_norm, title, role_family, loc, req_id)

        if canonical_id not in grouped:
            grouped[canonical_id] = []
        grouped[canonical_id].append({**obs, "company": comp_norm, "req_id": req_id})

    merged_records = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for canonical_id, sightings in grouped.items():
        # Pick primary authoritative source
        level_1_sightings = [s for s in sightings if is_level_1_source(s.get("source", "") or s.get("source_type", ""))]
        primary = level_1_sightings[0] if level_1_sightings else sightings[0]

        # Aggregate unique discovery sources
        disc_sources = []
        for s in sightings:
            src = s.get("source") or s.get("source_type") or "Direct"
            # Simplify source name
            s_name = "Workday" if "workday" in src.lower() else \
                     "Greenhouse" if "greenhouse" in src.lower() else \
                     "Lever" if "lever" in src.lower() else \
                     "Ashby" if "ashby" in src.lower() else \
                     "Eightfold" if "eightfold" in src.lower() else \
                     "SmartRecruiters" if "smartrecruiters" in src.lower() else \
                     "LinkedIn" if "linkedin" in src.lower() else \
                     "Naukri" if "naukri" in src.lower() else \
                     "Indeed" if "indeed" in src.lower() else \
                     "Wellfound" if "wellfound" in src.lower() else \
                     "Glassdoor" if "glassdoor" in src.lower() else \
                     "Direct ATS" if "ats" in src.lower() else src
                     
            if s_name not in disc_sources:
                disc_sources.append(s_name)

        # Dates
        first_seens = [s.get("first_seen") for s in sightings if s.get("first_seen")]
        last_seens = [s.get("last_seen") or s.get("last_verified") for s in sightings if (s.get("last_seen") or s.get("last_verified"))]
        posted_dates = [s.get("posted_date") for s in sightings if s.get("posted_date")]

        first_seen = min(first_seens) if first_seens else today_str
        last_seen = max(last_seens) if last_seens else today_str
        posted_date = min(posted_dates) if posted_dates else first_seen

        merged_rec = {
            **primary,
            "canonical_job_id": canonical_id,
            "primary_source": primary.get("ats") or primary.get("source") or "Company ATS",
            "discovery_sources": "; ".join(disc_sources),
            "job_url": primary.get("job_url") or primary.get("url") or "",
            "application_url": primary.get("application_url") or primary.get("job_url") or "",
            "first_seen": first_seen,
            "last_seen": last_seen,
            "posted_date": posted_date,
            "job_status": primary.get("job_status") or "ACTIVE",
            "sighting_count": len(sightings)
        }
        merged_records.append(merged_rec)

    return merged_records

if __name__ == "__main__":
    test_obs = [
        {
            "company": "Qualcomm India", "title": "Associate Engineer - Hardware / Embedded (Req 3084920)",
            "role_family": "Embedded Software", "location": "Hyderabad", "source": "Eightfold",
            "ats": "Eightfold", "job_url": "https://qualcomm.eightfold.ai/careers?pid=3084920", "first_seen": "2026-10-01"
        },
        {
            "company": "Qualcomm", "title": "Associate Engineer - Hardware / Embedded",
            "role_family": "Embedded Software", "location": "Hyderabad", "source": "LinkedIn",
            "ats": "Eightfold", "job_url": "https://www.linkedin.com/jobs/view/4470123456", "first_seen": "2026-10-02"
        },
        {
            "company": "Qualcomm Technologies Inc", "title": "Associate Engineer - Hardware (Req 3084920)",
            "role_family": "Embedded Software", "location": "Hyderabad", "source": "Naukri",
            "ats": "Eightfold", "job_url": "https://www.naukri.com/job-listings-qualcomm-3084920", "first_seen": "2026-10-02"
        }
    ]
    merged = merge_raw_observations(test_obs)
    print("Deduplication Test:")
    for m in merged:
        print(f"Canonical ID: {m['canonical_job_id']}")
        print(f"Company: {m['company']}")
        print(f"Primary Source: {m['primary_source']}")
        print(f"Discovery Sources: {m['discovery_sources']}")
        print(f"Job URL: {m['job_url']}")
        print(f"Sightings: {m['sighting_count']}")
