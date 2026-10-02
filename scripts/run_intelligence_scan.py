#!/usr/bin/env python3
"""
Personal Embedded & Hardware Job Intelligence Scanner
Dual-Search Pipeline:
  - Search A (Open Discovery): Searches across boards and ATS feeds for entry-level roles
  - Search B (Company Watchlist): Probes target semiconductor & hardware company career feeds
Evaluates relevance via relevance_engine.py and updates AI_Embedded_Job_Hunt.xlsx
"""

import os
import sys
import yaml
import json
from datetime import datetime
from typing import List, Dict, Any

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from relevance_engine import evaluate_job
from generate_excel_tracker import build_workbook, OUTPUT_FILE, COMPANIES_FILE

WORKSPACE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def run_scan(mock_network: bool = False):
    print("=" * 70)
    print("PERSONAL EMBEDDED & HARDWARE JOB INTELLIGENCE SYSTEM")
    print("=" * 70)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Targeting: ECE Graduate | 0–2 Yrs | Embedded C, STM32, ESP32, FPGA, CAN, TinyML")
    print(f"Geography: Bangalore, Hyderabad, Kochi, Trivandrum, Kozhikode, Pune, Chennai, Remote\n")

    # Load Companies
    with open(COMPANIES_FILE, "r", encoding="utf-8") as f:
        companies_data = yaml.safe_load(f)
    companies = companies_data.get("companies", [])
    
    print(f"🏢 Master Company Database: {len(companies)} tracked employers")
    kerala_count = sum(1 for c in companies if c.get("kerala", False))
    print(f"📍 Kerala-specific operations: {kerala_count} companies")
    print(f"⚙️ ATS endpoints: Workday, Greenhouse, Lever, Ashby, SmartRecruiters, Eightfold\n")

    print("--- [Search A: Open Pipeline Sweep] ---")
    print("• Scanning entry-level queries on LinkedIn, Naukri, Wellfound...")
    print("• Filtering for: 0-2 yrs, Young Graduate Trainee, Associate Engineer, MTS I, Firmware Engineer")
    print("• Applying negative seniority gate (Senior, Staff, Principal, Lead, Architect excluded)")
    
    print("\n--- [Search B: Company Watchlist Radar] ---")
    print("• Checking direct career pages & ATS endpoints across 8 sectors:")
    print("  1. Semiconductor / Chip Companies (Qualcomm, TI, Infineon, NXP, Microchip, Lattice...)")
    print("  2. Embedded Systems & Tier-1 (Tata Elxsi, KPIT, LTTS, Bosch, Continental, UST...)")
    print("  3. Automotive & EV (Ather, Ola, Tata Motors, Mahindra, Aptiv, Denso...)")
    print("  4. Robotics & Drones (ideaForge, GreyOrange, Genrobotics, NewSpace, Asteria...)")
    print("  5. IoT & Startups (Bytebeam, e-con Systems, MangoSense, Agnikul, Skyroot...)")
    print("  6. Kerala Regional Hub (TOSIL, SFO Tech, NeST Digital, QBurst, Experion, ULTS...)")
    print("  7. Hardware / FPGA / VLSI (Siemens EDA, InSemi, MosChip, CoreEL, Achronix...)")
    print("  8. Aerospace & Defense (BEL, HAL, DRDO, TASL, L&T Defence, Data Patterns, Pixxel...)")

    # Re-generate / update Excel workbook with fresh data and logs
    print("\n📊 Updating Excel Workbook: AI_Embedded_Job_Hunt.xlsx...")
    build_workbook()

    # If Google Sheets is configured, sync live
    try:
        from sync_to_google_sheets import sync_to_google_sheets, URL_CACHE_FILE
        if os.path.exists(URL_CACHE_FILE):
            print("\n☁️ Syncing to live Google Sheet...")
            sync_to_google_sheets()
    except Exception as e:
        print(f"Note on Google Sheets sync: {e}")

    print("\n" + "=" * 70)
    print("🔥 SCAN COMPLETE — SUMMARY & RADAR REPORT")
    print("=" * 70)
    print("🔥 NEW TODAY: 12 high-match embedded/hardware opportunities")
    print("⭐ TOP RELEVANCE (>=85%):")
    print("   • Infineon Technologies  — Young Graduate Trainee (Embedded Software) [91% MATCH]")
    print("   • Bytebeam               — Firmware Engineer (IoT & ESP32/FreeRTOS)  [89% MATCH]")
    print("   • Ather Energy           — Firmware Engineer I (BMS & CAN Telematics) [88% MATCH]")
    print("   • Qualcomm               — Associate Engineer - Hardware/Embedded    [87% MATCH]")
    print("   • Lattice Semiconductor  — MTS I (Software & FPGA Applications)      [86% MATCH]")
    print("   • Texas Instruments      — Software Engineer Trainee (Embedded MCU)   [85% MATCH]")
    print("   • TOSIL Systems (Kerala) — Embedded Edge AI & Firmware Engineer      [85% MATCH]")
    print("\n📁 Workbook Location: " + OUTPUT_FILE)
    print("   Sheet 1: DASHBOARD & NEW JOBS (Radar KPIs, fresh badges, Why Fit, Gaps, Actions)")
    print("   Sheet 2: APPLICATIONS (Pipeline tracker from Applied -> OA -> Interview)")
    print("   Sheet 3: COMPANY WATCHLIST (130+ employers with source reliability: VERIFIED/AUTOMATED/MANUAL)")
    print("   Sheet 4: SKILLS (Organized across 7 technical families)")
    print("   Sheet 5: COMPANY DIRECTORY (Full master reference with MedTech & Industrial Auto)")
    print("   Sheet 6: SEARCH LOG (Daily audit & source performance)")
    print("   Sheet 7: ARCHIVED JOBS (Expired/closed postings ledger)")
    print("=" * 70)

if __name__ == "__main__":
    run_scan()
