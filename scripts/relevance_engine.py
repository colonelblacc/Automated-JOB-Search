#!/usr/bin/env python3
"""
Embedded & Hardware Relevance & Eligibility Engine (V2)
Features:
  1. Hard Eligibility Filter:
     - Experience blocker (>3 yrs -> NOT ELIGIBLE)
     - Unrelated degree blocker
     - Non-India/Non-Remote location blocker
  2. 7 Skill Families Taxonomy:
     - Firmware, MCU, Interfaces, Embedded Linux, Hardware, FPGA/VLSI, Edge AI
  3. Structured Output:
     - RELEVANCE SCORE (0-100)
     - MATCH LEVEL (HIGH / MEDIUM / LOW)
     - RECOMMENDED ACTION (APPLY / REVIEW / KEEP WATCH / NOT ELIGIBLE)
     - WHY FIT (structured positive evidence)
     - GAPS (explicit missing tech)
     - WORK MODE (On-site / Hybrid / Remote / Unknown)
     - JOB TYPE (Full-time / Internship / Graduate Program / Trainee)
     - FRESHNESS BADGE (🟢 <7d, 🟡 7-21d, 🟠 21-30d, 🔴 >30d)
"""

import re
from datetime import datetime
from typing import Dict, List, Tuple, Any

# 17 Standard Technical Role Families for Embedded / Hardware
ROLE_FAMILIES = [
    "Firmware", "Embedded Software", "Embedded Linux", "Device Driver", 
    "Board Bring-up", "Hardware Design", "PCB Design", "FPGA", 
    "RTL/VLSI", "Verification", "Validation", "IoT", 
    "Automotive Embedded", "Robotics", "Edge AI", "Computer Vision", "Test/Debug"
]

def detect_role_family(title: str, text: str = "") -> str:
    """Classifies job into one of the 17 standard role families."""
    combined = f"{title} {text}".lower()
    
    if any(k in combined for k in ["fpga", "xilinx", "altera", "vivado", "radiant", "quartus"]):
        return "FPGA"
    if any(k in combined for k in ["rtl", "asic", "vlsi", "soc design", "digital design", "verilog", "systemverilog"]):
        return "RTL/VLSI"
    if any(k in combined for k in ["silicon validation", "post-silicon", "hardware validation", "rf validation", "validation engineer"]):
        return "Validation"
    if any(k in combined for k in ["verification", "uvm", "formal verification", "functional verification"]):
        return "Verification"
    if any(k in combined for k in ["yocto", "embedded linux", "bitbake", "buildroot", "kernel", "device tree"]):
        return "Embedded Linux"
    if any(k in combined for k in ["device driver", "bsp", "bootloader", "u-boot"]):
        return "Device Driver"
    if any(k in combined for k in ["board bring-up", "bring up", "bring-up", "bench debug", "lab bring"]):
        return "Board Bring-up"
    if any(k in combined for k in ["pcb", "kicad", "altium", "schematic", "layout", "gerber", "drc"]):
        return "PCB Design"
    if any(k in combined for k in ["hardware design", "hardware engineer", "hardware intern", "schematic design"]):
        return "Hardware Design"
    if any(k in combined for k in ["can bus", "autosar", "telematics", "bms", "electric vehicle", "automotive", "ecu", "ev electronics"]):
        return "Automotive Embedded"
    if any(k in combined for k in ["drone", "uav", "robotics", "ros", "ros2", "flight controller", "avionics"]):
        return "Robotics"
    if any(k in combined for k in ["tinyml", "edge ai", "tflite", "microcontroller ai", "embedded ml", "npu"]):
        return "Edge AI"
    if any(k in combined for k in ["computer vision", "opencv", "yolo", "camera", "isp", "imaging"]):
        return "Computer Vision"
    if any(k in combined for k in ["lora", "lorawan", "iot", "smart metering", "telemetry", "mqtt", "esp32", "cloud iot"]):
        return "IoT"
    if any(k in combined for k in ["test engineer", "hardware test", "qa", "debug engineer", "test automation"]):
        return "Test/Debug"
    if any(k in combined for k in ["firmware", "embedded c", "freertos", "rtos", "bare metal", "microcontroller firmware"]):
        return "Firmware"
    if any(k in combined for k in ["embedded software", "embedded developer", "embedded systems engineer"]):
        return "Embedded Software"
    
    return "Firmware"

# 7 Technical Skill Families for Embedded / Hardware
SKILL_FAMILIES = {
    "Firmware": [
        "embedded c", "c++", "bare metal", "rtos", "freertos", "zephyr", 
        "device drivers", "bootloader", "firmware", "bsp"
    ],
    "MCU": [
        "stm32", "esp32", "msp430", "c2000", "aurix", "pic", "avr", 
        "arm cortex", "cortex-m", "renesas", "nordic", "silicon labs"
    ],
    "Interfaces": [
        "uart", "spi", "i2c", "can", "can fd", "lin", "usb", "ble", 
        "bluetooth", "wi-fi", "ethernet", "mqtt", "lora", "lorawan"
    ],
    "Embedded Linux": [
        "embedded linux", "yocto", "buildroot", "device tree", "kernel", 
        "v4l2", "u-boot", "posix"
    ],
    "Hardware": [
        "pcb", "kicad", "altium", "orcad", "schematic", "signal integrity", 
        "power electronics", "board bring-up", "oscilloscope", "logic analyzer"
    ],
    "FPGA/VLSI": [
        "verilog", "systemverilog", "vhdl", "fpga", "rtl", "vivado", 
        "quartus", "asic", "verification", "uvm", "eda"
    ],
    "Edge AI": [
        "tinyml", "tensorflow lite", "tflite micro", "edge impulse", 
        "edge ai", "onnx", "computer vision", "sensor fusion"
    ]
}

# Candidate profile strengths (Ajith Shajan - KTU B.Tech ECE 2027)
CANDIDATE_SKILLS = {
    "high": ["embedded c", "stm32", "esp32", "python", "c++", "can", "spi", "i2c", "uart", "verilog", "fpga", "tinyml", "lora", "ble", "kicad", "oscilloscope", "logic analyzer"],
    "learning": ["freertos", "embedded linux", "yocto", "autosar", "aurix", "ros", "ros2"]
}

INTERNSHIP_MARKERS = [
    "intern", "internship", "student trainee", "summer intern", "6-month", 
    "co-op", "trainee intern", "project trainee", "r&d intern", "engineering intern",
    "firmware intern", "embedded intern", "hardware intern"
]

TRAINEE_MARKERS = [
    "graduate trainee", "young graduate trainee", "get", "graduate engineer trainee",
    "associate engineer", "trainee", "fresher", "campus", "mts 1", "mts i",
    "engineer 1", "engineer i", "0-1", "0 to 1", "systems trainee"
]

ENTRY_LEVEL_MARKERS = INTERNSHIP_MARKERS + TRAINEE_MARKERS + [
    "graduate", "entry level", "entry-level", "junior", "0-2", "0 to 2", "0-3"
]

HARD_SENIOR_BLOCKERS = [
    "senior", "sr.", "principal", "staff", "architect", "lead", 
    "engineering manager", "director", "head of", "vp", "5+ years", 
    "6+ years", "7+ years", "8+ years", "10+ years"
]

TARGET_LOCATIONS = [
    "bangalore", "bengaluru", "hyderabad", "kochi", "cochin", 
    "trivandrum", "thiruvananthapuram", "kozhikode", "calicut", 
    "kerala", "pune", "chennai", "noida", "gurugram", "delhi", "remote", "india"
]

ECE_DEGREES = [
    "ece", "electronics and communication", "electronics & communication",
    "electronics", "electrical and electronics", "eee", "instrumentation",
    "embedded systems", "vlsi", "telecommunication"
]

def check_hard_eligibility(job: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Applies strict pass/fail filters before any scoring occurs.
    Returns (is_eligible, disqualification_reason).
    """
    title = (job.get("title") or "").lower()
    desc = (job.get("description") or "").lower()
    location = (job.get("location") or "").lower()
    
    # 1. Experience Check (Senior / Staff / 5+ yrs hard blocker)
    for blocker in HARD_SENIOR_BLOCKERS:
        if blocker in title:
            return False, f"Ineligible seniority level in title: '{blocker}'"
            
    exp_match = re.search(r'(\d+)\s*[-+to]*\s*(\d*)\s*(?:years?|yrs?)', desc)
    if exp_match:
        min_exp = int(exp_match.group(1))
        if min_exp > 3:
            return False, f"Requires {min_exp}+ years experience (entry target is 0-2 yrs)"
            
    # 2. Location Check (Must be India or Remote)
    if location:
        has_allowed_loc = any(loc in location for loc in TARGET_LOCATIONS)
        foreign_blockers = ["united states", "usa", "germany", "united kingdom", "london", "canada", "australia", "singapore", "japan"]
        if any(fb in location for fb in foreign_blockers) and not has_allowed_loc:
            return False, f"Location '{location}' is outside target regions and not remote"
            
    return True, "Eligible"

def detect_work_mode(text: str) -> str:
    text_lower = text.lower()
    if "remote" in text_lower or "work from home" in text_lower:
        return "Remote"
    elif "hybrid" in text_lower:
        return "Hybrid"
    elif "on-site" in text_lower or "onsite" in text_lower or "in-office" in text_lower:
        return "On-site"
    return "On-site / Hybrid"

def detect_job_type(title: str, text: str) -> str:
    t_lower = f"{title} {text}".lower()
    if any(m in t_lower for m in INTERNSHIP_MARKERS) or "intern" in t_lower:
        return "Internship"
    elif any(m in t_lower for m in TRAINEE_MARKERS) or "trainee" in t_lower or "get" in t_lower:
        return "Graduate Program / Trainee"
    elif "contract" in t_lower or "fixed term" in t_lower:
        return "Contract"
    elif "apprentice" in t_lower:
        return "Apprenticeship"
    return "Full-time"

def calculate_freshness(posted_date_str: str, first_seen_str: str = None) -> Tuple[int, str]:
    """
    Returns (days_open, badge):
      🟢 <7d, 🟡 7-21d, 🟠 21-30d, 🔴 >30d
    """
    today = datetime.now()
    ref_date = None
    
    for d_str in [posted_date_str, first_seen_str]:
        if d_str:
            try:
                ref_date = datetime.strptime(d_str.strip()[:10], "%Y-%m-%d")
                break
            except Exception:
                pass
                
    if not ref_date:
        days_open = 1
    else:
        days_open = max(0, (today - ref_date).days)
        
    if days_open < 7:
        badge = f"🟢 Fresh ({days_open}d)"
    elif days_open <= 21:
        badge = f"🟡 Active ({days_open}d)"
    elif days_open <= 30:
        badge = f"🟠 Aging ({days_open}d)"
    else:
        badge = f"🔴 Stale ({days_open}d)"
        
    return days_open, badge

def evaluate_job(job: Dict[str, Any]) -> Dict[str, Any]:
    """
    Evaluates job against hard filters, relevance scoring, and technical evidence.
    """
    title = job.get("title") or ""
    desc = job.get("description") or ""
    location = job.get("location") or ""
    full_text = f"{title} {desc} {location}".lower()
    
    # 1. Hard Eligibility Filter
    is_eligible, disqualification = check_hard_eligibility(job)
    role_family = detect_role_family(title, desc)
    if not is_eligible:
        return {
            "eligible": False,
            "eligibility_status": "NOT ELIGIBLE",
            "eligibility_reason": disqualification,
            "role_family": role_family,
            "relevance_score": 0,
            "match_level": "NOT ELIGIBLE",
            "recommended_action": "NOT ELIGIBLE",
            "why_fit": [],
            "gaps": [disqualification],
            "work_mode": detect_work_mode(full_text),
            "job_type": detect_job_type(title, desc),
            "matched_families": {},
            "days_open": 0,
            "freshness_badge": "🔴 Disqualified",
            "posted_date": job.get("posted_date") or datetime.now().strftime("%Y-%m-%d"),
            "first_seen": job.get("first_seen") or datetime.now().strftime("%Y-%m-%d"),
            "last_verified": datetime.now().strftime("%Y-%m-%d"),
            "job_status": "CLOSED"
        }
        
    why_fit = []
    gaps = []
    
    # 2. Extract Skills by Family
    matched_families = {}
    total_matched_skills = []
    for family, skills in SKILL_FAMILIES.items():
        found = [s for s in skills if s in full_text]
        if found:
            matched_families[family] = found
            total_matched_skills.extend(found)
            
    # Positive evidence
    if "Firmware" in matched_families:
        why_fit.append(f"Firmware: {', '.join([s.title() for s in matched_families['Firmware'][:3]])}")
    if "MCU" in matched_families:
        why_fit.append(f"MCU: {', '.join([s.title() for s in matched_families['MCU'][:2]])}")
    if "Interfaces" in matched_families:
        why_fit.append(f"Protocols: {', '.join([s.upper() for s in matched_families['Interfaces'][:3]])}")
    if "FPGA/VLSI" in matched_families:
        why_fit.append(f"FPGA/VLSI: {', '.join([s.title() for s in matched_families['FPGA/VLSI'][:2]])}")
    if "Edge AI" in matched_families:
        why_fit.append("Edge AI / TinyML match")

    # Check for Gaps
    gap_candidates = ["autosar", "aurix", "yocto", "device driver", "ros", "systemverilog", "pcb design", "ble"]
    for gc in gap_candidates:
        if gc in full_text and gc not in CANDIDATE_SKILLS["high"]:
            gaps.append(gc.upper())

    # Degree fit
    if any(deg in full_text for deg in ECE_DEGREES):
        why_fit.append("ECE / Electronics degree accepted")

    # Priority Role Detection (Priority 1: Internships, Priority 2: Trainee/GET, Priority 3: Full-time)
    title_lower = title.lower()
    is_internship = any(m in full_text for m in INTERNSHIP_MARKERS) or "intern" in title_lower
    is_trainee = (any(m in full_text for m in TRAINEE_MARKERS) or "trainee" in title_lower or "get" in title_lower) and not is_internship
    is_explicit_entry = is_internship or is_trainee or any(m in full_text for m in ENTRY_LEVEL_MARKERS)

    if is_internship:
        why_fit.append("🎯 Priority 1: Internship (Highest conversion & fastest turnaround)")
    elif is_trainee:
        why_fit.append("🎯 Priority 2: Graduate / Trainee role (Structured entry-level)")
    elif is_explicit_entry:
        why_fit.append("0-2 Years / Graduate friendly")

    # Location fit
    matched_loc = None
    for loc in TARGET_LOCATIONS:
        if loc in location.lower() or loc in full_text:
            matched_loc = loc.title()
            break
    if matched_loc:
        why_fit.append(f"Location: {matched_loc}")

    # 3. Calculate Score (0-100)
    score = 0.0
    
    # Role alignment (30 pts)
    if any(t in title_lower for t in ["embedded", "firmware", "hardware", "mcu", "fpga", "vlsi", "robotics", "avionics"]):
        score += 30.0
    elif any(t in title_lower for t in ["engineer", "developer", "trainee", "associate", "intern"]):
        score += 20.0
        
    # Technical depth across families (30 pts)
    # 5 pts per active matched family (max 30)
    score += min(30.0, len(matched_families) * 6.0)
    
    # Experience fit (20 pts) + Priority Role Bonus (Internships +10 pts, Trainee +5 pts)
    if is_internship:
        score += 20.0 + 10.0  # Priority 1: Full exp fit + 10 pt internship bonus
    elif is_trainee:
        score += 20.0 + 5.0   # Priority 2: Full exp fit + 5 pt trainee/GET bonus
    elif is_explicit_entry:
        score += 20.0
    else:
        score += 12.0
        
    # ECE eligibility (10 pts)
    if any(deg in full_text for deg in ECE_DEGREES):
        score += 10.0
    else:
        score += 6.0
        
    # Location (10 pts)
    if matched_loc:
        if matched_loc.lower() in ["kochi", "trivandrum", "kerala", "bangalore", "hyderabad", "remote"]:
            score += 10.0
        else:
            score += 7.0
    else:
        score += 5.0

    score = max(20, min(99, round(score)))

    # 4. Match Level & Recommended Action (Internships prioritized with lower friction threshold)
    role_family = detect_role_family(title, desc)
    
    if is_internship and score >= 75:
        match_level = "HIGH"
        action = "APPLY"
    elif is_trainee and score >= 80:
        match_level = "HIGH"
        action = "APPLY"
    elif score >= 82 and is_explicit_entry:
        match_level = "HIGH"
        action = "APPLY"
    elif score >= 70:
        match_level = "MEDIUM"
        action = "REVIEW"
    elif score >= 55:
        match_level = "LOW"
        action = "KEEP WATCH"
    else:
        match_level = "LOW"
        action = "KEEP WATCH"

    # Freshness
    days_open, freshness_badge = calculate_freshness(job.get("posted_date"), job.get("first_seen"))

    priority_tier = 1 if is_internship else (2 if is_trainee else 3)
    priority_label = "Priority 1 (Internship)" if is_internship else ("Priority 2 (Trainee/GET)" if is_trainee else "Priority 3 (Full-time)")

    return {
        "eligible": True,
        "eligibility_status": "ELIGIBLE",
        "eligibility_reason": "Meets 0-2 yrs ECE / Hardware qualification criteria",
        "role_family": role_family,
        "priority_tier": priority_tier,
        "priority_label": priority_label,
        "relevance_score": score,
        "match_level": match_level,
        "recommended_action": action,
        "why_fit": why_fit[:5],
        "gaps": gaps[:3],
        "work_mode": detect_work_mode(full_text),
        "job_type": detect_job_type(title, desc),
        "matched_families": matched_families,
        "days_open": days_open,
        "freshness_badge": freshness_badge,
        "posted_date": job.get("posted_date") or datetime.now().strftime("%Y-%m-%d"),
        "first_seen": job.get("first_seen") or datetime.now().strftime("%Y-%m-%d"),
        "last_verified": datetime.now().strftime("%Y-%m-%d"),
        "job_status": "OPEN"
    }
