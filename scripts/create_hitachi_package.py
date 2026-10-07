import sys
import os
import subprocess
import shutil

# Ensure UTF-8 output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PARENT_DIR = r"c:\JOB SEARCH"
BASE_DOCX = os.path.join(PARENT_DIR, "AJITH SHAJAN - Resume EDit.docx")
HITACHI_DIR = os.path.join(PARENT_DIR, "HitachiEnergy")
OUTPUT_DOCX = os.path.join(HITACHI_DIR, "AJITH SHAJAN - Resume (Hitachi Energy GET).docx")
OUTPUT_PDF = os.path.join(HITACHI_DIR, "AJITH SHAJAN - Resume (Hitachi Energy GET).pdf")
OUTPUT_PDF_STD = os.path.join(HITACHI_DIR, "AJITH SHAJAN - Resume.pdf")
SCRIPT_IN_PARENT = os.path.join(PARENT_DIR, "format_hitachi_energy_resume.py")

os.makedirs(HITACHI_DIR, exist_ok=True)

script_content = r'''import sys, os
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document
from docx.shared import Pt

doc = Document("AJITH SHAJAN - Resume EDit.docx")
paras = doc.paragraphs

def set_run(run, text): run.text = text
def clear_runs_from(para, start_idx):
    for r in para.runs[start_idx:]: r.text = ""

# PARA 1 - Technical Skills line 1 (Embedded Systems, Firmware & Languages, 94 chars)
p1 = paras[1]
set_run(p1.runs[4], "Embedded C, Modern C++, Python, FreeRTOS, STM32 (ARM Cortex-M4), ESP32, Bare-Metal Drivers")
clear_runs_from(p1, 5)

# PARA 2 - Continuation (Hardware, Protocols, Instrumentation, 93 chars)
p2 = paras[2]
set_run(p2.runs[0], "KiCad PCB Layout, CAN FD, SPI, I2C, UART, LoRa (RF), DSOs, Logic Analyzers, ISO 7637-2, DFMEA")
clear_runs_from(p2, 1)

# PARA 5 - Interests (Embedded Systems, Industrial IoT, Telemetry, Power Systems, 81 chars)
p5 = paras[5]
set_run(p5.runs[4], "Embedded Firmware, Industrial IoT & Telemetry, Hardware Validation, Power Systems")
clear_runs_from(p5, 5)

# PARA 9 - Education CBSE 12th percentage
p9 = paras[9]
set_run(p9.runs[6], "90%")

# PARA 11 - Education CBSE 10th percentage
p11 = paras[11]
set_run(p11.runs[4], "93%")

# PARA 16 - ICFOSS Tech Used (88 chars)
p16 = paras[16]
set_run(p16.runs[1], "Embedded C, ESP32S3, FreeRTOS, Capacitive Touch Sensing, Hardware Abstraction, DSOs")
clear_runs_from(p16, 2)

# PARA 17 - ICFOSS Description (3 lines calibrated)
p17 = paras[17]
set_run(p17.runs[0],
    "Engineered real-time embedded firmware and digital sensor interfaces on ESP32S3 under FreeRTOS. "
    "Designed capacitive touch sensing signal conditioning circuits and deterministic state-machine logic. "
    "Validated hardware-software timing and signal integrity using digital storage oscilloscopes (DSOs) and logic analyzers.\t"
)
clear_runs_from(p17, 1)

# ------ PROJECT SLOT 1: Automotive BLE-to-CAN Hardware Gateway ------
p21 = paras[21]
set_run(p21.runs[0], "Automotive BLE-to-CAN Hardware Gateway")
set_run(p21.runs[1], "")
set_run(p21.runs[2], " ")
set_run(p21.runs[8], "1")

p22 = paras[22]
set_run(p22.runs[0], "Hardware Lead")
p22.runs[0].bold = True
set_run(p22.runs[1], "\t")
set_run(p22.runs[2], "2weeks")
p22.runs[2].italic = True
clear_runs_from(p22, 3)

p23 = paras[23]
can_tech = "KiCad 10.0 (4-Layer PCB), TCAN4550-Q1 (CAN FD), CC2340R5 (BLE 5.3), LM5164, TPS2116, ISO 7637-2, DFMEA"
set_run(p23.runs[0], "Technologies Used")
p23.runs[0].bold = True
p23.runs[0].italic = True
set_run(p23.runs[1], ":")
p23.runs[1].bold = True
p23.runs[1].italic = True
set_run(p23.runs[2], " ")
p23.runs[2].italic = True
set_run(p23.runs[3], can_tech)
p23.runs[3].bold = False
p23.runs[3].italic = True
clear_runs_from(p23, 4)

p24 = paras[24]
set_run(p24.runs[0],
    "Architected an industrial 4-layer (ENIG, 50\u03a9 CPWG, 90\u03a9 diff) BLE 5.3 to CAN FD gateway with 9V\u201336V ISO 7637-2 power conditioning. "
    "Engineered multi-rail power with LM5164 buck and TPS2116 <2\u03bcs battery failover; conducted formal AIAG/VDA DFMEA. "
    "Released manufacturing Gerbers and BOM (73 items, 109 components) with 0 DRC violations and 0 ERC violations."
)
clear_runs_from(p24, 1)

# ------ PROJECT SLOT 2: Custom STM32 Dev Board ------
p26 = paras[26]
set_run(p26.runs[0], "Custom STM32 Development Board & Bench Bring-Up ")
set_run(p26.runs[1], "\t")
set_run(p26.runs[2], "\t")
set_run(p26.runs[3], "Team ")
set_run(p26.runs[4], "S")
set_run(p26.runs[5], "ize")
set_run(p26.runs[6], ":")
set_run(p26.runs[7], " ")
set_run(p26.runs[8], "1")

p27 = paras[27]
set_run(p27.runs[0], "Hardware Designer")
p27.runs[0].bold = True
set_run(p27.runs[1], "\t")
set_run(p27.runs[2], "\t")
set_run(p27.runs[3], "2weeks")
p27.runs[3].italic = True

p28 = paras[28]
stm_tech = "KiCad (2-Layer), STM32F401RE (ARM Cortex-M4), 3.3V LDO, FreeRTOS, SPI, I2C, UART, SWD Debug"
set_run(p28.runs[0], "Technologies Use")
p28.runs[0].bold = True
p28.runs[0].italic = True
set_run(p28.runs[1], "d:")
p28.runs[1].bold = True
p28.runs[1].italic = True
set_run(p28.runs[2], " ")
p28.runs[2].italic = True
set_run(p28.runs[3], stm_tech)
p28.runs[3].bold = False
p28.runs[3].italic = True

p29 = paras[29]
set_run(p29.runs[0],
    "Designed and laid out a 2-layer STM32F401RE development board with LDO regulation and crystal routing (0 DRC violations and 0 ERC violations). "
    "Brought up bare-metal firmware under FreeRTOS; verified SPI, I2C, and UART bus timings on bench using digital storage oscilloscope (DSO) and logic analyzer."
)
clear_runs_from(p29, 1)

# ------ PROJECT SLOT 3: Edge AI Dynamic Braille System ------
p30 = paras[30]
set_run(p30.runs[0], "Edge AI based Dynamic Braille System")
set_run(p30.runs[1], "\t")
set_run(p30.runs[2], "\tTeam Size")
set_run(p30.runs[3], ":")
set_run(p30.runs[4], " ")
set_run(p30.runs[5], "4")

p31 = paras[31]
set_run(p31.runs[0], "Hardware Lead")
p31.runs[0].bold = True
set_run(p31.runs[1], "\t")
set_run(p31.runs[2], "\t")
set_run(p31.runs[3], "1week")
p31.runs[3].italic = True
clear_runs_from(p31, 4)

p32 = paras[32]
braille_tech = "Raspberry Pi 5, Arduino, Python, Edge ML (Gemma 2B), PaddleOCR, SG90 Servos, UART, TinyML"
set_run(p32.runs[0], "Technologies Use")
p32.runs[0].bold = True
p32.runs[0].italic = True
set_run(p32.runs[1], "d:")
p32.runs[1].bold = True
p32.runs[1].italic = True
set_run(p32.runs[2], " ")
p32.runs[2].italic = True
set_run(p32.runs[3], braille_tech)
p32.runs[3].bold = False
p32.runs[3].italic = True

p33 = paras[33]
set_run(p33.runs[0],
    "Architected hardware actuation interfacing Raspberry Pi 5 via UART to an ESP32 controlling servo arrays for tactile Braille output, integrating bilingual OCR and quantized on-device LLM."
)
clear_runs_from(p33, 1)

# ------ PROJECT SLOT 4: Landslide Tracker (Remote RF Telemetry & Sensor Fusion) ------
p34 = paras[34]
set_run(p34.runs[0], "Landslide Tracker")
set_run(p34.runs[1], " and Response ")
set_run(p34.runs[2], "System")

p35 = paras[35]
set_run(p35.runs[0], "Hardware Designer")
p35.runs[0].bold = True
set_run(p35.runs[1], "\t")
set_run(p35.runs[2], "\t")
set_run(p35.runs[3], "4days")
p35.runs[3].italic = True

p36 = paras[36]
ls_tech = "ESP-32, LoRa Module (Ra-02 433 MHz), Embedded C, Sensor Fusion, Anomaly Detection, Low-Power RF"
ls_desc = (
    "Designed an IoT sensing node performing multi-sensor fusion and threshold anomaly detection in C, cross-verifying telemetry over a 433 MHz LoRa mesh for real-time alerting."
)
set_run(p36.runs[1], ls_tech)
set_run(p36.runs[2], ".                                                                ")
set_run(p36.runs[4], ls_desc)
p36.runs[4].bold = False
p36.runs[4].italic = False
clear_runs_from(p36, 5)

# Header cleanup
sec = doc.sections[0]
for r in list(sec.header.paragraphs[1].runs):
    sec.header.paragraphs[1]._p.remove(r._r)
sec.header.paragraphs[1].text = ""

# Page 2 top spacing cleanup
ach_idx = None
ach_p = None
for i, p in enumerate(doc.paragraphs):
    if "ACHIEVEMENTS AND ACTIVITIES" in p.text:
        ach_idx = i
        ach_p = p
        break

if ach_idx is not None:
    for i in range(ach_idx - 1, 0, -1):
        prev_p = doc.paragraphs[i]
        if not prev_p.text.strip():
            prev_p._p.getparent().remove(prev_p._p)
        else:
            break
    ach_p.paragraph_format.page_break_before = True

# Enforce 9.0 pt across all body runs
for p in doc.paragraphs:
    for r in p.runs:
        r.font.size = Pt(9.0)

os.makedirs("HitachiEnergy", exist_ok=True)
doc.save("HitachiEnergy/AJITH SHAJAN - Resume (Hitachi Energy GET).docx")
print("Saved cleanly: HitachiEnergy/AJITH SHAJAN - Resume (Hitachi Energy GET).docx")
'''

with open(SCRIPT_IN_PARENT, "w", encoding="utf-8") as f:
    f.write(script_content)
print(f"Written format script: {SCRIPT_IN_PARENT}")

res = subprocess.run([sys.executable, SCRIPT_IN_PARENT], cwd=PARENT_DIR, capture_output=True, text=True, encoding="utf-8")
print(res.stdout)
if res.stderr:
    print("STDERR:", res.stderr)

# Convert to PDF
sys.path.insert(0, os.path.join(PARENT_DIR, "AI TOOL", "scripts"))
from convert_docx_to_pdf import convert
success = convert(OUTPUT_DOCX, OUTPUT_PDF)
if success:
    print(f"Generated PDF: {OUTPUT_PDF}")
    shutil.copy(OUTPUT_PDF, OUTPUT_PDF_STD)
    print(f"Copied to: {OUTPUT_PDF_STD}")

# Verify 2 pages
import fitz
doc_pdf = fitz.open(OUTPUT_PDF)
page_count = len(doc_pdf)
print(f"Verified PDF page count: {page_count}")

# Render PNGs
art_dir = r"C:\Users\colon\.gemini\antigravity-ide\brain\e2b3060d-2a4c-4f5e-80de-e591e8db47db"
for i, p in enumerate(doc_pdf):
    pix = p.get_pixmap(dpi=200)
    p_img = os.path.join(HITACHI_DIR, f"page_{i+1}.png")
    art_img = os.path.join(art_dir, f"hitachi_page_{i+1}.png")
    pix.save(p_img)
    pix.save(art_img)
    print(f"Rendered: {p_img}")

# Print bounding box metrics
p1_blocks = doc_pdf[0].get_text("blocks")
p2_blocks = doc_pdf[1].get_text("blocks")
print(f"Page 1 last content y: {p1_blocks[-1][3]:.2f} pt")
print(f"Page 2 first content y: {p2_blocks[0][1]:.2f} pt")

# Create Job Brief
brief_path = os.path.join(HITACHI_DIR, "JOB BRIEF - Hitachi Energy Graduate Engineer Trainee.txt")
brief_text = """================================================================================
JOB BRIEF: HITACHI ENERGY - GRADUATE ENGINEER TRAINEE (GET)
CAMPUS PLACEMENT DRIVE — MODEL ENGINEERING COLLEGE (BATCH OF 2027)
================================================================================

1. POSITION OVERVIEW:
   Company: Hitachi Energy India (https://www.hitachienergy.com)
   Specialization: Power Grids, High Voltage Transmission, Grid Automation, SCADA, Substations, Renewable Energy
   Role: Graduate Engineer Trainee (GET)
   Drive Category: On-Campus Recruitment Drive
   Internship Structure: 6-Month Internship with INR 20,000/month stipend followed by Full-Time GET Placement
   Eligible Batches: B.Tech 2027 Passing Out Batch
   Eligible Disciplines: Electronics & Communication Engineering (ECE), 
                         Electrical & Electronics Engineering (EEE), 
                         Computer Science Engineering (CSE)
   Work Locations: Bangalore / Chennai, India
   Registration Deadline: 11:00 PM on 11th October, 2026

2. CANDIDATE FIT & ELIGIBILITY AUDIT:
   - Minimum Academic Cutoff: CGPA >= 6.0 with 0 active backlogs.
   - Candidate (Ajith Shajan) Profile:
     * CGPA: 8.18 / 10 (Clearance with massive +2.18 CGPA margin)
     * Active Backlogs: 0
     * Branch: Electronics & Communication Engineering (Core Fit for Grid Automation & Hardware)
     * Board Percentages: CBSE 12th (90%), CBSE 10th (93%)
     * Priority Hierarchy: Matches Priority 1 Internship Pathway leading directly to FTE GET.

3. TARGET TECHNICAL COMPETENCIES REQUIRED BY HITACHI ENERGY:
   - Embedded Systems & Microcontrollers: STM32 (ARM Cortex-M4), ESP32, register-level C, interrupt handlers, FreeRTOS.
   - Industrial Communication & Fieldbuses: CAN Bus, CAN FD, Modbus, UART, SPI, I2C, industrial Ethernet, telemetry.
   - Circuit Protection & Hardware Robustness: 4-layer PCB design (KiCad), TVS diode transient suppression, 
     ISO 7637-2 load-dump protection, power distribution networks, LDO regulation, DFMEA risk mitigation.
   - Telemetry & Asset Monitoring: Long-range wireless mesh (433 MHz LoRa), multi-sensor data acquisition, 
     sensor fusion, threshold anomaly detection algorithms.
   - Instrumentation & Bring-Up: Hands-on bench verification using Digital Storage Oscilloscopes (DSOs) and Logic Analyzers.

4. RESUME ALIGNMENT & HIGHLIGHTS:
   - Project 1: Automotive BLE-to-CAN Hardware Gateway (4-layer PCB, ISO 7637-2 surge protection, 9-36V, TVS clamps, DFMEA, 0 DRC violations and 0 ERC violations).
   - Project 2: Custom STM32 Development Board & Bench Bring-Up (Crystal clock routing, LDO decoupling, FreeRTOS, digital storage oscilloscope (DSO) validation, 0 DRC violations and 0 ERC violations).
   - Project 3: Edge AI based Dynamic Braille System (Microcontroller UART actuation, deterministic hardware state machines).
   - Project 4: Landslide Tracker and Response System (433 MHz LoRa RF mesh, remote telemetry, sensor fusion, off-grid power optimization).
   - Core Certifications: Chairperson IEEE CAS MEC SB, FPGA based DSD using Verilog, IBM SkillsBuild Agentic AI.

5. SELECTION PROCESS & TIMELINE:
   - Step 1: Online Campus Registration & Form Submission before 11th October, 2026 (11:00 PM).
   - Step 2: Online Aptitude & Technical Screening Assessment.
   - Step 3: Technical Interview (Core Electrical/Electronics fundamentals, C programming, microcontroller peripherals, PCB layout).
   - Step 4: HR & Behavioral Round.
================================================================================
"""
with open(brief_path, "w", encoding="utf-8") as f:
    f.write(brief_text)
print(f"Created Job Brief: {brief_path}")

# Create Other Openings File
other_path = os.path.join(HITACHI_DIR, "Hitachi Energy - Other Openings.txt")
other_text = """================================================================================
HITACHI ENERGY — ACTIVE OPENINGS & TRACKED ROLES (2026-2027)
================================================================================

1. ON-CAMPUS CURRENT DRIVE (MEC BATCH OF 2027):
   - Role: Graduate Engineer Trainee (GET)
     Structure: 6-Month Internship @ INR 20,000/month + Full-Time GET Placement
     Locations: Bangalore / Chennai

2. EARLY CAREER & LATERAL TRACKS (PORTAL MONITORING):
   - Role: Associate Engineer — Substation Automation & Protection
     Location: Bangalore Technology Center
     Portal: https://www.hitachienergy.com/careers
   - Role: R&D Engineer Trainee — Power Electronics & Embedded Firmware
     Location: Chennai / Bangalore
   - Role: Control Systems Trainee — Grid Automation & SCADA
     Location: Vadodara / Bangalore

================================================================================
"""
with open(other_path, "w", encoding="utf-8") as f:
    f.write(other_text)
print(f"Created Other Openings: {other_path}")
