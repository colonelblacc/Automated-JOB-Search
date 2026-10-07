import sys
import os
import subprocess
import shutil

# Ensure UTF-8 output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PARENT_DIR = r"c:\JOB SEARCH"
BASE_DOCX = os.path.join(PARENT_DIR, "AJITH SHAJAN - Resume EDit.docx")
HCL_DIR = os.path.join(PARENT_DIR, "HCLTech")
OUTPUT_DOCX = os.path.join(HCL_DIR, "AJITH SHAJAN - Resume (HCLTech VLSI).docx")
OUTPUT_PDF = os.path.join(HCL_DIR, "AJITH SHAJAN - Resume (HCLTech VLSI).pdf")
OUTPUT_PDF_STD = os.path.join(HCL_DIR, "AJITH SHAJAN - Resume.pdf")
SCRIPT_IN_PARENT = os.path.join(PARENT_DIR, "format_hcltech_vlsi_resume.py")

os.makedirs(HCL_DIR, exist_ok=True)

script_content = r'''import sys, os
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document
from docx.shared import Pt

doc = Document("AJITH SHAJAN - Resume EDit.docx")
paras = doc.paragraphs

def set_run(run, text): run.text = text
def clear_runs_from(para, start_idx):
    for r in para.runs[start_idx:]: r.text = ""

# PARA 1 - Technical Skills line 1 (94 chars)
p1 = paras[1]
set_run(p1.runs[4], "Verilog HDL, Digital System Design (DSD), RTL Synthesis, Static Timing Analysis, KiCad Layout")
clear_runs_from(p1, 5)

# PARA 2 - Technical Skills line 2 (98 chars)
p2 = paras[2]
set_run(p2.runs[0], "FSM Design, Clock Domain Crossing (CDC), Logic Analyzers, DSOs, ModelSim, Icarus Verilog, Vivado")
clear_runs_from(p2, 1)

# PARA 5 - Interests (80 chars)
p5 = paras[5]
set_run(p5.runs[4], "VLSI Design, ASIC / FPGA RTL Architecture, Silicon Validation, Embedded Systems")
clear_runs_from(p5, 5)

# PARA 9 - Education CBSE 12th percentage
p9 = paras[9]
set_run(p9.runs[6], "90%")

# PARA 11 - Education CBSE 10th percentage
p11 = paras[11]
set_run(p11.runs[4], "93%")

# PARA 16 - ICFOSS Tech Used (88 chars)
p16 = paras[16]
set_run(p16.runs[1], "Embedded C, ESP32S3, PCB Design, Capacitive Sensing, Digital Logic FSMs, Lab Bring-Up")
clear_runs_from(p16, 2)

# PARA 17 - ICFOSS Description (3 lines calibrated)
p17 = paras[17]
set_run(p17.runs[0],
    "Engineered real-time embedded firmware and digital sensor interfaces on ESP32S3 under FreeRTOS. "
    "Designed capacitive touch sensing signal conditioning circuits and deterministic state-machine logic. "
    "Validated hardware-software timing and signal integrity using DSOs and logic analyzers.\t"
)
clear_runs_from(p17, 1)

# ------ PROJECT SLOT 1: Dual-Clock Asynchronous FIFO ------
p21 = paras[21]
set_run(p21.runs[0], "Dual-Clock Asynchronous FIFO Core")
set_run(p21.runs[1], "")
set_run(p21.runs[2], "")
set_run(p21.runs[3], "\tTeam ")
set_run(p21.runs[8], "1")

p22 = paras[22]
set_run(p22.runs[0], "RTL Design Lead")
set_run(p22.runs[1], "\t")
set_run(p22.runs[2], "2weeks")
p22.runs[0].bold = True
p22.runs[2].italic = True

p23 = paras[23]
fifo_tech = "Verilog HDL, Clock Domain Crossing (CDC), Gray Code, 2-FF Synchronizers, Icarus Verilog, GTKWave"
set_run(p23.runs[0], "Technologies Used")
p23.runs[0].bold = True
p23.runs[0].italic = True
set_run(p23.runs[1], ":")
p23.runs[1].bold = True
p23.runs[1].italic = True
set_run(p23.runs[2], " ")
p23.runs[2].italic = True
set_run(p23.runs[3], fifo_tech)
p23.runs[3].bold = False
p23.runs[3].italic = True
clear_runs_from(p23, 4)

p24 = paras[24]
set_run(p24.runs[0],
    "Designed a parameterized dual-clock FIFO in Verilog HDL to resolve clock domain crossing (CDC) across asynchronous clock domains. "
    "Implemented 2-stage flip-flop synchronizers and binary-to-Gray converters to eliminate metastability. "
    "Engineered wrap-around full/empty flag logic; verified via self-checking testbenches."
)
clear_runs_from(p24, 1)

# ------ PROJECT SLOT 2: Configurable UART Controller Core ------
p26 = paras[26]
set_run(p26.runs[0], "Configurable UART Controller Core ")
set_run(p26.runs[1], "\t")
set_run(p26.runs[2], "\t")
set_run(p26.runs[3], "Team ")
set_run(p26.runs[4], "S")
set_run(p26.runs[5], "ize")
set_run(p26.runs[6], ":")
set_run(p26.runs[7], " ")
set_run(p26.runs[8], "1")
clear_runs_from(p26, 9)

p27 = paras[27]
set_run(p27.runs[0], "RTL Design Engineer")
p27.runs[0].bold = True
set_run(p27.runs[1], "\t")
set_run(p27.runs[2], "\t")
set_run(p27.runs[3], "2weeks")
p27.runs[3].italic = True
clear_runs_from(p27, 4)

p28 = paras[28]
uart_tech = "Verilog HDL, FSM Design (Moore/Mealy), Baud Division, 16x Oversampling, RTL Simulation"
set_run(p28.runs[0], "Technologies Use")
p28.runs[0].bold = True
p28.runs[0].italic = True
set_run(p28.runs[1], "d:")
p28.runs[1].bold = True
p28.runs[1].italic = True
set_run(p28.runs[2], " ")
p28.runs[2].italic = True
set_run(p28.runs[3], uart_tech)
p28.runs[3].bold = False
p28.runs[3].italic = True

p29 = paras[29]
set_run(p29.runs[0],
    "Architected a modular UART core in Verilog HDL with Moore/Mealy FSMs for start, data, parity, and stop bit sequencing. "
    "Implemented a parameterized baud divider with 16x receiver oversampling; validated framing error detection in simulation."
)
clear_runs_from(p29, 1)

# ------ PROJECT SLOT 3: Automotive BLE-to-CAN Hardware Gateway ------
p30 = paras[30]
set_run(p30.runs[0], "Automotive BLE-to-CAN Hardware Gateway")
set_run(p30.runs[1], "\t")
set_run(p30.runs[2], "\tTeam Size")
set_run(p30.runs[3], ":")
set_run(p30.runs[4], " ")
set_run(p30.runs[5], "1")

p31 = paras[31]
set_run(p31.runs[0], "Hardware Lead")
p31.runs[0].bold = True
set_run(p31.runs[1], "\t")
set_run(p31.runs[2], "\t")
set_run(p31.runs[3], "2weeks")
p31.runs[3].italic = True
clear_runs_from(p31, 4)

p32 = paras[32]
can_tech = "KiCad 10.0 (4-Layer PCB), TCAN4550-Q1 (CAN FD), CC2340R5 (BLE 5.3), LM5164, TPS2116, ISO 7637-2, DFMEA"
set_run(p32.runs[0], "Technologies Use")
p32.runs[0].bold = True
p32.runs[0].italic = True
set_run(p32.runs[1], "d:")
p32.runs[1].bold = True
p32.runs[1].italic = True
set_run(p32.runs[2], " ")
p32.runs[2].italic = True
set_run(p32.runs[3], can_tech)
p32.runs[3].bold = False
p32.runs[3].italic = True

p33 = paras[33]
set_run(p33.runs[0],
    "Architected an automotive 4-layer BLE 5.3 to CAN FD gateway with 9V\u201336V ISO 7637-2 power conditioning and TPS2116 failover. "
    "Released production Gerbers and BOM with 0 DRC violations and 0 ERC violations."
)
clear_runs_from(p33, 1)

# ------ PROJECT SLOT 4: Custom STM32 Development Board & Bench Bring-Up ------
p34 = paras[34]
set_run(p34.runs[0], "Custom STM32 Development Board")
set_run(p34.runs[1], " & Bench Bring-Up")
set_run(p34.runs[2], "")

p35 = paras[35]
set_run(p35.runs[0], "Hardware Designer")
p35.runs[0].bold = True
set_run(p35.runs[1], "\t")
set_run(p35.runs[2], "\t")
set_run(p35.runs[3], "2weeks")
p35.runs[3].italic = True

p36 = paras[36]
stm_tech_4 = "KiCad (2-Layer), STM32F401RE (ARM Cortex-M4), 3.3V LDO, Crystal Routing, FreeRTOS, DSOs, Logic Analyzers"
stm_desc_4 = (
    "Designed a 2-layer STM32F401RE development board with crystal routing and LDO decoupling (0 DRC/ERC); validated bare-metal bus timings using digital storage oscilloscopes (DSOs)."
)
set_run(p36.runs[1], stm_tech_4)
set_run(p36.runs[2], ".                                                                ")
set_run(p36.runs[4], stm_desc_4)
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

os.makedirs("HCLTech", exist_ok=True)
doc.save("HCLTech/AJITH SHAJAN - Resume (HCLTech VLSI).docx")
print("Saved cleanly: HCLTech/AJITH SHAJAN - Resume (HCLTech VLSI).docx")
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
    p_img = os.path.join(HCL_DIR, f"page_{i+1}.png")
    art_img = os.path.join(art_dir, f"hcltech_page_{i+1}.png")
    pix.save(p_img)
    pix.save(art_img)
    print(f"Rendered: {p_img}")

# Print bounding box metrics
p1_blocks = doc_pdf[0].get_text("blocks")
p2_blocks = doc_pdf[1].get_text("blocks")
print(f"Page 1 last content y: {p1_blocks[-1][3]:.2f} pt")
print(f"Page 2 first content y: {p2_blocks[0][1]:.2f} pt")

# Create Job Brief
brief_path = os.path.join(HCL_DIR, "JOB BRIEF - HCL Technologies Graduate Engineer Trainee (VLSI).txt")
brief_text = """================================================================================
JOB BRIEF: HCL TECHNOLOGIES - GRADUATE ENGINEER TRAINEE (VLSI ROLE)
CAMPUS PLACEMENT DRIVE — MODEL ENGINEERING COLLEGE (BATCH OF 2027)
================================================================================

1. POSITION OVERVIEW:
   Company: HCL Technologies (HCLTech)
   Role: Graduate Engineer Trainee (GET) — VLSI Role
   Drive Category: On-Campus Recruitment Drive (Slot B1)
   Eligible Batches: B.Tech 2027 Passing Out Batch
   Eligible Disciplines: Electronics & Communication Engineering (ECE), 
                         Electronic & VLSI Engineering (EV), 
                         Electrical & Electronics Engineering (EEE)
                         [Note: Computer Science branches excluded from VLSI post]
   Compensation (CTC): INR 4.5 LPA
   Service Agreement: 1 Year Service Bond

2. CANDIDATE FIT & ELIGIBILITY:
   - Minimum Academic Cutoff: CGPA >= 7.0 with 0 active backlogs.
   - Candidate (Ajith Shajan) Profile:
     * CGPA: 8.18 / 10 (Clearance with safety buffer)
     * Active Backlogs: 0
     * Branch: Electronics & Communication Engineering (Core Fit)
     * Board Percentages: CBSE 12th (90%), CBSE 10th (93%)

3. TARGET TECHNICAL COMPETENCIES REQUIRED BY HCLTECH VLSI:
   - Digital System Design (DSD): Combinational & Sequential Logic, FSMs (Moore & Mealy).
   - HDL Programming: Verilog HDL, RTL modeling, structural/behavioral synthesis.
   - Clock Domain Crossing (CDC) & Timing: Metastability resolution, 2-FF synchronizers, Gray code, setup/hold constraints.
   - Asynchronous FIFO Design: Dual-clock domain pointers, empty/full generation.
   - Communication Protocols: UART, SPI, I2C, CAN, AXI/APB bus interfacing.
   - Verification Fundamentals: Testbench design, assertion basics, waveform inspection (ModelSim / GTKWave).
   - Physical Validation: PCB layout, signal integrity, bench instrumentation (DSOs, Logic Analyzers).

4. RESUME ALIGNMENT & HIGHLIGHTS:
   - Project 1: Dual-Clock Asynchronous FIFO with Gray-Code Synchronization (RTL Verilog, CDC, 2-FF, testbench).
   - Project 2: Configurable UART Controller Core & Baud Rate Generator (Moore/Mealy FSMs, 16x oversampling).
   - Project 3: Automotive BLE-to-CAN Hardware Gateway (4-layer PCB, 0 DRC/ERC, ISO 7637-2).
   - Project 4: Custom STM32 Development Board & Bench Bring-Up (Crystal routing, DSO/logic analyzer validation).
   - Core Certifications: FPGA based DSD using Verilog (C2S & MEC), Chairperson IEEE CAS MEC SB.

5. SELECTION PROCESS & TIMELINE:
   - Round 1: Online Technical & Aptitude Assessment (Digital Electronics, Verilog, C, Quantitative Aptitude).
   - Round 2: Technical Interview (FSM state diagrams, FIFO pointer math, setup/hold time violations, Verilog coding).
   - Round 3: HR & Leadership Interview.
================================================================================
"""
with open(brief_path, "w", encoding="utf-8") as f:
    f.write(brief_text)
print(f"Created Job Brief: {brief_path}")

# Create Other Openings File
other_path = os.path.join(HCL_DIR, "HCL Technologies - Other Openings.txt")
other_text = """================================================================================
HCL TECHNOLOGIES — ACTIVE OPENINGS & TRACKED ROLES (2026-2027)
================================================================================

1. ON-CAMPUS CURRENT DRIVE:
   - Role: Graduate Engineer Trainee (GET) — VLSI
   - Compensation: INR 4.5 LPA (Slot B1, 1-Year Service Bond)
   - Status: Active Campus Placement Registration (MEC Batch of 2027)

2. OFF-CAMPUS / LATERAL EARLY CAREER TRACKS (MONITORING):
   - Role: Associate Engineer — Silicon / Hardware Engineering
     URL: https://www.hcltech.com/careers
     Location: Bangalore / Chennai / Noida
   - Role: Graduate Engineer Trainee — Embedded & Automotive Systems
     Location: Bangalore / Hyderabad
   - Early Careers Hub: https://www.hcltech.com/careers/early-careers

================================================================================
"""
with open(other_path, "w", encoding="utf-8") as f:
    f.write(other_text)
print(f"Created Other Openings: {other_path}")
