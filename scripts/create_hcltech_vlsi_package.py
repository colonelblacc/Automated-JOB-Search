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

# PARA 1 - Technical Skills line 1 (VLSI, RTL & Digital Hardware biased, 95 chars)
p1 = paras[1]
set_run(p1.runs[4], "Verilog HDL, Digital System Design (DSD), FPGA Synthesis, STM32 (ARM Cortex-M4), KiCad PCB Layout")
clear_runs_from(p1, 5)

# PARA 2 - Technical Skills line 2 (Interfaces, Tools & Bring-Up, 98 chars)
p2 = paras[2]
set_run(p2.runs[0], "SPI, I2C, UART, CAN FD, Logic Analyzers, Digital Storage Oscilloscopes (DSOs), Vivado, ModelSim, EDA")
clear_runs_from(p2, 1)

# PARA 5 - Interests (VLSI & Digital focused, 96 chars)
p5 = paras[5]
set_run(p5.runs[4], "VLSI Design, ASIC / FPGA RTL Design, Digital Hardware Architecture, Silicon Bring-Up, Embedded Systems")
clear_runs_from(p5, 5)

# PARA 9 - Education CBSE 12th percentage
p9 = paras[9]
set_run(p9.runs[6], "90%")

# PARA 11 - Education CBSE 10th percentage
p11 = paras[11]
set_run(p11.runs[4], "93%")

# PARA 16 - ICFOSS Tech Used (102 chars)
p16 = paras[16]
set_run(p16.runs[1], "Embedded C, ESP32S3, PCB Design, Capacitive Touch Sensing, Digital Logic State Machines, Lab Bring-Up")
clear_runs_from(p16, 2)

# PARA 17 - ICFOSS Description (3 lines calibrated)
p17 = paras[17]
set_run(p17.runs[0],
    "Engineered real-time embedded firmware and digital sensor interfaces on ESP32S3 microcontrollers under FreeRTOS. "
    "Designed capacitive touch sensing signal conditioning circuits and deterministic state-machine logic. "
    "Debugged hardware-software interactions and signal integrity using digital storage oscilloscopes (DSOs) and logic analyzers.\t"
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
set_run(p22.runs[1], "\t")
set_run(p22.runs[2], "2weeks")
p22.runs[0].bold = True
p22.runs[2].italic = True

p23 = paras[23]
hw_gw_tech = "KiCad 10.0 (4-Layer PCB, ENIG), TCAN4550-Q1 (CAN FD), CC2340R5 (BLE 5.3), LM5164 Buck, TPS2116, TVS, DFMEA"
set_run(p23.runs[0], "Technologies Used")
p23.runs[0].bold = True
p23.runs[0].italic = True
set_run(p23.runs[1], ":")
p23.runs[1].bold = True
p23.runs[1].italic = True
set_run(p23.runs[2], " ")
p23.runs[2].italic = True
set_run(p23.runs[3], hw_gw_tech)
p23.runs[3].bold = False
p23.runs[3].italic = True
clear_runs_from(p23, 4)

p24 = paras[24]
set_run(p24.runs[0],
    "Architected an automotive 4-layer (ENIG, 50\u03a9 CPWG, 90\u03a9 diff) BLE 5.3 to CAN FD gateway (9\u201336V ISO 7637-2). "
    "Engineered multi-rail power with LM5164 buck, P-FET reverse clamp, and TPS2116 <2\u03bcs battery failover. "
    "Generated production-ready Gerbers and BOM (73 line items, 109 components) with 0 DRC violations and 0 ERC violations."
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
stm_tech = "KiCad (2-Layer), STM32F401RE, 3.3V LDO, FreeRTOS, SPI, I2C, UART, SWD Debug"
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
    "Designed and laid out a 2-layer STM32F401RE development board with LDO regulation and crystal routing (0 DRC/ERC). "
    "Brought up bare-metal with FreeRTOS; validated SPI, I2C, UART using digital storage oscilloscope (DSO) and logic analyzer."
)
clear_runs_from(p29, 1)

# ------ PROJECT SLOT 3: Edge AI Dynamic Braille (Interfacing & Digital Logic) ------
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

# ------ PROJECT SLOT 4: Landslide Tracker ------
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

    import fitz
    doc_pdf = fitz.open(OUTPUT_PDF)
    print(f"Verified PDF page count: {len(doc_pdf)}")
    for i, page in enumerate(doc_pdf):
        pix = page.get_pixmap(dpi=150)
        png_path = os.path.join(HCL_DIR, f"page_{i+1}.png")
        pix.save(png_path)
        print(f"Rendered: {png_path}")

    p1 = doc_pdf[0]
    p2 = doc_pdf[1]
    blocks1 = p1.get_text("blocks")
    blocks2 = p2.get_text("blocks")
    meaningful1 = [b for b in blocks1 if b[4].strip()]
    meaningful2 = [b for b in blocks2 if b[4].strip()]
    p1_last_y = max(b[3] for b in meaningful1)
    p2_first_y = min(b[1] for b in meaningful2)
    print(f"\nPage 1 last content y: {p1_last_y:.2f} pt")
    print(f"Page 2 first content y: {p2_first_y:.2f} pt")
else:
    print("PDF conversion failed!")

# Create Job Brief
brief_path = os.path.join(HCL_DIR, "JOB BRIEF - HCL Technologies Graduate Engineer Trainee (VLSI).txt")
brief_content = """================================================================================
JOB BRIEF: HCL TECHNOLOGIES — GRADUATE ENGINEER TRAINEE (VLSI / SILICON ENGINEERING)
================================================================================

1. POSITION DETAILS
--------------------------------------------------------------------------------
- Company: HCL Technologies (HCLTech)
- Target Division: Engineering and R&D Services (ERS) — Semiconductor & Silicon Engineering
- Role: Graduate Engineer Trainee (GET) — VLSI Role (Electronics Branches Exclusive)
- Drive: Campus Placement Drive (MEC Batch of 2027)
- Slot: B1
- Pay: ₹4.5 LPA
- Bond: 1 Year
- Eligibility Criteria: 7.0 CGPA and above with no active backlogs
- Eligible Branches: ECE / EV / EEE (Mechanical evaluated separately, CSE excluded from VLSI)
- Locations: PAN India (Key Semiconductor / VLSI Hubs: Bangalore, Chennai, Noida, Hyderabad)
- Official Portal: https://www.hcltech.com/careers

2. ROLE CONTEXT & VLSI TECHNICAL SCOPE
--------------------------------------------------------------------------------
HCLTech Engineering and R&D Services (ERS) is among the largest global semiconductor engineering
partners, delivering end-to-end silicon solutions:
- ASIC / FPGA Design: RTL design using Verilog / SystemVerilog, logic synthesis, FSM modeling.
- Pre-Silicon Verification: Testbench development, functional simulation, code coverage, UVM methodology.
- Physical Design (PD): Floorplanning, Placement & Routing (P&R), Clock Tree Synthesis (CTS), Static Timing Analysis (STA).
- Silicon Validation & Bring-Up: Post-silicon bench bring-up, lab debugging using DSOs and Logic Analyzers, JTAG/SWD validation.
- Embedded Firmware & Hardware Co-Design: Board support packages (BSP), register-level device drivers, hardware bring-up.

3. CANDIDATE PROFILE HIGHLIGHTS (AJITH SHAJAN)
--------------------------------------------------------------------------------
- Degree: B.Tech in Electronics and Communication Engineering (ECE), Model Engineering College (MEC).
- CGPA: 8.18 (Cutoff: 7.0 CGPA — Cleared with strong distinction).
- Leadership: Chairperson, IEEE Circuits and Systems Society (CAS) MEC Student Branch (Circuits & VLSI focus).
- Core Hardware:
  * Automotive 4-Layer BLE-to-CAN Hardware Gateway (4-layer ENIG PCB, ISO 7637-2, DFMEA, 0 DRC/ERC).
  * Custom STM32F401RE Development Board (Clock routing, LDO regulation, 0 DRC/ERC, DSO bring-up).
- Digital Stack: Verilog HDL, Digital System Design (DSD), SPI/I2C/UART/CAN bus protocols, Logic Analyzers.

4. PREPARATION CHEAT SHEET FOR TECHNICAL ROUNDS
--------------------------------------------------------------------------------
1. Digital Logic & Sequential Circuits:
   - Setup time (T_su), Hold time (T_h), Clock-to-Q delay, Metastability and 2-FF synchronizers.
   - Max operating frequency: F_max = 1 / (T_cq + T_comb + T_su - T_skew).
2. Verilog Fundamentals:
   - Blocking (=) vs Non-blocking (<=) statements; race conditions in synthesis.
   - Mealy FSM (output depends on input and state) vs Moore FSM (output depends only on state).
3. Static Timing Analysis (STA):
   - Setup slack = Required time - Arrival time (Slack >= 0 required).
   - Hold slack = Arrival time - Required time (Hold violation is catastrophic, cannot fix by slowing clock).
4. Semiconductor Physics & CMOS:
   - CMOS Inverter characteristics, dynamic power dissipation P = alpha * C * V^2 * f.

================================================================================
Generated via Career-Ops Command Center | October 2026
================================================================================
"""

with open(brief_path, "w", encoding="utf-8") as f:
    f.write(brief_content)
print(f"Created Job Brief: {brief_path}")

# Create Other Openings file
other_path = os.path.join(HCL_DIR, "HCL Technologies - Other Openings.txt")
other_content = """================================================================================
HCL TECHNOLOGIES (HCLTECH) — RELEVANT ENTRY-LEVEL & EARLY CAREER OPENINGS
================================================================================
1. Graduate Engineer Trainee (GET) - VLSI / ASIC Design (Campus Slot B1) [TARGET]
2. Associate Engineer - Embedded Systems & Firmware (Automotive / Medical ERS)
3. Junior Hardware Design Engineer - Board Bring-up & PCB Diagnostics
4. Silicon Validation Trainee - Lab Bring-Up & Pre-Silicon Emulation
5. Associate Software Engineer - Embedded C / C++ RTOS Devices

Explore continuous off-campus and corporate requirements:
https://www.hcltech.com/careers
================================================================================
"""
with open(other_path, "w", encoding="utf-8") as f:
    f.write(other_content)
print(f"Created Other Openings: {other_path}")
