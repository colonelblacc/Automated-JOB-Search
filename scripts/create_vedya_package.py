import sys
import os
import subprocess
import shutil

# Ensure UTF-8 output
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PARENT_DIR = r"c:\JOB SEARCH"
BASE_DOCX = os.path.join(PARENT_DIR, "AJITH SHAJAN - Resume EDit.docx")
VEDYA_DIR = os.path.join(PARENT_DIR, "VedyaLabs")
OUTPUT_DOCX = os.path.join(VEDYA_DIR, "AJITH SHAJAN - Resume (Vedya Labs Edge AI).docx")
OUTPUT_PDF = os.path.join(VEDYA_DIR, "AJITH SHAJAN - Resume (Vedya Labs Edge AI).pdf")
OUTPUT_PDF_STD = os.path.join(VEDYA_DIR, "AJITH SHAJAN - Resume.pdf")
SCRIPT_IN_PARENT = os.path.join(PARENT_DIR, "format_vedya_edge_ai_resume.py")

os.makedirs(VEDYA_DIR, exist_ok=True)

script_content = r'''import sys, os
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document
from docx.shared import Pt

doc = Document("AJITH SHAJAN - Resume EDit.docx")
paras = doc.paragraphs

def set_run(run, text): run.text = text
def clear_runs_from(para, start_idx):
    for r in para.runs[start_idx:]: r.text = ""

# PARA 1 - Technical Skills line 1 (Edge AI, ML & Embedded Languages, 94 chars)
p1 = paras[1]
set_run(p1.runs[4], "Python, Modern C++, Embedded C, PyTorch, TinyML, TFLite Micro, OpenCV, YOLOv8, Gemma 2B, FreeRTOS")
clear_runs_from(p1, 5)

# PARA 2 - Continuation (Hardware, Edge Neural Engines & Bring-Up, 95 chars)
p2 = paras[2]
set_run(p2.runs[0], "Raspberry Pi 5, ESP32S3, STM32, ESP-SR Neural AI, TensorRT/ONNX Basics, KiCad, DSOs, Logic Analyzers")
clear_runs_from(p2, 1)

# PARA 5 - Interests (Edge AI, Physical AI & Robotics, 80 chars)
p5 = paras[5]
set_run(p5.runs[4], "Edge AI & TinyML, Physical AI, Computer Vision, Embedded Systems, Robotics")
clear_runs_from(p5, 5)

# PARA 9 - Education CBSE 12th percentage
p9 = paras[9]
set_run(p9.runs[6], "90%")

# PARA 11 - Education CBSE 10th percentage
p11 = paras[11]
set_run(p11.runs[4], "93%")

# PARA 16 - ICFOSS Tech Used (88 chars)
p16 = paras[16]
set_run(p16.runs[1], "Embedded C, ESP32S3, FreeRTOS, Capacitive Touch Sensing, Edge Hardware Bring-Up, DSOs")
clear_runs_from(p16, 2)

# PARA 17 - ICFOSS Description (3 lines calibrated)
p17 = paras[17]
set_run(p17.runs[0],
    "Engineered real-time embedded firmware and digital sensor interfaces on ESP32S3 under FreeRTOS. "
    "Designed capacitive touch sensing signal conditioning circuits and deterministic state-machine logic. "
    "Validated hardware-software timing and signal integrity using DSOs and logic analyzers.\t"
)
clear_runs_from(p17, 1)

# ------ PROJECT SLOT 1: Edge AI Dynamic Braille System ------
p21 = paras[21]
set_run(p21.runs[0], "Edge AI based Dynamic Braille System")
set_run(p21.runs[1], "")
set_run(p21.runs[2], "")
set_run(p21.runs[3], "\tTeam ")
set_run(p21.runs[8], "4")

p22 = paras[22]
set_run(p22.runs[0], "Edge AI Developer")
p22.runs[0].bold = True
set_run(p22.runs[1], "\t")
set_run(p22.runs[2], "1week")
p22.runs[2].italic = True
clear_runs_from(p22, 3)

p23 = paras[23]
p23_tech = "Raspberry Pi 5, Arduino, Python, Edge ML (Gemma 2B), PaddleOCR, Tesseract, Multiprocessing, UART"
set_run(p23.runs[0], "Technologies Used")
p23.runs[0].bold = True
p23.runs[0].italic = True
set_run(p23.runs[1], ":")
p23.runs[1].bold = True
p23.runs[1].italic = True
set_run(p23.runs[2], " ")
p23.runs[2].italic = True
set_run(p23.runs[3], p23_tech)
p23.runs[3].bold = False
p23.runs[3].italic = True
clear_runs_from(p23, 4)

p24 = paras[24]
set_run(p24.runs[0],
    "Architected an on-device multimodal AI pipeline combining bilingual OCR (PaddleOCR) and quantized local LLM inference "
    "(Gemma 2B) on Embedded Linux. Optimized token latency, multithreaded queues, and serialized control outputs "
    "to physical Braille actuators for real-time tactile output."
)
clear_runs_from(p24, 1)

# ------ PROJECT SLOT 2: Adaptive LED Shadow Mapping (Physical AI / Vision) ------
p26 = paras[26]
set_run(p26.runs[0], "Adaptive LED Shadow Mapping ")
set_run(p26.runs[1], "\t")
set_run(p26.runs[2], "\t")
set_run(p26.runs[3], "Team ")
set_run(p26.runs[4], "S")
set_run(p26.runs[5], "ize")
set_run(p26.runs[6], ":")
set_run(p26.runs[7], " ")
set_run(p26.runs[8], "3")
clear_runs_from(p26, 9)

p27 = paras[27]
set_run(p27.runs[0], "Computer Vision Lead")
p27.runs[0].bold = True
set_run(p27.runs[1], "\t")
set_run(p27.runs[2], "1week")
p27.runs[2].italic = True
clear_runs_from(p27, 3)

p28 = paras[28]
p28_tech = "Python, YOLOv8, OpenCV, PyTorch, Raspberry Pi 5, WS2812B LED Matrix, Edge Inference"
set_run(p28.runs[0], "Technologies Use")
p28.runs[0].bold = True
p28.runs[0].italic = True
set_run(p28.runs[1], "d:")
p28.runs[1].bold = True
p28.runs[1].italic = True
set_run(p28.runs[2], " ")
p28.runs[2].italic = True
set_run(p28.runs[3], p28_tech)
p28.runs[3].bold = False
p28.runs[3].italic = True

p29 = paras[29]
set_run(p29.runs[0],
    "Engineered real-time vehicle detection and LED shadow mapping using optimized YOLOv8 on edge hardware. "
    "Developed coordinate mapping algorithms translating camera pixels to angular sectors of addressable LEDs with zero visual flicker."
)
clear_runs_from(p29, 1)

# ------ PROJECT SLOT 3: Automated Lecture Tracing System (Acoustic Neural AI) ------
p30 = paras[30]
set_run(p30.runs[0], "Automated Lecture Tracing System")
set_run(p30.runs[1], "\t")
set_run(p30.runs[2], "\tTeam Size")
set_run(p30.runs[3], ":")
set_run(p30.runs[4], " ")
set_run(p30.runs[5], "4")

p31 = paras[31]
set_run(p31.runs[0], "Embedded AI Lead")
p31.runs[0].bold = True
set_run(p31.runs[1], "\t")
set_run(p31.runs[2], "\t")
set_run(p31.runs[3], "1Month")
p31.runs[3].italic = True
clear_runs_from(p31, 4)

p32 = paras[32]
p32_tech = "C++, Python, ESP32S3, ESP-SR (Neural Audio AI), FreeRTOS, FastAPI, WebSocket, UART"
set_run(p32.runs[0], "Technologies Use")
p32.runs[0].bold = True
p32.runs[0].italic = True
set_run(p32.runs[1], "d:")
p32.runs[1].bold = True
p32.runs[1].italic = True
set_run(p32.runs[2], " ")
p32.runs[2].italic = True
set_run(p32.runs[3], p32_tech)
p32.runs[3].bold = False
p32.runs[3].italic = True

p33 = paras[33]
set_run(p33.runs[0],
    "Developed an intelligent embedded audio acquisition system using ESP32S3 streaming over WebSocket, "
    "featuring on-device acoustic neural wake-word detection (ESP-SR), DMA-driven I2S sampling, and FreeRTOS task scheduling."
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

os.makedirs("VedyaLabs", exist_ok=True)
doc.save("VedyaLabs/AJITH SHAJAN - Resume (Vedya Labs Edge AI).docx")
print("Saved cleanly: VedyaLabs/AJITH SHAJAN - Resume (Vedya Labs Edge AI).docx")
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
    p_img = os.path.join(VEDYA_DIR, f"page_{i+1}.png")
    art_img = os.path.join(art_dir, f"vedya_page_{i+1}.png")
    pix.save(p_img)
    pix.save(art_img)
    print(f"Rendered: {p_img}")

# Print bounding box metrics
p1_blocks = doc_pdf[0].get_text("blocks")
p2_blocks = doc_pdf[1].get_text("blocks")
print(f"Page 1 last content y: {p1_blocks[-1][3]:.2f} pt")
print(f"Page 2 first content y: {p2_blocks[0][1]:.2f} pt")

# Create Job Brief
brief_path = os.path.join(VEDYA_DIR, "JOB BRIEF - Vedya Labs Edge AI Engineer.txt")
brief_text = """================================================================================
JOB BRIEF: VEDYA LABS - EDGE AI ENGINEER
CAMPUS PLACEMENT DRIVE — MODEL ENGINEERING COLLEGE (BATCH OF 2027)
================================================================================

1. POSITION OVERVIEW:
   Company: Vedya Labs Private Limited (https://vedya.ai)
   Specialization: Edge AI, Embedded AI, Physical AI & Enterprise AI Solutions
   Industries Served: Semiconductors, Robotics, IoT & Industrial Automation
   Silicon / Hardware Ecosystem: Renesas, Microchip, NPUs, DSPs, Embedded GPUs & MCUs
   Role: Edge AI Engineer (Primary Post)
   Alternative Post: Robotics Engineer
   Drive Category: On-Campus Recruitment Drive (Slot B1/B2 Variable)
   Internship Offer: 6-Month Internship with INR 25,000/month stipend
   Full-Time Compensation: INR 5.0 - 7.0 LPA upon successful internship completion
   Eligible Batches: B.Tech 2027 Passing Out Batch
   Eligible Branches: CSE / CSBS / ECE / EV / EEE / ME
   Work Location: Hyderabad, Telangana, India (Primary Hub)
   Application Deadline: 10:00 PM on 8th October, 2026 (Tomorrow)

2. CANDIDATE FIT & ELIGIBILITY AUDIT:
   - Minimum Academic Cutoff: CGPA >= 7.5 with 0 active backlogs.
   - Candidate (Ajith Shajan) Profile:
     * CGPA: 8.18 / 10 (Clearance with robust safety margin)
     * Active Backlogs: 0
     * Branch: Electronics & Communication Engineering (ECE) with Minor in Machine Learning
     * Board Percentages: CBSE 12th (90%), CBSE 10th (93%)
     * Core Alignment: Priority-1 Internship Hierarchy match per Career-Ops guidelines.

3. TARGET TECHNICAL COMPETENCIES REQUIRED BY VEDYA LABS:
   - Edge AI & Quantization: Model pruning, quantization (INT8/FP16), TensorRT, ONNX Runtime, TFLite Micro, CMSIS-NN.
   - On-Device AI Architectures: Local LLM deployment (quantized Gemma 2B), multilingual edge OCR (PaddleOCR).
   - Computer Vision on Edge: YOLOv8 real-time object detection, OpenCV low-light enhancement, physical actuator mapping.
   - Acoustic Neural Processing: On-device wake-word detection, DMA-driven I2S audio sampling (ESP-SR).
   - Embedded Firmware & OS: Modern C++, Embedded C, Python, FreeRTOS task scheduling, UART/SPI/I2C communication.
   - Hardware Bench Validation: Raspberry Pi 5, ESP32S3, STM32 ARM Cortex-M4, Digital Storage Oscilloscopes (DSOs), Logic Analyzers.

4. RESUME ALIGNMENT & HIGHLIGHTS:
   - Project 1: Edge AI based Dynamic Braille System (Quantized Gemma 2B, PaddleOCR, RPi 5, UART, Multiprocessing).
   - Project 2: Adaptive LED Shadow Mapping (Physical AI, YOLOv8, PyTorch, OpenCV, real-time LED matrix actuation).
   - Project 3: Automated Lecture Tracing System (ESP-SR on-device acoustic neural engine, ESP32S3, FreeRTOS).
   - Project 4: Custom STM32 Development Board & Bench Bring-Up (2-layer KiCad, crystal clock routing, DSO bring-up).
   - Proven Edge & Robotics Leadership: Technical Lead of OMEGA Robotics Competition, Chairperson IEEE CAS MEC SB, Minor in ML.

5. SELECTION PROCESS & TIMELINE:
   - Step 1: Campus Placement Google Form registration before 8th October, 2026 (10:00 PM).
   - Step 2: Technical Assessment / Coding & Edge ML screening test.
   - Step 3: Technical Round (Deep-dive on quantization, inference latency, FreeRTOS pipelines, and Edge CV models).
   - Step 4: Final HR & Offer roll-out.
================================================================================
"""
with open(brief_path, "w", encoding="utf-8") as f:
    f.write(brief_text)
print(f"Created Job Brief: {brief_path}")

# Create Other Openings File
other_path = os.path.join(VEDYA_DIR, "Vedya Labs - Other Openings.txt")
other_text = """================================================================================
VEDYA LABS — ACTIVE OPENINGS & TRACKED ROLES (2026-2027)
================================================================================

1. ON-CAMPUS CURRENT DRIVE (MEC BATCH OF 2027):
   - Role 1: Edge AI Engineer
     Compensation: 6-Month Internship @ INR 25,000/mo + INR 5.0 - 7.0 LPA Full-Time
     Location: Hyderabad
   - Role 2: Robotics Engineer
     Compensation: 6-Month Internship @ INR 25,000/mo + INR 5.0 - 7.0 LPA Full-Time
     Location: Hyderabad

2. LATERAL & EARLY CAREER TRACKS (PORTAL TRACKING):
   - Role: Embedded AI Engineer / TinyML Specialist
     URL: https://vedya.ai
     Locations: Hyderabad / Bengaluru
   - Role: Physical AI & Computer Vision Engineer
     Locations: Hyderabad / Bengaluru

================================================================================
"""
with open(other_path, "w", encoding="utf-8") as f:
    f.write(other_text)
print(f"Created Other Openings: {other_path}")
