import sys
import os
import subprocess
import shutil

sys.stdout.reconfigure(encoding="utf-8")

PARENT_DIR = r"c:\JOB SEARCH"
BASE_DOCX = os.path.join(PARENT_DIR, "AJITH SHAJAN - Resume EDit.docx")
MISTRAL_DIR = os.path.join(PARENT_DIR, "MistralSolutions")
OUTPUT_DOCX = os.path.join(MISTRAL_DIR, "AJITH SHAJAN - Resume (Mistral HW Digital).docx")
OUTPUT_PDF  = os.path.join(MISTRAL_DIR, "AJITH SHAJAN - Resume (Mistral HW Digital).pdf")
OUTPUT_PDF_STD = os.path.join(MISTRAL_DIR, "AJITH SHAJAN - Resume - HW Digital.pdf")
SCRIPT_IN_PARENT = os.path.join(PARENT_DIR, "format_mistral_hw_digital_resume.py")

script_content = r'''import sys, os
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document
from docx.shared import Pt

doc = Document("AJITH SHAJAN - Resume EDit.docx")
paras = doc.paragraphs

def set_run(run, text): run.text = text
def clear_runs_from(para, start_idx):
    for r in para.runs[start_idx:]: r.text = ""

# PARA 1 - Technical Skills line 1 (Hardware-biased, 93 chars)
p1 = paras[1]
set_run(p1.runs[4], "KiCad (4-Layer PCB, HDI), STM32, ESP32, ARM Cortex-M4, Board Bring-Up, DSO, Logic Analyser")
clear_runs_from(p1, 5)

# PARA 2 - Technical Skills line 2 (Interfaces & Power, 97 chars)
p2 = paras[2]
set_run(p2.runs[0], "CAN FD, SPI, I2C, UART, LoRa (RF), DC/DC Power Design (Buck/LDO/TVS), SI/PDN/EMC, DFM, DFMEA")
clear_runs_from(p2, 1)

# PARA 5 - Interests (98 chars)
p5 = paras[5]
set_run(p5.runs[4], "Embedded Hardware Design, Multi-Layer PCB Layout, Mixed-Signal Systems, SI/EMC, Board Bring-Up")
clear_runs_from(p5, 5)

# PARA 16 - ICFOSS Tech Used (102 chars)
p16 = paras[16]
set_run(p16.runs[1], "Embedded C, ESP32S3, PCB Design, Capacitive Touch Sensing, Hardware Abstraction, Assistive Hardware")
clear_runs_from(p16, 2)

# PARA 17 - ICFOSS Description (3 lines calibrated)
p17 = paras[17]
set_run(p17.runs[0],
    "Engineered real-time embedded firmware and assistive hardware on ESP32S3 microcontrollers under FreeRTOS. "
    "Designed capacitive touch sensing signal conditioning circuits and deterministic state-machine logic. "
    "Debugged hardware-software interactions and signal integrity using DSOs and logic analyzers.\t"
)
clear_runs_from(p17, 1)

# ------ PROJECT SLOT 1: Automotive BLE-to-CAN Gateway ------
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
    "Conducted AIAG/VDA DFMEA and released Gerbers + BOM (73 items, 109 components) with 0 DRC/ERC violations."
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
    "Designed a 2-layer KiCad development board with STM32F401RE, LDO regulation, and crystal routing (0 DRC/ERC). "
    "Brought up bare-metal with FreeRTOS; validated SPI, I2C, UART using DSO and logic analyzer."
)
clear_runs_from(p29, 1)

# ------ PROJECT SLOT 3: Edge AI Dynamic Braille ------
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

os.makedirs("MistralSolutions", exist_ok=True)
doc.save("MistralSolutions/AJITH SHAJAN - Resume (Mistral HW Digital).docx")
print("Saved cleanly: MistralSolutions/AJITH SHAJAN - Resume (Mistral HW Digital).docx")
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

    # Remove stale page 3 png if exists
    p3_path = os.path.join(MISTRAL_DIR, "hw_digital_page_3.png")
    if os.path.exists(p3_path):
        os.remove(p3_path)

    import fitz
    doc_pdf = fitz.open(OUTPUT_PDF)
    print(f"Verified PDF page count: {len(doc_pdf)}")
    for i, page in enumerate(doc_pdf):
        pix = page.get_pixmap(dpi=150)
        png_path = os.path.join(MISTRAL_DIR, f"hw_digital_page_{i+1}.png")
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

# Job Brief creation
brief_path = os.path.join(MISTRAL_DIR, "JOB BRIEF - Mistral Solutions Embedded Hardware Digital.txt")
brief_content = """================================================================================
JOB BRIEF: MISTRAL SOLUTIONS — EMBEDDED HARDWARE (DIGITAL) — SENIOR ENGINEER / MODULE LEAD
================================================================================

1. TARGET POSITION DETAILS
--------------------------------------------------------------------------------
- Role: Embedded Hardware (Digital) — Senior Engineer / Module Lead
- Target Division: Embedded Product Engineering (Defense, Aerospace, Industrial, Automotive)
- Experience Level Required: 3 to 8 years (Note: fresh graduate applying to talent pool / entry consideration)
- Location: Bangalore, Karnataka, India
- General Talent Pool Form:
  https://docs.google.com/forms/d/e/1FAIpQLSfKkKzDSuLcnatjtZSJaLxdxXI7afKCBhgL7V5XLu2gbJSyFw/viewform
- Role-Specific Form:
  https://docs.google.com/forms/d/e/1FAIpQLSdXCkBGxh_l8CuzwNaudaOr3Xg_wQATHZsCQGRGb8TdVi5OGA/viewform?usp=pp_url&entry.1656541746=Embedded%20Hardware%20(Digital)
- Official JD Docx:
  https://mistralsolutions.com/wp-content/uploads/2026/07/5.Embedded-Hardware-Digital-Senior-Engineer-_Module-Lead-1.docx

2. OFFICIAL JD SUMMARY & REQUIREMENTS
--------------------------------------------------------------------------------
Mistral Solutions is looking for a Lead Hardware Engineer with strong expertise in
digital and mixed-signal hardware design for Consumer, Industrial, and Automotive markets:
- End-to-end hardware development: architecture, schematic capture, multi-layer layout,
  board bring-up, validation, debugging, and customer interactions.
- Microprocessors/FPGAs: Qualcomm, TI, NXP, NVIDIA, Xilinx, Altera, Lattice, STM32.
- High-Speed Interfaces: DDR3/4/5, LPDDR4/5, eMMC, UFS, CSI-2, PCIe, HDMI, DisplayPort, DSI, CSI.
- Wired Connectivity: Ethernet, USB, McASP.
- Power Supply: DC/DC Power supply design, LDOs, PoE-PD Design, Battery-operated product design.
- Test & Debug: Board Bring-Up, Interface testing, Level zero coding in C/Assembly for HW validation.
- Instruments: CROs / DSOs, Multimeters, Logic Analyzers, Spectrum Analyzers.
- PCB Design: Multi-layer PCB Layout, HDI Designs, DFM & DFA knowledge, interaction with PCB fab house.
- Analysis: Signal Integrity (SI), Power Distribution Network (PDN), Thermal, EMI/EMC compliance.

3. TAILORED RESUME HIGHLIGHTS (AJITH SHAJAN)
--------------------------------------------------------------------------------
- Project 1: Automotive 4-Layer BLE 5.3 to CAN FD Gateway
  * 4-layer stackup (Top Sig / GND / PWR / Bot Sig), ENIG finish, 50Ω CPWG, 90Ω diff pairs.
  * Wide-input 9V–36V power conditioning, ISO 7637-2 load dump protection, P-FET clamp.
  * TI LM5164 buck converter, TPS2116 priority power mux (<2µs battery failover).
  * Formal AIAG/VDA DFMEA; released production Gerbers + BOM (73 items, 109 components, 0 DRC/ERC).
- Project 2: Custom STM32F401RE Development Board & Bring-Up
  * 2-layer KiCad PCB design with CAN transceiver, crystal routing, power decoupling, SWD/JTAG debug.
  * Complete bench bring-up and validation of SPI, I2C, UART peripherals using DSO and logic analyzer.
- Project 3: Edge AI Dynamic Braille System
  * Hardware actuation pipeline interfacing Raspberry Pi 5 via UART to ESP32 controlling servo arrays.
- Project 4: Landslide Tracker and Response System
  * Low-power IoT sensing node with multi-sensor fusion, anomaly detection in C, 433 MHz LoRa mesh.
- Work Experience (ICFOSS):
  * Assistive hardware signal conditioning and capacitive touch hardware abstraction on ESP32S3 nodes.
- Technical Skills Matrix:
  * KiCad (4-Layer PCB, HDI), STM32, ESP32, ARM Cortex-M4, Board Bring-Up, DSO, Logic Analyser.
  * CAN FD, SPI, I2C, UART, LoRa (RF), DC/DC Power Design (Buck/LDO/TVS), SI/PDN/EMC, DFM, DFMEA.
  * FPGA-based DSD using Verilog certification (C2S & MEC).

4. HARDWARE TECHNICAL TALKING POINTS FOR MISTRAL INTERVIEWS
--------------------------------------------------------------------------------
- 4-Layer PCB Stackup: Top Signal (RF/high-speed) / L2 Continuous GND Plane / L3 Power Plane / Bottom Signal.
  Why? Continuous reference plane directly beneath Top layer minimizes loop inductance, prevents EMI,
  and enables tightly controlled 50Ω single-ended / 90Ω differential impedance.
- Power Supply Design:
  * LM5164 wide-VIN buck regulator: 9–36V automotive input stepped down to 3.3V master rail.
  * TPS2116 priority power multiplexer: Seamless switchover from vehicle battery to Li-Ion backup in <2µs.
  * Transient Protection: Bidirectional TVS diode + P-channel MOSFET reverse polarity clamp for ISO 7637-2 pulse compliance.
- Signal Integrity & High-Speed Routing:
  * 90Ω differential routing for CAN FD (5 Mbps) with 120Ω split termination and common-mode choke.
  * 50Ω coplanar waveguide with ground (CPWG) for 2.4 GHz BLE RF output with via fencing.
- Board Bring-Up Methodology:
  * Phase 1 (Cold Check): Impedance check across 3.3V, 5V, VBAT power rails to ground (check for shorts).
  * Phase 2 (Power-Up): Controlled current-limited bench supply, measure ripple and rail sequencing on DSO.
  * Phase 3 (Clock & Reset): Probe crystal oscillator oscillation (HSE/LSE) and NRST de-assertion.
  * Phase 4 (SWD/JTAG & Level-0 Code): Connect debug probe, flash bare-metal GPIO toggle / UART heartbeat, verify peripheral registers.
================================================================================
"""

with open(brief_path, "w", encoding="utf-8") as f:
    f.write(brief_content)
print(f"Generated Job Brief: {brief_path}")
