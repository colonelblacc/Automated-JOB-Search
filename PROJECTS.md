# Ajith Shajan — Master Projects Portfolio

A comprehensive, categorized catalog of engineering projects developed by **Ajith Shajan** (B.Tech in Electronics and Computer Engineering, Govt. Model Engineering College, Kochi).

---

## Quick Navigation
1. [Embedded Hardware & PCB Design](#1-embedded-hardware--pcb-design)
2. [Embedded Firmware, RTOS & Telemetry](#2-embedded-firmware-rtos--telemetry)
3. [Edge AI, TinyML & Computer Vision](#3-edge-ai-tinyml--computer-vision)
4. [Systems Software & Operating Systems](#4-systems-software--operating-systems)
5. [Assistive & Biomedical Technology](#5-assistive--biomedical-technology)
6. [Agentic AI & Enterprise Automation](#6-agentic-ai--enterprise-automation)
7. [Hackathon Awards & Competitions](#7-hackathon-awards--competitions)
8. [Role-to-Project Mapping Matrix](#8-role-to-project-mapping-matrix)

---

## 1. Embedded Hardware & PCB Design

### Project 1.1: STM32F401RE Custom Microcontroller PCB Development Board
* **Role:** Hardware Lead | **Team Size:** 1 | **Timeline:** 2 Weeks
* **Primary Tech Stack:** KiCad (Schematic + 2-Layer PCB Layout), STM32F401RETx (ARM Cortex-M4 @ 84 MHz), USB-C, 3.3V LDO, Crystal Oscillator, SWD Debug, SPI, I2C, UART, GPIO Breakout, Gerber / BOM
* **Overview:**
  Designed and fabricated a custom 2-layer microcontroller development board from scratch in KiCad, tailored for embedded peripherals evaluation, hardware validation, and RTOS firmware bring-up.
* **Key Engineering Highlights:**
  - Designed clean power distribution network (PDN) with 5V USB-C to 3.3V low-dropout regulation (LDO) and high-frequency decoupling capacitors placed adjacent to MCU VDD pins.
  - Routed high-speed crystal oscillator differential traces with ground isolation to guarantee clock stability at 84 MHz.
  - Implemented 4-pin SWD (Serial Wire Debug) interface for JTAG/ST-Link flashing and live runtime debugging.
  - Broken out all major MCU peripheral buses (SPI, I2C, UART, timers, GPIOs) to header pins for bench validation.
  - Achieved **0 DRC (Design Rule Check) and 0 ERC (Electrical Rule Check) violations** prior to fabrication export.
  - Brought up the physical board on bench; verified power rails and bus signal integrity using Digital Storage Oscilloscopes (DSOs).

---

### Project 1.2: Automotive BLE-to-CAN Hardware Gateway
* **Document Number:** `DOC-HW-BLE-CAN-001` (Revision 1.0 — Production Baseline)
* **Role:** Hardware Design Engineer | **Team Size:** 1 | **Timeline:** 2 Weeks
* **Primary Tech Stack:** KiCad 10.0 (Schematic & 4-Layer PCB Layout), TI CC2340R40/R52 (Arm Cortex-M0+ BLE 5.3 SoC), TI TCAN4550-Q1 (CAN FD SBC), TI LM5164, TI LMR36506, TI BQ24040, TI TPS2116, TI TPS63900, ISO 7637-2, CISPR 25, DFMEA, Gerbers
* **Overview:**
  Architected an automotive-grade telematics gateway bridging Bluetooth Low Energy (BLE 5.3) with high-speed CAN Flexible Data-Rate (CAN FD up to 5 Mbps) for connected-vehicle fleet tracking and edge diagnostics. Built on a 4-layer FR4 stackup operating continuously from a 9V–36V vehicle DC battery bus with integrated Li-Ion battery backup failover (<2µs), discrete RF front-end with ceramic chip antenna, and strict automotive transient suppression.
* **Key Engineering Highlights:**
  - **4-Layer High-Density Stackup:** Routed on a 4-layer stackup (JLC04161H-7628, 1.6mm, 1 oz outer / 0.5 oz inner Cu, ENIG) with a continuous Layer-2 ground plane and Layer-3 power distribution polygons, routing 50Ω CPWG RF traces and 90Ω differential pairs for CAN FD and USB.
  - **Automotive Radio & Controller:** Integrated TI CC2340R52 (Arm Cortex-M0+, 48 MHz) BLE 5.3 wireless SoC and TI TCAN4550-Q1 (AEC-Q100 Grade 1) CAN FD System Basis Chip over high-speed SPI with bidirectional level shifting and isolated ground pours.
  - **Vehicle Power Conditioning:** Designed robust 9V–36V vehicle power conditioning protected against reverse polarity via AEC-Q101 P-FET (SQ2361AEES, -60V Vds), 36V 400W TVS clamping (SMAJ36CA), and a 2A fast-acting fuse complying with ISO 7637-2 (Pulses 1, 2a, 3a, 3b, 5b) and ISO 16750-2.
  - **Subsystem Power Rail Isolation:** Implemented 3 isolated buck stages: TI LM5164 (100V, 1A COT) for 3.3V logic, and dual TI LMR36506 (65V, 0.6A) converters for 5V battery charging and clean ~6.0V CAN VSUP to isolate inductive switching noise from digital logic.
  - **Uninterrupted Battery Failover:** Engineered battery management featuring TI BQ24040 Li-Ion charger (NTC thermistor thermal foldback), TI TPS63900 buck-boost, and TI TPS2116 priority power multiplexer providing automatic <2µs hardware failover during vehicle power disconnect.
  - **RF Front-End & Bus Hardening:** Designed 50Ω CPWG RF feed with discrete Pi matching network, Johanson 2450AT18D0100E chip antenna with multi-layer keepout, Coilcraft ACT45B common-mode choke, Nexperia ESDCAN24-2BWY (±30kV) ESD clamps, and jumper-selectable 120.8Ω split termination.
  - **DFMEA & Industrial Sign-Off:** Formulated a 6-page AIAG/VDA DFMEA assessing failure modes, thermal dissipation, and stress deratings; released production Gerbers, NC drill data, and BOM (73 unique items, 109 components) with **0 ERC and 0 DRC violations**.

---

## 2. Embedded Firmware, RTOS & Telemetry

### Project 2.1: Automated Lecture Tracing System (LTS)
* **Role:** Embedded Firmware Lead | **Team Size:** 4 | **Timeline:** 1 Month
* **Primary Tech Stack:** C++, FreeRTOS, ESP32-S3 (Seeed XIAO Sense), ESP-SR (Neural Audio AI), FastAPI, WebSocket, UART, Wi-Fi, Google STT, Gemini AI
* **Overview:**
  Engineered an intelligent classroom audio acquisition and telemetry device that captures spoken lectures, performs local acoustic neural keyword detection, streams audio over WebSocket, and generates structured AI summaries.
* **Key Engineering Highlights:**
  - Implemented embedded C++ firmware running under **FreeRTOS** on the dual-core ESP32-S3 microcontroller.
  - Integrated **ESP-SR neural acoustic engine** for local, low-latency wake-word recognition without external cloud latency.
  - Utilized DMA-driven I2S audio sampling from digital MEMS microphones to eliminate CPU starvation.
  - Developed bi-directional real-time audio telemetry streaming over WebSocket with auto-reconnection and buffering.
  - Interfaced with a high-performance Python FastAPI backend orchestrating cloud speech-to-text and multimodal LLM summarization.
  - Verified bus timings, task scheduling constraints, and memory allocation using Logic Analyzers and FreeRTOS runtime stats.

---

### Project 2.2: Landslide Tracker and Emergency Early-Warning Response System
* **Role:** Embedded Systems Lead | **Team Size:** 4 | **Timeline:** 4 Days
* **Primary Tech Stack:** ESP32, LoRa (Ra-02 / SX1276 @ 433 MHz), Embedded C, Accelerometer (MPU6050), Soil Hygrometer, Sensor Fusion, Low-Power RF Mesh
* **Overview:**
  A rugged, solar-compatible disaster-detection embedded system designed for landslide-prone terrains that detects ground instability and floods, autonomously cross-verifies threats via peer-to-peer RF mesh, and broadcasts alarms.
* **Key Engineering Highlights:**
  - Implemented multi-sensor data acquisition reading 3-axis acceleration and moisture content via analog/I2C interfaces.
  - Developed threshold anomaly detection algorithms in embedded C to prevent false triggers from temporary tremors or rainfall.
  - Architected an autonomous peer-to-peer **433 MHz LoRa RF mesh** communicating across remote nodes without cellular or Wi-Fi dependency.
  - Established cross-verification consensus: nodes query adjacent sensors before escalating to high-priority disaster alarms.
  - Optimized power consumption with deep-sleep duty cycles for extended off-grid battery deployment.

---

## 3. Edge AI, TinyML & Computer Vision

### Project 3.1: Edge AI-Based Dynamic Braille System
* **Role:** Edge AI & Hardware Lead | **Team Size:** 4 | **Timeline:** 1 Week
* **Primary Tech Stack:** Python, PyTorch, Edge ML (Gemma 2B), PaddleOCR, Tesseract, Multiprocessing, Raspberry Pi 5, Arduino, UART Interface, SG90 Servos
* **Overview:**
  An autonomous, completely offline assistive learning station for visually impaired students that scans printed textbook pages, converts them into localized text, and drives dynamic mechanical Braille pins in real time.
* **Key Engineering Highlights:**
  - Built an end-to-end edge pipeline running on a Raspberry Pi 5 without internet or cloud connectivity.
  - Integrated bilingual OCR (PaddleOCR and Tesseract) supporting simultaneous English and Malayalam character recognition.
  - Deployed a quantized **Google Gemma 2B LLM** locally for contextual spelling correction, hyphen reconstruction, and summarization.
  - Engineered a multiprocessing queue in Python to decouple heavy AI inference from high-speed hardware UART serial transmission.
  - Programmed a downstream Arduino controller to translate characters into 6-dot Braille patterns driving an array of micro-servo actuators.

---

### Project 3.2: Adaptive LED Shadow Mapping System (Anti-Glare Automotive Vision)
* **Role:** Computer Vision & Edge Lead | **Team Size:** 3 | **Timeline:** 1 Week
* **Primary Tech Stack:** Python, PyTorch, YOLOv8, OpenCV, NumPy, Raspberry Pi 5, WS2812B Addressable LED Strip, Edge Deployment
* **Overview:**
  A real-time intelligent automotive headlamp system that detects oncoming and preceding vehicles at night and selectively shuts off individual LED matrix sectors to eliminate headlight glare for other drivers while illuminating the roadway.
* **Key Engineering Highlights:**
  - Implemented low-light image enhancement algorithms in OpenCV to maximize contrast in dark driving environments.
  - Deployed an optimized **YOLOv8 deep learning model** running real-time vehicle detection and bounding-box tracking.
  - Developed spatial coordinate mapping algorithms translating 2D camera pixels to discrete angular sectors of an addressable WS2812B LED strip.
  - Formatted dynamic shadow masks with zero visual flicker to actively blackout illumination around incoming vehicle windshields.

---

## 4. Systems Software & Operating Systems

### Project 4.1: PazhamOS — Bare-Metal x86_64 Operating System Kernel
* **Role:** Systems Programmer | **Team Size:** 1 | **Timeline:** Independent Project
* **Primary Tech Stack:** x86_64 Assembly (NASM), C, Custom Two-Stage Bootloader, GDT, IDT, VGA Text Mode Driver, Hardware I/O Ports, QEMU, GNU Make
* **Overview:**
  A bare-metal, 64-bit operating system kernel developed from scratch to study low-level CPU execution, memory structures, hardware registers, and hardware abstraction.
* **Key Engineering Highlights:**
  - Authored a custom Master Boot Record (MBR) bootloader transitioning CPU execution from 16-bit Real Mode to 32-bit Protected Mode and finally into 64-bit Long Mode.
  - Initialized the Global Descriptor Table (GDT), 64-bit Page Tables (PML4), and Interrupt Descriptor Table (IDT).
  - Wrote a bare-metal VGA text mode screen driver with scrolling, color formatting, and cursor control via direct memory-mapped I/O (`0xB8000`).
  - Implemented low-level port I/O routines (`inb`/`outb`) in assembly for keyboard scanning and programmable interrupt controller (PIC 8259) remapping.
  - Automated compilation, disk image generation, and QEMU virtualized execution via Makefiles.

---

## 5. Assistive & Biomedical Technology

### Project 5.1: Real-Time Assistive Embedded Hardware & Tactile Mapping (ICFOSS Internship)
* **Role:** Embedded Engineering Intern | **Organization:** International Centre for Free and Open Source Software (ICFOSS)
* **Primary Tech Stack:** Embedded C, FreeRTOS, ESP32-S3, Capacitive Touch Sensing, Hardware Abstraction Layers (HAL), Digital Storage Oscilloscopes (DSOs), Logic Analyzers
* **Overview:**
  Conducted research and firmware engineering on specialized assistive hardware platforms under the Government of Kerala's open-source hardware initiative.
* **Key Engineering Highlights:**
  - Researched and implemented capacitive touch sensor matrix algorithms on ESP32-S3 microcontrollers for interactive tactile Braille maps.
  - Developed deterministic, interrupt-driven state machines under FreeRTOS ensuring millisecond-level response times for user inputs.
  - Architected memory-cell assistive hardware concepts to support cognitively impaired individuals with structured daily memory retention routines.
  - Performed rigorous bench verification of peripheral bus signals (I2C, SPI, UART) and signal integrity using oscilloscopes and logic analyzers.

---

## 6. Agentic AI & Enterprise Automation

### Project 6.1: Meeting Intelligence Hub (Agentic AI Multi-Agent Orchestrator)
* **Role:** AI Developer | **Team Size:** 3 | **Timeline:** 2 Weeks
* **Primary Tech Stack:** Python, LangGraph, Agentic AI, FastAPI, Docker, Multi-Agent Orchestration, REST APIs
* **Overview:**
  An enterprise-grade autonomous multi-agent platform that processes unstructured meeting audio and text, performs multi-perspective analysis, extracts action items, and generates verifiable executive briefs.
* **Key Engineering Highlights:**
  - Built stateful agent graphs using **LangGraph** with supervisor-worker architectures and conditional routing.
  - Designed specialized autonomous agents: Transcription Verifier, Action Item Extractor, Technical Synthesizer, and Risk Assessor.
  - Deployed containerized microservices via Docker exposing asynchronous REST endpoints through FastAPI.

---

## 7. Hackathon Awards & Competitions

| Honor / Award | Event / Organizer | Project / Focus Area |
| :--- | :--- | :--- |
| 🥈 **2nd Runner-Up** | **HACKEFX 2.0** (National Hackathon by Electrifex) | Hardware & IoT Innovation |
| 🥉 **3rd Prize** | **Hardware Hackathon** (ECSA, CUSAT) | Embedded Hardware & Sensor Interfacing |
| 🥇 **1st Prize** | **Electrovation** (Mixed Signals MEC SB) | Hardware Prototyping & Circuits |
| 🌟 **Finalist** | **GEN AI Hackathon** (Google Developers Group, Kochi) | Generative AI & Edge LLM Systems |
| 🤖 **Technical Lead** | **OMEGA Robotics Competition** (IEEE RAS & MEC SB) | Robotics Hardware & Autonomous Arena Design |

---

## 8. Role-to-Project Mapping Matrix

When tailoring resumes or interviewing for specific technical roles, use this matrix to select the most impactful projects:

| Target Job Profile | Recommended Top 4 Projects | Key Buzzwords to Emphasize |
| :--- | :--- | :--- |
| **Embedded Firmware / RTOS Engineer** | 1. STM32 Custom PCB Board<br>2. Automated Lecture Tracing (ESP-SR)<br>3. ICFOSS Embedded Firmware<br>4. Landslide Tracker (LoRa) | FreeRTOS, Bare-Metal, C/C++, STM32 HAL, ESP32-S3, UART/SPI/I2C, DSO, Board Bring-Up |
| **Embedded Hardware / PCB Design Engineer** | 1. STM32 Custom PCB Board<br>2. Automotive CAN FD Gateway<br>3. ICFOSS Assistive Hardware<br>4. Landslide Tracker Node | KiCad 2-Layer PCB, ISO 7637-2, DFMEA, Schematics, DRC/ERC 0 errors, DSOs, LDOs |
| **Edge AI / TinyML / Computer Vision Engineer** | 1. Edge AI Dynamic Braille (Gemma 2B)<br>2. Adaptive LED Shadow Mapping (YOLOv8)<br>3. Automated Lecture Tracing (ESP-SR)<br>4. Meeting Intelligence Hub | TinyML, Gemma 2B, PyTorch, YOLOv8, OpenCV, Raspberry Pi 5, On-Device LLM |
| **Automotive Electronics / Telematics Engineer** | 1. Automotive CAN FD Gateway<br>2. STM32 Custom PCB Board<br>3. Adaptive LED Shadow Mapping<br>4. Landslide Tracker (LoRa) | CAN FD, CAN Bus, ISO 7637-2 Load Dump, TVS Protection, DFMEA, KiCad, Automotive Safety |
| **Aerospace, Defense & RF Systems Trainee** | 1. Landslide Tracker (433 MHz LoRa)<br>2. STM32 Custom PCB Board<br>3. ICFOSS Firmware & Validation<br>4. Automated Lecture Tracing | RF Telemetry, Sensor Fusion, Low-Power RF, Signal Integrity, Oscilloscope, Validation |

---
*Maintained in Career-Ops for Ajith Shajan | Last Updated: October 2026*
