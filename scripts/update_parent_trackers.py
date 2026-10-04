import os

PARENT_DIR = r"c:\JOB SEARCH"
STATUS_FILE = os.path.join(PARENT_DIR, "JOB STATUS.md")
LINKS_FILE = os.path.join(PARENT_DIR, "JOB LINKS.txt")

# 1. Update JOB STATUS.md
with open(STATUS_FILE, "r", encoding="utf-8") as f:
    status_content = f.read()

new_row = "| 26 | Qualcomm | Interim Engineering Intern_2027_SW | ✅ Qualcomm/ | ✅ Done | ✅ Done | ✅ Done | ⏳ TO APPLY | 2026-10-04 |\n"
target_needle = "| 25 | Ather Energy |"

if "| 26 | Qualcomm |" not in status_content:
    lines = status_content.splitlines(keepends=True)
    new_lines = []
    inserted = False
    for line in lines:
        new_lines.append(line)
        if target_needle in line and not inserted:
            new_lines.append(new_row)
            inserted = True
    if not inserted:
        new_lines.append(new_row)
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        f.writelines(new_lines)
    print("Updated JOB STATUS.md with Qualcomm entry.")
else:
    print("Qualcomm already in JOB STATUS.md")

# 2. Update JOB LINKS.txt
with open(LINKS_FILE, "r", encoding="utf-8") as f:
    links_content = f.read()

qualcomm_entry = """
QUALCOMM INDIA: [TO APPLY 2026-10-04]
Role: Interim Engineering Intern_2027_SW (Job ID: 446719785836)
Portal / Application URL: https://careers.qualcomm.com/careers/job/446719785836?domain=qualcomm.com&hl=en
Locations: Hyderabad, Telangana, India / Bangalore, Karnataka, India / Chennai, Tamil Nadu, India / Noida, Uttar Pradesh, India
Target Batch: 2027 Batch (B.Tech in Electronics & Communication Engineering / Computer Science)
Notes: Full application package complete. Tailored 2-page resume highlighting PazhamOS (x86_64 bare-metal kernel, paging, bootloader), Automated Lecture Tracing System (FreeRTOS, DMA-driven I2S audio, acoustic neural wake-word), Automotive BLE-to-CAN Hardware Gateway (4-layer PCB, SPI drivers), and Landslide LoRa mesh. Prepared comprehensive job brief with technical breakdown across Embedded/BSP, Audio/Multimedia, and Wireless tracks. Folder: Qualcomm/
"""

if "QUALCOMM INDIA" not in links_content:
    with open(LINKS_FILE, "a", encoding="utf-8") as f:
        f.write(qualcomm_entry)
    print("Updated JOB LINKS.txt with Qualcomm entry.")
else:
    print("Qualcomm already in JOB LINKS.txt")
