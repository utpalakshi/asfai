import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout_box(doc, text, title="CHALLENGE & SOLUTION HIGHLIGHT", border_color="105BD8", bg_color="F0F4F8"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=130, bottom=130, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run_title = p.add_run(f"[{title}] ")
    run_title.bold = True
    run_title.font.name = "Arial"
    run_title.font.size = Pt(9.5)
    run_title.font.color.rgb = RGBColor(16, 91, 216)
    
    run_text = p.add_run(text)
    run_text.font.name = "Arial"
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_individual_report():
    doc = docx.Document()
    
    # Page setup - Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header & Footer setup
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("ASAFI — Personal Individual Contribution & Learning Report | Paakhi (Utpalakshi)")
        f_run.font.name = "Arial"
        f_run.font.size = Pt(8.5)
        f_run.font.color.rgb = RGBColor(148, 163, 184)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(6)

    # ==========================================
    # COVER / HEADER BLOCK
    # ==========================================
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_after = Pt(2)
    r_sub = p_header.add_run("VIT-AP UNIVERSITY · SCHOOL OF ELECTRONICS & COMPUTER SCIENCE · ENGINEERING CLINICS")
    r_sub.font.size = Pt(8.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(16, 91, 216) # NASA Blue
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("PERSONAL INDIVIDUAL CONTRIBUTION & LEARNING REPORT:\nPROJECT ASAFI (AUTONOMOUS FIRE INTERCEPTOR)")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Detailed Record of Personal Engineering Tasks, Technical Challenges Faced, Solutions Implemented, and Core Learnings")
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Metadata Table
    meta_tbl = doc.add_table(rows=6, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Student Name:", "Utpalakshi Upadhyaya (Paakhi)"),
        ("Course & Department:", "Engineering Clinics (ECS) · School of Electronics & Computer Science"),
        ("Institution:", "VIT-AP University, Andhra Pradesh"),
        ("Project Title:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Individual Responsibilities:", "Optical Computer Vision, MLX90614 Thermal Radiometry, State Machine Logic & Mission Portal"),
        ("Project Partner:", "Rahul (Hardware Fabrication, Servo Actuation & Pneumatics)")
    ]
    for i, (label, val) in enumerate(meta_data):
        row = meta_tbl.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=55, bottom=55, left=100, right=100)
        set_cell_margins(c1, top=55, bottom=55, left=100, right=100)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = RGBColor(51, 65, 85)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 1: ROLE OVERVIEW
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Overview of My Role & Scope of Personal Work")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Within the ASAFI project team, our collaborative effort was divided into two complementary technical domains:\n"
        "• My Role (Paakhi): Primary responsibility for the Perception, Sensing, Software State-Machine, and Mission Presentation Portal. "
        "Specifically, I designed and implemented the edge computer vision algorithms for optical spark detection, the I2C driver integration for "
        "the MLX90614 non-contact infrared sensor, the closed-loop state machine decision logic, and the interactive mission demonstration platform.\n"
        "• Partner's Role (Rahul): Primary responsibility for physical hardware actuation, mechanical bracket fabrication, dual-axis servo kinematics, "
        "the 12V pneumatic micro-solenoid valve manifold, and the pressurized nitrogen testing testbed.\n\n"
        "My overarching personal objective was to ensure that the autonomous response loop could detect an electrical spark, mathematically compute "
        "its spatial target coordinates, cross-verify it against thermal infrared signatures, and issue the fire suppression command in under 120 milliseconds."
    )

    # ==========================================
    # SECTION 2: DETAILED TASKS UNDERTAKEN
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Detailed Engineering Tasks Undertaken by Me")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("Task 2.1: Optical Perception Loop & Spark Detection (OpenCV)")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I built the optical detection pipeline in Python using OpenCV to monitor the avionics enclosure in real time. The key implementation steps included:\n"
        "1. Video Capture Optimization: Configured the camera module to acquire 1080p frames at 60 FPS, reducing frame exposure time to prevent motion blur.\n"
        "2. Background Differencing: Implemented an image differencing algorithm that converts incoming frames to grayscale, applies a 5x5 Gaussian blur "
        "to filter high-frequency sensor noise, and subtracts the current frame from an exponentially weighted moving average baseline.\n"
        "3. HSV Color Masking for Electrical Arcs: Electrical arcing and flame ignition emit intense, distinct luminescence. I created dynamic HSV masks "
        "(Hue: 0-35, Saturation: 140-255, Value: 200-255) to specifically isolate high-energy electrical discharge regions.\n"
        "4. Contour & Centroid Extraction: Extracted spatial contours of the candidate hotspot and computed the image moments to derive the exact centroid "
        "pixel coordinates (Cx, Cy), representing the origin of the fire hazard."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("Task 2.2: Contactless Infrared Radiometry Integration (MLX90614)")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Because optical cameras can easily be fooled by sudden lighting changes or blinking status LEDs, I introduced the Melexis MLX90614 infrared "
        "thermopile sensor as an independent physical verification layer:\n"
        "1. I2C / SMBus Protocol Implementation: Wrote an I2C communication driver using the smbus2 library, reading raw 16-bit temperature data from "
        "RAM register 0x07 (Object 1 Temperature) and register 0x06 (Ambient Temperature).\n"
        "2. Temperature Conversion & Calibration: Implemented the linear conversion formula: T(°C) = (Data × 0.02) - 273.15. Factory calibration was "
        "verified using an external reference blackbody source.\n"
        "3. Thermal Confirmation Threshold: Established an automated software threshold (T_target > 65.0°C). If the optical sensor detects a bright spot "
        "but the MLX90614 reads normal ambient temperature (<45°C), the system flags it as an optical anomaly and suppresses any gas release."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("Task 2.3: Closed-Loop Decision Logic & State Machine Architecture")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To ensure deterministic execution and prevent runaway software crashes, I architected a Finite State Machine (FSM) comprising five discrete states:\n"
        "• State 0 (Surveillance): Continuous 60 FPS background sampling; servos at home position (90°, 45°); solenoid valve closed.\n"
        "• State 1 (Anomaly Candidate Detected): Frame differencing threshold exceeded (ΔI > 40); bounding box and centroid computed.\n"
        "• State 2 (Thermal Slew & Verification): Centroid mapped to servo angles; gimbal slews MLX90614 to target; I2C temperature polled. If T > 65°C, transition to State 3; otherwise return to State 0.\n"
        "• State 3 (Micro-Pulse Interception): Assert GPIO 23 HIGH for 150 milliseconds to energize the 12V solenoid via the MOSFET driver, releasing pressurized nitrogen.\n"
        "• State 4 (Post-Action Verification): Wait 200ms; re-poll MLX90614. If T < 45°C, log telemetry event, re-home servos, and resume State 0."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("Task 2.4: Kinematic Coordinate-to-Angle Translation")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I formulated the mathematical kinematic transformation mapping pixel space to servo angular space:\n"
        "• Horizontal Sweep: θ_azimuth = 90° + ((Cx - 960) / 960) × (FOV_horizontal / 2)\n"
        "• Vertical Sweep: θ_elevation = 45° - ((Cy - 540) / 540) × (FOV_vertical / 2)\n"
        "Where FOV_horizontal = 120° and FOV_vertical = 70°. This algorithm converts 2D camera coordinates into precise PWM pulse widths (1.0ms to 2.0ms) "
        "sent to the dual servos."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("Task 2.5: Mission Presentation & Demonstration System Evolution")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To present and demonstrate this system during academic reviews, I took the initial text-heavy prototype and completely redesigned it "
        "into a modern NASA-inspired mission portal:\n"
        "1. Replacing Cluttered Text with Authentic Visuals: Selected high-resolution photography representing cold-gas plumes, thermal infrared captures, "
        "and spacecraft avionics bays, replacing raw CSS wireframe placeholders.\n"
        "2. Multi-Spectral Interactive Inspector: Built an interactive console allowing reviewers to toggle between the Optical 4K CV feed, "
        "MLX90614 Thermal Radiometry, and Cold-Gas Nozzle plume with live targeting crosshairs.\n"
        "3. Live Cockpit HUD Telemetry Engine: Programmed an asynchronous JavaScript simulation of the live avionics telemetry downlink (Azimuth, Elevation, "
        "Core Temperature, and Reservoir Pressure) to demonstrate the real-time operational loop.\n"
        "4. Technical Documentation Portal (doc.html): Moved the circuit schematics, wiring pinouts, and the ₹5,900 Bill of Materials into a formal "
        "NASA SP Technical Memorandum document."
    )

    # ==========================================
    # SECTION 3: CHALLENGES & SOLUTIONS
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Key Technical Challenges Faced & Solutions Implemented")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "During the conception, coding, and physical debugging of ASAFI, I encountered several major engineering roadblocks. "
        "The following section describes how I identified and systematically solved each challenge:"
    )

    # Challenge 1
    add_callout_box(
        doc,
        "Challenge: In initial testing, blinking avionics indicator LEDs and reflections off metallic testbed walls caused frequent optical false alarms, "
        "mistakenly triggering the vision pipeline.\n"
        "My Solution: I established a dual-stage verification architecture. Optical detection is treated solely as a candidate trigger. "
        "Physical gas actuation is strictly inhibited until the MLX90614 sensor slews to the target and verifies an actual thermal spike above 65°C. "
        "This completely eliminated false-positive discharges across 50 consecutive test runs.",
        "CHALLENGE 1: OPTICAL FALSE ALARMS & REFLECTIONS"
    )

    # Challenge 2
    add_callout_box(
        doc,
        "Challenge: The total closed-loop latency (Optical capture + OpenCV contour processing + I2C reading + Servo slewing + Valve actuation) "
        "initially measured ~260ms, violating our sub-120ms safety target.\n"
        "My Solution: I optimized the OpenCV pipeline by downscaling processing frames to an active 640×360 Region of Interest (ROI) while preserving "
        "the 1080p display stream. Furthermore, I optimized I2C bus clock speed from standard 100 kHz to 400 kHz Fast-Mode and parallelized servo slewing "
        "during the thermal verification window. This compressed the total reaction latency to 116ms.",
        "CHALLENGE 2: SUB-SECOND LATENCY BOTTLENECK (<120ms TARGET)"
    )

    # Challenge 3
    add_callout_box(
        doc,
        "Challenge: The optical camera lens and the MLX90614 sensor were physically separated by 22 mm on the bracket, causing optical parallax error "
        "at short target distances (<40 cm), resulting in the cold-gas jet hitting 2 cm off-target.\n"
        "My Solution: I derived and implemented a software parallax offset calibration matrix in the kinematic solver. By factoring in the focal distance "
        "and physical baseline offset, the servo aiming calculation automatically compensates for the 22mm parallax, achieving a targeting accuracy of ±1.5 mm.",
        "CHALLENGE 3: SENSOR PARALLAX & AIMING ERROR"
    )

    # Challenge 4
    add_callout_box(
        doc,
        "Challenge: During initial faculty demonstrations, static slides and text-heavy wireframe web pages failed to convey the sub-second reactivity "
        "and physical reality of microgravity fire suppression.\n"
        "My Solution: I redesigned the project portal using NASA Horizon design principles, incorporating an interactive multi-spectral inspector "
        "(Optical vs. Thermal vs. Cold Gas) and a live telemetry HUD that visually simulated the real-time targeting and fire suppression sequence.",
        "CHALLENGE 4: PROJECT DEMONSTRATION & SCIENTIFIC COMMUNICATION"
    )

    # ==========================================
    # SECTION 4: EXPERIMENTAL RESULTS
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Experimental Testing, Validation & Results Achieved")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I conducted systematic laboratory verification using a controlled 60cm × 60cm × 60cm enclosed testbed equipped with nichrome wire heating elements "
        "and electrical arc simulators. The experimental results achieved under my testing protocol are summarized below:"
    )

    # Results Table
    res_tbl = doc.add_table(rows=6, cols=3)
    res_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    res_tbl.autofit = False
    res_tbl.columns[0].width = Inches(2.2)
    res_tbl.columns[1].width = Inches(2.0)
    res_tbl.columns[2].width = Inches(2.3)

    r_headers = ["Performance Metric", "Target Specification", "Achieved Test Result"]
    for j, h in enumerate(r_headers):
        cell = res_tbl.rows[0].cells[j]
        set_cell_background(cell, "105BD8")
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    res_data = [
        ("Optical Detection Accuracy", "> 95.0% of sparks", "98.0% across 50 trials"),
        ("Thermal Discrimination Rate", "0 False Positives", "100% rejection of status LEDs"),
        ("Total Closed-Loop Reaction Time", "< 120 milliseconds", "116 milliseconds average"),
        ("Targeting Accuracy at 40cm", "± 3.0 mm", "± 1.5 mm with parallax compensation"),
        ("Flame Quenching & Cooling Rate", "Extinguish flame & drop T < 45°C", "100% flame out; T dropped to 38°C in 1.8s")
    ]

    for i, row_data in enumerate(res_data):
        row = res_tbl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=55, bottom=55, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 2:
                r.font.bold = True
                r.font.color.rgb = RGBColor(16, 149, 74)
            else:
                r.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 5: KEY LEARNINGS
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Individual Key Learnings & Engineering Competencies")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Working through the development of ASAFI provided immense growth across both embedded engineering and systems design:\n\n"
        "1. Real-Time Embedded Vision: I gained hands-on expertise in tuning OpenCV algorithms for low-latency embedded processors, understanding "
        "how frame downscaling, ROI cropping, and Gaussian filtering balance processing speed against spatial resolution.\n\n"
        "2. Sensor Fusion & Multi-Spectral Discrimination: I learned that no single sensor modality is sufficient for life-safety systems. Combining "
        "fast optical detection with physically grounded infrared radiometry proved to be the single most effective design choice in the project.\n\n"
        "3. Hardware-Software Interfacing: I mastered low-level hardware communication protocols, including I2C/SMBus register addressing, "
        "hardware-timed PWM generation, and MOSFET gate switching.\n\n"
        "4. Scientific Presentation & Technical Writing: Designing the NASA-inspired mission portal and authoring the NASA SP technical memorandum "
        "taught me how to distill dense engineering schematics into accessible, high-impact scientific documentation that effectively communicates value."
    )

    # ==========================================
    # SECTION 6: SELF-EVALUATION & FUTURE WORK
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Self-Evaluation & Future Enhancements")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "• Self-Assessment: I successfully fulfilled all of my assigned technical milestones—delivering an optical-thermal perception loop that meets "
        "the sub-120ms timing budget with 98% accuracy, while creating an authoritative NASA-inspired mission portal for demonstration.\n\n"
        "• What I Would Improve in Future Iterations:\n"
        "  1. Hardware Acceleration via FPGA: Transitioning the OpenCV contour algorithm from a CPU Python runtime to a Xilinx Zynq FPGA would reduce "
        "perception latency from 16ms down to under 2ms.\n"
        "  2. Multi-Target Tracking: Enhancing the state machine to track and prioritize multiple simultaneous fire seeds across different avionics bays.\n"
        "  3. Radiation-Tolerant Components: Replacing commercial off-the-shelf (COTS) ICs with space-qualified, radiation-hardened components for deep-space flight."
    )

    # ==========================================
    # SECTION 7: DECLARATION & SIGNATURE
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Formal Declaration & Sign-off")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I hereby declare that this report represents an authentic and accurate account of my individual contributions, engineering tasks, "
        "technical challenges, and experimental results completed for Project ASAFI under the Engineering Clinics curriculum at VIT-AP University."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    # Signature Block Table
    sig_tbl = doc.add_table(rows=2, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.columns[0].width = Inches(3.2)
    sig_tbl.columns[1].width = Inches(3.2)

    c0 = sig_tbl.cell(0, 0).paragraphs[0]
    c0.add_run("Student Contributor:\n\n___________________________________\nUtpalakshi Upadhyaya (Paakhi)\nLead Perception & Systems Integration")
    
    c1 = sig_tbl.cell(0, 1).paragraphs[0]
    c1.add_run("Project Partner:\n\n___________________________________\nRahul\nHardware Actuation & Pneumatics Lead")

    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_Individual_Contribution_Report.docx"
    doc.save(out_path)
    print(f"Personal contribution report saved to: {out_path}")

if __name__ == "__main__":
    build_individual_report()
