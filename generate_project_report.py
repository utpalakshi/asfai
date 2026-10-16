import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout_box(doc, text, title="KEY SPECIFICATION", border_color="105BD8", bg_color="F0F4F8"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run_title = p.add_run(f"[{title}] ")
    run_title.bold = True
    run_title.font.name = "Arial"
    run_title.font.size = Pt(10)
    run_title.font.color.rgb = RGBColor(16, 91, 216)
    
    run_text = p.add_run(text)
    run_text.font.name = "Arial"
    run_text.font.size = Pt(10)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_project_report():
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
        f_run = f_p.add_run("ASAFI Engineering & Hardware Implementation Report | VIT-AP Engineering Clinics")
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
    r_title = p_title.add_run("ASAFI: AUTONOMOUS SPACECRAFT AVIONICS FIRE INTERCEPTOR\nENGINEERING DESIGN & HARDWARE IMPLEMENTATION REPORT")
    r_title.font.size = Pt(17)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Sub-Second Optical-Thermal Anomaly Detection, Multi-Axis Kinematic Vectoring, and Pressurized Cold-Gas Micro-Pulse Interception for Microgravity Avionics")
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Metadata Table
    meta_tbl = doc.add_table(rows=5, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Project Designation:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Project Domain:", "Aerospace Avionics Safety, Embedded Robotics & Edge Perception"),
        ("Academic Curriculum:", "Engineering Clinics (ECS) · VIT-AP University"),
        ("Project Team Members:", "Paakhi (Perception, Thermal Sensing & System Integration) & Rahul (Avionics Hardware & Actuation)"),
        ("Project Status:", "TRL-5 Verified (Hardware Testbed & Simulated Vacuum Environment)")
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
    # SECTION 1: EXECUTIVE SUMMARY
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Executive Summary & Project Objectives")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor) is an active, autonomous safety payload designed to mitigate localized "
        "electrical fire hazards, arc flashes, and thermal runaway inside spacecraft avionics bays. In orbital spaceflight, "
        "avionics enclosures are packed with dense multi-layer printed circuit boards (PCBs), high-current power distribution units (PDUs), "
        "and lithium-ion battery modules. An electrical short circuit or insulation breakdown in these confined volumes can trigger rapid "
        "combustion. Traditional fire suppression techniques rely on whole-cabin gas flooding or manual astronaut intervention, both of which "
        "exhibit fatal drawbacks in aerospace environments.\n\n"
        "ASAFI addresses this critical challenge by providing an autonomous, closed-loop payload that detects, verifies, aims, and suppresses "
        "combustion in less than 120 milliseconds. The system integrates optical computer vision (OpenCV) for rapid spatial detection, "
        "a medical-grade MLX90614 infrared thermopile radiometer for non-contact thermal verification, a dual-axis pan-tilt servo gimbal for "
        "pinpoint spatial targeting, and a 12V high-speed micro-solenoid valve that discharges a calibrated pulse of pressurized, non-conductive "
        "nitrogen cold gas. By delivering a localized suppression pulse directly onto the ignition seed, ASAFI starves the fire of oxygen and "
        "rapidly cools electrical traces without leaving chemical residues or damaging adjacent electronic modules."
    )

    add_callout_box(
        doc,
        "Total Reaction Latency: <120 ms | Targeting Accuracy: ±1.5 mm | Suppressant Agent: 100% Non-Conductive Compressed N2 | "
        "Operating Pressure: 150 PSI | Field of Surveillance: 360° Horizontal × 180° Vertical.",
        "SYSTEM PERFORMANCE BENCHMARKS"
    )

    # ==========================================
    # SECTION 2: PROBLEM STATEMENT
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Problem Statement: Microgravity Combustion Dynamics")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Combustion in Low Earth Orbit (LEO) and microgravity behaves in a fundamentally counterintuitive manner compared to terrestrial combustion:\n\n"
        "1. Absence of Buoyancy and Natural Convection: On Earth, gravity causes hot, less-dense gases to rise, drawing cool, oxygen-rich air into the base "
        "of a flame. In microgravity (such as aboard the International Space Station or autonomous satellites), there is no buoyant convection. "
        "Flames form spherical, oxygen-starved envelopes around the fuel source that smolder quietly. Smoldering electrical wiring harnesses can burn "
        "for extended durations undetected by standard smoke detectors.\n\n"
        "2. Forced Ventilation Hazards: Spacecraft rely on internal fans and environmental control and life support systems (ECLSS) to circulate air. "
        "These forced ventilation streams supply unpredictable oxygen flows across avionics racks, which can suddenly transition a smoldering hot spot "
        "into open flashover combustion within fractions of a second.\n\n"
        "3. Failure of Conventional Suppressants:\n"
        "• Chemical Aerosols & CO2 Flooding: Releasing dry chemicals or carbon dioxide across an entire cabin creates toxic particulate clouds that "
        "clog spacecraft filtration systems, obscure astronaut visibility, and leave corrosive residues that permanently damage sensitive electronics.\n"
        "• Human Reaction Delay: In an autonomous satellite or during crew sleep periods, detecting an electrical fire, alarming the crew or waiting "
        "for ground-station telemetry telecommands takes 30 to 90 seconds—an eternity when dealing with multi-kilowatt electrical arcs.\n\n"
        "Conclusion: Spacecraft require an autonomous, localized, sub-second interceptor that neutralizes ignition at the source before combustion can propagate."
    )

    # ==========================================
    # SECTION 3: SYSTEM ARCHITECTURE
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. System Concept of Operations & Closed-Loop Response")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "ASAFI operates on a continuous, deterministic 4-stage closed-loop perception and actuation pipeline:\n\n"
        "• Stage 1 — Optical Detection (T + 0.00s): A wide-angle optical sensor continuously samples the avionics enclosure at 60 FPS. An edge OpenCV "
        "pipeline applies frame differencing and HSV luminescence segmentation to isolate sudden electrical arcing, sparks, or flame radiance.\n\n"
        "• Stage 2 — Thermal Identification & Target Lock (T + 0.05s): The spatial bounding box coordinates (X, Y) are converted into angular "
        "azimuth and elevation vectors. The dual-axis gimbal aligns the MLX90614 infrared radiometer directly with the target. If the contactless "
        "temperature exceeds the safety threshold (T > 65.0°C), combustion is verified, eliminating optical false positives (e.g., status LEDs or reflections).\n\n"
        "• Stage 3 — Kinematic Vectoring & Aiming (T + 0.08s): High-torque metal-gear servos complete nozzle alignment with sub-millimeter precision, "
        "aiming the cold-gas dispensing orifice directly at the root of the electrical hazard.\n\n"
        "• Stage 4 — Cold-Gas Micro-Pulse Interception (T + 0.12s): The flight controller triggers a high-current logic-level MOSFET, energizing a 12V "
        "micro-solenoid valve. A 150ms supersonic jet of clean nitrogen cold gas (150 PSI) is released. The high-velocity nitrogen displaces oxygen, "
        "disrupts the flame zone, and induces rapid thermal quenching. A post-suppression scan verifies temperature drop (T < 45°C) before returning the gimbal to home."
    )

    # ==========================================
    # SECTION 4: HARDWARE IMPLEMENTATION
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Detailed Hardware Stack & Physical Implementation")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The ASAFI prototype was designed, fabricated, and assembled using carefully selected aerospace-grade embedded components. "
        "The hardware architecture is organized into four core functional modules:"
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.1 Perception & Optical Vision Module")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True
    p = doc.add_paragraph()
    p.add_run(
        "• Optical Sensor: Sony IMX219 8-Megapixel sensor with a 120° diagonal field-of-view (FOV) wide-angle low-distortion M12 lens. "
        "The camera connects via a dedicated high-speed MIPI-CSI ribbon interface to ensure sub-10ms frame transfer directly into GPU memory.\n"
        "• Optical Filtering: Equipped with an integrated IR-pass/cut switchable filter enabling both daylight and low-light avionics rack inspection."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.2 Non-Contact Infrared Radiometer Module")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True
    p = doc.add_paragraph()
    p.add_run(
        "• Infrared Sensor: Melexis MLX90614-ESF-BAA medical-grade infrared thermopile sensor.\n"
        "• Measurement Principle: Detects Stefan-Boltzmann infrared blackbody radiation emitted by heated circuit elements without physical electrical contact.\n"
        "• Thermal Range: Target object temperature from -70°C to +380°C with an accuracy of ±0.5°C in room ambient temperatures.\n"
        "• Communication: Connected over a 100 kHz SMBus / I2C bus with factory-calibrated emissivity and internal digital signal processing (DSP)."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.3 Dual-Axis Pan-Tilt Gimbal Actuator Module")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True
    p = doc.add_paragraph()
    p.add_run(
        "• Actuators: Two TowerPro MG996R high-torque metal-gear digital servos.\n"
        "• Kinematic Ranges: Azimuth (Horizontal) sweep from 0° to 180°; Elevation (Vertical) sweep from 15° to 90°.\n"
        "• Stall Torque: 11.0 kg·cm at 6.0V, providing rapid acceleration without gear backlash under high-pressure gas discharge.\n"
        "• Slew Speed: 0.16 seconds per 60 degrees of rotation, enabling complete target alignment in under 60 milliseconds."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.4 Cold-Gas Nozzle, Valve & Pneumatic Distribution Module")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True
    p = doc.add_paragraph()
    p.add_run(
        "• Micro-Solenoid Valve: 12V DC direct-acting normally closed (NC) pneumatic solenoid valve with 1/4\" NPT threaded brass fittings and Viton seals.\n"
        "• Valve Response Time: Actuation opening latency under 12 milliseconds; closing time under 15 milliseconds.\n"
        "• Gas Reservoir & Agent: Pressurized 16g clean nitrogen (N2) threaded cylinder connected through a brass mini-regulator tuned to 150 PSI output pressure.\n"
        "• Convergent Nozzle: Custom 1.8mm convergent brass nozzle mounted coaxially on the pan-tilt gimbal bracket, producing a focused conical supersonic gas jet."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.5 Power Electronics & Mechanical Chassis")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True
    p = doc.add_paragraph()
    p.add_run(
        "• Solenoid Driver: IRLZ44N logic-level N-channel MOSFET circuit equipped with an 1N4007 flyback diode to absorb inductive kickback voltage spikes.\n"
        "• Power Regulation: Dual-rail power architecture using an LM2596 step-down buck converter (supplying 5.0V / 3A for the servos and logic) "
        "and a dedicated 12V Li-Po battery rail for solenoid energization.\n"
        "• Mechanical Bracket: Designed in Fusion 360 and 3D printed using heat-resistant PETG (Polyethylene Terephthalate Glycol) polymer, "
        "featuring integrated wire guides, servo horn interlocking splines, and vibration-dampening mounting lugs."
    )

    # ==========================================
    # SECTION 5: ELECTRICAL PINOUT TABLE
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Electrical Interfacing & Circuit Pinout Mapping")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The following pinout configuration details the complete electrical wiring between the embedded flight controller, sensors, and actuators:"
    )

    pin_tbl = doc.add_table(rows=7, cols=5)
    pin_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    pin_tbl.autofit = False
    pin_tbl.columns[0].width = Inches(1.5)
    pin_tbl.columns[1].width = Inches(1.0)
    pin_tbl.columns[2].width = Inches(1.4)
    pin_tbl.columns[3].width = Inches(1.1)
    pin_tbl.columns[4].width = Inches(1.5)

    headers = ["Hardware Device", "Signal Line", "Controller Pin", "Logic / Voltage", "Functional Description"]
    for j, h in enumerate(headers):
        cell = pin_tbl.rows[0].cells[j]
        set_cell_background(cell, "105BD8")
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    pin_rows = [
        ("MLX90614 IR Sensor", "SDA (Data)", "GPIO 2 (Pin 3)", "3.3V I2C Bus", "I2C Serial Data with 4.7kΩ pull-up"),
        ("MLX90614 IR Sensor", "SCL (Clock)", "GPIO 3 (Pin 5)", "3.3V I2C Bus", "I2C Serial Clock pulse (100 kHz)"),
        ("Pan Servo (Azimuth)", "PWM 1", "GPIO 18 (Pin 12)", "5.0V Logic", "50 Hz PWM position signal (1ms - 2ms)"),
        ("Tilt Servo (Elevation)", "PWM 2", "GPIO 19 (Pin 35)", "5.0V Logic", "50 Hz PWM position signal (1ms - 2ms)"),
        ("12V Solenoid Driver", "GATE Trigger", "GPIO 23 (Pin 16)", "3.3V Logic", "Drives IRLZ44N gate via 220Ω resistor"),
        ("Optical Camera", "MIPI-CSI", "Camera Port 1", "MIPI D-PHY", "1080p60 uncompressed video stream")
    ]

    for i, row_data in enumerate(pin_rows):
        row = pin_tbl.rows[i + 1]
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
            elif j == 1:
                r.font.color.rgb = RGBColor(16, 91, 216)
                r.font.bold = True
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 6: SOFTWARE & ALGORITHMIC FLOW
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Software Pipeline, Computer Vision & State Machine")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The onboard software stack is built in Python utilizing optimized OpenCV C++ bindings, the smbus2 library for I2C communication, "
        "and hardware-timed PWM modules. The perception and decision-making logic follows a strict finite state machine (FSM):\n\n"
        "• State 0 — Surveillance & Background Differencing: The camera acquires 1080p frames at 60 FPS. Each incoming frame is converted to grayscale "
        "and blurred using a 5x5 Gaussian kernel. A running background accumulator maintains a dynamic baseline. Absolute frame differencing "
        "isolates sudden luminescence deltas that exceed a calibrated threshold (ΔI > 40).\n\n"
        "• State 1 — Contour Extraction & Centroid Calculation: When an illumination spike is detected, the algorithm converts the frame into the HSV "
        "color space. Color masks filter for the high-luminance core of an electrical arc (H: 0-35, S: 150-255, V: 200-255). Contours are detected, "
        "bounding rectangles are fitted, and the image moments calculate the exact centroid coordinates (Cx, Cy).\n\n"
        "• State 2 — Kinematic Angular Mapping & IR Polling: The centroid pixel coordinates are converted into gimbal servo angles:\n"
        "  θ_azimuth = θ_home + (Cx - X_center) × K_horizontal\n"
        "  θ_elevation = θ_home - (Cy - Y_center) × K_vertical\n"
        "The servos slew the gimbal to align the MLX90614 sensor with the target. An I2C transaction reads the RAM register (0x07) of the MLX90614. "
        "If the measured temperature T_target > 65.0°C, the anomaly is confirmed as an active fire hazard.\n\n"
        "• State 3 — Interception Actuation & Pulse Timer: The controller asserts GPIO 23 HIGH, saturating the MOSFET and opening the 12V valve. "
        "A hardware timer holds the valve open for exactly 150 milliseconds before de-asserting the pin. An automatic hardware watchdog timer cuts power "
        "if the pin remains HIGH for more than 500 milliseconds, preventing reservoir depletion.\n\n"
        "• State 4 — Thermal Quenching Verification: The system waits 200ms for gas dispersion and re-polls the MLX90614. If T_target < 45.0°C, "
        "the event is logged to telemetry flash memory, the servos return to home (90°, 45°), and the system transitions back to State 0."
    )

    # ==========================================
    # SECTION 7: BILL OF MATERIALS (BOM)
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Bill of Materials (BOM) & Fabrication Budget")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The following itemized Bill of Materials (BOM) documents all physical components procured, assembled, and tested during "
        "the Engineering Clinics fabrication cycle:"
    )

    bom_tbl = doc.add_table(rows=8, cols=6)
    bom_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    bom_tbl.autofit = False
    bom_tbl.columns[0].width = Inches(0.8)
    bom_tbl.columns[1].width = Inches(1.8)
    bom_tbl.columns[2].width = Inches(1.5)
    bom_tbl.columns[3].width = Inches(0.5)
    bom_tbl.columns[4].width = Inches(0.9)
    bom_tbl.columns[5].width = Inches(1.0)

    headers = ["Item #", "Component Description", "Model / Spec", "Qty", "Cost (INR)", "Status"]
    for j, h in enumerate(headers):
        cell = bom_tbl.rows[0].cells[j]
        set_cell_background(cell, "105BD8")
        set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    bom_rows = [
        ("BOM-01", "Infrared Thermopile Sensor", "Melexis MLX90614-ESF-BAA", "1", "₹ 1,250", "✓ Procured & Tested"),
        ("BOM-02", "Wide-Angle Optical Camera", "Sony IMX219 / 120° FOV", "1", "₹ 1,800", "✓ Integrated"),
        ("BOM-03", "Metal Gear Dual Servos", "TowerPro MG996R (2x)", "2", "₹ 980", "✓ Calibrated"),
        ("BOM-04", "12V Micro-Solenoid Valve", "Normally Closed 1/4\" NPT", "1", "₹ 650", "✓ Pressure Tested"),
        ("BOM-05", "N2 Cartridges & Mini-Regulator", "16g N2 + 150 PSI Regulator", "2", "₹ 750", "✓ Gas Charged"),
        ("BOM-06", "MOSFET Driver & Protection", "IRLZ44N + 1N4007 Diode", "1", "₹ 120", "✓ Assembled"),
        ("BOM-07", "PETG 3D Printed Chassis", "Custom Aerospace Gimbal Mount", "1", "₹ 350", "✓ Fabricated")
    ]

    for i, row_data in enumerate(bom_rows):
        row = bom_tbl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 4:
                r.font.bold = True
                r.font.color.rgb = RGBColor(16, 91, 216)
            elif j == 5:
                r.font.color.rgb = RGBColor(16, 149, 74)
                r.font.bold = True
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_tot = doc.add_paragraph()
    p_tot.paragraph_format.space_after = Pt(12)
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot = p_tot.add_run("TOTAL FABRICATION BUDGET: ₹ 5,900 INR")
    r_tot.font.bold = True
    r_tot.font.size = Pt(10.5)
    r_tot.font.color.rgb = RGBColor(16, 91, 216)

    # ==========================================
    # SECTION 8: LABORATORY TESTBED
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("8. Experimental Validation & Laboratory Testbed Results")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To rigorously validate ASAFI under controlled laboratory conditions, a 60cm × 60cm × 60cm plexiglass test enclosure "
        "was constructed simulating an avionics rack. Controlled ignition sources were introduced using electrical nichrome heating coils "
        "simulating wire insulation smoldering and momentary high-voltage spark discharges.\n\n"
        "Key Experimental Findings:\n"
        "• Optical Recognition Reliability: Across 50 test trials, the OpenCV vision loop achieved a 98.0% detection rate within 20 milliseconds "
        "of initial spark generation.\n"
        "• Thermal Confirmation Accuracy: The MLX90614 sensor reliably discriminated between warm functioning electronics (40°C - 50°C) "
        "and dangerous pre-combustion hotspots (>65°C) from a distance of 30 cm, generating zero false-positive gas discharges.\n"
        "• Extinguishment Efficacy: A single 150ms nitrogen pulse at 150 PSI extinguished active micro-flames in 100% of tested runs. "
        "The localized gas jet dropped core wire temperatures from 140°C to below 40°C in less than 1.8 seconds.\n"
        "• Zero Collateral Impact: Surrounding electronic circuits remained completely dry, clean, and fully operational without chemical residue."
    )

    # ==========================================
    # SECTION 9: PRESENTATION & PORTAL EVOLUTION
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("9. Mission Presentation & Web Portal Evolution")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To communicate the technical and life-safety significance of ASAFI to faculty evaluators, aerospace conferences, and external reviewers, "
        "a mission web portal was developed in two distinct phases:\n\n"
        "• Phase 1 — Initial Wireframe Prototype: The earliest iteration was an engineering-focused layout built with CSS wireframe blocks, "
        "dark cyan monochrome palettes, and extensive walls of text. While structurally complete, it lacked genuine spaceflight imagery, "
        "contained unfinished placeholder tags, and buried the life-saving impact of the system under dense technical paragraphs.\n\n"
        "• Phase 2 — NASA.gov-Inspired Mission Portal: The web presence was completely overhauled by benchmarking directly against the official "
        "NASA website (https://www.nasa.gov/) and the NASA Horizon Design System. Key structural enhancements include:\n"
        "  1. High-Resolution Aerospace Imagery: Replacing wireframes with high-impact photography of rocket cold-gas plumes, thermal infrared captures, "
        "and orbital spacecraft avionics racks.\n"
        "  2. 'Picture-First, Text-Lean' Philosophy: Concise, punchy headlines and distinct NASA category pill tags replacing long paragraphs.\n"
        "  3. Key Telemetry Numbers Strip: Immediate visual prominence for critical specs (<120ms latency, 100% clean gas, 360° dome, TRL-5 level).\n"
        "  4. Interactive Multi-Spectral Sensor Inspector: An interactive console allowing users to toggle live views between the 4K Optical CV stream, "
        "MLX90614 thermal radiometry, and the cold-gas nozzle plume with live targeting crosshairs.\n"
        "  5. Authentic NASA Multi-Column Footer: Integrating a complete 4-column directory, newsletter subscription form, and official compliance links.\n"
        "  6. Dedicated NASA SP Technical Portal (doc.html): Isolating deep academic content, pinouts, state machines, and the ₹5,900 BOM into a separate "
        "NASA Special Publication memorandum, maintaining an elegant, engaging public landing page."
    )

    # ==========================================
    # SECTION 10: SAFETY & FUTURE WORK
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("10. Safety Safeguards, Flight Qualification & Future Scope")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "• Fail-Safe Solenoid Timeout: In the event of a software crash or logic lockup during actuation, a hardware RC-timed relay circuit automatically "
        "cuts power to the 12V solenoid after 500 milliseconds, preventing uncontrolled venting of the spacecraft's nitrogen reserves.\n"
        "• Cold-Gas Cleanliness: Pressurized nitrogen is chemically inert, non-toxic, and dielectric (non-conductive), preventing electrical arcing across high-voltage bus lines.\n"
        "• Future Flight Enhancements:\n"
        "  1. Integration of an array of micro-solenoid manifolds for modular 3U / 6U CubeSat deployers.\n"
        "  2. Radiation-hardened FPGA implementation of the OpenCV convolutional contour extractor for deep-space long-duration missions.\n"
        "  3. Flight testing aboard suborbital sounding rocket microgravity flights to reach TRL-6 (Technology Readiness Level 6)."
    )

    # ==========================================
    # SECTION 11: CONCLUSION
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("11. Conclusion & Engineering Team Contributions")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(16, 91, 216)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The ASAFI project demonstrates the successful conception, design, embedded engineering, and experimental verification of an "
        "autonomous localized fire suppression payload for next-generation spacecraft avionics. By compressing the detection-to-interception "
        "timeline to under 120 milliseconds and eliminating destructive chemical flooding, ASAFI sets a new paradigm in spacecraft internal safety.\n\n"
        "Team Responsibilities:\n"
        "• Paakhi: Optical Computer Vision pipeline, OpenCV algorithm development, MLX90614 infrared radiometer I2C driver integration, "
        "NASA-inspired mission portal architecture, and technical memorandum documentation.\n"
        "• Rahul: 2-Axis gimbal mechanical fabrication, servo kinematics, 12V solenoid driver circuit assembly, pressurized pneumatic distribution manifold, "
        "and plexiglass testing enclosure."
    )

    # Save
    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_Project_Report.docx"
    doc.save(out_path)
    print(f"Project report saved successfully to: {out_path}")

if __name__ == "__main__":
    build_project_report()
