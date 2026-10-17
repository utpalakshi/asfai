import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout_box(doc, text, title="DESIGN NOTE", border_color="113166", bg_color="F0F4F8"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
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
    run_title.font.color.rgb = RGBColor(17, 49, 102) # NASA Blue
    
    run_text = p.add_run(text)
    run_text.font.name = "Arial"
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_report():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("ASAFI (NASA SP-2026) | Progress & Individual Contribution Report — Paakhi & Rahul")
        f_run.font.name = "Arial"
        f_run.font.size = Pt(8.5)
        f_run.font.color.rgb = RGBColor(148, 163, 184)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(6)

    # ==========================================
    # TITLE & INSTITUTIONAL HEADER BLOCK
    # ==========================================
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_after = Pt(2)
    r_sub = p_header.add_run("VIT-AP UNIVERSITY · SCHOOL OF ELECTRONICS & COMPUTER SCIENCE\nENGINEERING CLINICS (ECS) | FINAL MILESTONE EVALUATION")
    r_sub.font.size = Pt(9)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(17, 49, 102)
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("COMPREHENSIVE PROGRESS & INDIVIDUAL CONTRIBUTION REPORT\nPROJECT ASAFI (AUTONOMOUS SPACECRAFT AVIONICS FIRE INTERCEPTOR)")
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(11, 13, 23) # Deep Space Navy
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("A Technical Product Management & UI/UX Design Synthesis: NASA Design Research, Interactive Front-End Architecture, Vector Schematic Synthesis, and Academic Milestone Readiness")
    r_sub.font.size = Pt(10.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Clean Metadata Table (No URLs, No Scope Statement, Strictly Executive Metadata)
    meta_tbl = doc.add_table(rows=4, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Project Title:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Lead Author & UI/UX Designer:", "Paakhi (Utpalakshi Upadhyaya) — Lead UI/UX Designer, NASA Research & Brand Strategist"),
        ("Front-End Development Co-Author:", "Rahul — Front-End Web Developer & Technical Documentation Compiler"),
        ("Evaluation Context:", "Engineering Clinics (ECS) Project Expo & Final Milestone Evaluation")
    ]
    for i, (label, val) in enumerate(meta_data):
        row = meta_tbl.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=50, bottom=50, left=90, right=90)
        set_cell_margins(c1, top=50, bottom=50, left=90, right=90)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(51, 65, 85)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)
        r1.font.color.rgb = RGBColor(17, 49, 102) if "Paakhi" in val else RGBColor(30, 41, 59)
        if "Paakhi" in val:
            r1.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 1
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 1: PROJECT OVERVIEW & ROLE BREAKDOWN")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.1 Executive Overview")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The Autonomous Spacecraft Avionics Fire Interceptor (ASAFI) represents an active, closed-loop safety payload engineered to mitigate localized "
        "electrical fire hazards, arc tracking, and thermal runaway events within sealed spacecraft avionics bays and pressurized satellite compartments. "
        "In microgravity environments, combustion behaves in an anomalous, non-intuitive manner due to the complete absence of natural buoyant convection. "
        "Hot combustion gases do not rise, and cooler oxygenated air is not naturally drawn into the flame boundary. Instead, flames form spherical, "
        "oxygen-starved envelopes that smolder silently along wiring harnesses, circuit traces, and power distribution units (PDUs). Conventional spacecraft "
        "mitigation techniques—such as whole-cabin halon/chemical flooding or manual crew extinguisher deployment—introduce catastrophic collateral risks, "
        "including toxic atmospheric contamination, corrosive residue deposition across sensitive electronics, and prohibitive human reaction delays (30 to 90 seconds)."
    )

    p = doc.add_paragraph()
    p.add_run(
        "ASAFI solves this critical aerospace vulnerability by executing an autonomous, multi-spectral closed-loop detection and suppression sequence in under "
        "1.5 seconds (<1.5s). Operating at the edge via a Raspberry Pi 5 (8GB) flight computer, the system ingests a 30 FPS 720p/1080p optical stream, "
        "cross-correlates luminescence anomalies with a Long-Wave Infrared (LWIR) radiometer (FLIR Lepton 3.5 / MLX90614), vectors a high-torque dual-axis "
        "pan-tilt gimbal to target coordinates, and discharges a non-conductive, pressurized nitrogen cold-gas micro-pulse (150 PSI) through a convergent "
        "supersonic nozzle."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Clarification of Project Scope:\n"
        "This progress and contribution report focuses exclusively on non-hardware and non-backend core deliverables. The core hardware engineering "
        "(circuit soldering, mechanical plumbing, embedded microcontrollers) and backend machine learning inference were handled by external technical collaborators. "
        "The entire operational mandate documented herein centers on front-end user interface and user experience (UI/UX) engineering, NASA portal visual "
        "research, digital brand identity, technical documentation synthesis, and milestone evaluation readiness."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.2 Detailed Contribution & Responsibility Matrix")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To achieve a world-class demonstration portal and exhaustive technical documentation, responsibilities were structured to maximize the synergy "
        "between Paakhi's design and research leadership and Rahul's front-end technical execution:"
    )

    # Detailed Matrix Table
    mat_tbl = doc.add_table(rows=5, cols=3)
    mat_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    mat_tbl.autofit = False
    mat_tbl.columns[0].width = Inches(1.5)
    mat_tbl.columns[1].width = Inches(2.5)
    mat_tbl.columns[2].width = Inches(2.5)

    headers = ["Functional Domain", "Paakhi (Lead UI/UX & Research)", "Rahul (Front-End & Tech Writer)"]
    for j, h in enumerate(headers):
        cell = mat_tbl.rows[0].cells[j]
        set_cell_background(cell, "113166")
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    matrix_rows = [
        ("NASA Research & Benchmark", "Spearheaded research into NASA.gov and the Horizon Design System to establish an authentic aerospace aesthetic.", "Reviewed technical standards and aligned subsystem specifications with NASA SP publication formats."),
        ("Information Architecture & UI", "Conceived full wireframe progression, user flows, component hierarchies, and visual storytelling.", "Translated wireframes into semantic HTML5 and modular CSS3/Tailwind utility architecture."),
        ("Design Tokens & Visual Identity", "Curated color tokens (Deep Space Navy, NASA Blue, Avionics Amber), typography scales, glassmorphism, and HUD elements.", "Implemented responsive layout grids, CSS transitions, hover card zooms, and high-contrast styling."),
        ("Interactive Dynamics & Docs", "Planned interactive user journeys for multi-spectral sensor inspection and simulated cockpit HUD consoles.", "Programmed JavaScript state engines (script.js), live fluctuating telemetry loops, lightbox modals, and compiled doc.html.")
    ]

    for i, row_data in enumerate(matrix_rows):
        row = mat_tbl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 1:
                r.font.color.rgb = RGBColor(17, 49, 102)
                r.font.bold = True
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 2
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 2: NASA PORTAL RESEARCH, UI/UX METHODOLOGY & WEB DESIGN")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.1 NASA Visual Identity Research & Concept Ideation")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A critical initiative led by Paakhi was conducting extensive visual and architectural research into NASA’s official web ecosystem (NASA.gov) "
        "and the Horizon Design System. Standard university engineering presentations frequently suffer from generic, uninspired web templates, "
        "unformatted text blocks, and disjointed styling. To command immediate authority during evaluation, Paakhi established that ASAFI’s digital interface "
        "must look and operate like an authentic NASA flight-qualified mission control center and publication portal."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Through iterative comparative analysis of NASA flight project portals (such as Artemis, James Webb Space Telescope, and the Saffire fire safety experiments), "
        "Paakhi developed a customized design token system specifically tailored for ASAFI:\n"
        "• Deep Space Navy (#0b0d17): Serves as the primary canvas background. Represents the vacuum of Low Earth Orbit while eliminating eye fatigue during prolonged evaluator review.\n"
        "• NASA Authority Blue (#113166) & Orbital Blue (#105bd8): Used for primary action buttons, verified flight badges, and key navigation states, conveying institutional trust and engineering rigor.\n"
        "• Avionics Amber (#f39c12): Applied to thermal anomaly callouts, warning indicators, and infrared radiometer readouts, immediately communicating elevated pre-combustion states.\n"
        "• Emergency Signal Red (#e74c3c): Reserved for immediate fire pulse alerts, critical target locks, and the stylized vector chevron in the ASAFI meatball insignia.\n"
        "• Crisp Telemetry White (#f8f9fa): High-contrast typography layer engineered for pin-sharp legibility across digital monitors and overhead classroom projectors."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Typography & Visual Language:\n"
        "Paakhi established a clear typographic hierarchy pairing modern geometric sans-serif (Inter) for fluid editorial prose and explanatory summaries "
        "with technical monospace typefaces (JetBrains Mono, Consolas) for raw telemetry metrics, latency timestamps, GPIO pinouts, and algorithmic state identifiers. "
        "The visual identity was further elevated through tactical aerospace motifs: precision 64px orthogonal background gridlines, subtle glassmorphic card backgrounds "
        "(rgba fills with backdrop-filter blur), sub-pixel high-contrast borders, animated status indicators, and SVG targeting crosshairs."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 Core Structural Sections & User Experience Architecture")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Paakhi structured and planned seven modular sections across index.html, which were then brought to life through front-end development:"
    )

    sections_info = [
        ("1. Tactical HUD Navigation & Meatball Emblem Header: ", 
         "Features a bespoke SVG emblem merging the iconic NASA meatball with ASAFI flight lettering, horizontal category navigation links, active status pills ('● LEO ORBIT READY' and 'TRL-5 VERIFIED'), and a direct link to the technical specification portal (doc.html)."),
        ("2. NASA Saffire Alignment & Microgravity Fire Physics Briefing: ", 
         "Contextualizes the mission within NASA’s Spacecraft Fire Safety (Saffire) research. Clearly articulates the absence of natural buoyant convection, the spherical shape of microgravity flames, and the risks of cabin flooding."),
        ("3. High-Impact Telemetry Performance Metrics Bar: ", 
         "Four prominent telemetry cards highlighted by colored left border accents: '<1.5s Execution Cycle', '100% Clean N2 Gas', '360° × 180° Surveillance Dome', and 'TRL-5 Flight Qualified'."),
        ("4. Core Subsystem Architecture Story Cards: ", 
         "Transforms complex hardware schematics into clean visual story cards detailing Edge Perception, LWIR Radiometry, 2-Axis Gimbal Kinematics, and the Cold-Gas Solenoid Nozzle."),
        ("5. Interactive Multi-Spectral Sensor Inspection Console: ", 
         "A signature interactive feature allowing reviewers to switch between the Optical 4K/720p Stream, the MLX90614 Thermal Radiometry view, and the Cold-Gas Nozzle Plume. Updates HUD targeting reticles, operating parameters, and specifications dynamically."),
        ("6. 3-Stage Closed-Loop Response Sequence: ", 
         "Photographic cards breaking down the rapid response cycle: 01 DETECT (T+0.00s) -> 02 ACQUIRE (T+0.05s) -> 03 INTERCEPT (T+0.12s), highlighting the sub-second timeline."),
        ("7. Simulated Cockpit HUD & Dynamic Telemetry Terminal: ", 
         "A live targeting radar canvas with an animated reticle, live fluctuating telemetry readouts (Azimuth, Elevation, Core Temperature, Pressure), and a real-time event log terminal.")
    ]

    for title, desc in sections_info:
        p = doc.add_paragraph()
        r1 = p.add_run(title)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(17, 49, 102)
        r2 = p.add_run(desc)
        r2.font.color.rgb = RGBColor(51, 65, 85)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.3 Iterative UI/UX Evolution (Trial & Error Design Process)")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The platform reached its current state through four distinct evolutionary phases driven by continuous heuristic evaluation and usability feedback:\n"
        "• Stage 1 (Initial Text-Heavy Wireframe): The earliest prototype was an engineering-focused layout characterized by dense text blocks, static lists, "
        "and an abstract 3D CSS box labeled 'ASAFI'. Internal testing revealed high cognitive fatigue, poor mobile responsiveness, and an inability for external "
        "evaluators to grasp the sub-second speed of the system.\n"
        "• Stage 2 (NASA Picture-First Shift): Paakhi executed a major design pivot benchmarked against NASA.gov. We purged 70% of redundant text, "
        "replaced placeholder wireframe boxes with curated high-resolution aerospace photography, introduced distinct category badges, and adopted an asymmetrical editorial grid.\n"
        "• Stage 3 (Interactive Dynamics Integration): Working collaboratively, Rahul implemented JavaScript interactive states (script.js), including full-screen "
        "media lightbox modals, interactive multi-spectral sensor inspection, and an asynchronous telemetry simulator that updates realistic sensor readouts every 1.8 seconds.\n"
        "• Stage 4 (Portal Decoupling): To preserve visual elegance while ensuring rigorous academic compliance, Paakhi decoupled the architecture: index.html was "
        "reserved for the public mission briefing and live telemetry, while deep specifications, software state machines, and electrical pinouts were migrated into a dedicated "
        "NASA SP-2026 Technical Specification Portal (doc.html)."
    )

    # ==========================================
    # SECTION 3
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 3: TECHNICAL DOCUMENTATION & SYSTEM DIAGRAM SYNTHESIS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.1 Architecture & Flowchart Visual Synthesis")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A primary responsibility led by Paakhi was synthesizing complex embedded engineering concepts into polished, presentation-grade vector diagrams. "
        "Raw hand-drawn schematics and text notes were translated into structured visual assets:\n"
        "• Interactive Vector System Architecture Diagram: Developed an SVG diagram embedded directly in doc.html. It cleanly illustrates the full hardware dataflow: "
        "Optical Camera (MIPI-CSI 30 FPS) → Edge OpenCV Perception → MLX90614 IR Sensor (I2C 400kHz Fast-Mode) → Flight Controller Logic → "
        "2-Axis Gimbal PWM Timer & 12V Solenoid Valve MOSFET Gate Trigger.\n"
        "• Software State Machine Flowchart: Mapped the complete algorithmic state progression (State 0: Surveillance → State 1: Anomaly Detected → "
        "State 2: Thermal Slew & Verification → State 3: Micro-Pulse Interception → State 4: Post-Action Cooling Verification) with color-coded logic paths "
        "and automatic watchdog timeouts."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.2 NASA SP-2026 Technical Specification Standards")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To meet the rigorous evaluation standards of the Engineering Clinics curriculum, all technical documentation was aligned with formal "
        "NASA Special Publication (NASA SP) guidelines. In doc.html, we structured:\n"
        "• Formal Document Control Block: Series designation NASA SP-2026-09, distribution notice (PUBLIC RELEASE), project ID (ASAFI-LEO-01), and author attributions.\n"
        "• Subsystem Engineering Specifications: Detailed tabular specifications for the optical sensor (IMX219, 1080p60, 120° FOV), infrared radiometer "
        "(MLX90614, -70°C to +380°C, ±0.5°C accuracy), 2-axis servos (MG996R, 11 kg·cm torque, 0.16s/60°), and solenoid valve (12V DC, 150 PSI, <12ms actuation).\n"
        "• Electrical Interfacing & Pinout Reference Table: A comprehensive 5-column table mapping hardware devices, signal lines, controller GPIO pins, "
        "logic voltages, and operational functions.\n"
        "• Embedded Print Stylesheet: Integrated dedicated CSS print rules (@media print) that automatically format clean page breaks, high-contrast borders, "
        "and crisp tabular lines for physical print evaluation."
    )

    # Pinout Table (Technical Depth, Non-BOM)
    pin_tbl = doc.add_table(rows=7, cols=5)
    pin_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    pin_tbl.autofit = False
    pin_tbl.columns[0].width = Inches(1.5)
    pin_tbl.columns[1].width = Inches(1.0)
    pin_tbl.columns[2].width = Inches(1.4)
    pin_tbl.columns[3].width = Inches(1.1)
    pin_tbl.columns[4].width = Inches(1.5)

    pin_headers = ["Hardware Module", "Signal Line", "Controller Pin", "Logic / Voltage", "Functional Description"]
    for j, h in enumerate(pin_headers):
        cell = pin_tbl.rows[0].cells[j]
        set_cell_background(cell, "113166")
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    pin_rows = [
        ("MLX90614 IR Sensor", "SDA (Data)", "GPIO 2 (Pin 3)", "3.3V I2C Bus", "Serial data line with 4.7kΩ pull-up"),
        ("MLX90614 IR Sensor", "SCL (Clock)", "GPIO 3 (Pin 5)", "3.3V I2C Bus", "Serial clock line (400 kHz Fast-Mode)"),
        ("Pan Servo (Azimuth)", "PWM 1", "GPIO 18 (Pin 12)", "5.0V Logic", "50 Hz hardware PWM position signal"),
        ("Tilt Servo (Elevation)", "PWM 2", "GPIO 19 (Pin 35)", "5.0V Logic", "50 Hz hardware PWM position signal"),
        ("12V Solenoid Driver", "GATE Trigger", "GPIO 23 (Pin 16)", "3.3V Logic", "Drives IRLZ44N gate via 220Ω resistor"),
        ("Optical Camera", "MIPI-CSI", "Camera Port 1", "MIPI D-PHY", "High-speed raw 1080p60 video stream")
    ]

    for i, row_data in enumerate(pin_rows):
        row = pin_tbl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(8)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 1:
                r.font.color.rgb = RGBColor(17, 49, 102)
                r.font.bold = True
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 4
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 4: DELIVERABLES CHECKLIST & EXPO READINESS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Through disciplined collaboration and clear division of labor, all front-end, visual, and documentation deliverables have been completed "
        "and validated for the Engineering Clinics Final Evaluation and Project Expo:"
    )

    deliv_items = [
        ("✓ Interactive Mission Landing Platform (index.html): ", "Responsive aerospace landing portal featuring tactical HUD navbar, hero briefing, key metrics strip, multi-spectral sensor inspection console, 3-stage visual timeline, live cockpit HUD telemetry simulator, and NASA multi-column footer."),
        ("✓ NASA SP-2026 Technical Specification Portal (doc.html): ", "Formal technical memorandum structured to NASA Special Publication standards with embedded SVG system architecture diagrams, electrical pinout reference tables, software state machines, and print stylesheets."),
        ("✓ Custom Aerospace Design System (styles.css): ", "Bespoke design tokens (Deep Space Navy, NASA Blue, Avionics Amber, Emergency Red), glassmorphism HUD cards, 64px space gridlines, and card hover animations."),
        ("✓ Interactive Front-End Engine (script.js): ", "Clean JavaScript modules powering the multi-spectral sensor switcher, asynchronous fluctuating cockpit telemetry engine, full-screen lightbox modal, and responsive mobile navigation drawer."),
        ("✓ Vector Schematics & System Flowcharts: ", "High-clarity SVG architecture diagram and software state machine mapping States 0 to 4 with watchdog timers."),
        ("✓ Exhibition Presentation Deck & Poster Suite: ", "High-resolution 300 DPI project posters, widescreen presentation slide decks, and live scripted demonstration flows formulated for judging panel evaluation.")
    ]

    for title, desc in deliv_items:
        p = doc.add_paragraph()
        r1 = p.add_run(title)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(16, 149, 74) # Green
        r2 = p.add_run(desc)
        r2.font.color.rgb = RGBColor(51, 65, 85)

    p_conc = doc.add_paragraph()
    p_conc.paragraph_format.space_before = Pt(8)
    r = p_conc.add_run("Concluding Statement on Leadership & Technical Synthesis:\n")
    r.font.bold = True
    r.font.color.rgb = RGBColor(17, 49, 102)
    p_conc.add_run(
        "This project report represents the successful realization of a comprehensive UI/UX and technical communication strategy for Project ASAFI. "
        "Through Paakhi’s design leadership, NASA visual benchmarking, and systematic information architecture, combined with Rahul’s front-end code implementation, "
        "our team transformed dense, complex aerospace engineering into an intuitive, visually captivating, and scientifically unassailable digital platform. "
        "The resulting web portal and technical documentation reflect the highest standards of academic rigor, professional product design, and milestone readiness."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    # Formal Sign-off
    sig_tbl = doc.add_table(rows=2, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.columns[0].width = Inches(3.2)
    sig_tbl.columns[1].width = Inches(3.2)

    c0 = sig_tbl.cell(0, 0).paragraphs[0]
    c0.add_run("Lead Author & UI/UX Designer:\n\n___________________________________\nPaakhi (Utpalakshi Upadhyaya)\nLead UI/UX Designer & NASA Research Lead\nProject ASAFI · VIT-AP University")
    
    c1 = sig_tbl.cell(0, 1).paragraphs[0]
    c1.add_run("Co-Author & Front-End Developer:\n\n___________________________________\nRahul\nLead Front-End Developer & Technical Writer\nProject ASAFI · VIT-AP University")

    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_Progress_and_Contribution_Report.docx"
    doc.save(out_path)
    print(f"Report successfully generated and saved to: {out_path}")

if __name__ == "__main__":
    build_report()
