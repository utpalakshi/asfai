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

def add_callout_box(doc, text, title="DESIGN & SPECIFICATION HIGHLIGHT", border_color="113166", bg_color="F0F4F8"):
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
        f_run = f_p.add_run("ASAFI (NASA SP-2026) | Individual Contribution & Progress Report — Utpalakshi Upadhyaya")
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
    r_title = p_title.add_run("INDIVIDUAL PROGRESS & CONTRIBUTION REPORT\nPROJECT ASAFI (AUTONOMOUS SPACECRAFT AVIONICS FIRE INTERCEPTOR)")
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(11, 13, 23) # Deep Space Navy
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Technical Product Management, NASA Visual Ecosystem Research, Front-End UI/UX Architecture, and NASA SP-2026 Technical Specification Portal Synthesis")
    r_sub.font.size = Pt(10.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Executive Metadata Table
    meta_tbl = doc.add_table(rows=3, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Project Title:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Author & Role:", "Utpalakshi Upadhyaya — Lead UI/UX Designer, NASA Research & Visual Brand Strategist"),
        ("Evaluation Context:", "Engineering Clinics (ECS) Project Expo & Milestone Review")
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
        r1.font.color.rgb = RGBColor(17, 49, 102) if "Utpalakshi" in val else RGBColor(30, 41, 59)
        if "Utpalakshi" in val:
            r1.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 1
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 1: EXECUTIVE OVERVIEW & MY ROLE")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.1 Project Overview")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The Autonomous Spacecraft Avionics Fire Interceptor (ASAFI) represents an active, closed-loop safety payload engineered to mitigate localized "
        "electrical fire hazards, arc tracking, and thermal runaway events within sealed spacecraft avionics bays and pressurized satellite compartments. "
        "In microgravity environments, combustion behaves in an anomalous, non-intuitive manner due to the complete absence of natural buoyant convection: "
        "hot combustion gases do not rise, and cooler oxygenated air is not naturally drawn into the flame boundary. Instead, flames form spherical, "
        "oxygen-starved envelopes that smolder silently along wiring harnesses, circuit traces, and power distribution units (PDUs). Conventional spacecraft "
        "mitigation techniques—such as whole-cabin halon/chemical flooding or manual crew extinguisher deployment—introduce catastrophic collateral risks, "
        "including toxic atmospheric contamination, corrosive residue deposition across sensitive electronics, and prohibitive human reaction delays (30 to 90 seconds)."
    )

    p = doc.add_paragraph()
    p.add_run(
        "ASAFI circumvents these systemic vulnerabilities by executing an autonomous, multi-spectral closed-loop detection and suppression sequence in under "
        "1.5 seconds (<1.5s overall cycle, with sub-120ms reaction latency). Operating at the edge, the system ingests an optical video stream (30 FPS 720p/1080p), "
        "cross-correlates luminescence anomalies with a contactless Long-Wave Infrared (LWIR) thermopile radiometer (MLX90614), vectors a high-torque dual-axis "
        "pan-tilt gimbal to target coordinates, and discharges a non-conductive, pressurized nitrogen cold-gas micro-pulse (150 PSI) through a convergent "
        "supersonic nozzle."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Clarification of My Individual Scope:\n"
        "This progress and contribution report documents exclusively my individual deliverables as the Lead UI/UX Designer and Technical Documentation "
        "Architect. The core physical hardware assembly and embedded neural network weights were engineered by technical teammates. My operational mandate "
        "centered entirely on front-end user experience (UI/UX) engineering, NASA portal visual research, digital brand identity, technical documentation "
        "synthesis, and the complete architectural design of both the public showcase platform and the NASA SP-2026 Technical Specification Portal."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.2 Summary of My Individual Contributions")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Throughout the project lifecycle, I conceptualized, designed, and delivered the following key contributions:\n"
        "• NASA Visual Research & Benchmarking: Spearheaded visual and structural research into NASA’s official web ecosystem (NASA.gov) and the NASA Horizon "
        "Design System to establish an authentic aerospace aesthetic rather than a generic academic website.\n"
        "• Information Architecture & UI/UX Design: Planned all wireframes, user journeys, card layouts, and responsive component hierarchies across the main platform, "
        "ensuring complex sub-second aerospace operations are instantly digestible for faculty evaluators.\n"
        "• Visual Brand Identity & Tokens: Formulated the cohesive color token framework (Deep Space Navy, NASA Blue, Avionics Amber, Emergency Signal Red), "
        "glassmorphic card styling, monospace telemetry readouts, tactical HUD targeting indicators, and the custom ASAFI meatball insignia.\n"
        "• NASA Technical Specification Portal Design (doc.html): Architected, formatted, and synthesized the dedicated technical memorandum portal containing "
        "deep system architecture flowcharts, the 5-State Software Machine, the 4 core subsystem specifications, and the 5-column electrical pinout table.\n"
        "• Presentation & Expo Assets: Designed high-resolution 300 DPI exhibition posters, widescreen presentation slide decks, and live scripted demonstration "
        "flows for the final Engineering Clinics Project Expo."
    )

    # ==========================================
    # SECTION 2
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 2: NASA PORTAL RESEARCH, DESIGN SYSTEM & LANDING PAGE ARCHITECTURE")
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
        "A critical initiative I spearheaded was conducting deep visual and structural research into NASA’s official digital ecosystem (NASA.gov) "
        "and the NASA Horizon Design System. Standard university engineering presentations frequently suffer from generic, uninspired web templates, "
        "unformatted text blocks, and disjointed styling. To command immediate authority during evaluation, I established that ASAFI’s digital interface "
        "must look and operate like an authentic NASA flight-qualified mission control center and publication portal."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Through comparative analysis of NASA flight project portals (including Artemis, James Webb Space Telescope, and the Saffire fire safety experiments), "
        "I formulated a customized design token system specifically tailored for ASAFI:\n"
        "• Deep Space Navy (#0b0d17): Serves as the primary canvas background. Represents the vacuum of Low Earth Orbit while eliminating eye fatigue during prolonged evaluator review.\n"
        "• NASA Authority Blue (#113166) & Orbital Blue (#105bd8): Used for primary action buttons, verified flight badges, and key navigation states, conveying institutional trust and engineering rigor.\n"
        "• Avionics Amber (#f39c12): Applied to thermal anomaly callouts, warning indicators, and infrared radiometer readouts, immediately communicating elevated pre-combustion states.\n"
        "• Emergency Signal Red (#e74c3c): Reserved for immediate fire pulse alerts, critical target locks, and the stylized vector chevron in the ASAFI meatball insignia.\n"
        "• Crisp Telemetry White (#f8f9fa): High-contrast typography layer engineered for pin-sharp legibility across digital monitors and overhead classroom projectors."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Typography & Visual Motifs:\n"
        "I established a strict typographic hierarchy pairing modern geometric sans-serif (Inter) for fluid editorial prose and explanatory summaries with technical "
        "monospace typefaces (JetBrains Mono, Consolas) for raw telemetry metrics, latency timestamps, GPIO pinouts, and algorithmic state identifiers. "
        "The visual identity was further elevated through tactical aerospace motifs: precision 64px orthogonal background gridlines, subtle glassmorphic card backgrounds "
        "(rgba fills with backdrop-filter blur), sub-pixel high-contrast borders, animated status indicators, and SVG targeting crosshairs."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 Core Structural Sections Designed for the Main Platform (index.html)")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I structured and planned seven modular sections across index.html, optimizing visual hierarchy and cognitive pacing for faculty review:"
    )

    sections_info = [
        ("1. Tactical HUD Navigation & Meatball Emblem Header: ", 
         "Features a bespoke SVG emblem merging the iconic NASA meatball with ASAFI flight lettering, horizontal category navigation links, active status pills ('● LEO ORBIT READY' and 'TRL-5 VERIFIED'), and a direct link to the technical specification portal (doc.html)."),
        ("2. Mission Briefing & Problem Statement: ", 
         "Contextualizes the mission within NASA’s Spacecraft Fire Safety (Saffire) research. Clearly articulates the absence of natural buoyant convection, the spherical shape of microgravity flames, and the risks of cabin flooding."),
        ("3. High-Impact Telemetry Performance Metrics Bar: ", 
         "Four prominent telemetry cards highlighted by colored left border accents: '<1.5s Execution Cycle', '100% Clean N2 Gas', '360° × 180° Surveillance Dome', and 'TRL-5 Flight Qualified'."),
        ("4. Subsystem Architecture Visual Cards: ", 
         "Transforms complex hardware schematics into clean visual story cards detailing Edge Perception, LWIR Radiometry, 2-Axis Gimbal Kinematics, and the Cold-Gas Solenoid Nozzle."),
        ("5. Interactive Multi-Spectral Sensor Inspection Console: ", 
         "A signature interactive feature allowing reviewers to switch between the Optical 4K/720p Stream, the MLX90614 Thermal Radiometry view, and the Cold-Gas Nozzle Plume. Updates HUD targeting reticles, operating parameters, and specifications dynamically."),
        ("6. 3-Stage Closed-Loop Response Sequence: ", 
         "Photographic cards breaking down the rapid response cycle: 01 DETECT (T+0.00s) -> 02 ACQUIRE (T+0.05s) -> 03 INTERCEPT (T+0.12s), highlighting the sub-second timeline."),
        ("7. Simulated Cockpit HUD & Telemetry Console: ", 
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
        "The platform reached its current high standard through three deliberate evolutionary phases driven by usability auditing and peer feedback:\n"
        "• Initial Prototype Audit: The earliest prototype was an engineering-focused layout characterized by dense text blocks, static lists, "
        "and an abstract 3D CSS box. Usability testing revealed severe cognitive fatigue, poor mobile responsiveness, and an inability for external evaluators "
        "to grasp the sub-second speed and physical reality of the system.\n"
        "• NASA Editorial Pivot: I executed a major design pivot benchmarked against NASA.gov. I purged 70% of redundant text, replaced placeholder boxes with "
        "curated high-resolution aerospace photography, introduced distinct category badges (PROPULSION, IR SPECTRUM, LEO READY), and adopted an asymmetrical editorial grid.\n"
        "• Interactive Dynamics & Portal Decoupling: I conceptualized interactive states (multi-spectral sensor switcher, full-screen lightbox modal, and cockpit telemetry "
        "simulation) to bring the hardware alive digitally. Crucially, I decoupled deep engineering specifications from the landing page, migrating them into a dedicated "
        "NASA SP-2026 Technical Specification Portal (doc.html) to ensure academic rigor without compromising visual elegance."
    )

    # ==========================================
    # SECTION 3
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 3: NASA TECHNICAL SPECIFICATION PORTAL (doc.html) SYNTHESIS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A major accomplishment was designing, authoring, and structuring doc.html as an authentic NASA Special Publication Technical Memorandum (NASA SP-2026-09). "
        "This portal serves as the academic and technical bedrock of the project, providing evaluators with complete engineering transparency."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.1 System Architecture & Telemetry Pipeline")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I designed the visual and logical representation of the two-tier verification architecture:\n"
        "• Tier 1 (Optical Stream): Optical Camera captures 1080p frames at 60 FPS via MIPI-CSI → OpenCV Edge Perception pipeline executes background frame differencing "
        "and HSV color segmentation to identify candidate spark contours and compute centroid coordinates (Cx, Cy).\n"
        "• Tier 2 (Infrared Radiometer): The centroid is converted into servo angles → The dual-axis gimbal slews the MLX90614 infrared radiometer to the target → "
        "An I2C transaction reads object temperature registers. If T_target > 65.0°C, target lock is confirmed → Flight Computer triggers GPIO 23 → "
        "IRLZ44N MOSFET energizes the 12V valve nozzle, releasing a 150 PSI nitrogen micro-pulse."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.2 Subsystem Engineering Specifications (Structured Tables & Layout)")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I organized the four core flight hardware modules into structured, high-density specification cards within the technical specification portal:"
    )

    specs_tbl = doc.add_table(rows=5, cols=3)
    specs_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    specs_tbl.autofit = False
    specs_tbl.columns[0].width = Inches(1.5)
    specs_tbl.columns[1].width = Inches(2.2)
    specs_tbl.columns[2].width = Inches(2.8)

    s_headers = ["Subsystem Module", "Primary Component", "Verified Engineering Specifications"]
    for j, h in enumerate(s_headers):
        cell = specs_tbl.rows[0].cells[j]
        set_cell_background(cell, "113166")
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    subsystems_data = [
        ("Subsystem 01: Optical Vision", "Sony IMX219 (MIPI-CSI)", "1920×1080 @ 60 FPS; 120° Wide-Angle FOV; HSV Color Masking & Grayscale; ~16ms processing latency; rapid luminescence contour trigger."),
        ("Subsystem 02: Contactless IR", "Melexis MLX90614-ESF-BAA", "Target range: -70°C to +380°C; Ambient range: -40°C to +125°C; Accuracy: ±0.5°C; Resolution: 0.02°C; I2C/SMBus interface (100–400 kHz Fast-Mode)."),
        ("Subsystem 03: Pan-Tilt Gimbal", "TowerPro MG996R Servos (2x)", "0° to 180° Azimuth sweep; 15° to 90° Elevation sweep; 11.0 kg·cm metal-gear torque; Slew speed: 0.16s / 60°; 50 Hz hardware PWM position control."),
        ("Subsystem 04: Solenoid Interceptor", "12V Micro-Solenoid Valve", "12V DC operating voltage (MOSFET switched); < 12ms opening actuation; Pressurized clean nitrogen (100–150 PSI); 500ms automatic safety hardware cutoff.")
    ]

    for i, row_data in enumerate(subsystems_data):
        row = specs_tbl.rows[i + 1]
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

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.3 Software State Machine Sequence")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "I formulated and detailed the five-stage software state machine layout in doc.html, establishing clear transition criteria and fail-safe watchdog timers:\n"
        "• STATE 0 (Surveillance / Nominal): Background sampling at 1080p @ 60 FPS. Servos held at home position (90°, 45°); 12V valve closed. Dynamic baseline maintained.\n"
        "• STATE 1 (Anomaly Candidate Detected): OpenCV frame differencing isolates sudden illumination delta (ΔI > 40). Bounding box and centroid (Cx, Cy) computed.\n"
        "• STATE 2 (Thermal Co-Validation & Slew): Gimbal slews MLX90614 sensor toward calculated target azimuth and elevation. I2C reads object temperature. "
        "If T_target > 65.0°C, target confirmed; otherwise system returns to State 0.\n"
        "• STATE 3 (Interception Pulse Discharge): Flight controller asserts GPIO 23 HIGH for exactly 150 milliseconds. 12V valve fires pressurized nitrogen cold-gas jet.\n"
        "• STATE 4 (Post-Action Verification): System waits 200ms for gas dispersion and re-polls MLX90614. If T_target < 45.0°C, event is logged to telemetry memory, "
        "servos return to home, and system resumes State 0."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.4 Wiring & Electrical Pinout Reference Table")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To provide complete electrical interfacing transparency, I structured the 5-column electrical pinout reference table:"
    )

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
        ("MLX90614 IR Sensor", "SDA (Data)", "GPIO 2 (Pin 3)", "3.3V I2C Bus", "Serial data line with 4.7kΩ pull-up resistor"),
        ("MLX90614 IR Sensor", "SCL (Clock)", "GPIO 3 (Pin 5)", "3.3V I2C Bus", "Serial clock line operating at 400 kHz Fast-Mode"),
        ("Pan Servo (Azimuth)", "PWM 1", "GPIO 18 (Pin 12)", "5.0V Logic", "50 Hz hardware PWM position signal (1.0ms–2.0ms)"),
        ("Tilt Servo (Elevation)", "PWM 2", "GPIO 19 (Pin 35)", "5.0V Logic", "50 Hz hardware PWM position signal (1.0ms–2.0ms)"),
        ("12V Solenoid Driver", "GATE Trigger", "GPIO 23 (Pin 16)", "3.3V Logic", "Drives IRLZ44N gate via 220Ω resistor"),
        ("Optical Camera", "MIPI-CSI", "Camera Port 1", "MIPI D-PHY", "High-speed raw 1080p60 uncompressed video stream")
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
        "Through disciplined technical leadership, research benchmarking, and systematic information architecture, all front-end, visual, and documentation "
        "deliverables have been fully realized and verified for the Engineering Clinics Final Evaluation and Project Expo:"
    )

    deliv_items = [
        ("✓ Main Landing Platform UI/UX Architecture (index.html): ", "Full-bleed aerospace mission portal featuring tactical HUD navbar, hero briefing, key telemetry metrics strip, multi-spectral sensor inspection console, 3-stage visual timeline, live cockpit HUD telemetry simulator, and NASA multi-column footer."),
        ("✓ NASA SP-2026 Technical Specification Portal (doc.html): ", "Formal technical memorandum structured to NASA Special Publication standards with embedded vector system architecture diagrams, electrical pinout reference tables, 5-state software machines, and print stylesheets."),
        ("✓ Custom Aerospace Design System (styles.css): ", "Bespoke design tokens (Deep Space Navy, NASA Blue, Avionics Amber, Emergency Signal Red), glassmorphism HUD cards, 64px space gridlines, and card hover animations."),
        ("✓ Interactive Multi-Spectral Sensor & Telemetry Console Layout: ", "Engineered dynamic user flows allowing reviewers to inspect real-time optical vs. thermal vs. cold-gas states with simulated live targeting crosshairs."),
        ("✓ Flowcharts, System Architecture Diagrams & Pinout Tables: ", "High-clarity SVG architecture schematic and software state machine mapping States 0 to 4 with watchdog timers."),
        ("✓ Exhibition Presentation Posters, Slide Deck Design & Demo Scripts: ", "High-resolution 300 DPI project posters, widescreen presentation slide decks, and live scripted demonstration flows formulated for judging panel evaluation.")
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
    r = p_conc.add_run("Concluding Statement on Individual Leadership & Technical Synthesis:\n")
    r.font.bold = True
    r.font.color.rgb = RGBColor(17, 49, 102)
    p_conc.add_run(
        "My work on Project ASAFI demonstrates the vital role of UI/UX engineering, visual brand strategy, and structured technical documentation "
        "in bringing advanced aerospace concepts to life. By grounding the visual identity in NASA’s official web frameworks and systematically decoupling "
        "the public showcase from the deep technical memorandum, I ensured that ASAFI is communicated with immediate visual impact while maintaining "
        "uncompromising academic and engineering rigor. The resulting digital platform and documentation stand fully prepared for the Engineering Clinics "
        "evaluation, reflecting the highest standards of technical product management and user experience design."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    # Formal Sign-off
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_sig = p_sig.add_run("Submitted by:\n\n___________________________________\nUtpalakshi Upadhyaya\nLead UI/UX Designer, NASA Research & Visual Brand Strategist\nProject ASAFI · School of Electronics & Computer Science\nVIT-AP University")
    r_sig.font.size = Pt(9.5)
    r_sig.font.color.rgb = RGBColor(30, 41, 59)

    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_Individual_Contribution_Report.docx"
    doc.save(out_path)
    print(f"Report successfully generated and saved to: {out_path}")

if __name__ == "__main__":
    build_report()
