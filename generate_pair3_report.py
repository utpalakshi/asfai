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

def add_callout_box(doc, text, title="EXECUTIVE CALLOUT", border_color="113166", bg_color="F0F4F8"):
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
    run_title.font.color.rgb = RGBColor(17, 49, 102) # NASA Blue
    
    run_text = p.add_run(text)
    run_text.font.name = "Arial"
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_pair3_report():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("ASAFI (NASA SP-2026) | Pair 3 Progress & Contribution Report (Paakhi & Rahul)")
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
    # HEADER / COVER BLOCK
    # ==========================================
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_after = Pt(2)
    r_sub = p_header.add_run("VIT-AP UNIVERSITY · SCHOOL OF ELECTRONICS & COMPUTER SCIENCE · ENGINEERING CLINICS (ECS)")
    r_sub.font.size = Pt(8.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(17, 49, 102)
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("ASAFI: AUTONOMOUS SPACECRAFT AVIONICS FIRE INTERCEPTOR\nPAIR 3 PROGRESS & COMPREHENSIVE CONTRIBUTION REPORT")
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(11, 13, 23) # Deep Space Navy
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("UI/UX Design Methodology, Interactive Web Architecture, NASA SP-2026 Technical Specification Portal Synthesis, and ECS Financial Reimbursement Audit")
    r_sub.font.size = Pt(10.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Metadata Table
    meta_tbl = doc.add_table(rows=6, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Project Designation:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Reporting Sub-Team:", "Pair 3: Front-End UI/UX, Documentation & Financial Administration"),
        ("Team Members:", "Paakhi (Lead UI/UX Designer & Finance Lead) & Rahul (Lead Front-End Developer & Tech Writer)"),
        ("Evaluation Context:", "Engineering Clinics (ECS) Final Milestone Evaluation & Project Expo"),
        ("Live Web Portal:", "https://utpalakshi.github.io/asfai/ (Landing Page) & doc.html (Technical Spec)"),
        ("Project Scope Statement:", "Strictly Non-Hardware / Non-Backend Core: Full Front-End Architecture, Visual Hierarchy, NASA SP Documentation Synthesis, and Financial Billing Compliance")
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
        r1.font.color.rgb = RGBColor(17, 49, 102) if "https" in val or "Pair 3" in val else RGBColor(30, 41, 59)
        if "https" in val or "Pair 3" in val:
            r1.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 1
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 1: PROJECT OVERVIEW & PAIR 3 ROLE DEFINITION")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.1 Executive Summary")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The Autonomous Spacecraft Avionics Fire Interceptor (ASAFI) represents an advanced flight-safety payload conceived to address "
        "one of the most critical vulnerabilities in human and autonomous spaceflight: localized electrical ignition, arc tracking, and smoldering "
        "combustion within pressurized avionics bays and enclosed CubeSat enclosures. Operating in microgravity, where the complete absence "
        "of natural buoyant convection causes spherical, oxygen-starved flames to smolder undetected, ASAFI introduces a sub-second autonomous "
        "interception mechanism. By orchestrating a multi-spectral sensory loop—combining optical computer vision (30 FPS 720p ingestion) "
        "with Long-Wave Infrared (LWIR) thermography—and triggering a precision dual-axis pan-tilt gimbal armed with a cold-gas micro-pulse "
        "actuator, ASAFI achieves a verified end-to-end response cycle in under 1.5 seconds (<1.5s), neutralizing ignition before thermal runaway occurs."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Within the overarching multi-subsystem engineering structure of the ASAFI project, the technical labor was partitioned across specialized sub-teams:\n"
        "• Pair 1: Embedded AI Edge Pipeline & Hardware Actuation Core (Raspberry Pi 5, Lepton 3.5 LWIR, OpenCV inference, solenoid mechanics).\n"
        "• Pair 2: Ground Station Telemetry Engine & Backend Sensor Networking.\n"
        "• Pair 3 (Paakhi & Rahul): Front-End User Experience (UI/UX), Interactive Web Architecture, NASA SP-2026 Documentation Synthesis, "
        "Public Presentation Infrastructure, and Financial Billing / ECS Reimbursement Compliance.\n\n"
        "Pair 3’s explicit mandate is non-hardware and non-backend core. Our entire operational scope centers on the front-end user experience, "
        "interactive digital platform development, aesthetic and visual brand identity, complex technical system synthesis into digestible digital "
        "media, and rigorous administrative compliance for the Engineering Clinics (ECS) evaluation panel."
    )

    add_callout_box(
        doc,
        "Pair 3 Mandate: Full ownership of the public-facing digital platform (index.html), the NASA SP-2026 Technical Specification Portal (doc.html), "
        "schematic and flowchart visual synthesis, demonstration media decks, and the formal GST invoice audit for institutional ECS reimbursement.",
        "MANDATE SPECIFICATION"
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.2 Division of Responsibilities (Pair 3: Paakhi & Rahul)")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To ensure operational efficiency and exhaustive coverage of the evaluation criteria, labor within Pair 3 was precisely partitioned "
        "based on specialized skill sets while maintaining tight collaborative synchronization:"
    )

    p_paakhi = doc.add_paragraph()
    r = p_paakhi.add_run("Paakhi — Lead UI/UX Designer, Visual Strategist & Financial Documentation Manager:\n")
    r.font.bold = True
    r.font.color.rgb = RGBColor(17, 49, 102)
    p_paakhi.add_run(
        "• Information Architecture & Layout Strategy: Formulated the wireframe progression and user journey from initial text-cluttered prototypes "
        "to the final NASA Horizon-inspired editorial layout.\n"
        "• Design System Curation: Established the design token framework (Deep Space Navy, NASA Blue, Avionics Amber, Emergency Red), typography scales, "
        "glassmorphism overlays, and tactical HUD crosshair indicators.\n"
        "• Component Hierarchy & Visual Assets: Planned the multi-spectral sensor inspection console, 3-stage visual intercept sequence, and high-impact "
        "presentation posters for the ECS Project Expo.\n"
        "• Financial & Billing Management: Spearheaded the physical and digital collection, verification, and audit of all procurement receipts, "
        "tax invoices, and GST compliance logs to establish the formal ₹5,900 ECS reimbursement portfolio."
    )

    p_rahul = doc.add_paragraph()
    r = p_rahul.add_run("Rahul — Front-End Web Developer, Layout Implementer & Technical Memorandum Compiler:\n")
    r.font.bold = True
    r.font.color.rgb = RGBColor(17, 49, 102)
    p_rahul.add_run(
        "• Front-End Code Implementation: Transformed Paakhi’s visual designs and wireframes into clean, semantic, responsive HTML5, CSS3, and JavaScript code.\n"
        "• Interactive Component Development: Programmed the interactive sensor array switcher (script.js), the asynchronous simulated cockpit HUD "
        "telemetry loop with real-time value fluctuations, and the high-resolution media lightbox modal.\n"
        "• Technical Documentation Synthesis: Authored and structured doc.html in strict alignment with NASA Special Publication (NASA SP-2026) standards, "
        "integrating SVG system flowcharts, pinout tables, and software state machines.\n"
        "• Responsive Optimization: Implemented mobile breakpoint styling, touch target optimization, and CSS grid scaling across desktop, tablet, and mobile displays."
    )

    p_shared = doc.add_paragraph()
    r = p_shared.add_run("Shared Collaborative Deliverables:\n")
    r.font.bold = True
    p_shared.add_run(
        "• Iterative Design Reviews: Conducted joint UX audit sessions to identify usability bottlenecks, eliminate cognitive clutter, and refine visual contrast.\n"
        "• ECS Expo Presentation Deck: Jointly authored the project slide deck, scripted demonstration flow, and exhibition pitch tailored for faculty evaluators.\n"
        "• Final Submission Verification: Executed cross-browser functional testing, GitHub Pages cloud deployment, and print-layout verification."
    )

    # ==========================================
    # SECTION 2
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 2: UI/UX DESIGN METHODOLOGY & ITERATIVE WEB DEVELOPMENT")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.1 Initial Design Vision & Concept Ideation")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The primary design vision was to engineer a high-impact, responsive, web-based mission portal that transcended conventional academic website "
        "templates. Faculty evaluators, industry judges, and external engineering reviewers required an immediate, visceral understanding of ASAFI’s "
        "mission-critical capabilities. We established an aesthetic benchmark rooted in NASA’s official web ecosystem (NASA.gov) and the Horizon Design System, "
        "blending editorial visual storytelling with aerospace mission control precision."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Core Design Tokens & Aesthetic Specifications:\n"
        "• Color Palette:\n"
        "  - Deep Space Navy (#0b0d17): Base canvas background, evoking the vacuum of Low Earth Orbit and eliminating evaluator eye fatigue.\n"
        "  - NASA Blue (#113166) & Bright Blue (#105bd8): Primary brand accents, action buttons, and active status indicators denoting institutional authority.\n"
        "  - Avionics Amber (#f39c12): Warning threshold indicators, thermal anomaly callouts, and sensor telemetry highlights.\n"
        "  - Emergency Signal Red (#e74c3c): Critical intervention alerts, solenoid fire pulse indicators, and the stylized vector chevron.\n"
        "  - Crisp Telemetry White (#f8f9fa): Primary typography layer engineered for ultra-high contrast against dark backgrounds.\n"
        "• Typography Hierarchy: Paired modern geometric sans-serif (Inter) for fluid editorial prose with monospace typefaces (JetBrains Mono, Consolas) "
        "for raw telemetry data, GPIO pinouts, and algorithmic execution timeframes.\n"
        "• Tactical Visual Language: Leveraged subtle glassmorphism (rgba background fills with backdrop-filter blur), precision 64px space grid overlays, "
        "sub-pixel high-contrast borders, animated status beacons, and tactical HUD targeting crosshairs."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 Core Structural Sections Developed")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Pair 3 engineered seven modular structural sections across the landing platform, each designed to resolve a specific user journey requirement:"
    )

    sections_data = [
        ("1. Tactical HUD Header & Navigation Bar: ", 
         "Features the custom-crafted ASAFI Meatball emblem, a clean horizontal navigation menu, and dynamic status pills: '● LEO ORBIT READY' and 'TRL-5 VERIFIED'. Engineered with a sticky blur backdrop and accessible mobile drawer navigation."),
        ("2. Mission Briefing & Problem Alignment: ", 
         "Contextualizes ASAFI within NASA’s Spacecraft Fire Safety (Saffire) research program. Directly communicates the physics of microgravity combustion—the absence of buoyant convection and the catastrophic risks of whole-cabin flooding."),
        ("3. Key Performance Metrics Bar: ", 
         "Four high-impact telemetry metric callouts featuring colored left borders: '<1.5s Execution Cycle', '100% Clean N2 Gas', '360° × 180° Surveillance Dome', and 'TRL-5 Flight Qualified'."),
        ("4. Core System Architecture Walkthrough: ", 
         "Transforms complex hardware schematics into clean visual story cards detailing Edge Perception, LWIR Radiometry, 2-Axis Gimbal Kinematics, and the Cold-Gas Solenoid Nozzle."),
        ("5. Interactive Multi-Spectral Sensor Inspector: ", 
         "A signature interactive feature allowing reviewers to switch between the Optical 4K/720p Stream, the MLX90614 Thermal Radiometry view, and the Cold-Gas Nozzle Plume. Updates HUD targeting reticles, operating parameters, and specifications dynamically."),
        ("6. 3-Stage Closed-Loop Response Sequence: ", 
         "Photographic cards breaking down the rapid response cycle: 01 DETECT (T+0.00s) -> 02 ACQUIRE (T+0.05s) -> 03 INTERCEPT (T+0.12s), highlighting the sub-second timeline."),
        ("7. Simulated Cockpit HUD & Telemetry Console: ", 
         "A live targeting radar canvas with an animated reticle, live fluctuating telemetry readouts (Azimuth, Elevation, Core Temperature, Pressure), and a real-time event log terminal."),
        ("8. Authentic Multi-Column NASA Footer: ", 
         "Four structured link directories (Missions, Technology, Multimedia, Technical Docs), a functional Mission Alerts newsletter form, social links, and formal institutional/regulatory disclaimers.")
    ]

    for title, desc in sections_data:
        p = doc.add_paragraph()
        r1 = p.add_run(title)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(17, 49, 102)
        r2 = p.add_run(desc)
        r2.font.color.rgb = RGBColor(51, 65, 85)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.3 Iterative Improvements & Trial-and-Error Process")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The current platform is the culmination of rigorous iterative design cycles. The initial prototype suffered from significant UX deficiencies:\n"
        "• Trial 1 (Early Wireframe Phase): The initial layout relied on text-heavy descriptions, raw code blocks, and an abstract 3D CSS cube labeled 'ASAFI'. "
        "Evaluator feedback indicated high cognitive fatigue, an inability to understand the sub-second timing, and lack of visual proof.\n"
        "• Iteration 2 (NASA Editorial Pivot): Paakhi conducted an in-depth audit of NASA.gov, extracting the 'Picture-First, Text-Lean' methodology. "
        "We eliminated verbose paragraphs in favor of bold 1-2 sentence captions, integrated curated high-resolution aerospace photography, "
        "and introduced distinct category badges.\n"
        "• Iteration 3 (Interactive Dynamics): Rahul developed the JavaScript state switcher for sensor comparisons, the full-screen media lightbox modal, "
        "and the asynchronous telemetry fluctuation loop, injecting realism into the demonstration.\n"
        "• Iteration 4 (Separation of Concerns): To ensure academic rigor without cluttering the landing page, we decoupled deep specifications into "
        "a separate dedicated portal: doc.html (NASA SP-2026 Technical Memorandum), complete with an interactive SVG architecture schematic and itemized BOM."
    )

    # ==========================================
    # SECTION 3
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 3: TECHNICAL DOCUMENTATION, FLOWCHARTS & SCHEMATIC SYNTHESIS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.1 Visual Asset Creation & System Diagrams")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A cornerstone of Pair 3’s role was acting as the translation bridge between raw engineering data (from Pair 1 and Pair 2) and polished visual assets. "
        "Technical schematics provided in hand-drawn notes or raw code were synthesized into professional graphic representations:\n"
        "• Interactive SVG System Architecture Diagram: Engineered an uncompressed, responsive vector diagram embedded directly in doc.html. "
        "Illustrates the complete dataflow: Optical Camera (MIPI-CSI) → Edge Computer Vision → MLX90614 IR Sensor (I2C) → Flight Computer Logic → "
        "2-Axis Gimbal PWM & 12V Solenoid MOSFET Gate Trigger.\n"
        "• Software State Machine Flowchart: Mapped the algorithmic transitions through States 0 to 4 (Surveillance → Candidate Anomaly → "
        "Thermal Slew & Lock → Interception Pulse → Post-Action Cooling Verification) with color-coded logic paths and fail-safe watchdog timeouts.\n"
        "• Presentation Slide Deck & Poster Assets: Formatted high-resolution 300 DPI graphics, typography banners, and project summary cards for the ECS Expo exhibition."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.2 Technical Memorandum Formatting (NASA SP-2026 Standards)")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To satisfy the academic evaluation criteria of VIT-AP Engineering Clinics, all written documentation was structured according to formal "
        "NASA Special Publication (NASA SP) standards. In doc.html, we established:\n"
        "• Formal Document Control Block: Series designation NASA SP-2026-09, distribution notice (PUBLIC RELEASE), project ID (ASAFI-LEO-01), "
        "and author attributions.\n"
        "• Subsystem Engineering Specifications: Detailed tabular specifications for the optical sensor (IMX219, 1080p60, 120° FOV), infrared radiometer "
        "(MLX90614, -70°C to +380°C, ±0.5°C accuracy), 2-axis servos (MG996R, 11 kg·cm torque, 0.16s/60°), and solenoid valve (12V DC, 150 PSI, <12ms actuation).\n"
        "• Electrical Interfacing & Pinout Reference Table: A comprehensive 5-column table mapping hardware devices, signal lines, controller GPIO pins, "
        "logic voltages, and operational functions.\n"
        "• Embedded Print Engine: Formatted dedicated CSS @media print stylesheets enabling evaluators to generate an immaculate paper or PDF report "
        "with a single click."
    )

    # ==========================================
    # SECTION 4
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 4: FINANCIAL BILLING, BOM & ECS REIMBURSEMENT READINESS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.1 Bill of Materials (BOM) Verification")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Pair 3 undertook the exhaustive task of verifying, cataloging, and formatting the physical Bill of Materials (BOM) to ensure absolute financial "
        "transparency and audit-readiness for the Engineering Clinics evaluation committee. The table below represents the finalized, verified hardware breakdown:"
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

    headers = ["Item #", "Subsystem / Component", "Part / Specification", "Qty", "Cost (INR)", "Invoice Status"]
    for j, h in enumerate(headers):
        cell = bom_tbl.rows[0].cells[j]
        set_cell_background(cell, "113166")
        set_cell_margins(cell, top=70, bottom=70, left=60, right=60)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    bom_rows = [
        ("BOM-01", "Infrared Thermopile Radiometer", "Melexis MLX90614-ESF-BAA (I2C)", "1", "₹ 1,250", "GST Invoice Verified"),
        ("BOM-02", "Wide-Angle Optical Sensor", "Sony IMX219 / 120° FOV M12", "1", "₹ 1,800", "GST Invoice Verified"),
        ("BOM-03", "Dual-Axis Metal Gear Servos", "TowerPro MG996R High-Torque (2x)", "2", "₹ 980", "GST Invoice Verified"),
        ("BOM-04", "12V Micro-Solenoid Gas Valve", "Normally Closed 1/4\" NPT Brass", "1", "₹ 650", "GST Invoice Verified"),
        ("BOM-05", "Pressurized N2 Gas & Regulator", "16g N2 Cartridge + Brass Regulator", "2", "₹ 750", "Retail Bill Verified"),
        ("BOM-06", "MOSFET Driver & Protection", "IRLZ44N FET + 1N4007 Flyback Diode", "1", "₹ 120", "Component Tax Invoice"),
        ("BOM-07", "PETG 3D Printed Chassis", "Custom Aerospace Gimbal Mount", "1", "₹ 350", "Lab Fabrication Slip")
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
            r.font.size = Pt(8)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 4:
                r.font.bold = True
                r.font.color.rgb = RGBColor(17, 49, 102)
            elif j == 5:
                r.font.color.rgb = RGBColor(16, 149, 74)
                r.font.bold = True
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    p_tot = doc.add_paragraph()
    p_tot.paragraph_format.space_after = Pt(10)
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot = p_tot.add_run("TOTAL AUDITED EXPENDITURE: ₹ 5,900.00 INR")
    r_tot.font.bold = True
    r_tot.font.size = Pt(10)
    r_tot.font.color.rgb = RGBColor(17, 49, 102)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.2 Financial Receipts & GST Invoice Audit")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(30, 41, 59)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Led by Paakhi, Pair 3 established an institutional financial audit folder to guarantee seamless reimbursement processing by the VIT-AP finance office:\n"
        "• Invoice Verification & Itemization: Every physical receipt and digital PDF tax invoice was cross-checked against vendor GSTIN numbers, HSN/SAC codes, "
        "and physical hardware delivery slips.\n"
        "• Procurement Compliance: Verified that all electronic components (sensors, microcontrollers, servos) adhered to institutional procurement guidelines, "
        "with matching quantities and exact tax amounts.\n"
        "• Administrative Dossier: Created a structured billing portfolio containing: (1) Summary Reimbursement Claim Form, (2) Certified Bill of Materials, "
        "(3) Chronologically ordered original GST invoices, and (4) Photographic inventory proof of procured hardware."
    )

    # ==========================================
    # SECTION 5
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("SECTION 5: SUMMARY OF DELIVERABLES & ECS EXPO READINESS")
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(17, 49, 102)
    r.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Pair 3 has successfully fulfilled 100% of its designated milestones, establishing an exhaustive, production-ready portfolio for the Engineering Clinics "
        "Final Evaluation and Project Expo. The completed deliverables checklist includes:"
    )

    deliv_data = [
        ("✓ Interactive Mission Landing Portal (index.html): ", "Published live on GitHub Pages (https://utpalakshi.github.io/asfai/), featuring full-bleed hero banner, interactive multi-spectral sensor inspector, 3-stage response sequence, simulated cockpit HUD, and full NASA-style multi-column footer."),
        ("✓ Technical Specification Portal (doc.html): ", "Published live (https://utpalakshi.github.io/asfai/doc.html), structured in NASA SP-2026 format with embedded vector architecture diagrams, complete pinouts, state machines, and printable formatting."),
        ("✓ Design System & Stylesheet (styles.css): ", "Curated aerospace color tokens, responsive CSS grid, card zoom micro-interactions, HUD targeting reticle, and custom NASA branding."),
        ("✓ Front-End JavaScript Engine (script.js): ", "Clean, modular code powering the interactive sensor switcher, real-time simulated telemetry engine, media lightbox modal, and mobile navigation drawer."),
        ("✓ System Flowcharts & Schematic Assets: ", "Clean SVG architecture diagram and software FSM diagrams translating complex Pair 1/Pair 2 engineering data into intuitive visual blocks."),
        ("✓ Audited Financial Portfolio & BOM: ", "Complete ₹5,900 itemized BOM with verified GST tax invoices and physical inventory slips organized for institutional ECS reimbursement approval."),
        ("✓ Project Presentation Deck & Expo Script: ", "Exhibition pitch, presentation slides, and live demonstration script formulated for judging panel evaluation.")
    ]

    for title, desc in deliv_data:
        p = doc.add_paragraph()
        r1 = p.add_run(title)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(16, 149, 74)
        r2 = p.add_run(desc)
        r2.font.color.rgb = RGBColor(51, 65, 85)

    p_collab = doc.add_paragraph()
    p_collab.paragraph_format.space_before = Pt(8)
    r = p_collab.add_run("Statement on Team Collaboration & Project Synthesis:\n")
    r.font.bold = True
    r.font.color.rgb = RGBColor(17, 49, 102)
    p_collab.add_run(
        "The accomplishments of Pair 3 reflect a highly disciplined, synergistic partnership between Paakhi and Rahul. By maintaining clear lines "
        "of responsibility while actively challenging and refining each other’s deliverables, we succeeded in taking complex, disparate aerospace "
        "concepts and synthesizing them into a unified, visually captivating, and academically unassailable mission portal. Pair 3 stands fully prepared "
        "for the Engineering Clinics evaluation, confident that our digital platform, technical documentation, and administrative records reflect "
        "the highest standards of engineering excellence."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    # Formal Signatures
    sig_tbl = doc.add_table(rows=2, cols=2)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.columns[0].width = Inches(3.2)
    sig_tbl.columns[1].width = Inches(3.2)

    c0 = sig_tbl.cell(0, 0).paragraphs[0]
    c0.add_run("Prepared & Submitted by:\n\n___________________________________\nUtpalakshi Upadhyaya (Paakhi)\nLead UI/UX Designer & Financial Documentation Lead\nPair 3, Project ASAFI")
    
    c1 = sig_tbl.cell(0, 1).paragraphs[0]
    c1.add_run("Co-Submitted by:\n\n___________________________________\nRahul\nLead Front-End Developer & Technical Writer\nPair 3, Project ASAFI")

    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_Pair3_Contribution_Report.docx"
    doc.save(out_path)
    print(f"Pair 3 Report successfully saved to: {out_path}")

if __name__ == "__main__":
    build_pair3_report()
