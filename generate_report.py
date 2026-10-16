import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_callout_box(doc, text, title="NOTE / HIGHLIGHT", border_color="105BD8", bg_color="F0F4F8"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border only
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

def build_report():
    doc = docx.Document()
    
    # Page setup - Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header & Footer setup
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        f_run = f_p.add_run("ASAFI UI/UX Engineering & Design Evolution Report | VIT-AP Engineering Clinics")
        f_run.font.name = "Arial"
        f_run.font.size = Pt(8.5)
        f_run.font.color.rgb = RGBColor(148, 163, 184)
        
    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.2
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
    r_title = p_title.add_run("PROJECT CONTRIBUTION & UI/UX DESIGN REPORT:\nREDESIGNING THE ASAFI AEROSPACE MISSION PORTAL")
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("From Technical Wireframe to NASA.gov-Inspired Visual Landing Page & Dedicated Flight Documentation Portal")
    r_sub.font.size = Pt(11.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Metadata Table
    meta_tbl = doc.add_table(rows=4, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    meta_tbl.columns[0].width = Inches(2.2)
    meta_tbl.columns[1].width = Inches(4.3)
    
    meta_data = [
        ("Student Contributor & Role:", "Utpalakshi (Paakhi) — Lead UI/UX Designer & Frontend Engineer"),
        ("Project Title:", "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor)"),
        ("Team Collaborators:", "Paakhi (UI/UX, CV & Perception) & Rahul (Avionics Hardware & Actuation)"),
        ("Live Deployment URL:", "https://utpalakshi.github.io/asfai/ (GitHub Pages)")
    ]
    for i, (label, val) in enumerate(meta_data):
        row = meta_tbl.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        
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
        r1.font.color.rgb = RGBColor(16, 91, 216) if "http" in val else RGBColor(30, 41, 59)
        if "http" in val:
            r1.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 1: EXECUTIVE SUMMARY
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Executive Summary & Overview of Contributions")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "This project report documents the design, user experience (UI/UX) engineering, and architectural evolution of the web portal for "
        "ASAFI (Autonomous Spacecraft Avionics Fire Interceptor), developed under the Engineering Clinics curriculum at VIT-AP University. "
        "As the lead UI/UX designer and frontend engineer on this project, my primary mandate was to transform an initial highly technical, "
        "text-cluttered prototype into a world-class, visual-first mission landing page modeled directly after the modern editorial standards "
        "of NASA.gov and the NASA Web Design System (Horizon)."
    )

    p = doc.add_paragraph()
    p.add_run(
        "Key contributions executed during this development lifecycle include:\n"
        "• Formulation of the aerospace user problem statement and information architecture (IA).\n"
        "• Redesigning the landing page from a text-heavy prototype into a high-impact, photo-dominant interface with minimal, precise copy.\n"
        "• Integrating core visual design concepts from NASA.gov (full-bleed hero, asymmetrical editorial grid, telemetry numbers strip, and high-contrast dark space theme).\n"
        "• Designing and implementing an interactive multi-spectral sensor inspection console (Optical 4K CV vs. MLX90614 Thermal IR vs. Cold-Gas Plume).\n"
        "• Building an authentic, multi-column NASA footer with agency directories, mission alert subscription, and academic/legal disclosures.\n"
        "• Establishing a dedicated NASA Special Publication (SP-2026-09) Technical Memorandum portal (doc.html) containing system architecture diagrams, wiring pinouts, state machines, and Bill of Materials (BOM).\n"
        "• Managing version control with Git and publishing the completed portal live to the global web via GitHub Pages."
    )

    add_callout_box(
        doc,
        "The live interactive mission portal is fully deployed and accessible on any device at: https://utpalakshi.github.io/asfai/. "
        "The complete technical documentation manual is available at: https://utpalakshi.github.io/asfai/doc.html.",
        "LIVE ACCESS POINT"
    )

    # ==========================================
    # SECTION 2: PROBLEM STATEMENT
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("2. Problem Statement: Engineering Challenge & UI/UX Challenge")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    h2 = doc.add_heading(level=2)
    r_h2 = h2.add_run("2.1 The Aerospace Engineering Problem")
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(30, 41, 59)
    r_h2.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "In Low Earth Orbit (LEO) and microgravity environments, fire behaves fundamentally differently than on Earth. Without natural gravity, "
        "there is no buoyant convection—hot gases do not rise, and fresh oxygen is not automatically pulled into a combustion zone. "
        "Instead, electrical short circuits or battery thermal runaway inside sealed spacecraft avionics bays form localized, spherical, "
        "oxygen-starved smoldering hot spots that can silently spread throughout wire harnesses.\n\n"
        "Traditional spacecraft safety mechanisms present critical vulnerabilities:\n"
        "1. Cabin Flooding Suppressants: Total flooding of avionics bays with halon or chemical aerosols causes corrosive residue and electrical shorts across healthy adjacent avionics racks.\n"
        "2. Reaction Latency: Crewed emergency procedures or ground-station telecommand intervention take 30 to 90 seconds—prohibitive when dealing with microsecond electronic arcs.\n"
        "3. Solution: ASAFI solves this through localized sub-second detection (<120ms) using optical computer vision, contactless MLX90614 infrared pyrometry, a 2-axis gimbal, and a focused cold-gas nitrogen pulse."
    )

    h2 = doc.add_heading(level=2)
    r_h2 = h2.add_run("2.2 The UI/UX & Web Communication Challenge")
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(30, 41, 59)
    r_h2.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "While the underlying aerospace engineering concept was robust, the project suffered from a severe UI/UX communication gap. "
        "The initial web prototype had several distinct UX shortcomings that hindered evaluation:\n"
        "• Cognitive Overload & Text Clutter: Long technical paragraphs, unformatted lists, and repetitive technical descriptions created fatigue.\n"
        "• Absence of Real Imagery: The interface relied on synthetic CSS wireframe boxes, cartoonish polygon shapes, and empty bracket placeholders (e.g., '[INSERT MODEL / RESOLUTION]').\n"
        "• Weak Visual Hierarchy: Critical life-safety performance metrics (such as the sub-120ms latency and non-conductive gas rating) were buried deep inside text blocks.\n"
        "• Incomplete User Experience: There was no comprehensive footer, no clear navigational journey, and no clear separation between a public-facing stakeholder landing page and rigorous academic documentation."
    )

    # ==========================================
    # SECTION 3: EVOLUTION CHRONOLOGY
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("3. Design Evolution: Phase 1 (Initial Prototype) vs. Phase 2 (NASA Redesign)")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To systematically resolve these deficiencies, I structured the UI/UX overhaul into two distinct phases, transitioning "
        "the website from an internal developer wireframe into a modern, production-grade aerospace portal."
    )

    # Comparative Table
    comp_tbl = doc.add_table(rows=7, cols=3)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_tbl.autofit = False
    comp_tbl.columns[0].width = Inches(1.5)
    comp_tbl.columns[1].width = Inches(2.5)
    comp_tbl.columns[2].width = Inches(2.5)

    headers = ["UX Attribute", "Phase 1: Initial Prototype", "Phase 2: NASA.gov Redesign"]
    for j, h in enumerate(headers):
        cell = comp_tbl.rows[0].cells[j]
        set_cell_background(cell, "105BD8")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    comp_rows = [
        ("Design Philosophy", "Text-heavy technical spec with dark cyan CSS wireframes.", "Photo-dominant editorial portal ('More pictures, less text, high precision')."),
        ("Hero Banner", "Small split layout with an abstract CSS 3D box labeled 'ASAFI'.", "Full-bleed high-res orbital view with NASA typography and punchy action CTAs."),
        ("Visuals & Assets", "Zero photographic imagery; CSS simulated wireframe craft.", "Curated high-res space, rocket plume, thermal IR, and avionics photography."),
        ("Performance Metrics", "Buried in text bullet points and code tags.", "Prominent Key Metrics Bar (<120ms, 100% Clean N2, 360° Dome, TRL-5 Qualified)."),
        ("Sensor Inspection", "Static text buttons switching text paragraphs in a box.", "Interactive Multi-Spectral Inspector with live reticle and Optical vs. IR comparison."),
        ("Footer Architecture", "Single line text disclaimer: 'ASAFI // Aerospace prototype'.", "Authentic 4-column NASA directory, newsletter subscription, and compliance disclosures.")
    ]

    for i, row_data in enumerate(comp_rows):
        row = comp_tbl.rows[i + 1]
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif j == 1:
                r.font.color.rgb = RGBColor(100, 116, 139)
            else:
                r.font.color.rgb = RGBColor(16, 91, 216)
                r.font.bold = (i == 0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 4: NASA.GOV DESIGN INSPIRATION
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("4. UI/UX Reference: Adopting the NASA.gov Design Language")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "The redesign was directly benchmarked against NASA’s official flagship website (https://www.nasa.gov/) and the "
        "NASA Horizon Design System. NASA's web ecosystem is renowned for communicating complex spaceflight engineering "
        "through visual storytelling. The specific design patterns adapted include:"
    )

    nasa_points = [
        ("1. Full-Bleed Hero Banner & High-Impact Typography: ", 
         "NASA's homepage opens with breathtaking full-screen imagery overlaid with high-contrast, uppercase headings and pill badges. For ASAFI, I implemented an orbital horizon view paired with the bold headline 'AUTONOMOUS AVIONICS DEFENSE' and a sharp, two-line value proposition."),
        ("2. Asymmetrical Editorial Photo Grid: ", 
         "Instead of repetitive cards, NASA utilizes dynamic grid layouts with dominant featured highlights and companion cards. I designed a 5-card responsive grid showcasing cold-gas plume firing, MLX90614 infrared heat spectrum, 2-axis gimbal kinematics, OpenCV edge vision, and CubeSat bay integration."),
        ("3. Visual Telemetry Bar: ", 
         "NASA presents complex mission numbers in clean, legible statistical strips. I engineered a 4-metric strip with colored left accents highlighting sub-120ms latency, 100% clean agent status, 360° coverage, and TRL-5 qualification."),
        ("4. Interactive Media Lightbox: ", 
         "Allowing visitors to inspect high-resolution photography with NASA metadata overlays without leaving the page."),
        ("5. Authentic Multi-Column NASA Footer: ", 
         "A four-column structured directory covering Missions, Technology, Multimedia Archive, and Technical Docs, integrated with a newsletter subscription form and official administrative links (FOIA, Privacy, Section 508 Accessibility, No FEAR Act).")
    ]

    for title, desc in nasa_points:
        p = doc.add_paragraph()
        r1 = p.add_run(title)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(16, 91, 216)
        r2 = p.add_run(desc)
        r2.font.color.rgb = RGBColor(51, 65, 85)

    # ==========================================
    # SECTION 5: UI/UX CONTRIBUTIONS & ARCHITECTURE
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("5. Detailed UI/UX Engineering & Technical Implementations")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    h2 = doc.add_heading(level=2)
    r_h2 = h2.add_run("5.1 Design Tokens & Color Palette")
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(30, 41, 59)
    r_h2.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To establish authentic aerospace branding, I defined custom CSS variables and Tailwind utility configurations:\n"
        "• Deep Space Black (#060911): The primary backdrop, reducing eye fatigue and giving visual pop to images.\n"
        "• NASA Blue (#105BD8) & Hover (#0B4DBF): Primary brand and action buttons, symbolizing institutional authority.\n"
        "• NASA Red (#E03C31): Used on the custom SVG Meatball chevron and emergency intervention triggers.\n"
        "• Cyan Telemetry (#00E5FF): Active HUD reticle lines, crosshairs, and digital data readouts.\n"
        "• Solar Warning (#FF8A00): Thermal threshold alerts and MLX90614 infrared temperature states.\n"
        "• Status Green (#00E676): System nominal, LEO-ready verification indicators."
    )

    h2 = doc.add_heading(level=2)
    r_h2 = h2.add_run("5.2 Interactive Sensor Inspection Suite (script.js)")
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(30, 41, 59)
    r_h2.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A major UX innovation was replacing boring bulleted lists with a dynamic multi-spectral sensor inspection widget. "
        "Visitors can toggle between three discrete operational modes:\n"
        "1. Optical Camera (4K CV Stream): Displays real-time visual tracking of wire harnesses with 60 FPS OpenCV bounding boxes.\n"
        "2. MLX90614 Infrared Radiometer: Switches the view to contactless thermal imaging resolving 0.02°C heat spikes at 74.2°C.\n"
        "3. Cold-Gas Interceptor Plume: Shows the high-speed 12V solenoid nozzle firing pressurized nitrogen at 150 PSI.\n"
        "The widget dynamically updates the high-res image, HUD targeting reticle, telemetry badges, and technical specifications via JavaScript."
    )

    h2 = doc.add_heading(level=2)
    r_h2 = h2.add_run("5.3 Cockpit HUD Telemetry Engine")
    r_h2.font.size = Pt(12)
    r_h2.font.color.rgb = RGBColor(30, 41, 59)
    r_h2.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To provide mission-control authenticity, I coded an asynchronous JavaScript telemetry loop that realistically simulates "
        "live sensor downlink telemetry. The HUD displays real-time fluctuating readings:\n"
        "• Azimuth: 123.8° – 124.2° (Target Lock tracking)\n"
        "• Elevation: 37.9° – 38.3° (Gimbal compensation)\n"
        "• Core Temperature: 73.8°C – 74.5°C (Simulated electrical thermal spike)\n"
        "• Solenoid Pressure: 148.5 – 150.2 PSI (Nitrogen reservoir state)\n"
        "• Downlink Latency: 114ms – 121ms (Validating sub-120ms execution requirement)"
    )

    # ==========================================
    # SECTION 6: TECHNICAL DOCUMENTATION PORTAL
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("6. Separation of Concerns: The NASA SP Documentation Portal (doc.html)")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "A critical design decision in resolving the clutter problem was establishing a dedicated second page: the "
        "NASA SP Technical Memorandum (doc.html). This kept the landing page engaging and photo-centric, while providing "
        "professors, reviewers, and engineering evaluators with exhaustive technical depth in standard NASA documentation format."
    )

    p = doc.add_paragraph()
    p.add_run(
        "The technical documentation portal includes:\n"
        "1. Document Control Block: NASA SP-2026-09 classification and author attributions (Paakhi & Rahul).\n"
        "2. Interactive SVG System Architecture Diagram: Visualizing the complete data flow from Optical Camera → OpenCV Perception → MLX90614 IR Sensor → Flight Computer → Pan-Tilt Gimbal & 12V Solenoid Valve.\n"
        "3. Complete Hardware Specification & Pinout Table: Mapping I2C lines (SDA/SCL on GPIO 2/3), PWM servo lines (GPIO 18/19), and 12V MOSFET gate trigger (GPIO 23).\n"
        "4. Software State Machine: Documenting States 0 through 4 (Surveillance → Candidate Anomaly → Thermal Slew → Pulse Discharge → Post-Action Cooling Verification).\n"
        "5. Bill of Materials (BOM) Table: Itemized hardware procurement log totaling ₹5,900 INR with vendor part numbers for Engineering Clinics reimbursement.\n"
        "6. One-Click Print / PDF Engine: Embedded print styling allowing evaluators to export a formatted paper report instantly."
    )

    # ==========================================
    # SECTION 7: DEPLOYMENT & VERIFICATION
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("7. Git Version Control & Live GitHub Pages Deployment")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "To ensure professional delivery beyond a local machine environment ('localhost'), I managed the complete deployment pipeline:\n"
        "• Git Repository Setup: Initialized Git version control on branch 'main', configured a clean '.gitignore' for OS and editor artifacts, and committed structured modular code.\n"
        "• GitHub Remote Repository: Linked the project to https://github.com/utpalakshi/asfai.git and successfully pushed the master codebase.\n"
        "• Continuous Cloud Deployment: Enabled GitHub Pages with root branch deployment and automated HTTPS enforcement.\n"
        "• Universal Cross-Device Access: The resulting web application is live and accessible globally without local runtime dependencies."
    )

    # Summary Table of Links
    link_tbl = doc.add_table(rows=3, cols=2)
    link_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    link_tbl.autofit = False
    link_tbl.columns[0].width = Inches(2.2)
    link_tbl.columns[1].width = Inches(4.3)

    links_data = [
        ("Public Landing Page:", "https://utpalakshi.github.io/asfai/"),
        ("Technical Spec Portal:", "https://utpalakshi.github.io/asfai/doc.html"),
        ("Source Code Repository:", "https://github.com/utpalakshi/asfai")
    ]
    for i, (label, url) in enumerate(links_data):
        row = link_tbl.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(url)
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(16, 91, 216)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==========================================
    # SECTION 8: CONCLUSION & LEARNINGS
    # ==========================================
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("8. Conclusion & Key Engineering Learnings")
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(16, 91, 216)
    r_h1.font.bold = True

    p = doc.add_paragraph()
    p.add_run(
        "Working on the UI/UX architecture of ASAFI demonstrated that for complex aerospace systems, presentation and user interface "
        "design are as vital as physical hardware. By analyzing and adopting the NASA Horizon design language, our team successfully "
        "transformed a dense, hard-to-read academic prototype into an engaging, authoritative, and scientifically rigorous mission portal.\n\n"
        "Key takeaways from this design cycle include:\n"
        "• Visual Hierarchy Drives Comprehension: Leading with high-resolution imagery and bold statistics enables evaluators to grasp the mission's sub-second life-safety value within seconds.\n"
        "• Modular Separation of Concerns: Separating the public narrative (index.html) from formal engineering documentation (doc.html) achieves both broad engagement and technical rigor without compromising either.\n"
        "• Micro-Interactions Reinforce Realism: Simulated telemetry fluctuations, live reticles, and multi-spectral image switchers immerse the user in the operational reality of spacecraft avionics control.\n"
        "• Professional Delivery: Moving beyond localhost to live GitHub Pages hosting empowers global demonstration, seamless collaboration, and high-impact academic defense."
    )

    # Save
    out_path = r"C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\ASAFI_UI_UX_Project_Report.docx"
    doc.save(out_path)
    print(f"Report saved successfully to: {out_path}")

if __name__ == "__main__":
    build_report()
