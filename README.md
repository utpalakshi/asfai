# ASAFI // Autonomous Spacecraft Avionics Fire Interceptor
### NASA.gov-Inspired Mission Web Portal & Technical Documentation

Designed as a modern, image-forward aerospace project portal inspired by **[NASA.gov](https://www.nasa.gov/)** and the **NASA Web Design System (Horizon)**.

---

## 🚀 Key Features

1. **Picture-Heavy, Text-Lean Editorial Design**:
   - High-resolution aerospace photography, launch plume captures, thermal imaging, and avionics bays.
   - Minimal, ultra-precise technical copy replacing long paragraphs.
   - Distinct NASA category badges (`MISSION BRIEFING`, `PROPULSION // INTERCEPT`, `IR SPECTRUM`, `LEO READY`).

2. **Full Authentic NASA Footer**:
   - Meatball insignia with mission slogan.
   - 4-column structured directory (Missions & Flights, Technology & Payloads, Multimedia Archive, Technical Documents).
   - Mission Alert / Newsletter subscription box with real-time feedback.
   - Standard administrative/federal disclosures (Privacy Policy, FOIA, No FEAR Act, Section 508 Accessibility, Inspector General).

3. **NASA SP Technical Documentation Portal (`doc.html`)**:
   - Formatted in official NASA Special Publication (SP) Technical Memorandum style.
   - Interactive SVG System Architecture Diagram.
   - Subsystem breakdowns (Optical CV, MLX90614, 2-Axis Gimbal, 12V Solenoid).
   - Software State Machine (States 0 through 4).
   - Hardware Pinout & Wiring Table.
   - Bill of Materials (BOM) & ECS Reimbursement Tracker.
   - Engineering Clinics team roster (Utpalakshi & Rahul).
   - One-click "Print / Save PDF" formatting.

4. **Interactive Mission Features**:
   - **Interactive Sensor Array Viewer**: Toggle between Optical 4K CV, MLX90614 Infrared, and Cold-Gas Nozzle.
   - **High-Resolution Media Lightbox**: Click any visual card to expand into a NASA-styled lightbox with technical metadata.
   - **Simulated HUD Cockpit Telemetry**: Dynamic live fluctuations of Azimuth, Elevation, Core Temp, and Pressure.
   - **Gallery Category Filtering**: Quickly filter between Sensors, Propulsion, and Avionics.

---

## 📂 Project Structure

```
asafi-space-portal/
├── index.html       # NASA-style visual landing page with full footer
├── doc.html         # NASA SP Technical Memorandum & Engineering Specs
├── styles.css       # Custom aerospace styles, HUD reticles, and NASA palette
├── script.js        # Lightbox, live telemetry simulator, sensor switcher
└── README.md        # Project guide and documentation
```

---

## 💻 How to View Locally

Simply double-click `index.html` or open it with any web browser:
```powershell
# Open with default browser (PowerShell)
Start-Process "C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\index.html"
```

Or view the technical documentation directly:
```powershell
Start-Process "C:\Users\Utpalakshi upadhyaya\.gemini\antigravity\scratch\asafi-space-portal\doc.html"
```
