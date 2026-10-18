/**
 * ASAFI // NASA-Inspired Web Portal Scripts
 * Interactive Lightbox, Telemetry Engine, and Sensor Inspection
 */

document.addEventListener('DOMContentLoaded', () => {
  // 1. Mobile Menu Toggle
  const mobileMenuBtn = document.getElementById('mobileMenuBtn');
  const mobileNav = document.getElementById('mobileNav');

  if (mobileMenuBtn && mobileNav) {
    mobileMenuBtn.addEventListener('click', () => {
      const isHidden = mobileNav.classList.contains('hidden');
      mobileNav.classList.toggle('hidden');
      mobileMenuBtn.setAttribute('aria-expanded', isHidden ? 'true' : 'false');
    });
  }

  // 2. Lightbox Modal
  const lightboxModal = document.getElementById('lightboxModal');
  const lightboxImg = document.getElementById('lightboxImg');
  const lightboxTitle = document.getElementById('lightboxTitle');
  const lightboxDesc = document.getElementById('lightboxDesc');
  const lightboxTag = document.getElementById('lightboxTag');
  const closeLightboxBtn = document.getElementById('closeLightbox');

  function openLightbox(src, title, desc, tag) {
    if (!lightboxModal) return;
    lightboxImg.src = src;
    lightboxImg.alt = title;
    lightboxTitle.textContent = title;
    lightboxDesc.textContent = desc;
    lightboxTag.textContent = tag || 'ASAFI MULTIMEDIA';
    
    lightboxModal.classList.remove('hidden');
    requestAnimationFrame(() => {
      lightboxModal.classList.add('open');
    });
    document.body.style.overflow = 'hidden';
  }

  function closeLightbox() {
    if (!lightboxModal) return;
    lightboxModal.classList.remove('open');
    setTimeout(() => {
      lightboxModal.classList.add('hidden');
      document.body.style.overflow = '';
    }, 200);
  }

  if (closeLightboxBtn) {
    closeLightboxBtn.addEventListener('click', closeLightbox);
  }

  if (lightboxModal) {
    lightboxModal.addEventListener('click', (e) => {
      if (e.target === lightboxModal || e.target.classList.contains('lightbox-backdrop')) {
        closeLightbox();
      }
    });
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && lightboxModal && lightboxModal.classList.contains('open')) {
      closeLightbox();
    }
  });

  document.querySelectorAll('.open-lightbox').forEach((el) => {
    el.addEventListener('click', (e) => {
      e.preventDefault();
      const src = el.getAttribute('data-img');
      const title = el.getAttribute('data-title');
      const desc = el.getAttribute('data-desc');
      const tag = el.getAttribute('data-tag');
      openLightbox(src, title, desc, tag);
    });
  });

  // 3. Interactive Sensor Array Switcher (Optical vs IR vs Gas Pulse)
  const sensorTabs = document.querySelectorAll('.sensor-tab');
  const sensorDisplayImg = document.getElementById('sensorDisplayImg');
  const sensorDisplayBadge = document.getElementById('sensorDisplayBadge');
  const sensorDisplayTitle = document.getElementById('sensorDisplayTitle');
  const sensorDisplayDesc = document.getElementById('sensorDisplayDesc');
  const sensorDisplaySpec = document.getElementById('sensorDisplaySpec');

  const sensorData = {
    optical: {
      src: "https://images.unsplash.com/photo-1518770660439-4636190af475?q=80&w=1600&auto=format&fit=crop",
      badge: "OPTICAL SENSING // 4K CV STREAM",
      title: "High-Speed Avionics Visual Tracking",
      desc: "Edge-computed OpenCV detection runs at 60 FPS, isolating localized spark discharges and smoke particulate signatures.",
      spec: "Latency: <16.6ms | Resolution: 1080p60 | Algorithm: HSV Contour & Frame Diff"
    },
    thermal: {
      src: "https://images.unsplash.com/photo-1509228468518-180dd4864904?q=80&w=1600&auto=format&fit=crop",
      badge: "INFRARED RADIOMETRY // MLX90614",
      title: "Contactless Dual-Zone Pyrometry",
      desc: "Melexis MLX90614 medical-grade thermopile validates localized thermal gradient without electrical contact.",
      spec: "Range: -70°C to +380°C | Accuracy: ±0.5°C | Thermal Resolution: 0.02°C"
    },
    plume: {
      src: "./assets/coldspray.jpg",
      badge: "COLD-GAS INTERCEPT // 12V SOLENOID",
      title: "Sub-Second Micro-Pulse Interception",
      desc: "Clean-agent nitrogen burst dispenses directly onto anomalous wiring, suppressing combustion before runaway.",
      spec: "Valve Speed: <12ms open | Pressure: 150 PSI | Agent: Non-Conductive N2"
    }
  };

  sensorTabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      const mode = tab.getAttribute('data-sensor');
      if (!sensorData[mode]) return;

      sensorTabs.forEach(t => {
        t.classList.remove('active', 'border-blue-500', 'bg-blue-600/20', 'text-white');
        t.classList.add('border-white/10', 'bg-white/5', 'text-slate-400');
      });

      tab.classList.add('active', 'border-blue-500', 'bg-blue-600/20', 'text-white');
      tab.classList.remove('border-white/10', 'bg-white/5', 'text-slate-400');

      if (sensorDisplayImg) {
        sensorDisplayImg.style.opacity = '0.4';
        setTimeout(() => {
          sensorDisplayImg.src = sensorData[mode].src;
          sensorDisplayBadge.textContent = sensorData[mode].badge;
          sensorDisplayTitle.textContent = sensorData[mode].title;
          sensorDisplayDesc.textContent = sensorData[mode].desc;
          sensorDisplaySpec.textContent = sensorData[mode].spec;
          sensorDisplayImg.style.opacity = '1';
        }, 150);
      }
    });
  });

  // 4. Live HUD Telemetry Fluctuation Simulator
  const liveTemp = document.getElementById('liveTemp');
  const liveAzimuth = document.getElementById('liveAzimuth');
  const liveElevation = document.getElementById('liveElevation');
  const livePressure = document.getElementById('livePressure');
  const liveLatency = document.getElementById('liveLatency');

  if (liveTemp) {
    setInterval(() => {
      // Fluctuations around nominal baseline
      const tempVal = (74.0 + (Math.random() * 0.5 - 0.25)).toFixed(1);
      const azVal = (124.0 + (Math.random() * 0.4 - 0.2)).toFixed(1);
      const elVal = (38.0 + (Math.random() * 0.4 - 0.2)).toFixed(1);
      const pressVal = (149.5 + (Math.random() * 0.8 - 0.4)).toFixed(1);
      const latVal = Math.floor(116 + Math.random() * 6);

      liveTemp.textContent = `${tempVal}°C`;
      liveAzimuth.textContent = `${azVal}°`;
      liveElevation.textContent = `${elVal}°`;
      livePressure.textContent = `${pressVal} PSI`;
      if (liveLatency) liveLatency.textContent = `${latVal} ms`;
    }, 1800);
  }

  // 5. Media Gallery Filtering
  const filterBtns = document.querySelectorAll('.gallery-filter-btn');
  const galleryItems = document.querySelectorAll('.gallery-item');

  filterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const filter = btn.getAttribute('data-filter');

      filterBtns.forEach(b => {
        b.classList.remove('bg-blue-600', 'text-white');
        b.classList.add('bg-white/5', 'text-slate-400');
      });
      btn.classList.add('bg-blue-600', 'text-white');
      btn.classList.remove('bg-white/5', 'text-slate-400');

      galleryItems.forEach(item => {
        const cat = item.getAttribute('data-category');
        if (filter === 'all' || cat === filter) {
          item.classList.remove('hidden');
        } else {
          item.classList.add('hidden');
        }
      });
    });
  });

  // 6. Mission Updates Newsletter Submission
  const newsletterForm = document.getElementById('newsletterForm');
  const newsletterMsg = document.getElementById('newsletterMsg');

  if (newsletterForm && newsletterMsg) {
    newsletterForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const emailInput = newsletterForm.querySelector('input[type="email"]');
      if (emailInput && emailInput.value) {
        newsletterMsg.classList.remove('hidden');
        newsletterForm.reset();
        setTimeout(() => {
          newsletterMsg.classList.add('hidden');
        }, 5000);
      }
    });
  }
});
