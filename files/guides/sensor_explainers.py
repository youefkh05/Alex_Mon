"""'How does it work' explainer slides for the turbidity (SEN0710) and dissolved-oxygen (SEN0681) sensors.

Imported by build_sensor_slides.py; each entry is inserted just before that sensor's own slide.
"""


def _circ(x, y, r, fill="#d9c9a3"):
    return f'<circle cx="{x}" cy="{y}" r="{r}" fill="{fill}" stroke="#8a7a58" stroke-width=".8"/>'


# ------------------------------------------------------------------ turbidity diagram
_clear = "".join(_circ(x, y, 3.2) for x, y in [(180, 150), (262, 166), (350, 170)])
_cloudy_pts = [(150, 128), (172, 176), (196, 142), (214, 190), (236, 158), (248, 132), (262, 176), (274, 150), (290, 188),
               (306, 138), (322, 166), (338, 196), (356, 146), (372, 176), (250, 200), (226, 118), (300, 120), (200, 205)]
_cloudy = "".join(f'<g class="tc-drift" style="animation-delay:-{i * .37:.2f}s">{_circ(x, y, 3.6 if i % 3 else 4.6)}</g>'
                  for i, (x, y) in enumerate(_cloudy_pts))
_scat_clear = '<path d="M262 160 L260 98" stroke="#fff3b0" stroke-width="1.6" opacity=".55"/>'
_scat_cloudy = "".join(
    f'<path d="M{x} {y} L260 98" stroke="#fff3b0" stroke-width="1.8" class="tc-scat" style="animation-delay:-{i * .21:.2f}s"/>'
    for i, (x, y) in enumerate([(236, 158), (248, 150), (262, 162), (274, 152), (290, 160), (256, 172), (270, 168)]))

TURB_SVG = f"""<svg viewBox="0 0 520 270" role="img" aria-label="Turbidity sensor: an infrared beam crosses the water; particles scatter some light sideways to a detector at 90 degrees. Clear water scatters little (low NTU); cloudy water scatters a lot (high NTU).">
<defs><linearGradient id="tcW" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a9cc0" stop-opacity=".45"/><stop offset="1" stop-color="#05324f" stop-opacity=".85"/></linearGradient></defs>
<rect x="40" y="52" width="440" height="190" rx="16" fill="url(#tcW)" stroke="rgba(160,235,255,.45)" stroke-width="2"/>
<rect x="44" y="150" width="40" height="22" rx="4" fill="#cfd8dc" stroke="#fff"/><text x="64" y="144" text-anchor="middle" font-size="9" fill="#ffd27a" font-weight="700">IR LED</text>
<rect x="238" y="62" width="44" height="26" rx="4" fill="#cfd8dc" stroke="#fff"/><text x="260" y="56" text-anchor="middle" font-size="9" fill="#8cf8f5" font-weight="700">DETECTOR (90°)</text>
<path d="M84 161 H430" stroke="#ffd27a" stroke-width="10" opacity=".14"/>
<path d="M84 161 H430" stroke="#ffd27a" stroke-width="3" stroke-dasharray="10 8" class="tc-beam"/>
<g class="tc-clear">{_clear}{_scat_clear}</g>
<g class="tc-cloudy">{_cloudy}{_scat_cloudy}</g>
<g class="tc-clear"><rect x="100" y="206" width="320" height="26" rx="13" fill="#04263f" stroke="#a9d8e4"/><text x="260" y="223" text-anchor="middle" font-size="11.5" font-weight="700" fill="#eafcff">Clear water · little light scatters → low NTU</text></g>
<g class="tc-cloudy"><rect x="100" y="206" width="320" height="26" rx="13" fill="#04263f" stroke="#ffd27a"/><text x="260" y="223" text-anchor="middle" font-size="11.5" font-weight="700" fill="#ffe3a3">Cloudy water · much light scatters → high NTU</text></g>
<text x="440" y="108" text-anchor="end" font-size="9" fill="#a9d8e4" letter-spacing="1">TURBIDITY (NTU)</text>
<rect x="340" y="114" width="100" height="10" rx="5" fill="rgba(255,255,255,.15)"/>
<rect x="340" y="114" width="100" height="10" rx="5" fill="#ffd27a" class="tc-meter" style="transform-box:fill-box;transform-origin:left"/>
</svg>"""


# ------------------------------------------------------------------ dissolved-oxygen diagram
def _o2(x, y, d):
    return (f'<g transform="translate({x} {y})"><g class="do-drift" style="animation-delay:-{d:.2f}s">'
            '<circle r="5" cx="-5" fill="#2de2e6" fill-opacity=".9" stroke="#fff" stroke-opacity=".6"/>'
            '<circle r="5" cx="5" fill="#2de2e6" fill-opacity=".9" stroke="#fff" stroke-opacity=".6"/></g></g>')


_o2_low = "".join(_o2(x, y, i * .8) for i, (x, y) in enumerate([(120, 110), (400, 120), (130, 196)]))
_o2_high = "".join(_o2(x, y, i * .5) for i, (x, y) in enumerate(
    [(120, 110), (400, 120), (130, 196), (190, 80), (330, 84), (410, 190), (200, 200), (350, 196), (150, 150), (380, 160), (300, 140), (220, 130)]))

DO_SVG = f"""<svg viewBox="0 0 520 270" role="img" aria-label="Fluorescence oxygen sensor: blue light makes a dye film glow red; oxygen molecules touching the film quench the glow. Less oxygen means a brighter, longer glow; more oxygen means a dimmer glow.">
<defs><linearGradient id="doW" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a9cc0" stop-opacity=".45"/><stop offset="1" stop-color="#05324f" stop-opacity=".85"/></linearGradient>
<radialGradient id="doG"><stop offset="0" stop-color="#ff5a4a" stop-opacity=".95"/><stop offset="1" stop-color="#ff5a4a" stop-opacity="0"/></radialGradient></defs>
<rect x="40" y="30" width="440" height="212" rx="16" fill="url(#doW)" stroke="rgba(160,235,255,.45)" stroke-width="2"/>
<rect x="238" y="170" width="44" height="30" rx="6" fill="#cfd8dc" stroke="#fff"/><rect x="236" y="152" width="48" height="20" rx="6" fill="#8d5b9e" stroke="#fff"/>
<text x="260" y="148" text-anchor="middle" font-size="9" fill="#ffd27a" font-weight="700">FLUORESCENT FILM</text>
<path d="M260 150 V96" stroke="#4aa3ff" stroke-width="4" stroke-dasharray="6 6" class="tc-beam"/><text x="274" y="110" font-size="9" fill="#9cc8ff" font-weight="700">blue light</text>
<g class="tc-clear"><circle cx="260" cy="130" r="58" fill="url(#doG)" class="do-glow"/>{_o2_low}</g>
<g class="tc-cloudy"><circle cx="260" cy="130" r="26" fill="url(#doG)" opacity=".55"/>{_o2_high}</g>
<g class="tc-clear"><rect x="90" y="206" width="340" height="26" rx="13" fill="#04263f" stroke="#a9d8e4"/><text x="260" y="223" text-anchor="middle" font-size="11.5" font-weight="700" fill="#eafcff">Little O₂ · film glows bright and long → low reading</text></g>
<g class="tc-cloudy"><rect x="90" y="206" width="340" height="26" rx="13" fill="#04263f" stroke="#ffd27a"/><text x="260" y="223" text-anchor="middle" font-size="11.5" font-weight="700" fill="#ffe3a3">Lots of O₂ · glow quenched, dims fast → high reading</text></g>
<g font-size="10" fill="#bfeaf5"><circle cx="64" cy="54" r="5" fill="#2de2e6"/><circle cx="74" cy="54" r="5" fill="#2de2e6"/><text x="86" y="58">O₂ molecule</text></g>
</svg>"""


def _explainer(kicker, title, sub, svg, fx, steps_title, steps, right_title, right_text, flow, check_title, checks, foot):
    st = "".join(f"<li><b>{i}</b><span>{s}</span></li>" for i, s in enumerate(steps, 1))
    node = lambda a, b: f'<div class="formula-node">{a}<br><span style="color:var(--muted);font-weight:500">{b}</span></div>'
    fl = '<div class="formula-arrow">+</div>'.join(node(a, b) for a, b in flow[:-1]) + '<div class="formula-arrow">→</div>' + node(*flow[-1])
    cols = "1fr auto 1fr auto 1fr" if len(flow) == 3 else "1fr auto 1fr"
    ck = "".join(f"<li>{c}</li>" for c in checks)
    return f"""<section class="slide salinity-slide gs-explain">
  <div class="water-ribbons" aria-hidden="true"><span></span><span></span><span></span><span></span></div>
  <div class="slide-inner">
    <span class="kicker rise" style="--i:0">{kicker}</span>
    <h2 class="rise" style="--i:1">{title}</h2>
    <p class="sub rise" style="--i:2">{sub}</p>
    <div class="salinity-layout">
      <article class="salinity-panel panel rise" style="--i:3" data-fx="{fx}">
        <div class="cell-demo">{svg}</div>
        <h3>{steps_title}</h3>
        <ol class="cell-steps">{st}</ol>
      </article>
      <article class="salinity-panel panel rise" style="--i:4">
        <h3>{right_title}</h3>
        <p>{right_text}</p>
        <div class="formula-flow" style="grid-template-columns:{cols}">{fl}</div>
        <h3 style="margin-top:.9rem">{check_title}</h3>
        <ul>{ck}</ul>
      </article>
    </div>
    <p class="slide-footnote rise" style="--i:5">{foot}</p>
  </div>
</section>"""


EXPL_TURB = _explainer(
    "Sensor technology · Sedimentation (turbidity)",
    'How does a <span class="hl">turbidity sensor</span> measure sediment?',
    "The probe shines an invisible infrared beam through the water. Sediment particles scatter some of that light sideways, and a detector at 90° counts it: the cloudier the water, the more light it sees.",
    TURB_SVG, "sediment", "How cloudiness is measured",
    ["An infrared LED sends a narrow beam into the water.",
     "Clear water lets the beam pass; the detector (at 90°, out of the beam) sees almost nothing.",
     "Suspended sediment scatters light sideways into the detector. More particles, more scattered light.",
     "The sensor converts that light into turbidity in NTU (0–1000 NTU for the SEN0710)."],
    "From light to recorded data",
    "The probe turns the scattered-light signal into an NTU value and sends it, with a temperature reading, over RS485 Modbus to the controller in the buoy.",
    [("Scattered light", "90° detector"), ("Turbidity", "NTU value")], "What to check for ALEX-MON",
    ["Clean the optical window: biofouling reads as false turbidity.",
     "Mount clear of the seabed, walls and bubbles, which scatter light too.",
     "NTU is a cloudiness index, not mass of sediment; compare it with water samples to convert.",
     "Confirm the plastic body is rated for long seawater immersion."],
    '<b style="color:var(--gold)">Note:</b> turbidity is read as relative change. A storm or dredging pulse shows as a jump in NTU; the sensor does not say what the sediment is made of.')

EXPL_DO = _explainer(
    "Sensor technology · Pollution (dissolved oxygen)",
    'How does a <span class="hl">dissolved-oxygen sensor</span> work?',
    "A fluorescence sensor lights a special dye film with blue light and watches the red glow it gives back. Oxygen touching the film switches the glow off faster, so the glow tells us how much oxygen is dissolved.",
    DO_SVG, "bubbles", "How oxygen is measured",
    ["Blue light hits a dye film on the sensor tip. The dye glows red.",
     "Dissolved oxygen molecules that touch the film take the energy, so the glow fades faster (quenching).",
     "The sensor measures how bright or long-lasting the glow is. More oxygen, shorter glow.",
     "It combines that with temperature and salinity to report dissolved oxygen in mg/L and % saturation."],
    "From glow to recorded data",
    "The probe works out oxygen from the glow, corrected for temperature and the salinity you set, then sends mg/L, % saturation and temperature over RS485 Modbus.",
    [("Glow", "brightness / lifetime"), ("Temperature", "Compensation"), ("Dissolved O₂", "mg/L · % sat.")], "What to check for ALEX-MON",
    ["Set salinity to the site value (about 36–40 ‰), not the default 30‰.",
     "Replace the fluorescent film about once a year and clean off biofouling.",
     "Keep bubbles off the film; response time is up to 60 s.",
     "A sustained drop below the usual 3–15 mg/L site range points to organic pollution or sewage inflow."],
    '<b style="color:var(--gold)">Note:</b> the film does not use up oxygen and needs no stirring, so it suits a still seabed; a slow reading is normal.')

EXPLAINERS = {"SEN0710": EXPL_TURB, "SEN0681": EXPL_DO}

CSS = """
.gs-explain .cell-demo svg{width:100%;height:auto;display:block}
.gs-explain .tc-beam{animation:tcDash 1s linear infinite}
@keyframes tcDash{to{stroke-dashoffset:-18}}
.gs-explain .tc-clear{animation:tcA 10s ease-in-out infinite}
.gs-explain .tc-cloudy{animation:tcB 10s ease-in-out infinite}
@keyframes tcA{0%,42%{opacity:1}52%,92%{opacity:0}100%{opacity:1}}
@keyframes tcB{0%,42%{opacity:0}52%,92%{opacity:1}100%{opacity:0}}
.gs-explain .tc-drift,.gs-explain .do-drift{animation:tcDrift 3.4s ease-in-out infinite alternate}
@keyframes tcDrift{from{transform:translate(-5px,-3px)}to{transform:translate(6px,4px)}}
.gs-explain .tc-scat{animation:tcFlick 1.3s ease-in-out infinite;opacity:.8}
@keyframes tcFlick{0%,100%{opacity:.25}50%{opacity:.95}}
.gs-explain .tc-meter{animation:tcMeter 10s ease-in-out infinite;transform:scaleX(.08)}
@keyframes tcMeter{0%,42%{transform:scaleX(.08)}52%,92%{transform:scaleX(.72)}100%{transform:scaleX(.08)}}
.gs-explain .do-glow{animation:doGlow 2.2s ease-in-out infinite alternate;transform-box:fill-box;transform-origin:center}
@keyframes doGlow{from{transform:scale(.9);opacity:.8}to{transform:scale(1.08);opacity:1}}
"""
