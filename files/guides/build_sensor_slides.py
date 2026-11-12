"""Generate one deck slide per selected sensor (guide layout, no price) and splice them into index.html.

Data mirrors build_guide.py (Arabic guide); this file holds the English deck copy.
Run from anywhere:  python build_sensor_slides.py
Re-running replaces the block between the SENSOR-SLIDES markers, so it is safe to repeat.
"""
import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INDEX = ROOT / "index.html"
IMG = "files/guides/img/"
e = html.escape

DIY_SVG = """<svg viewBox="0 0 520 220" role="img" aria-label="Tilt current meter: a float tethered to the seabed leans with the current; the lean angle gives speed.">
<defs><linearGradient id="diyw" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#d4eef5"/><stop offset="1" stop-color="#a9dbe8"/></linearGradient></defs>
<rect width="520" height="220" fill="url(#diyw)"/><rect y="190" width="520" height="30" fill="#d9c9a3"/><rect x="225" y="176" width="70" height="16" rx="3" fill="#9aa4ad"/>
<path d="M260 176 L335 78" stroke="#4a5b6b" stroke-width="3"/><line x1="260" y1="176" x2="260" y2="30" stroke="#7f93ad" stroke-width="1.5" stroke-dasharray="4 4"/>
<g transform="rotate(-37 335 78)"><rect x="315" y="30" width="40" height="96" rx="18" fill="#0b2a4a"/><rect x="322" y="56" width="26" height="28" rx="4" fill="#14b8a6"/><text x="335" y="75" text-anchor="middle" font-size="9" fill="#fff">IMU</text></g>
<path d="M260 120 A56 56 0 0 1 290 128" stroke="#c0392b" stroke-width="2" fill="none" stroke-dasharray="3 3"/><text x="296" y="122" font-size="15" fill="#c0392b" font-weight="700">θ</text>
<g fill="#0b6b7a"><path d="M30 80 h110 v-9 l24 18 -24 18 v-9 h-110z"/><path d="M30 140 h80 v-8 l20 15 -20 15 v-8 h-80z"/></g>
<text x="85" y="62" text-anchor="middle" font-size="14" fill="#0b2a4a" font-weight="700">Current</text>
<text x="420" y="36" text-anchor="middle" font-size="12" fill="#0b2a4a">Float + 2" housing</text><text x="420" y="54" text-anchor="middle" font-size="12" fill="#0b2a4a">tilt angle θ → speed</text></svg>"""

SENSORS = [
    dict(group="Salinity & conductivity", brand="SensorsX", model="SX-CCFNA", sub="4-electrode conductivity / salinity probe",
         tags=["Salinity", "4 electrodes"], img="sx-ccfna.webp",
         desc="A digital probe that measures conductivity, salinity, TDS and temperature together with four electrodes, and sends readings over RS-485 (Modbus).",
         tiles=[("0.001 – 100 mS/cm", "Conductivity (high-range option)"), ("0 – 80 ppt", "Salinity range"), ("±1 %", "Accuracy (or 0.01 mS/cm)"),
                ("RS-485", "Modbus interface"), ("9 – 24 V", "Supply · 0.2 W"), ("IP68", "Φ22 × 175.5 mm")],
         steps=["Order the high-range option: 0–100 mS/cm / 0–80 ppt (the 0–5 mS/cm option is not for seawater).",
                "Mount on the seabed frame and wire the RS-485 cable to the collector.",
                "Reads conductivity, salinity and temperature continuously over Modbus.",
                "Clean off biofouling regularly and recalibrate (one or two points)."],
         pros=["4 readings in one probe", "Ni + PEEK + titanium body suits seawater", "Low power: 0.2 W", "Factory calibrated"],
         cons=["High-range option must be requested", "Standard cable only 10 m", "Operating depth not in the datasheet", "Calibration method not described"],
         fit=("In range", "ok"), team=("Conditional", "cond"),
         links=[("Product page", "https://sensorsx.com/products/sx-ccfna-series")]),
    dict(group="Salinity & conductivity", brand="WizSensor", model="WS-OM-MUL-17A", sub="CTD probe: conductivity · salinity · temperature · depth",
         tags=["CTD", "Salinity"], img="ws-om-mul-17a.webp",
         desc="A CTD that measures conductivity, temperature and depth and calculates salinity, with RS-485 Modbus RTU output. The depth reading also tracks water level above the site.",
         tiles=[("0 – 100 mS/cm", "Conductivity (±1 %)"), ("0 – 80 ppt", "Salinity (calculated, ±1 ppt)"), ("±0.05 °C", "Temperature (−5 to 50 °C)"),
                ("0 – 200 m", "Depth (±0.1 % FS)"), ("RS-485", "Modbus RTU · 9600 8N1"), ("6 – 36 V", "Supply (≈ 20 mA at 12 V)")],
         steps=["Request a quote; confirm cable length (30 m standard) and pressure rating ≥ 10 m.",
                "Mount on the frame; never hang the unit from the signal cable.",
                "Sends temperature, conductivity, salinity and depth live over Modbus.",
                "Clean every 4–8 weeks and recalibrate (temperature and conductivity, 1–2 points)."],
         pros=["Best technical match for site salinity and temperature", "Also measures depth / water level", "Low power suits the solar buoy", "High temperature accuracy ±0.05 °C"],
         cons=["Price not published", "316L body: confirm long-term seawater immersion", "Datasheet sent on request only", "Salinity accuracy ±1 ppt"],
         fit=("In range", "ok"), team=("Suitable", "yes"),
         links=[("Product page", "https://wizsensor.com/pro/ws-om-mul-17a-ctd-sensor/")]),
    dict(group="Salinity & conductivity", brand="Dragino", model="DR-EC200", sub="4-electrode conductivity probe for seawater",
         tags=["Salinity", "4 electrodes"], img="dr-ec200-drawing.png", img_class="drawing",
         desc="A digital four-electrode probe for high-conductivity water such as seawater. It measures conductivity, salinity and temperature and sends them over RS-485.",
         tiles=[("1 – 200,000 µS/cm", "Conductivity (±1 % FS)"), ("0 – 70 PSU", "Salinity (0.1 resolution)"), ("±0.5 °C", "Temperature (−5 to 80 °C)"),
                ("RS-485", "Modbus-RTU · 9600"), ("7 – 30 V", "Supply (< 0.5 W)"), ("0.6 MPa", "Pressure · IP68 (not connector)")],
         steps=["Extend the cable (5 m standard) so it reaches the surface.",
                "Two-point Modbus calibration: 1,413 µS/cm then 111,310 µS/cm.",
                "Reads conductivity, salinity and temperature continuously.",
                "Recalibrate every 3 months and clean off biofouling."],
         pros=["Cheapest salinity sensor that fits the sites", "Designed for seawater", "Calibration steps well documented", "Flexible 7–30 V supply"],
         cons=["Range about 3× larger than needed (over-ranged)", "Cable only 5 m", "Needs calibration every 3 months", "Body dimensions not stated in text"],
         fit=("Over-ranged", "over"), team=("Suitable", "yes"),
         links=[("Product page", "https://www.antratek.com/dr-ec200-water-ec-sensor-for-wqs")]),
    dict(group="Water current", brand="DIY build", model="Tilt Current Meter", sub="Tilt-angle current meter (BNO085)",
         tags=["Current", "DIY"], svg=DIY_SVG,
         desc="A float tethered to the seabed leans with the current. A motion sensor (IMU) in a watertight housing measures the lean angle and direction, and calibration converts the angle to speed.",
         tiles=[("By calibration", "Speed range (needed 0–1 m/s)"), ("0 – 360°", "Direction (tilt-compensated compass)"), ("BNO085", "Motion sensor (Adafruit)"),
                ("18650", "Battery (≈ 1 month logging)"), ("60–70 ft", "Depth tested leak-free (≈ 20 m)"), ("SD card", "Logging (add RS-485 for live data)")],
         steps=["Fit the IMU, processor and battery inside a 2-inch housing.",
                "Tie the float by a line to a concrete base on the seabed.",
                "Calibrate by towing at known speeds (GPS) to convert angle to speed.",
                "Keep iron away from the compass; add RS-485 to stream live."],
         pros=["By far the cheapest option", "Parts easy to buy", "Adaptable to project needs", "No rotating parts"],
         cons=["Accuracy unknown before calibration", "No live data without modification", "Marine growth on the float changes readings", "Not yet tested on site"],
         fit=("Unknown", "unk"), team=("Conditional", "cond"),
         links=[("Project page", "https://hackaday.io/project/158083-cheap-underwater-tilt-current-sensor")]),
    dict(group="Water current", brand="Nortek", model="Signature 500", sub="5-beam ADCP profiler · 500 kHz",
         tags=["Current", "Profiler"], img="signature500.png",
         desc="A five-beam acoustic Doppler profiler that measures current in layers through the water column, with built-in temperature, pressure and compass; output over Ethernet or RS-232/422.",
         tiles=[("1 – 5 m/s", "Velocity range (adjustable)"), ("0.3 % ±0.3 cm/s", "Accuracy"), ("60 – 70 m", "Profile range (cells 0.5–4 m)"),
                ("300 m", "Max depth rating"), ("16 – 128 GB", "Memory"), ("12 – 48 V", "Supply (8 W at 1 Hz)")],
         steps=["Set velocity range to 1 m/s and cell size to 0.5 m.",
                "Mount on a seabed frame facing up.",
                "Records on its own, or connect a cable for live reading (Ethernet / RS-422).",
                "Process the data with Nortek software after recovery or live."],
         pros=["Very high accuracy", "Measures every layer of the water column", "Temperature, pressure and compass built in", "Up to 128 GB memory"],
         cons=["Built for deep water (60–70 m range)", "0.5 m blanking zone in front of the sensor", "Large and heavy (Ø228 mm, 6.4 kg)", "High price, on request"],
         note="Better fit for 2–8 m depth: Signature 1000 (range 25–30 m, cells from 0.2 m, HR mode from 0.1 m).",
         fit=("In range", "ok"), team=("Not reviewed", "unk"),
         links=[("Product page", "https://www.nortekgroup.com/products/signature-500")]),
    dict(group="Water current & waves", brand="Nortek", model="AWAC 600 kHz", sub="Current and wave profiler · 2nd generation",
         tags=["Current", "Waves"], img="awac600.png",
         desc="An acoustic Doppler profiler that measures current in layers and also measures waves with AST, with temperature and pressure sensors.",
         tiles=[("1 – 5 m/s", "Velocity range (adjustable)"), ("±1 % ±0.5 cm/s", "Accuracy"), ("50 m", "Profile range (cells 0.5–8 m)"),
                ("300 m", "Max depth rating"), ("up to 60 m", "Wave measurement"), ("12 – 48 V", "Supply (external batteries)")],
         steps=["Set velocity range to 1 m/s and choose the smallest cell (0.5 m).",
                "Mount on a seabed frame with the external battery case (540 / 1800 Wh).",
                "Records current and waves, or connect Ethernet / RS-422 for live data.",
                "Analyse current and wave data with Nortek software."],
         pros=["Current and waves in one instrument", "Rugged and proven in coastal waters", "Temperature and pressure built in", "Ethernet connection"],
         cons=["50 m range far above site depth", "0.5–8 m cells too coarse for 2–8 m depth", "Needs an external battery case", "High price, on request"],
         note="Better fit for 2–8 m depth: AWAC 1 MHz (range 30 m, cells from 0.25 m).",
         fit=("In range", "ok"), team=("Not reviewed", "unk"),
         links=[("Product page", "https://www.nortekgroup.com/products/awac2-600-khz")]),
    dict(group="Turbidity · Sedimentation", brand="DFRobot", model="SEN0710", sub="RS485 water-quality turbidity sensor",
         tags=["Turbidity", "Sediment"], img="sen0710.webp",
         desc="A digital sensor that measures how cloudy the water is (suspended particles and sediment) using 90° infrared light scattering, with a temperature reading and RS485 Modbus-RTU output.",
         tiles=[("0 – 1000 NTU", "Turbidity (0.1 NTU resolution)"), ("±5 % FS", "Accuracy at 25 °C"), ("0 – 60 °C", "Temperature (±0.5 °C)"),
                ("RS485", "Modbus-RTU · 4800 default"), ("10 – 30 V", "Supply · 0.2 W"), ("IP68", "0.6 MPa · 5 m cable")],
         steps=["Mount on the frame with the sensing head clear of the seabed and nearby walls.",
                "Wire: brown VCC, black GND, yellow 485-A, blue 485-B to the collector.",
                "Reads turbidity and temperature continuously (response ≤ 30 s).",
                "Clean the optical window of biofouling; recalibrate with 2 points if needed."],
         pros=["Measures suspended sediment (turbidity) directly", "Factory calibrated, 2-point recalibration", "Same RS485 interface as the other sensors", "Low power: 0.2 W"],
         cons=["Seawater compatibility not stated (plastic body)", "No maximum depth stated (0.6 MPa only)", "Operating temperature 0–40 °C", "No published site turbidity data to compare"],
         note="Ask DFRobot to confirm the plastic body for long-term seawater immersion before buying.",
         fit=("Unknown", "unk"), team=("Not reviewed", "unk"),
         links=[("Wiki", "https://wiki.dfrobot.com/sen0710/")]),
    dict(group="Dissolved oxygen · Pollution", brand="DFRobot", model="SEN0681", sub="RS485 fluorescence dissolved-oxygen sensor (seawater)",
         tags=["Dissolved O₂", "Fluorescence"], img="sen0681.webp",
         desc="A fluorescence dissolved-oxygen sensor (no stirring, consumes no oxygen), seawater version with default salinity 30‰, with a temperature reading and RS485 Modbus-RTU output.",
         tiles=[("0 – 20 mg/L", "Dissolved oxygen (0–200 % sat.)"), ("±3 % FS", "Accuracy (±0.5 °C at 25 °C)"), ("0.01 mg/L", "Resolution"),
                ("RS485", "Modbus-RTU"), ("10 – 30 V", "Supply · 0.2 W"), ("IP68", "0.6 MPa · 5 m cable")],
         steps=["Set salinity to the site value (about 36–40 ‰) instead of the default 30‰.",
                "Mount on the frame with the fluorescent film facing free water, no trapped bubbles.",
                "Reads oxygen (mg/L and % saturation) and temperature over Modbus-RTU (response ≤ 60 s).",
                "Replace the fluorescent film yearly; clean biofouling and calibrate per DFRobot guide."],
         pros=["Seawater version", "Fluorescence: no oxygen use, no stirring", "0–20 mg/L covers the 3–15 mg/L site range", "Same RS485 interface as the other sensors"],
         cons=["Fluorescent film lasts 1 year", "Accuracy ±3 % FS", "Plastic body: confirm long-term seawater use", "No maximum depth stated (0.6 MPa only)"],
         fit=("In range", "ok"), team=("Not reviewed", "unk"),
         links=[("Wiki", "https://wiki.dfrobot.com/SEN0681")]),
]

CSS = """
/* sensor slides (guide layout, no price) */
.gs-slide .slide-inner{max-width:1180px}
.gs-head{display:flex;align-items:flex-end;justify-content:space-between;gap:1rem;flex-wrap:wrap}
.gs-brand{font-size:.8rem;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:var(--gold)}
.gs-tags{display:flex;gap:.45rem;flex-wrap:wrap}
.gs-tag{font-size:.72rem;font-weight:700;padding:.25em .8em;border-radius:999px;border:1px solid var(--line2);color:var(--aqua);background:rgba(4,38,68,.45)}
.gs-grid{display:grid;grid-template-columns:minmax(0,.9fr) minmax(0,1.4fr);gap:1rem;margin-top:.9rem;align-items:start}
.gs-left,.gs-right{display:grid;gap:.8rem}
.gs-img{height:200px;border-radius:16px;background:#fff;display:flex;align-items:center;justify-content:center;overflow:hidden;border:1px solid var(--line2);position:relative}
.gs-img img{max-width:100%;max-height:190px;object-fit:contain}
.gs-img img.drawing{max-height:185px}
.gs-img svg{width:100%;height:100%}
.gs-desc{font-size:.88rem;line-height:1.55;color:var(--muted);border-left:3px solid var(--aqua);padding:.2rem .8rem}
.gs-fit{display:flex;gap:.5rem;align-items:center;flex-wrap:wrap;font-size:.76rem;color:var(--faint)}
.gs-pill{font-weight:800;padding:.2em .75em;border-radius:999px;font-size:.74rem}
.gs-pill.ok,.gs-pill.yes{background:rgba(55,211,153,.18);color:#7af0c0;border:1px solid rgba(55,211,153,.5)}
.gs-pill.over,.gs-pill.cond{background:rgba(255,210,122,.16);color:var(--gold);border:1px solid rgba(255,210,122,.5)}
.gs-pill.unk{background:rgba(160,235,255,.1);color:var(--muted);border:1px solid var(--line2)}
.gs-note{font-size:.76rem;color:#ffe3a3;background:rgba(255,210,122,.1);border:1px solid rgba(255,210,122,.3);border-radius:10px;padding:.45rem .7rem;line-height:1.4}
.gs-tiles{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.55rem}
.gs-tile{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:.55rem .6rem;text-align:center}
.gs-tile b{display:block;font-size:.98rem;color:var(--ink);font-family:Georgia,serif}
.gs-tile span{display:block;font-size:.68rem;color:var(--muted);margin-top:.15rem;line-height:1.3}
.gs-steps{display:grid;grid-template-columns:1fr 1fr;gap:.45rem .8rem;margin:0;padding:0;list-style:none}
.gs-steps li{display:flex;gap:.5rem;font-size:.76rem;line-height:1.4;color:var(--muted);align-items:flex-start}
.gs-steps li i{font-style:normal;flex:none;width:1.3rem;height:1.3rem;border-radius:50%;background:var(--teal);color:#012a3a;font-weight:800;font-size:.72rem;display:flex;align-items:center;justify-content:center}
.gs-h{font-size:.72rem;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:var(--aqua);margin-bottom:.4rem}
.gs-pc{display:grid;grid-template-columns:1fr 1fr;gap:.7rem}
.gs-box{border-radius:12px;padding:.6rem .8rem;border-top:3px solid}
.gs-box.pros{background:rgba(55,211,153,.1);border-color:var(--kelp)}
.gs-box.cons{background:rgba(255,122,92,.1);border-color:var(--coral)}
.gs-box h4{font-size:.78rem;margin-bottom:.3rem}
.gs-box.pros h4{color:#7af0c0}.gs-box.cons h4{color:#ff9d7a}
.gs-box ul{margin:0;padding-left:1.1rem;font-size:.74rem;line-height:1.45;color:var(--muted)}
.gs-link{font-size:.74rem;color:var(--aqua)}
@media(max-width:900px){.gs-grid,.gs-pc,.gs-steps{grid-template-columns:1fr}.gs-tiles{grid-template-columns:1fr 1fr}}
"""


def slide(n, total, s):
    img = s["svg"] if "svg" in s else f'<img class="{s.get("img_class", "")}" src="{IMG}{s["img"]}" alt="{e(s["model"])}" loading="lazy">'
    tags = "".join(f'<span class="gs-tag">{e(t)}</span>' for t in s["tags"])
    tiles = "".join(f'<div class="gs-tile"><b>{e(v)}</b><span>{e(k)}</span></div>' for v, k in s["tiles"])
    steps = "".join(f"<li><i>{k}</i><span>{e(t)}</span></li>" for k, t in enumerate(s["steps"], 1))
    pros = "".join(f"<li>{e(t)}</li>" for t in s["pros"])
    cons = "".join(f"<li>{e(t)}</li>" for t in s["cons"])
    links = " · ".join(f'<a class="gs-link" href="{e(u)}" target="_blank" rel="noopener">{e(k)} ↗</a>' for k, u in s["links"])
    note = f'<div class="gs-note">{e(s["note"])}</div>' if s.get("note") else ""
    fit, fc = s["fit"]
    team, tc = s["team"]
    return f"""<section class="slide gs-slide salinity-slide" data-sensor="{e(s['model'])}">
  <div class="water-ribbons" aria-hidden="true"><span></span><span></span><span></span><span></span></div>
  <div class="slide-inner">
    <span class="kicker rise" style="--i:0">Selected sensor {n} of {total} · {e(s['group'])}</span>
    <div class="gs-head"><h2 class="rise" style="--i:1"><span class="gs-brand">{e(s['brand'])}</span><br><span class="hl">{e(s['model'])}</span></h2><div class="gs-tags rise" style="--i:1">{tags}</div></div>
    <p class="sub rise" style="--i:2">{e(s['sub'])}</p>
    <div class="gs-grid">
      <div class="gs-left">
        <div class="gs-img panel rise" style="--i:3" data-fx="photo">{img}</div>
        <div class="gs-desc rise" style="--i:4">{e(s['desc'])}</div>
        <div class="gs-fit rise" style="--i:4"><span>Range fit:</span><span class="gs-pill {fc}">{e(fit)}</span><span>Team review:</span><span class="gs-pill {tc}">{e(team)}</span></div>
        {note}
      </div>
      <div class="gs-right">
        <div class="gs-tiles rise" style="--i:5" data-fx="specs">{tiles}</div>
        <div class="rise" style="--i:6" data-fx="steps"><div class="gs-h">How it works in the buoy</div><ol class="gs-steps">{steps}</ol></div>
        <div class="gs-pc rise" style="--i:7" data-fx="proscons"><div class="gs-box pros"><h4>Advantages</h4><ul>{pros}</ul></div><div class="gs-box cons"><h4>Limitations</h4><ul>{cons}</ul></div></div>
        <div class="rise" style="--i:8">{links}</div>
      </div>
    </div>
  </div>
</section>"""


def main():
    total = len(SENSORS)
    block = "<!-- SENSOR-SLIDES:START (generated by files/guides/build_sensor_slides.py) -->\n" + "\n\n".join(
        slide(i, total, s) for i, s in enumerate(SENSORS, 1)) + "\n<!-- SENSOR-SLIDES:END -->"
    css = "/* SENSOR-CSS:START */" + CSS + "/* SENSOR-CSS:END */"
    t = INDEX.read_text(encoding="utf-8")
    if "<!-- SENSOR-SLIDES:START" in t:
        t = re.sub(r"<!-- SENSOR-SLIDES:START.*?<!-- SENSOR-SLIDES:END -->", lambda m: block, t, flags=re.S)
        t = re.sub(r"/\* SENSOR-CSS:START \*/.*?/\* SENSOR-CSS:END \*/", lambda m: css, t, flags=re.S)
    else:
        anchor = '\n\n</div>\n\n<div class="dots" id="dots"></div>'
        assert anchor in t, "anchor not found"
        t = t.replace(anchor, "\n\n" + block + anchor, 1)
        t = t.replace("</style>", css + "\n</style>", 1)
    n = len(re.findall(r'<section class="slide', t))
    t = re.sub(r'(<span id="total">)\d+(</span>)', lambda m: m.group(1) + str(n) + m.group(2), t, count=1)
    INDEX.write_text(t, encoding="utf-8")
    print("slides:", n)


if __name__ == "__main__":
    main()
