"""Generate the wireless-communication slides (overview + cellular, LoRa, Wi-Fi/radio) and splice them into index.html.

Slides go right after the sensor slides (SENSOR-SLIDES:END). Re-running replaces the block between the COMM markers.
Run from anywhere:  python build_comm_slides.py
"""
import html
import re
from pathlib import Path

INDEX = Path(__file__).resolve().parents[2] / "index.html"
e = html.escape


# ---------------------------------------------------------------- diagram
def arcs(x, y, delay0=0.0):
    """Three radio arcs opening to the right of an antenna tip."""
    out = ""
    for k, r in enumerate((16, 30, 44)):
        out += (f'<path class="cm-arc" style="animation-delay:{delay0 + k * .35:.2f}s" d="M{x + 8} {y - r} A{r} {r} 0 0 1 {x + 8} {y + r}" '
                'fill="none" stroke="#8cf8f5" stroke-width="3" stroke-linecap="round"/>')
    return out


def buoy(bx, label):
    return f"""<g>
<line x1="{bx}" y1="176" x2="{bx}" y2="262" stroke="#a9d8e4" stroke-width="2.5" stroke-dasharray="5 4"/>
<rect x="{bx - 22}" y="256" width="18" height="14" rx="3" fill="#cfd8dc" stroke="#fff"/><rect x="{bx + 4}" y="256" width="18" height="14" rx="3" fill="#cfd8dc" stroke="#fff"/>
<text x="{bx}" y="292" text-anchor="middle" font-size="11" fill="#a9d8e4">sensors · cable</text>
<line x1="{bx}" y1="132" x2="{bx}" y2="82" stroke="#eafcff" stroke-width="3"/><circle cx="{bx}" cy="80" r="5" fill="#ffd27a"/>
<rect x="{bx - 17}" y="132" width="34" height="42" rx="7" fill="#0b2a4a" stroke="#2de2e6" stroke-width="2"/>
<rect x="{bx - 10}" y="140" width="20" height="12" rx="2" fill="#14b8a6"/>
<ellipse cx="{bx}" cy="176" rx="40" ry="10" fill="#ffd27a" stroke="#fff" stroke-opacity=".6"/>
<text x="{bx}" y="206" text-anchor="middle" font-size="12" font-weight="700" fill="#ffe3a3">{e(label)}</text></g>"""


def node(kind, nx, label):
    if kind == "tower":
        icon = (f'<path d="M{nx - 18} 176 L{nx} 72 L{nx + 18} 176 M{nx - 12} 140 H{nx + 12} M{nx - 7} 108 H{nx + 7} M{nx - 12} 140 L{nx + 7} 108 M{nx + 12} 140 L{nx - 7} 108" '
                'fill="none" stroke="#eafcff" stroke-width="2.5"/>'
                f'<circle cx="{nx}" cy="70" r="5" fill="#ff7a5c"/>')
        top = (nx, 70)
    elif kind == "gateway":
        icon = (f'<rect x="{nx - 24}" y="136" width="48" height="40" rx="6" fill="#0b2a4a" stroke="#2de2e6" stroke-width="2"/>'
                f'<circle cx="{nx - 10}" cy="156" r="3" fill="#37d399"/><circle cx="{nx}" cy="156" r="3" fill="#ffd27a"/><circle cx="{nx + 10}" cy="156" r="3" fill="#37d399"/>'
                f'<line x1="{nx}" y1="136" x2="{nx}" y2="92" stroke="#eafcff" stroke-width="3"/><circle cx="{nx}" cy="90" r="5" fill="#ff7a5c"/>')
        top = (nx, 90)
    else:  # dish
        icon = (f'<line x1="{nx}" y1="176" x2="{nx}" y2="100" stroke="#eafcff" stroke-width="4"/>'
                f'<path d="M{nx + 6} 78 Q{nx - 22} 100 {nx + 6} 122" fill="rgba(45,226,230,.2)" stroke="#2de2e6" stroke-width="3"/>'
                f'<circle cx="{nx - 12}" cy="100" r="4" fill="#ff7a5c"/>')
        top = (nx - 12, 100)
    land = f'<rect x="{nx - 70}" y="176" width="140" height="124" fill="#b79b6b" opacity=".75"/>'
    lab = f'<text x="{nx}" y="206" text-anchor="middle" font-size="12" font-weight="700" fill="#2b1d0e">{e(label)}</text>'
    return land + icon + lab, top


def cloud(cx, cy, label):
    return (f'<g><path d="M{cx - 44} {cy + 14} a18 18 0 0 1 8 -34 a26 26 0 0 1 48 -6 a20 20 0 0 1 8 40z" fill="rgba(45,226,230,.18)" stroke="#8cf8f5" stroke-width="2.5"/>'
            f'<text x="{cx}" y="{cy + 5}" text-anchor="middle" font-size="12" font-weight="700" fill="#eafcff">{e(label)}</text></g>')


def dashboard(dx, dy):
    return (f'<g><rect x="{dx}" y="{dy}" width="120" height="78" rx="8" fill="#062c46" stroke="#2de2e6" stroke-width="2.5"/>'
            f'<path d="M{dx + 12} {dy + 58} L{dx + 36} {dy + 40} L{dx + 56} {dy + 50} L{dx + 80} {dy + 24} L{dx + 108} {dy + 34}" fill="none" stroke="#ffd27a" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>'
            f'<rect x="{dx + 12}" y="{dy + 10}" width="30" height="8" rx="3" fill="#37d399"/><rect x="{dx + 48}" y="{dy + 10}" width="20" height="8" rx="3" fill="#8cf8f5"/>'
            f'<rect x="{dx + 40}" y="{dy + 78}" width="40" height="6" rx="3" fill="#8e99a3"/>'
            f'<text x="{dx + 60}" y="{dy + 106}" text-anchor="middle" font-size="12" font-weight="700" fill="#eafcff">Dashboard</text></g>')


def packet(path, dur, delay=0.0, color="#ffd27a"):
    return (f'<circle r="5" fill="{color}" stroke="#fff" stroke-width="1.5"><animateMotion dur="{dur}s" begin="{delay}s" repeatCount="indefinite" path="{path}"/>'
            f'<animate attributeName="opacity" values="0;1;1;0" dur="{dur}s" begin="{delay}s" repeatCount="indefinite"/></circle>')


def scene(buoy_label, kind, node_label, cloud_label="Internet", bx=120, nx=470):
    nd, (tx, ty) = node(kind, nx, node_label)
    cx, cy = 660, 80
    dx, dy = 770, 46
    p1 = f"M{bx} 80 L{tx} {ty}"
    p2 = f"M{tx} {ty} L{cx - 40} {cy}"
    p3 = f"M{cx + 40} {cy} L{dx} {dy + 40}"
    return f"""<svg viewBox="0 0 900 310" role="img" aria-label="The buoy sends sensor data through the air to shore, then over the internet to the dashboard.">
<defs><linearGradient id="cmSea" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#1a9cc0" stop-opacity=".55"/><stop offset="1" stop-color="#05324f" stop-opacity=".9"/></linearGradient></defs>
<rect x="0" y="176" width="{nx - 70}" height="134" fill="url(#cmSea)"/>
<path class="cm-wave" d="M0 178 q22 -8 44 0 t44 0 t44 0 t44 0 t44 0 t44 0 t44 0 t44 0 t44 0" stroke="#aafcff" stroke-opacity=".7" stroke-width="2.5" fill="none"/>
{nd}{buoy(bx, buoy_label)}{arcs(bx, 80)}
<path d="{p1}" stroke="rgba(255,210,122,.3)" stroke-width="2" stroke-dasharray="6 6" fill="none"/>
<path d="{p2}" stroke="rgba(160,235,255,.35)" stroke-width="2" stroke-dasharray="3 6" fill="none"/>
<path d="{p3}" stroke="rgba(160,235,255,.35)" stroke-width="2" stroke-dasharray="3 6" fill="none"/>
{cloud(cx, cy, cloud_label)}{dashboard(dx, dy)}
{packet(p1, 2.6)}{packet(p1, 2.6, 1.3)}{packet(p2, 1.6, 2.6, "#8cf8f5")}{packet(p3, 1.2, 4.2, "#37d399")}
</svg>"""


# ---------------------------------------------------------------- slides
CSS = """
/* wireless communication slides */
.cm-slide .slide-inner{max-width:1180px}
.cm-slide .cell-demo svg{width:100%;height:auto;display:block}
.cm-arc{animation:cmArc 1.6s ease-out infinite;opacity:0}
@keyframes cmArc{0%{opacity:0}25%{opacity:.95}100%{opacity:0}}
.cm-wave{animation:cmWave 3s linear infinite}
@keyframes cmWave{to{transform:translateX(-44px)}}
.cm-layout{display:grid;grid-template-columns:1.25fr 1fr;gap:1rem;margin-top:1rem;align-items:stretch}
.cm-stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.55rem;margin-top:.8rem}
.cm-flow{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.7rem;margin-top:1rem}
.cm-step{padding:.8rem .9rem;border:1px solid var(--line2);border-radius:14px;background:rgba(2,33,57,.5);font-size:.86rem;line-height:1.45;color:var(--muted)}
.cm-step b{display:block;font-family:Georgia,serif;font-size:1.02rem;color:var(--ink);margin-bottom:.25rem}
.cm-step i{font-style:normal;display:inline-flex;width:1.5rem;height:1.5rem;border-radius:50%;background:var(--teal);color:#012a3a;font-weight:800;font-size:.8rem;align-items:center;justify-content:center;margin-right:.45rem}
.cm-pc{display:grid;grid-template-columns:1fr 1fr;gap:.7rem;margin-top:.8rem}
@media(max-width:900px){.cm-layout,.cm-flow,.cm-pc{grid-template-columns:1fr}.cm-stats{grid-template-columns:1fr 1fr}}
"""


def li(items):
    return "".join(f"<li>{e(t)}</li>" for t in items)


def overview():
    steps = [("Sensors read", "The probes underwater send readings up the cable to the buoy (RS-485)."),
             ("Controller packs", "A small board in the buoy saves the readings and packs them into short messages."),
             ("Radio link", "The buoy's radio sends the message through the air to shore."),
             ("Internet → dashboard", "A tower or gateway on shore forwards it over the internet to the live dashboard.")]
    flow = "".join(f'<div class="cm-step"><b><i>{n}</i>{e(a)}</b>{e(b)}</div>' for n, (a, b) in enumerate(steps, 1))
    return f"""<section class="slide salinity-slide cm-slide" data-comm="overview">
  <div class="water-ribbons" aria-hidden="true"><span></span><span></span><span></span><span></span></div>
  <div class="slide-inner">
    <span class="kicker rise" style="--i:0">Data link · From the sea to the dashboard</span>
    <h2 class="rise" style="--i:1">Wireless <span class="hl">communication</span></h2>
    <p class="sub rise" style="--i:2">The sensors stay underwater on a cable. The buoy has no wire to land, so it sends the readings through the air to shore, and from there over the internet to the dashboard.</p>
    <div class="salinity-panel panel rise" style="--i:3;margin-top:1rem" data-fx="radio"><div class="cell-demo">{scene("Buoy + radio", "tower", "Shore receiver")}</div></div>
    <div class="cm-flow rise" style="--i:4">{flow}</div>
    <p class="slide-footnote rise" style="--i:5">Three ways to make the radio link follow: <b style="color:var(--gold)">cellular</b>, <b style="color:var(--gold)">LoRa</b> and a <b style="color:var(--gold)">Wi-Fi / radio link</b>. The readings are tiny, so all three can carry them.</p>
  </div>
</section>"""


METHODS = [
    dict(key="cellular", kicker="Data link · Option 1 of 3", title='<span class="hl">Cellular</span> (4G / LTE-M)',
         sub="The buoy carries a small modem with a SIM card, like a phone. It sends the readings to the nearest mobile tower, which forwards them over the internet.",
         scene=("4G modem + SIM", "tower", "Mobile tower"),
         stats=[("Few – 20 km", "Range: to the nearest tower"), ("Bursts 1–2 A", "Power: sleeps between sends"), ("SIM plan", "Small monthly data fee"), ("All sites", "If signal reaches")],
         steps=["The controller wakes up every few minutes and reads the sensors.",
                "The modem connects to the nearest mobile tower and sends the readings.",
                "The tower forwards them over the internet to the dashboard."],
         pros=["Nothing to install on shore", "Works anywhere there is mobile signal", "Two-way: we can send commands back"],
         cons=["Signal over the sea must be checked on site", "Short power peaks: needs a good battery", "Monthly SIM fee"],
         foot="<b style=\"color:var(--gold)\">Best for ALEX-MON:</b> the first choice for every site. Test the signal on a boat at each site before buying."),
    dict(key="lora", kicker="Data link · Option 2 of 3", title='<span class="hl">LoRa</span> long-range radio',
         sub="A low-power radio built for small messages over long distances. The buoy sends to a LoRa gateway on shore, which is connected to the internet.",
         scene=("LoRa radio", "gateway", "LoRa gateway (shore)"),
         stats=[("5 – 15 km", "Range over water"), ("Very low", "Power: runs for months"), ("No fees", "No SIM, own gateway"), ("Canopus · Heracleion", "Farther sites")],
         steps=["The buoy sends a short message (a few readings) on the LoRa radio.",
                "A gateway on shore, with an antenna up high, hears it.",
                "The gateway forwards it over the internet to the dashboard."],
         pros=["Very low power: suits a solar buoy", "Long range with a small antenna", "No SIM or monthly fee"],
         cons=["Only small messages (no photos or video)", "Needs a powered gateway on shore", "Check Egyptian radio rules (NTRA) before use"],
         foot="<b style=\"color:var(--gold)\">Best for ALEX-MON:</b> the backup for Canopus (~2 km) and Heracleion (~6–7 km) if cellular signal is weak."),
    dict(key="wifi", kicker="Data link · Option 3 of 3", title='<span class="hl">Wi-Fi / radio</span> point-to-point link',
         sub="Two directional antennas aimed at each other, one on the buoy and one on shore, make a private wireless link, like a long Wi-Fi cable through the air.",
         scene=("Buoy + dish antenna", "dish", "Shore antenna"),
         stats=[("1 – 5 km", "Range with directional antennas"), ("Medium", "Power: radio kept on"), ("Hardware only", "No fees"), ("Qaitbay · Harbour", "Near-shore sites")],
         steps=["A directional antenna on the buoy points at the shore antenna.",
                "The two radios form a fast private link over the water.",
                "The shore side connects to the internet and the dashboard."],
         pros=["Fast: can send much more data, even photos", "No fees", "Simple, cheap hardware"],
         cons=["The buoy rocks, so aiming is hard to keep", "Needs clear line of sight to a shore antenna", "Shorter range than LoRa; uses more power"],
         foot="<b style=\"color:var(--gold)\">Best for ALEX-MON:</b> Qaitbay and the Eastern Harbour, which are within about 200 m of the shore."),
]


def method(n, m):
    return f"""<section class="slide salinity-slide cm-slide" data-comm="{m['key']}">
  <div class="water-ribbons" aria-hidden="true"><span></span><span></span><span></span><span></span></div>
  <div class="slide-inner">
    <span class="kicker rise" style="--i:0">{e(m['kicker'])}</span>
    <h2 class="rise" style="--i:1">{m['title']}</h2>
    <p class="sub rise" style="--i:2">{e(m['sub'])}</p>
    <div class="cm-layout">
      <div>
        <div class="salinity-panel panel rise" style="--i:3" data-fx="radio"><div class="cell-demo">{scene(*m['scene'])}</div></div>
        <div class="cm-stats rise" style="--i:4">{"".join(f'<div class="gs-tile"><b>{e(a)}</b><span>{e(b)}</span></div>' for a, b in m['stats'])}</div>
      </div>
      <div class="salinity-panel panel rise" style="--i:5">
        <h3>How it works</h3>
        <ol class="cell-steps">{"".join(f"<li><b>{i}</b><span>{e(s)}</span></li>" for i, s in enumerate(m['steps'], 1))}</ol>
        <div class="cm-pc"><div class="gs-box pros"><h4>Advantages</h4><ul>{li(m['pros'])}</ul></div><div class="gs-box cons"><h4>Limitations</h4><ul>{li(m['cons'])}</ul></div></div>
      </div>
    </div>
    <p class="slide-footnote rise" style="--i:6">{m['foot']}</p>
  </div>
</section>"""


def main():
    block = "<!-- COMM-SLIDES:START (generated by files/guides/build_comm_slides.py) -->\n" + "\n\n".join(
        [overview()] + [method(i, m) for i, m in enumerate(METHODS, 1)]) + "\n<!-- COMM-SLIDES:END -->"
    css = "/* COMM-CSS:START */" + CSS + "/* COMM-CSS:END */"
    t = INDEX.read_text(encoding="utf-8")
    if "<!-- COMM-SLIDES:START" in t:
        t = re.sub(r"<!-- COMM-SLIDES:START.*?<!-- COMM-SLIDES:END -->", lambda m: block, t, flags=re.S)
        t = re.sub(r"/\* COMM-CSS:START \*/.*?/\* COMM-CSS:END \*/", lambda m: css, t, flags=re.S)
    else:
        anchor = "<!-- SENSOR-SLIDES:END -->"
        assert anchor in t, "sensor block not found"
        t = t.replace(anchor, anchor + "\n\n" + block, 1)
        t = t.replace("/* SENSOR-CSS:END */", "/* SENSOR-CSS:END */\n" + css, 1)
    n = len(re.findall(r'<section class="slide', t))
    t = re.sub(r'(<span id="total">)\d+(</span>)', lambda m: m.group(1) + str(n) + m.group(2), t, count=1)
    INDEX.write_text(t, encoding="utf-8")
    print("slides:", n)


if __name__ == "__main__":
    main()
