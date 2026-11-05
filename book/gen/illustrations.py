"""Flat illustrations for chapter covers and opening cases, rendered from SVG with Playwright."""
import math
import os
import sys

NAVY, AMBER, TEAL, RUST, MIST = "#14213D", "#DFA03F", "#2C7A7B", "#B5532A", "#DDEAEA"


def target(cx, cy, r):
    rings = [(r, NAVY), (r * .78, "#FFFFFF"), (r * .56, TEAL), (r * .34, "#FFFFFF"), (r * .14, AMBER)]
    return "".join(f'<circle cx="{cx}" cy="{cy}" r="{rr:.0f}" fill="{c}"/>' for rr, c in rings)


def dots(points, color, r=16):
    return "".join(f'<circle cx="{x}" cy="{y}" r="{r}" fill="{color}" stroke="#FFFFFF" stroke-width="5"/>' for x, y in points)


SVGS = {
    # Chapter 10 – bias (cluster away from centre) and noise (scatter)
    "ch11_cover": (1280, 960, target(420, 480, 300) + dots([(520, 330), (560, 360), (500, 370), (545, 310), (580, 335)], RUST)
                   + target(980, 480, 220) + dots([(900, 380), (1080, 450), (960, 600), (1040, 330), (930, 520)], AMBER, 13)),
    "ch11_case": (1240, 800,
                  f'<rect x="120" y="170" width="380" height="470" rx="28" fill="{MIST}"/><rect x="740" y="170" width="380" height="470" rx="28" fill="{MIST}"/>'
                  f'<rect x="170" y="230" width="280" height="26" rx="13" fill="{NAVY}"/><rect x="170" y="290" width="200" height="22" rx="11" fill="{TEAL}"/>'
                  f'<rect x="790" y="230" width="280" height="26" rx="13" fill="{NAVY}"/><rect x="790" y="290" width="200" height="22" rx="11" fill="{TEAL}"/>'
                  f'<text x="310" y="500" font-family="Arial" font-weight="bold" font-size="130" fill="{TEAL}" text-anchor="middle">9.5</text>'
                  f'<text x="930" y="500" font-family="Arial" font-weight="bold" font-size="130" fill="{RUST}" text-anchor="middle">16</text>'
                  f'<circle cx="620" cy="400" r="70" fill="{AMBER}"/><text x="620" y="440" font-family="Arial" font-weight="bold" font-size="110" fill="{NAVY}" text-anchor="middle">?</text>'),
    # Chapter 11 – a chart with one highlighted bar and a speech bubble
    "ch13_cover": (1280, 960,
                   f'<rect x="200" y="140" width="880" height="600" rx="30" fill="{MIST}"/>'
                   + "".join(f'<rect x="{300 + i * 150}" y="{640 - h}" width="100" height="{h}" rx="12" fill="{c}"/>'
                             for i, (h, c) in enumerate([(180, "#9FB7BC"), (250, "#9FB7BC"), (210, "#9FB7BC"), (420, RUST), (230, "#9FB7BC")]))
                   + f'<rect x="260" y="640" width="760" height="10" fill="{NAVY}"/><rect x="600" y="740" width="80" height="120" fill="{NAVY}"/><rect x="460" y="850" width="360" height="40" rx="20" fill="{NAVY}"/>'
                   f'<path d="M820 70 h320 a30 30 0 0 1 30 30 v120 a30 30 0 0 1 -30 30 h-210 l-60 60 v-60 h-50 a30 30 0 0 1 -30 -30 v-120 a30 30 0 0 1 30 -30z" fill="{AMBER}"/>'
                   f'<rect x="860" y="120" width="240" height="22" rx="11" fill="{NAVY}"/><rect x="860" y="170" width="170" height="22" rx="11" fill="{NAVY}"/>'),
    "ch13_case": (1240, 800, "__ROSE__"),
    # Chapter 12 – A/B test and monitoring
    "ch14_cover": (1280, 960,
                   f'<rect x="120" y="200" width="460" height="420" rx="30" fill="{MIST}"/><rect x="700" y="200" width="460" height="420" rx="30" fill="{MIST}"/>'
                   f'<text x="350" y="330" font-family="Arial" font-weight="bold" font-size="110" fill="{NAVY}" text-anchor="middle">A</text>'
                   f'<text x="930" y="330" font-family="Arial" font-weight="bold" font-size="110" fill="{NAVY}" text-anchor="middle">B</text>'
                   f'<rect x="200" y="420" width="300" height="60" rx="30" fill="#9FB7BC"/><rect x="200" y="510" width="200" height="40" rx="20" fill="#9FB7BC"/>'
                   f'<rect x="780" y="420" width="300" height="60" rx="30" fill="{TEAL}"/><rect x="780" y="510" width="260" height="40" rx="20" fill="{TEAL}"/>'
                   f'<circle cx="1110" cy="640" r="85" fill="{AMBER}"/><path d="M1070 640 l30 32 l55 -62" stroke="{NAVY}" stroke-width="22" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
                   f'<path d="M180 760 C 400 700, 520 820, 700 760 S 1000 700, 1100 760" stroke="{RUST}" stroke-width="18" fill="none" stroke-linecap="round"/>'),
    "ch14_case": (1240, 800,
                  f'<rect x="150" y="150" width="400" height="500" rx="30" fill="{MIST}"/><rect x="690" y="150" width="400" height="500" rx="30" fill="{MIST}"/>'
                  f'<rect x="200" y="220" width="300" height="34" rx="17" fill="{NAVY}"/><rect x="200" y="280" width="220" height="22" rx="11" fill="#9FB7BC"/>'
                  f'<rect x="740" y="220" width="300" height="34" rx="17" fill="{NAVY}"/><rect x="740" y="270" width="300" height="34" rx="17" fill="{NAVY}"/><rect x="740" y="330" width="220" height="22" rx="11" fill="#9FB7BC"/>'
                  f'<rect x="230" y="520" width="60" height="80" fill="#9FB7BC"/><rect x="320" y="490" width="60" height="110" fill="#9FB7BC"/>'
                  f'<rect x="770" y="470" width="60" height="130" fill="{TEAL}"/><rect x="860" y="420" width="60" height="180" fill="{TEAL}"/><rect x="950" y="380" width="60" height="220" fill="{RUST}"/>'),
    # Chapter 13 – six connected steps around a decision
    "ch18_cover": (1280, 960, "__CYCLE__"),
    "ch18_case": (1240, 800,
                  f'<path d="M220 200 h260 l50 60 h500 a30 30 0 0 1 30 30 v360 a30 30 0 0 1 -30 30 h-810 a30 30 0 0 1 -30 -30 v-420 a30 30 0 0 1 30 -30z" fill="{TEAL}"/>'
                  f'<rect x="300" y="300" width="640" height="300" rx="20" fill="#FFFFFF"/>'
                  + "".join(f'<rect x="350" y="{340 + i * 55}" width="{w}" height="24" rx="12" fill="{c}"/>' for i, (w, c) in enumerate([(420, NAVY), (520, "#9FB7BC"), (360, "#9FB7BC"), (470, "#9FB7BC")]))
                  + f'<circle cx="940" cy="600" r="80" fill="{AMBER}"/><path d="M905 600 l26 28 l50 -56" stroke="{NAVY}" stroke-width="20" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'),
}


def rose():
    out = []
    radii = [(330, TEAL), (300, TEAL), (250, TEAL), (210, TEAL), (180, TEAL), (160, "#9FB7BC"),
             (140, "#9FB7BC"), (120, "#9FB7BC"), (150, "#9FB7BC"), (190, RUST), (240, RUST), (280, RUST)]
    step = math.pi / 6
    for i, (r, c) in enumerate(radii):
        a0, a1 = i * step - math.pi / 2, (i + 1) * step - math.pi / 2
        x0, y0 = 620 + r * math.cos(a0), 420 + r * math.sin(a0)
        x1, y1 = 620 + r * math.cos(a1), 420 + r * math.sin(a1)
        out.append(f'<path d="M620 420 L{x0:.0f} {y0:.0f} A{r} {r} 0 0 1 {x1:.0f} {y1:.0f} Z" fill="{c}" stroke="#FFFFFF" stroke-width="6"/>')
    return "".join(out)


def cycle():
    out = []
    pts = [(640 + 300 * math.cos(i * math.pi / 3 - math.pi / 2), 480 + 300 * math.sin(i * math.pi / 3 - math.pi / 2)) for i in range(6)]
    for x, y in pts:
        out.append(f'<line x1="640" y1="480" x2="{x:.0f}" y2="{y:.0f}" stroke="{NAVY}" stroke-width="14"/>')
    for i, ((x, y), c) in enumerate(zip(pts, [TEAL, TEAL, AMBER, RUST, NAVY, "#6B4E71"])):
        out.append(f'<circle cx="{x:.0f}" cy="{y:.0f}" r="95" fill="{c}"/>'
                   f'<text x="{x:.0f}" y="{y + 36:.0f}" font-family="Arial" font-weight="bold" font-size="100" fill="#FFFFFF" text-anchor="middle">{i + 1}</text>')
    out.append(f'<circle cx="640" cy="480" r="120" fill="{AMBER}" stroke="{NAVY}" stroke-width="14"/>'
               f'<path d="M585 480 l40 42 l75 -84" stroke="{NAVY}" stroke-width="26" fill="none" stroke-linecap="round" stroke-linejoin="round"/>')
    return "".join(out)


SVGS["ch13_case"] = (1240, 800, rose())
SVGS["ch18_cover"] = (1280, 960, cycle())


def render(out_dir, names=None):
    """Write each SVG next to its PNG and rasterise with Node Playwright (global install)."""
    import json
    import subprocess
    os.makedirs(out_dir, exist_ok=True)
    jobs = []
    for name, (w, h, body) in SVGS.items():
        if names and name not in names:
            continue
        svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
               f'<rect width="{w}" height="{h}" fill="#FFFFFF"/>{body}</svg>')
        path = os.path.abspath(os.path.join(out_dir, name + ".svg"))
        open(path, "w").write(svg)
        jobs.append({"svg": path, "png": path[:-4] + ".png", "w": w, "h": h})
    js = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rasterize.js")
    subprocess.run(["node", js, json.dumps(jobs)], check=True,
                   env={**os.environ, "NODE_PATH": "/opt/node22/lib/node_modules"})


if __name__ == "__main__":
    render(sys.argv[1] if len(sys.argv) > 1 else "img")
