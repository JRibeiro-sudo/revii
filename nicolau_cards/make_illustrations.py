"""Renders detailed illustrations to images/<key>.png (1024x1024) — free alternative to AI photos.
Each drawing uses the card's tile_color as background. Skips cards that already have an image.

Usage: python make_illustrations.py [--force] [key ...]
"""
import json, pathlib, sys
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).resolve().parent

# Shared defs: soft drop shadow + glass/shine gradients. viewBox 0 0 100 100.
DEFS = """
<filter id="sh" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="2.2"/></filter>
<filter id="soft" x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation=".6"/></filter>
<linearGradient id="glass" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity=".55"/><stop offset=".25" stop-color="#fff" stop-opacity=".15"/><stop offset=".75" stop-color="#cfe6f5" stop-opacity=".2"/><stop offset="1" stop-color="#9fc3dc" stop-opacity=".55"/></linearGradient>
<linearGradient id="shine" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity="0"/><stop offset=".5" stop-color="#fff" stop-opacity=".9"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient>
"""

def shadow(cx=50, cy=90, rx=26, ry=4, op=.22):
    return f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="#000" opacity="{op}" filter="url(#sh)"/>'

ART = {
"agua": shadow() + """
<defs><linearGradient id="wat" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8fd3ff"/><stop offset="1" stop-color="#3a9ad9"/></linearGradient>
<linearGradient id="watS" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity=".35"/><stop offset=".3" stop-color="#fff" stop-opacity="0"/><stop offset=".85" stop-color="#1d6fa8" stop-opacity=".25"/></linearGradient></defs>
<path d="M30 14h40l-4.6 72a4 4 0 0 1-4 3.8H38.6a4 4 0 0 1-4-3.8z" fill="url(#glass)" stroke="#a9c9de" stroke-width=".8"/>
<path d="M31.9 36h36.2l-3.1 49.6a3 3 0 0 1-3 2.8H38a3 3 0 0 1-3-2.8z" fill="url(#wat)"/>
<path d="M31.9 36h36.2l-3.1 49.6a3 3 0 0 1-3 2.8H38a3 3 0 0 1-3-2.8z" fill="url(#watS)"/>
<ellipse cx="50" cy="36" rx="18.1" ry="2.6" fill="#c8ecff"/><ellipse cx="50" cy="36" rx="18.1" ry="2.6" fill="none" stroke="#fff" stroke-opacity=".8" stroke-width=".6"/>
<ellipse cx="50" cy="14" rx="20" ry="2.4" fill="none" stroke="#b9d6e8" stroke-width=".9"/>
<path d="M35 18l2.4 64" stroke="#fff" stroke-width="2.6" stroke-linecap="round" opacity=".75"/>
<path d="M39.5 20l1.4 30" stroke="#fff" stroke-width="1" stroke-linecap="round" opacity=".6"/>
<g fill="#fff" opacity=".7"><circle cx="56" cy="52" r="1"/><circle cx="60" cy="66" r=".8"/><circle cx="47" cy="74" r=".7"/><circle cx="58" cy="78" r="1.1"/></g>
""",
"leite": shadow() + """
<defs><linearGradient id="milk" x1="0" x2="1"><stop offset="0" stop-color="#f4f1ea"/><stop offset=".3" stop-color="#fffefb"/><stop offset=".8" stop-color="#ede8dd"/><stop offset="1" stop-color="#d9d2c4"/></linearGradient></defs>
<path d="M30 14h40l-4.6 72a4 4 0 0 1-4 3.8H38.6a4 4 0 0 1-4-3.8z" fill="url(#glass)" stroke="#bccbd6" stroke-width=".8"/>
<path d="M31.3 28h37.4l-3.7 57.6a3 3 0 0 1-3 2.8H38a3 3 0 0 1-3-2.8z" fill="url(#milk)"/>
<ellipse cx="50" cy="28" rx="18.7" ry="2.8" fill="#fffdf8"/><ellipse cx="50" cy="28" rx="18.7" ry="2.8" fill="none" stroke="#e5dfd2" stroke-width=".6"/>
<path d="M32 24.6c0 2 1.5 3 3 3" fill="none" stroke="#fff" stroke-width="1.6" opacity=".9"/>
<ellipse cx="50" cy="14" rx="20" ry="2.4" fill="none" stroke="#c7d4de" stroke-width=".9"/>
<path d="M35 18l2.3 64" stroke="#fff" stroke-width="2.6" stroke-linecap="round" opacity=".8"/>
<path d="M63 34l-2.2 48" stroke="#c9c1b1" stroke-width="2" stroke-linecap="round" opacity=".5"/>
""",
"biberon": shadow(rx=20) + """
<defs><linearGradient id="tea" x1="0" x2="1"><stop offset="0" stop-color="#f2c98a"/><stop offset=".4" stop-color="#fbe2b6"/><stop offset="1" stop-color="#d9a660"/></linearGradient>
<linearGradient id="ring" x1="0" x2="1"><stop offset="0" stop-color="#3f9ed6"/><stop offset=".35" stop-color="#8fd0f5"/><stop offset="1" stop-color="#2475ad"/></linearGradient>
<linearGradient id="bot" x1="0" x2="1"><stop offset="0" stop-color="#f6f9fc"/><stop offset=".3" stop-color="#fff"/><stop offset="1" stop-color="#d7e1ea"/></linearGradient>
<linearGradient id="bmilk" x1="0" x2="1"><stop offset="0" stop-color="#f6efe0"/><stop offset=".35" stop-color="#fffaf0"/><stop offset="1" stop-color="#e6dac2"/></linearGradient></defs>
<path d="M44 26c0-6 1.2-10 3-12.4a3.6 3.6 0 0 1 6 0c1.8 2.4 3 6.4 3 12.4z" fill="url(#tea)" opacity=".95"/>
<ellipse cx="50" cy="13" rx="2.4" ry="1.6" fill="#f7d9a6"/>
<path d="M40 25h20l1.5 3h-23z" fill="#f0c27e"/>
<rect x="34" y="27" width="32" height="11" rx="3.5" fill="url(#ring)"/>
<path d="M34 31h32" stroke="#fff" stroke-opacity=".35" stroke-width="1"/>
<path d="M36 38h28v41a9 9 0 0 1-9 9H45a9 9 0 0 1-9-9z" fill="url(#bot)" stroke="#c9d6e1" stroke-width=".7"/>
<path d="M36.6 54h26.8v25a8.4 8.4 0 0 1-8.4 8.4H45a8.4 8.4 0 0 1-8.4-8.4z" fill="url(#bmilk)"/>
<ellipse cx="50" cy="54" rx="13.4" ry="1.5" fill="#fffdf6" stroke="#eadfca" stroke-width=".4"/>
<g stroke="#7aaed3" stroke-width=".9" stroke-linecap="round"><path d="M57 46h5M59 51h3M57 56h5M59 61h3M57 66h5M59 71h3M57 76h5"/></g>
<path d="M39.5 41v38" stroke="#fff" stroke-width="2.4" stroke-linecap="round" opacity=".9"/>
<g fill="#f6a5c0"><path d="M43 70c-2.4-2.2-3.6-3.4-3.6-4.8a1.8 1.8 0 0 1 3.6-.6 1.8 1.8 0 0 1 3.6.6c0 1.4-1.2 2.6-3.6 4.8z"/></g>
""",
"papa": shadow(cy=86, rx=34, ry=4.5) + """
<defs><linearGradient id="bowl" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7ec3f2"/><stop offset="1" stop-color="#2f7fc0"/></linearGradient>
<linearGradient id="bowlS" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity=".35"/><stop offset=".35" stop-color="#fff" stop-opacity="0"/><stop offset="1" stop-color="#0b3f6b" stop-opacity=".3"/></linearGradient>
<radialGradient id="pur" cx=".45" cy=".4" r=".7"><stop offset="0" stop-color="#ffc27a"/><stop offset=".6" stop-color="#f59a3c"/><stop offset="1" stop-color="#d9741f"/></radialGradient>
<linearGradient id="spn" x1="0" x2="1"><stop offset="0" stop-color="#ffd84d"/><stop offset="1" stop-color="#f0a800"/></linearGradient></defs>
<path d="M12 52h76c0 18-16 32-38 32S12 70 12 52z" fill="url(#bowl)"/>
<path d="M12 52h76c0 18-16 32-38 32S12 70 12 52z" fill="url(#bowlS)"/>
<path d="M38 84h24l-2 4H40z" fill="#2a6fa8"/>
<ellipse cx="50" cy="52" rx="38" ry="9" fill="#cfe9fb"/>
<ellipse cx="50" cy="53" rx="34.5" ry="7.2" fill="url(#pur)"/>
<path d="M30 52c6-3 14 1 20-1s12-3 18 0" fill="none" stroke="#ffd7a3" stroke-width="1.6" stroke-linecap="round" opacity=".8"/>
<path d="M38 56c5 1.5 10 1 14-.5" fill="none" stroke="#c9661a" stroke-width="1" stroke-linecap="round" opacity=".5"/>
<path d="M18 60c2 7 7 12 14 15" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" opacity=".55"/>
<path d="M56 52 80 20" stroke="url(#spn)" stroke-width="5.5" stroke-linecap="round"/>
<path d="M57.5 51 79 22.5" stroke="#fff" stroke-width="1.2" stroke-linecap="round" opacity=".6"/>
<ellipse cx="52" cy="54" rx="7" ry="3.6" fill="#f0a800"/><ellipse cx="52" cy="53.4" rx="6" ry="2.8" fill="#ffb347"/>
""",
}

def main():
    args = sys.argv[1:]
    force = "--force" in args
    only = [a for a in args if not a.startswith("--")]
    data = json.loads((HERE / "cards.json").read_text(encoding="utf-8"))
    cards = sorted(data["cards"], key=lambda c: (c["page"], c["position"]))
    todo = [c for c in cards if c["key"] in ART and (not only or c["key"] in only)]
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": 1024, "height": 1024})
        for i, c in enumerate(todo, 1):
            out = HERE / c["image_file"]
            if out.exists() and not force:
                print(f"{i}/{len(todo)} {out.name} (já existe, saltado)"); continue
            out.parent.mkdir(parents=True, exist_ok=True)
            svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="1024" height="1024">'
                   f'<defs>{DEFS}</defs><rect width="100" height="100" fill="{c["tile_color"]}"/>'
                   f'<g transform="translate(50 50) scale(.74) translate(-50 -52)">{ART[c["key"]]}</g></svg>')
            pg.set_content(f'<html><body style="margin:0">{svg}</body></html>')
            pg.screenshot(path=str(out), clip={"x": 0, "y": 0, "width": 1024, "height": 1024})
            print(f"{i}/{len(todo)} {out.name}")
        b.close()

if __name__ == "__main__":
    main()
