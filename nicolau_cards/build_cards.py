"""Builds Nicolau_resimli_kartlar.pdf from cards.json.
Uses images/<key>.png when present; otherwise falls back to the original drawing (art.py).
Requires: pip install playwright && python -m playwright install chromium"""
import json, html, asyncio, pathlib
HERE=pathlib.Path(__file__).resolve().parent
exec((HERE/"art.py").read_text(encoding="utf-8"))
D=json.loads((HERE/"cards.json").read_text(encoding="utf-8"))
E=html.escape
def P(s): return E(s).replace("^",'<sup class="nz">n</sup>')
def card(c):
    w,p,t,note=c["pt"],c["pron"],c["tr"],c["sound"]
    fs=20 if len(w)<=10 else (16.5 if len(w)<=14 else 14.5)
    ps=14 if len(p)<=16 else 12
    img=HERE/c["image_file"]
    for ext in (".png",".jpg",".jpeg",".webp"):
        cand=img.with_suffix(ext)
        if cand.exists(): img=cand; break
    if img.exists():
        pic=f'<img src="{img.as_uri()}" style="width:100%;height:100%;object-fit:cover;border-radius:10px">'
    else:
        pic=f'<svg viewBox="0 0 100 100" width="104" height="104"><g stroke="{O}" stroke-width="2.4" stroke-linejoin="round" stroke-linecap="round">{A[c["key"]]}</g></svg>'
    nt=f' <span class="ns">“{E(note)}”</span>' if note else ''
    return (f'<div class="cd"><div class="tile" style="background:{c["tile_color"]}">{pic}</div>'
            f'<div class="pw" style="font-size:{fs}px">{E(w)}</div><div class="pp" style="font-size:{ps}px">{P(p)}</div><div class="tt">{E(t)}{nt}</div></div>')
CSS="""
@page{size:A4;margin:0} html,body{margin:0;padding:0;background:#FFFDF9}
body{font-family:'Source Sans 3','Segoe UI',sans-serif;color:#1E1B18}
.pg{width:794px;height:1123px;box-sizing:border-box;padding:34px 40px 28px;display:flex;flex-direction:column;overflow:hidden;break-after:page}
.pg:last-child{break-after:auto}
.hd{text-align:center;padding-bottom:6px;border-bottom:1.5px solid}
.hd .a{font-family:'Fraunces',Georgia,serif;font-size:20px;font-weight:700;letter-spacing:.02em}
.hd .b{font-size:13px;color:#5A534B;margin-top:1px}
.cg{margin-top:12px;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));grid-template-rows:repeat(5,186px);gap:10px}
.cd{box-sizing:border-box;border:1.5px solid #E2D8C8;border-radius:14px;background:#fff;padding:7px 6px 6px;display:flex;flex-direction:column;align-items:center;text-align:center;overflow:hidden}
.tile{width:100%;height:116px;border-radius:10px;display:flex;align-items:center;justify-content:center;overflow:hidden}
.pw{font-family:'Fraunces',Georgia,serif;font-weight:700;line-height:22px;margin-top:6px;white-space:nowrap}
.pp{font-weight:700;color:#1F3F75;line-height:18px;white-space:nowrap}
.tt{font-size:12.5px;line-height:16px;color:#5A534B;white-space:nowrap}
.ns{color:#A2441F;font-style:italic}
sup.nz{font-size:.68em;line-height:0;vertical-align:.55em;margin-left:.5px;font-weight:700;color:#A2441F}
.foot{margin-top:auto;display:flex;justify-content:space-between;font-size:11.5px;color:#5A534B}
"""
FOOT='Mavi satırı Türkçe gibi okuyun · BÜYÜK hece = vurgu · <sup class="nz">n</sup> = genizden · hr = gırtlaktan r · ı = çok kısa'
FONTS='<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,700&family=Source+Sans+3:wght@400;600;700&display=swap" rel="stylesheet">'
NP=len(D["pages"])
out=[]
for pg in D["pages"]:
    cs=sorted([c for c in D["cards"] if c["page"]==pg["page"]],key=lambda c:c["position"])
    out.append(f'<div class="pg"><div class="hd" style="border-color:{pg["color"]}"><div class="a" style="color:{pg["color"]}">{pg["title_tr"]}</div><div class="b">{pg["title_pt"]} · Nicolau ile Portekizce</div></div><div class="cg">{"".join(card(c) for c in cs)}</div><div class="foot"><div>{FOOT}</div><div>{pg["page"]} / {NP}</div></div></div>')
doc=f'<!doctype html><html lang="tr"><head><meta charset="utf-8"><title>Nicolau ile Portekizce</title>{FONTS}<style>{CSS}</style></head><body>{"".join(out)}</body></html>'
(HERE/"cards.html").write_text(doc,encoding="utf-8")
async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        b=await p.chromium.launch(); pg=await b.new_page()
        await pg.goto((HERE/"cards.html").as_uri()); await pg.wait_for_timeout(1500); await pg.evaluate("document.fonts.ready")
        await pg.pdf(path=str(HERE/"Nicolau_resimli_kartlar.pdf"),width="210mm",height="297mm",print_background=True)
        await b.close()
asyncio.run(main())
print("OK ->",HERE/"Nicolau_resimli_kartlar.pdf")
