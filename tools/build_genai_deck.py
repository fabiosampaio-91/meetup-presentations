#!/usr/bin/env python3
"""
Lisbon GenAI Community meetup deck builder: deck.json -> PPTX + PDF + web deck.

Usage:  python3 tools/build_genai_deck.py "Lisbon GenAI Presentations/<dated folder>" [--no-publish]

The synvert xgeeks company deck (build_company_deck.py) re-skinned with the Lisbon GenAI visual
identity from the meetup poster (warm red-orange -> yellow -> lime gradient, black hero cards with
the corner tab, white framed cards with a hard shadow, coral accents) and re-ordered for a GenAI
meetup night:

  1 banner  2 the community  3 sponsors & partners  4 We're synvert xgeeks  5 accelerate
  6 people  7 expertise  8 clients  9 communities  10 KCD (in KCD Porto colours)
  11 Talk with us  12 thanks to the sponsors (hand-over to the venue sponsor)

Same Canvas/backends/PDF/preview/publish pipeline as build_deck.py; the slide background and the
gradient "Talk with us" panel are generated here from brand.gradient into assets/derived/.
"""
from __future__ import annotations
import json, shutil, sys
from pathlib import Path
import numpy as np
from PIL import Image

TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))
from build_deck import (  # noqa: E402
    Canvas, Run, SLIDE_W, SLIDE_H,
    convert_pdf_libreoffice, print_pdf_chrome, pdf_info, brand_font_embedded, render_previews, screenshot_web,
    publish_docs,
)
from build_company_deck import kicker_pill  # noqa: E402

MARGIN = 0.6
TOTAL = 12


# ----------------------------------------------------------------------------- generated brand images

def gradient_rgb(w, h, stops, tilt=0.08):
    """Poster gradient: left-to-right colour stops, slightly faster towards the bottom (as on the poster)."""
    x = np.linspace(0, 1, w, dtype=np.float32)[None, :]
    y = np.linspace(0, 1, h, dtype=np.float32)[:, None]
    t = np.clip((x + tilt * y) / (1 + tilt), 0, 1)
    pos = [s[0] for s in stops]
    cols = np.array([[int(s[1][i:i + 2], 16) for i in (1, 3, 5)] for s in stops], np.float32)
    out = np.stack([np.interp(t, pos, cols[:, k]) for k in range(3)], axis=-1)
    return Image.fromarray(out.astype(np.uint8), "RGB")


def ensure_brand_images(deck, mdir: Path):
    b = deck["brand"]
    bg = mdir / b["background"]
    bg.parent.mkdir(parents=True, exist_ok=True)
    gradient_rgb(2400, 1350, b["gradient"]).save(bg, quality=90)
    tw = deck["talk_with_us"]
    shape = Image.open(mdir / tw["panel_shape"]).convert("RGBA")
    px = tw["panel"]
    w = int(px["w"] * 160)
    h = int(round(w * shape.height / shape.width))
    # the exported shape carries faint (alpha < 64) shadow-filter lines around it: drop them
    mask = shape.getchannel("A").point(lambda v: 0 if v < 64 else v).resize((w, h), Image.LANCZOS)
    panel = gradient_rgb(w, h, b["gradient"]).convert("RGBA")
    panel.putalpha(mask)
    out = mdir / "assets" / "derived" / "talk-panel-gradient.png"
    panel.save(out)
    return out, px["w"] * h / w


# ----------------------------------------------------------------------------- GenAI design language

def framed(c, x, y, w, h, fill, line, shadow, radius=0.2, line_w=2.5):
    """Card with outline + hard offset shadow (the poster's date card / the AWS + xgeeks card language)."""
    d = c.brand.get("shadow_offset", 0.09)
    c.rect(x + d, y + d, w, h, shadow, radius)
    c.rect(x, y, w, h, fill, radius, line, line_w)


def hero(c, x, y, w, h, tab=True):
    """The poster's black hero card, with the small square corner tab of the xgeeks title motif."""
    C = c.C
    if tab:
        t = 0.42
        c.rect(x, y - t * 0.55, t, t, C["black"], radius=0.1)
    framed(c, x, y, w, h, C["black"], C["black"], C["black"])


def kicker(c, text, y=0.55, size=15, bg=None, fg=None):
    C = c.C
    return kicker_pill(c, MARGIN, y, text, bg or C["black"], fg or C["white"], size=size)


def logo_fit(c, path, x, y, w, h, area):
    """Place a logo centred in (x, y, w, h) at a common optical size (equal area), capped by the box."""
    im = Image.open(path)
    ar = im.width / im.height
    lw = (area * ar) ** 0.5
    lh = lw / ar
    if lw > w:
        lw, lh = w, w / ar
    if lh > h:
        lw, lh = h * ar, h
    c.image(path, x + (w - lw) / 2, y + (h - lh) / 2, lw, lh, fit="fill")


def footer(c, d, color=None):
    col = color or c.C["ink"]
    c.text(MARGIN, 7.08, 9, 0.3, d["footer"], 10, col, 600, valign="middle", check=False)
    c.text(11.6, 7.08, 1.13, 0.3, f"{c.n} / {TOTAL}", 10, col, 600, "right", "middle", check=False)


# ----------------------------------------------------------------------------- slides

def draw_deck(c: Canvas, d: dict, mdir: Path, panel_png: Path, panel_h: float):
    C = c.C
    P = lambda p: mdir / p  # noqa: E731

    # ------------------------------------------------------------- 1. BANNER (the #16 poster, as is)
    c.new_slide(bg=None)
    c.image(P(d["banner"]), 0, 0, SLIDE_W, SLIDE_H, fit="fill")

    # ------------------------------------------------------- 2. WHAT ABOUT LISBON GENAI COMMUNITY
    c.new_slide()
    g = d["group_info"]
    tab = 0.42
    tw = c.m.width_in(g["title_prefix"] + g["name"], 26, 700) + 0.75
    c.rect(MARGIN, 0.55 - tab * 0.55, tab, tab, C["black"], radius=0.1)
    c.rect(MARGIN, 0.55, tw, 0.85, C["black"], radius=0.42)
    c.text(MARGIN + 0.35, 0.55, tw - 0.4, 0.85,
           [[Run(g["title_prefix"], 700, C["white"]), Run(g["name"], 700, C["coral"])]],
           26, C["white"], 700, valign="middle", check=False)
    hero(c, MARGIN, 1.85, 4.55, 4.9, tab=False)
    c.image(P(g["visual"]), MARGIN + 0.3, 2.1, 3.95, 4.4)
    x0, cw, gap = 5.55, 7.18, 0.3
    tile_w, tile_h, row_gap = (cw - gap) / 2, 1.42, 0.24
    for i, s in enumerate(g["stats"]):
        r, k = divmod(i, 2)
        x, y = x0 + k * (tile_w + gap), 1.85 + r * (tile_h + row_gap)
        c.card(x, y, tile_w, tile_h, framed=True)
        c.text(x + 0.3, y + 0.1, tile_w - 0.6, 0.72, s["value"], 30, C["black"], 700, valign="middle", check=False)
        c.text(x + 0.3, y + 0.82, tile_w - 0.6, 0.5, s["label"], 12, C["muted_on_light"], 600, ls=1.0)
    y = 1.85 + 2 * tile_h + row_gap + 0.3
    c.card(x0, y, cw, 6.75 - y, framed=True)
    c.text(x0 + 0.3, y + 0.14, cw - 0.6, 0.86, g["about"], 13.5, C["ink"], 600, valign="middle", ls=1.1)
    c.text(x0 + 0.3, y + 1.0, cw - 0.6, 0.3, g["sponsor_line"], 11, C["muted_on_light"], 500, valign="middle")
    footer(c, d)

    # ------------------------------------------------------------------- 3. SPONSORS & PARTNERS
    c.new_slide()
    sp = d["sponsors_partners"]
    tw = c.m.width_in(sp["title"], 30, 700) + 0.8
    c.card(MARGIN, 0.45, tw, 0.95, framed=True)
    c.text(MARGIN + 0.35, 0.45, tw - 0.5, 0.95, sp["title"], 30, C["black"], 700, valign="middle")
    sec_gap = 0.9
    sec_w = (SLIDE_W - 2 * MARGIN - sec_gap) / 2
    for si, (label, items) in enumerate(((sp["sponsors_label"], sp["sponsors"]), (sp["partners_label"], sp["partners"]))):
        sx = MARGIN + si * (sec_w + sec_gap)
        c.label(sx, 1.8, sec_w, label, color=C["black"])
        n, gap = len(items), 0.3
        w = (sec_w - gap * (n - 1)) / n
        for i, it in enumerate(items):
            x = sx + i * (w + gap)
            c.card(x, 2.2, w, 3.1, framed=True)
            logo_fit(c, P(it["logo"]), x + 0.3, 2.45, w - 0.6, 1.75, area=1.15)
            c.rect(x + 0.3, 4.4, w - 0.6, 0.012, C["grey"], radius=0)
            c.text(x + 0.2, 4.5, w - 0.4, 0.62, it["name"], 15, C["black"], 700, "center", "middle")
    c.rect(MARGIN + sec_w + sec_gap / 2 - 0.01, 1.8, 0.02, 3.5, C["black"], radius=0)
    c.text(MARGIN, 5.75, SLIDE_W - 2 * MARGIN, 0.9, sp["thanks"], 17, C["black"], 600, "center", "middle", ls=1.1)
    footer(c, d)

    # ------------------------------------------------------- 4. WE'RE synvert xgeeks + GlobalLogic
    c.new_slide()
    wa = d["we_are"]
    hero(c, MARGIN, 0.75, SLIDE_W - 2 * MARGIN, 6.0)
    c.text(1.6, 1.2, 6, 0.9, wa["kicker"], 34, C["white"], 700, check=False)
    c.image(P(wa["logo"]), 1.45, 1.95, 8.6, 3.55, align="left")
    c.rect(1.6, 5.75, 10.13, 0.014, C["muted_on_dark"], radius=0)
    c.text(1.6, 5.95, 3.6, 0.6, wa["part_of"], 12, C["muted_on_dark"], 700, letter_spacing=1.5, valign="middle")
    c.rect(8.43, 5.9, 3.3, 0.88, C["white"], radius=0.16)
    c.image(P(wa["globallogic_logo"]), 8.6, 5.97, 2.96, 0.74)
    footer(c, d)

    # ------------------------------------------------------------------------------ 5. ACCELERATE
    c.new_slide()
    ac = d["accelerate"]
    hero(c, MARGIN, 1.0, 7.55, 3.6)
    c.text(MARGIN + 0.45, 1.3, 6.7, 2.1, ac["headline"], 30, C["white"], 700, ls=1.08)
    c.text(MARGIN + 0.45, 3.5, 6.7, 0.85, ac["subhead"], 15, C["coral"], 600, ls=1.15)
    c.pill(MARGIN, 5.0, 0.5, ac["badge"], C["white"], C["black"], 13, 700)
    c.image(P(ac["icon"]), 8.55, 1.1, 4.2, 4.5)
    c.rect(10.45, 6.0, 2.28, 0.82, C["white"], radius=0.2)
    c.image(P(d["logos"]["synvert_light"]), 10.6, 6.08, 1.98, 0.66)
    footer(c, d)

    # ------------------------------------------------------------------------- 6. PEOPLE (whoami)
    c.new_slide()
    pe = d["people"]
    kicker(c, pe["kicker"])
    tw_, th_, gap = 2.8, 2.2, 0.3
    for i, s in enumerate(pe["stats"]):
        r, k = divmod(i, 2)
        x, y = MARGIN + k * (tw_ + gap), 1.85 + r * (th_ + gap)
        c.card(x, y, tw_, th_, framed=True)
        c.text(x + 0.15, y + 0.3, tw_ - 0.3, 1.05, s["value"], 36, C["black"], 700, "center", "middle", check=False)
        c.text(x + 0.15, y + 1.38, tw_ - 0.3, 0.5, s["label"], 14, C["muted_on_light"], 600, "center", "middle")
    px_ = MARGIN + 2 * tw_ + gap + 0.45
    pw = SLIDE_W - MARGIN - px_
    c.card(px_, 1.85, pw, 4.7, framed=True)
    c.image(P(pe["photo"]), px_ + 0.14, 1.99, pw - 0.28, 4.42, fit="cover", radius=0.12)
    c.image(P(pe["icon"]), 11.45, 0.3, 1.3, 1.45)
    footer(c, d)

    # ------------------------------------------------------------------------------ 7. EXPERTISE
    c.new_slide()
    ex = d["expertise"]
    kicker(c, ex["kicker"])
    n, gap = len(ex["areas"]), 0.35
    aw = (SLIDE_W - 2 * MARGIN - gap * (n - 1)) / n
    for i, a in enumerate(ex["areas"]):
        x = MARGIN + i * (aw + gap)
        c.card(x, 2.05, aw, 4.5, framed=True)
        c.text(x + 0.35, 2.3, 2, 0.9, f"{i + 1:02d}", 40, C["red"], 700, valign="middle", check=False)
        c.text(x + 0.35, 4.0, aw - 0.7, 2.25, a, 25, C["black"], 700, valign="bottom", ls=1.08)
    c.image(P(ex["icon"]), 11.05, 0.2, 1.05, 1.98)
    footer(c, d)

    # -------------------------------------------------------------------------------- 8. CLIENTS
    c.new_slide()
    cl = d["clients"]
    kicker(c, cl["kicker"])
    c.image(P(cl["icon"]), 11.4, 0.22, 1.35, 1.15)
    logos = cl["logos"]
    cols, rows, gap = 6, 3, 0.2
    gw = (SLIDE_W - 2 * MARGIN - gap * (cols - 1)) / cols
    gh = (5.25 - gap * (rows - 1)) / rows
    last_n = len(logos) - cols * (rows - 1)
    for i, lg in enumerate(logos):
        lg = lg if isinstance(lg, dict) else {"file": lg}
        r, k = divmod(i, cols)
        off = (cols - last_n) * (gw + gap) / 2 if r == rows - 1 else 0
        x, y = MARGIN + off + k * (gw + gap), 1.55 + r * (gh + gap)
        # "dark": the logo's own artwork is white-on-dark (keeps its colours on a black plate)
        framed(c, x, y, gw, gh, C["black"] if lg.get("dark") else C["white"], C["black"], C["black"],
               radius=0.14, line_w=1.5)
        logo_fit(c, P(f"assets/logos/{lg['file']}"), x + 0.18, y + 0.2, gw - 0.36, gh - 0.4,
                 area=0.62 * lg.get("scale", 1.0))
    footer(c, d)

    # ---------------------------------------------------------------------------- 9. COMMUNITIES
    c.new_slide()
    cg = d["communities"]
    kicker(c, cg["kicker"])
    c.text(MARGIN, 1.28, 10.2, 0.5, cg["tagline"], 14, C["black"], 600, ls=1.05, valign="middle")
    c.image(P(cg["icon"]), 11.25, 0.3, 1.5, 0.85)
    cols, gap = 3, 0.27
    gw = (SLIDE_W - 2 * MARGIN - gap * (cols - 1)) / cols
    gh = (4.95 - gap * 2) / 3
    for i, com in enumerate(cg["items"]):
        r, k = divmod(i, cols)
        x, y = MARGIN + k * (gw + gap), 1.9 + r * (gh + gap)
        dark = com["name"] == cg["highlight"]
        if dark:
            framed(c, x, y, gw, gh, C["black"], C["black"], C["black"])
        else:
            c.card(x, y, gw, gh, framed=True)
        logo = com.get("logo_dark") if dark else com["logo"]
        c.image(P(f"assets/logos/{logo}"), x + gw * 0.1, y + 0.1, gw * 0.8, gh - 0.55)
        c.text(x + 0.15, y + gh - 0.45, gw - 0.3, 0.38, com["name"], 12, C["white"] if dark else C["black"], 700,
               "center", "middle", ls=0.95)
    footer(c, d)

    # ----------------------------------------------------------- 10. KCD (KCD Porto raffle colours)
    k = d["kcd"]
    K = k["colors"]
    c.new_slide(bg=K["bg"])
    kicker(c, k["title"], size=17, bg=K["title"], fg=C["white"])
    lw = 7.95
    framed(c, MARGIN, 1.6, lw, 2.15, K["blue"], K["ink"], K["ink"])
    c.image(P(k["lockup"]), MARGIN + 0.28, 1.84, 3.2, 1.67)
    c.text(MARGIN + 3.8, 1.8, lw - 4.05, 1.0, k["event_name"], 19, C["white"], 700, valign="bottom", ls=1.05)
    c.text(MARGIN + 3.8, 2.88, lw - 4.05, 0.45, k["when_where"], 15, C["white"], 600, valign="middle")
    c.text(MARGIN + 3.8, 3.3, lw - 4.05, 0.3, k["hashtag"], 10.5, K["light"], 600, valign="middle")
    c.text(MARGIN, 3.98, 6, 0.32, k["keynotes_label"].upper(), 11, K["title"], 700, letter_spacing=1.5, valign="middle")
    kn = k["keynotes"]
    s = (lw - 0.2 * (len(kn) - 1)) / len(kn)
    for i, photo in enumerate(kn):
        x = MARGIN + i * (s + 0.2)
        c.rect(x + 0.06, 4.4 + 0.06, s, s, K["ink"], radius=0.12)
        c.image(P(photo), x, 4.4, s, s, fit="cover", radius=0.12)
    rx, rw = MARGIN + lw + 0.3, SLIDE_W - MARGIN - (MARGIN + lw + 0.3)
    framed(c, rx, 1.6, rw, 4.4 + s - 1.6, C["white"], K["ink"], K["ink"])
    q = 2.3
    c.qr("kcd-site", k["url"], q, rx + (rw - q) / 2, 1.78)
    c.text(rx + 0.15, 4.12, rw - 0.3, 0.32, k["qr_label"], 12.5, K["ink"], 700, "center", "middle")
    c.text(rx + 0.15, 4.5, rw - 0.3, 0.5, k["discount_label"], 25, K["title"], 700, "center", "middle")
    c.text(rx + 0.15, 5.02, rw - 0.3, 0.26, k["code_label"].upper(), 10, K["muted"], 700, "center", "middle", letter_spacing=1.2)
    c.pill(rx + rw / 2, 5.36, 0.56, k["discount_code"], K["blue"], C["white"], 16, 700, align="center", w=rw - 0.5)
    footer(c, d, color=K["ink"])

    # -------------------------------------------------------------------------- 11. TALK WITH US
    c.new_slide(bg="#FFFFFF")
    tk = d["talk_with_us"]
    pn = tk["panel"]
    c.image(panel_png, pn["x"], pn["y"], pn["w"], panel_h, fit="fill")
    # content stays inside the panel's main body (x 0.55-9.52, y 0.8-5.49 for this shape at w=10.25)
    kicker_pill(c, 1.05, 1.2, tk["kicker"], C["black"], C["white"], size=15)
    c.text(1.05, 1.8, 6.0, 1.35, tk["headline"], 32, C["black"], 700, ls=1.05, valign="middle")
    c.card(1.05, 3.3, 1.95, 1.95, framed=True)
    c.qr("talk-with-us", tk["url"], 1.75, 1.15, 3.4)
    for i, it in enumerate(tk["items"]):
        y = 3.48 + i * 0.58
        c.image(P(it["icon"]), 3.45, y, 0.42, 0.42)
        c.text(4.03, y - 0.05, 3.2, 0.52, it["label"], 15.5, C["black"], 600, valign="middle")
    c.image(P(tk["illustration"]), 6.75, 1.65, 6.05, 3.45)
    c.text(MARGIN, 5.98, 4, 0.32, tk["site"], 12.5, C["black"], 700, valign="middle")
    x = MARGIN
    for so in tk["socials"]:
        tw = c.m.width_in(so["handle"], 12, 500)
        w = tw + 0.95
        c.rect(x, 6.38, w, 0.46, C["white"], radius=0.12, line=C["black"], line_w=1.25)
        c.image(P(so["icon"]), x + 0.17, 6.49, 0.26, 0.24)
        c.text(x + 0.55, 6.38, tw + 0.3, 0.46, so["handle"], 12, C["black"], 500, valign="middle", check=False)
        x += w + 0.2
    c.image(P(d["logos"]["synvert_light"]), 11.0, 6.05, 1.75, 0.75, align="right")
    footer(c, d)

    # ----------------------------------------------------------- 12. THANKS TO THE MEETUP SPONSORS
    c.new_slide()
    cs = d["closing"]
    c.text(MARGIN, 0.9, SLIDE_W - 2 * MARGIN, 1.9, cs["text"], 46, C["black"], 700, "center", "middle", ls=1.05)
    cw_, ch_ = 6.6, 2.75
    c.card((SLIDE_W - cw_) / 2, 3.25, cw_, ch_, framed=True)
    logo_fit(c, P(cs["logo"]), (SLIDE_W - cw_) / 2 + 0.55, 3.55, cw_ - 1.1, ch_ - 0.6, area=6.0)
    footer(c, d)

    assert c.n == TOTAL, f"expected {TOTAL} slides, drew {c.n}"


# ----------------------------------------------------------------------------- pipeline

def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    mdir = Path(sys.argv[1]).resolve()
    repo = mdir.parent.parent
    deck = json.loads((mdir / "deck.json").read_text(encoding="utf-8"))
    slug = deck["slug"]
    panel_png, panel_h = ensure_brand_images(deck, mdir)
    c = Canvas(deck, mdir)
    draw_deck(c, deck, mdir, panel_png, panel_h)
    pptx_path = mdir / f"{slug}.pptx"
    web_dir = mdir / "web"; web_dir.mkdir(exist_ok=True)
    html_path = web_dir / "index.html"
    c.save(pptx_path, html_path)
    print(f"PPTX  {pptx_path}")
    print(f"WEB   {html_path} ({html_path.stat().st_size / 1e6:.2f} MB)")
    for w in c.warnings:
        print("WARN ", w)

    preview = mdir / "preview"; preview.mkdir(exist_ok=True)
    brand_font = deck["brand"]["font"].replace(" ", "")
    pdf_path = convert_pdf_libreoffice(pptx_path)
    lo_pages, lo_fonts = pdf_info(pdf_path)
    assert lo_pages == c.n, f"LibreOffice PDF has {lo_pages} pages, expected {c.n}"
    lo_check = preview / "libreoffice-check.pdf"
    shutil.copy2(pdf_path, lo_check)
    render_previews(lo_check, preview, prefix="pptx")
    engine = "libreoffice"
    if not brand_font_embedded(brand_font, lo_fonts):
        engine = "chrome"
        print_pdf_chrome(html_path, pdf_path)
    pages, fonts = pdf_info(pdf_path)
    assert pages == c.n, f"PDF has {pages} pages, expected {c.n}"
    render_previews(pdf_path, preview, prefix="slide")
    print(f"PDF   {pdf_path}  pages={pages} (slides={c.n})  engine={engine}  fonts={fonts}")
    if engine == "chrome":
        print(f"NOTE  LibreOffice embedded {lo_fonts} instead of {brand_font}; PDF printed from the web deck instead. "
              f"LibreOffice render kept for geometry check: {lo_check}")
    if screenshot_web(html_path, preview, c.n):
        print(f"WEB screenshots -> {preview}/web-XX.png")
    if "--no-publish" not in sys.argv:
        dest = publish_docs(repo, mdir, deck, html_path, pdf_path)
        print(f"DOCS  {dest}")


if __name__ == "__main__":
    main()
