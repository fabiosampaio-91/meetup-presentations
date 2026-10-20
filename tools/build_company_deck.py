#!/usr/bin/env python3
"""
synvert xgeeks generic company deck builder: deck.json -> PPTX + PDF + web deck.

Usage:  python3 tools/build_company_deck.py "<Community> Presentations/<dated folder>"

Reuses the shared Canvas/backends/pipeline from build_deck.py (same PPTX+web dual
geometry, same LibreOffice/Chrome PDF pipeline, same preview + GitHub Pages publish
step) but with its own 10-slide layout, since this deck's slide types (sponsors &
partners lockup, brand + GlobalLogic lockup, clients/community logo grids, reusable
"What About <community>" and "Tech Talks Formats" placeholder slides) are unrelated
to the fixed 9-slide meetup-event template in build_deck.py.
"""
from __future__ import annotations
import json, sys
from pathlib import Path

TOOLS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS_DIR))
from build_deck import (  # noqa: E402
    Canvas, Run, SLIDE_W, SLIDE_H,
    convert_pdf_libreoffice, print_pdf_chrome, pdf_info, brand_font_embedded, render_previews, screenshot_web,
    publish_docs, CHROME,
)

MARGIN = 0.6


def kicker_pill(c, x, y, text, bg, fg, size=15, weight=700, tab=0.32, pill_h=0.6, w=None, pad=0.28):
    """The xgeeks deck's own title treatment: a small square tab overlapping a
    rounded pill ('$ kubectl apply -f ...', 'What About Viseu Tech Talks', ...)."""
    C = c.C
    c.rect(x, y - tab * 0.55, tab, tab, bg, radius=0.08)
    tw = c.m.width_in(text, size, weight)
    w = w or (tw + 2 * pad)
    c.rect(x, y, w, pill_h, bg, radius=pill_h / 2)
    c.text(x + pad * 0.7, y, w - pad, pill_h, text, size, fg, weight, "left", "middle", check=False)
    return w


def block_title(c, text, size=28):
    """A dark, bold section title used under a kicker pill (not the AWS title_card)."""
    c.text(MARGIN, 1.28, 11.5, 0.55, text, size, c.C["navy"], 700, valign="middle")


def draw_deck(c: Canvas, d: dict, mdir: Path):
    C = c.C
    total = 10
    A = lambda p: f"assets/logos/{p}"  # noqa: E731

    def footer(dark):
        col = C["muted_on_dark"] if dark else C["muted_on_light"]
        c.text(MARGIN, 7.08, 9, 0.3, d.get("footer", ""), 10, col, 500, valign="middle", check=False)
        c.text(11.6, 7.08, 1.13, 0.3, f"{c.n} / {total}", 10, col, 500, "right", "middle", check=False)

    # ---------------------------------------------------------------- 1. COVER
    c.new_slide(bg="gradient")
    cov = d["cover"]
    c.placeholder(1.4, 1.35, SLIDE_W - 2.8, SLIDE_H - 2.7, cov["label"], cov["sublabel"], radius=0.22)

    # ------------------------------------------------- 2. SPONSORS & PARTNERS
    c.new_slide(bg="gradient")
    sp = d["sponsors_partners"]
    c.card(MARGIN, 0.45, c.m.width_in(sp["title"], 30, 700) + 0.8, 0.95, framed=True)
    c.text(MARGIN + 0.35, 0.45, 9, 0.95, sp["title"], 30, C["black"], 700, valign="middle")

    col_w = 7.7
    c.label(MARGIN, 1.75, 6, sp["sponsors_label"], color=C["white"])
    n_s = sp["sponsor_slots"]
    gap = 0.25
    sw = (col_w - gap * (n_s - 1)) / n_s
    for i in range(n_s):
        x = MARGIN + i * (sw + gap)
        c.card(x, 2.1, sw, 3.05, framed=True)
        c.placeholder(x + 0.25, 2.35, sw - 0.5, 1.55, "SPONSOR", "logo goes here")
        c.placeholder(x + 0.25, 4.05, sw - 0.5, 0.85, "Name · kind", None)
    c.rect(MARGIN + col_w + 0.5, 1.75, 0.014, 3.4, C["muted_on_dark"], radius=0)
    px0 = MARGIN + col_w + 0.85
    pcol_w = SLIDE_W - MARGIN - px0
    c.label(px0, 1.75, pcol_w, sp["partners_label"], color=C["white"])
    n_p = sp["partner_slots"]
    pw = (pcol_w - gap * (n_p - 1)) / n_p
    for i in range(n_p):
        x = px0 + i * (pw + gap)
        c.card(x, 2.1, pw, 3.05, framed=True)
        c.placeholder(x + 0.2, 2.35, pw - 0.4, 1.55, "PARTNER", "logo goes here")
        c.placeholder(x + 0.2, 4.05, pw - 0.4, 0.85, "Name · kind", None)
    c.text(MARGIN, 5.85, SLIDE_W - 2 * MARGIN, 0.9, sp["thanks"], 15, C["white"], 500, "center", "middle", ls=1.1)
    footer(True)

    # -------------------------------------------------------- 3. ABOUT xgeeks
    c.new_slide(bg="#FFFFFF")
    ab = d["about"]
    kicker_pill(c, MARGIN, 0.55, "$ whoami", C["teal"], C["black"])
    block_title(c, ab["title"])
    stats = ab["stats"]
    gap = 0.3
    n = len(stats)
    sw = (7.35 - gap * (n - 1)) / n
    for i, s in enumerate(stats):
        x = MARGIN + i * (sw + gap)
        c.card(x, 2.15, sw, 2.05, framed=True)
        c.text(x + 0.1, 2.15, sw - 0.2, 1.15, s["value"], 30, C["navy"], 700, "center", "middle", check=False)
        c.text(x + 0.1, 3.35, sw - 0.2, 0.65, s["label"], 13, C["muted_on_light"], 500, "center", "middle")
    c.card(MARGIN, 4.55, 7.35, 2.05, framed=True)
    c.image(mdir / ab["photo"], MARGIN + 0.12, 4.67, 7.11, 1.81, fit="cover", radius=0.1)
    c.card(8.25, 2.15, 4.48, 4.45, framed=True)
    c.text(8.55, 2.4, 3.9, 3.35,
           "synvert xgeeks is a team of engineers who accelerate enterprise success "
           "with AI-native and Cloud Native engineering — turning DevOps, data and "
           "AI into real business impact for the world's leading brands.",
           16.5, C["navy"], 600, valign="middle", ls=1.2)
    c.image(mdir / "assets/logos/synvert-xgeeks-navy.png", 8.55, 6.05, 2.1, 0.42)
    footer(False)

    # ------------------------------------------- 4. WE'RE synvert xgeeks + GL
    c.new_slide(bg="#002E33")
    wa = d["we_are"]
    c.text(2.0, 1.5, 6, 1.1, wa["kicker"], 34, C["white"], 700, check=False)
    c.image(mdir / wa["logo"], 1.85, 2.25, 8.8, 3.65)
    c.rect(2.0, 6.15, 9.33, 0.014, C["muted_on_dark"], radius=0)
    c.text(2.0, 6.35, 3.6, 0.55, "PART OF", 12, C["muted_on_dark"], 700, letter_spacing=1.5, valign="middle")
    c.card(6.0, 6.28, 5.33, 0.98, framed=True)
    c.image(mdir / wa["globallogic_logo"], 6.28, 6.44, 4.1, 0.66)
    footer(True)

    # ------------------------------------------------------------ 5. ACCELERATE
    c.new_slide(bg="#FFFFFF")
    ac = d["accelerate"]
    tab = 0.42
    c.rect(MARGIN, 1.0 - tab * 0.55, tab, tab, C["teal"], radius=0.1)
    c.card(MARGIN, 1.0, 7.55, 3.55, framed=True)
    c.rect(MARGIN, 1.0, 7.55, 3.55, C["teal"], radius=0.2)
    c.text(MARGIN + 0.45, 1.3, 6.7, 2.05, ac["headline"], 29, C["navy"], 700, ls=1.08)
    c.text(MARGIN + 0.45, 3.55, 6.7, 0.85, ac["subhead"], 14, C["navy"], 500, ls=1.15)
    c.image(mdir / ac["icon"], 8.55, 1.55, 4.2, 4.2, align="center")
    c.pill(MARGIN, 4.95, 0.48, ac["badge"], C["navy"], C["white"], 13, 700)
    c.image(mdir / "assets/logos/synvert-xgeeks-navy.png", 10.75, 6.35, 2.0, 0.58)
    footer(False)

    # ------------------------------------------------------------- 6. WHAT WE DO
    c.new_slide(bg="#FFFFFF")
    wd = d["what_we_do"]
    kicker_pill(c, MARGIN, 0.55, wd["kicker"], C["teal"], C["black"], size=16, w=6.6)
    c.text(MARGIN, 1.35, 8, 0.5, wd["subtag"], 13.5, C["muted_on_light"], 600, letter_spacing=0.3, valign="middle")
    areas = wd["areas"]
    n = len(areas)
    gap = 0.3
    aw = (SLIDE_W - 2 * MARGIN - gap * (n - 1)) / n
    for i, a in enumerate(areas):
        x = MARGIN + i * (aw + gap)
        c.card(x, 2.15, aw, 4.05, framed=True)
        c.text(x + 0.25, 2.15, aw - 0.5, 4.05, a["title"], 21, C["navy"], 700, valign="middle", ls=1.1)
    c.image(mdir / wd["icon"], 10.85, 0.35, 1.9, 1.9)
    footer(False)

    # --------------------------------------------------------------- 7. CLIENTS
    c.new_slide(bg="#FFFFFF")
    cl = d["clients"]
    kicker_pill(c, MARGIN, 0.5, cl["kicker"], C["teal"], C["black"], size=15, w=4.6)
    logos = cl["logos"]
    cols, rows = 6, 3
    gap = 0.18
    gw = (SLIDE_W - 2 * MARGIN - gap * (cols - 1)) / cols
    gh = (5.55 - gap * (rows - 1)) / rows
    y0 = 1.4
    last_row_n = len(logos) - cols * (rows - 1)
    for i, name in enumerate(logos):
        r, cIdx = divmod(i, cols)
        row_n = last_row_n if r == rows - 1 else cols
        row_offset = (cols - row_n) * (gw + gap) / 2
        x = MARGIN + row_offset + cIdx * (gw + gap)
        y = y0 + r * (gh + gap)
        c.card(x, y, gw, gh, framed=False)
        c.rect(x, y, gw, gh, None, radius=0.14, line=C["black"], line_w=1.25)
        c.image(mdir / A(name), x + gw * 0.1, y + gh * 0.14, gw * 0.8, gh * 0.72, fit="contain")
    footer(False)

    # ------------------------------------------------ 8. WHAT ABOUT <COMMUNITY>
    c.new_slide(bg="#FFFFFF")
    wc = d["what_about_community"]
    title_text = [[Run(wc["title_prefix"], 700, C["navy"]), Run(wc["community_placeholder"], 700, C["white"], italic=True)]]
    tab = 0.42
    c.rect(MARGIN, 0.55 - tab * 0.55, tab, tab, C["teal"], radius=0.1)
    c.rect(MARGIN, 0.55, 8.4, 0.85, C["teal"], radius=0.42)
    c.text(MARGIN + 0.35, 0.55, 8.0, 0.85, title_text, 26, C["navy"], 700, valign="middle", check=False)
    c.card(MARGIN, 1.75, 6.0, 4.85, framed=True)
    c.placeholder(MARGIN + 0.25, 2.0, 5.5, 4.35, wc["image_label"], wc["image_sublabel"])
    c.card(6.95, 1.75, 5.78, 4.85, framed=True)
    stat_h = 4.35 / len(wc["stats"])
    for i, s in enumerate(wc["stats"]):
        y = 2.0 + i * stat_h
        c.text(7.2, y, 5.3, stat_h, [f"•  {s}"], 16, C["navy"], 600, valign="middle", check=False)
    footer(False)

    # ---------------------------------------------------------- 9. TECH TALKS FORMATS
    c.new_slide(bg="#FFFFFF")
    tf = d["tech_talks_formats"]
    kicker_pill(c, MARGIN, 0.55, tf["title"], C["teal"], C["black"], size=17, w=4.1)
    blocks = tf["blocks"]
    gap = 0.4
    bw = (SLIDE_W - 2 * MARGIN - gap) / 2
    for i, blk in enumerate(blocks):
        x = MARGIN + i * (bw + gap)
        c.rect(x, 1.75, bw, 0.14, C["teal"], radius=0)
        c.card(x, 1.89, bw, 4.7, framed=True)
        if blk.get("logo"):
            c.image(mdir / blk["logo"], x + bw / 2 - 1.6, 2.25, 3.2, 1.7, align="center")
        else:
            c.placeholder(x + 0.4, 2.25, bw - 0.8, 1.7, blk.get("logo_placeholder", "Logo goes here"), None)
        c.rect(x + 0.4, 4.15, bw - 0.8, 0.012, C["grey"], radius=0)
        c.text(x + 0.4, 4.35, bw - 0.8, 2.05, blk["description_placeholder"], 14, C["muted_on_light"], 500, ls=1.2)
    footer(False)

    # ------------------------------------------ 10. $ kubectl apply -f ./community.yaml
    c.new_slide(bg="#FFFFFF")
    cg = d["community_grid"]
    kicker_pill(c, MARGIN, 0.5, cg["kicker"], C["teal"], C["black"], size=15, w=4.9)
    c.text(MARGIN, 1.28, 9.6, 0.55, cg["tagline"], 13.5, C["muted_on_light"], 600, ls=1.1, valign="middle")
    c.image(mdir / cg["icon"], 11.55, 0.3, 1.2, 1.2)
    communities = cg["communities"]
    cols, rows = 3, 3
    gap = 0.26
    gw = (SLIDE_W - 2 * MARGIN - gap * (cols - 1)) / cols
    gh = (5.2 - gap * (rows - 1)) / rows
    y0 = 1.75
    for i, com in enumerate(communities):
        r, cIdx = divmod(i, cols)
        x = MARGIN + cIdx * (gw + gap)
        y = y0 + r * (gh + gap)
        c.card(x, y, gw, gh, framed=True)
        c.image(mdir / A(com["logo"]), x + gw * 0.08, y + gh * 0.1, gw * 0.84, gh * 0.62, fit="contain")
        c.text(x + 0.15, y + gh - 0.5, gw - 0.3, 0.4, com["name"], 11.5, C["navy"], 700, "center", "middle", ls=0.95)
    footer(False)

    assert c.n == total, f"expected {total} slides, drew {c.n}"


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    mdir = Path(sys.argv[1]).resolve()
    repo = mdir.parent.parent
    deck = json.loads((mdir / "deck.json").read_text(encoding="utf-8"))
    slug = deck["slug"]
    c = Canvas(deck, mdir)
    draw_deck(c, deck, mdir)
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
    import shutil
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
