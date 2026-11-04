#!/usr/bin/env python3
"""
Multi-community meetup night deck builder (Cloud Native Porto x AWS UG Porto x ...): deck.json -> PPTX + PDF + web.

Usage:  python3 tools/build_cnp_deck.py "Cloud Native Porto Presentations/<dated folder>" [--no-publish]

The Lisbon GenAI deck (build_genai_deck.py) without the synvert xgeeks company slides, re-skinned
with the meetup banner (vertical blue -> cyan gradient, white cards with the black outline and hard
shadow, blue name pills) and with two slides per hosting community:

  1 banner  2..2k  per community: "What About <Community>" + leadership team
  then  sponsors & partners  ·  KCD (in KCD Porto colours)  ·  agenda  ·  one slide per sponsor
  (the last one is the hand-over to tonight's host)

Same Canvas/backends/PDF/preview/publish pipeline as build_deck.py; the slide background is
generated here from brand.gradient into assets/derived/.
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
from build_genai_deck import framed, logo_fit  # noqa: E402

MARGIN = 0.6


def total_slides(d):
    return 1 + 2 * len(d["communities"]) + 3 + len(d["sponsor_slides"])


# ----------------------------------------------------------------------------- generated brand images

def ensure_background(deck, mdir: Path):
    """Banner background: a linear top-to-bottom gradient between the sampled stops."""
    b = deck["brand"]
    out = mdir / b["background"]
    out.parent.mkdir(parents=True, exist_ok=True)
    w, h = 2400, 1350
    pos = [s[0] for s in b["gradient"]]
    cols = np.array([[int(s[1][i:i + 2], 16) for i in (1, 3, 5)] for s in b["gradient"]], np.float32)
    t = np.linspace(0, 1, h, dtype=np.float32)
    col = np.stack([np.interp(t, pos, cols[:, k]) for k in range(3)], axis=-1)
    Image.fromarray(np.repeat(col[:, None, :], w, axis=1).astype(np.uint8), "RGB").save(out, quality=92)


# ----------------------------------------------------------------------------- design language

def title_pill(c, runs, y=0.55, size=24):
    """Black tab + pill title (xgeeks/GenAI motif) with mixed-colour runs, e.g. 'What About' + name."""
    C = c.C
    tab = 0.42
    tw = sum(c.m.width_in(r.text, size, 700) for r in runs) + 0.75
    c.rect(MARGIN, y - tab * 0.55, tab, tab, C["black"], radius=0.1)
    c.rect(MARGIN, y, tw, 0.8, C["black"], radius=0.4)
    c.text(MARGIN + 0.35, y, tw - 0.4, 0.8, [runs], size, C["white"], 700, valign="middle", check=False)


def footer(c, d, total, color=None):
    col = color or c.C["ink"]
    c.text(MARGIN, 7.08, 9.5, 0.3, d["footer"], 10, col, 600, valign="middle", check=False)
    c.text(11.6, 7.08, 1.13, 0.3, f"{c.n} / {total}", 10, col, 600, "right", "middle", check=False)


def avatar(c, path, x, y, dia):
    c.ellipse(x - 0.03, y - 0.03, dia + 0.06, dia + 0.06, c.C["black"])
    c.image(path, x, y, dia, dia, fit="cover", radius=dia / 2)


# ----------------------------------------------------------------------------- slides

def draw_deck(c: Canvas, d: dict, mdir: Path):
    C = c.C
    P = lambda p: mdir / p  # noqa: E731
    total = total_slides(d)

    # ------------------------------------------------------------- 1. BANNER (as published)
    c.new_slide(bg=None)
    c.image(P(d["banner"]), 0, 0, SLIDE_W, SLIDE_H, fit="fill")

    for g in d["communities"]:
        # ------------------------------------------------------- WHAT ABOUT <COMMUNITY>
        c.new_slide()
        title_pill(c, [Run("What About ", 700, C["white"]), Run(g["name"], 700, C["cyan"])])
        hx, hy, hw, hh = MARGIN, 1.85, 4.55, 4.9
        c.card(hx, hy, hw, hh, framed=True)
        c.image(P(g["visual"]), hx + 0.3, hy + 0.3, hw - 0.6, 2.75)
        c.rect(hx + 0.3, hy + 3.3, hw - 0.6, 0.012, C["grey_line"], radius=0)
        q = 1.25
        c.qr(g["qr_name"], g["join_url"], q, hx + 0.22, hy + hh - q - 0.2)
        c.text(hx + 1.6, hy + 3.55, hw - 1.8, 0.42, g["join_label"], 15, C["black"], 700, valign="middle")
        c.text(hx + 1.6, hy + 3.97, hw - 1.8, 0.62, g["join_short"], 10.5, C["muted_on_light"], 500, ls=1.05,
               link=g["join_url"])
        x0, cw, gap = 5.55, 7.18, 0.3
        tile_w, tile_h, row_gap = (cw - gap) / 2, 1.2, 0.24
        for i, s in enumerate(g["stats"]):
            r, k = divmod(i, 2)
            x, y = x0 + k * (tile_w + gap), 1.85 + r * (tile_h + row_gap)
            c.card(x, y, tile_w, tile_h, framed=True)
            c.text(x + 0.3, y + 0.1, tile_w - 0.6, 0.68, s["value"], 30, C["blue"], 700, valign="middle", check=False)
            c.text(x + 0.3, y + 0.72, tile_w - 0.6, 0.4, s["label"], 12, C["muted_on_light"], 600, ls=1.0)
        y = 1.85 + 2 * tile_h + row_gap + 0.3
        h = 6.75 - y
        c.card(x0, y, cw, h, framed=True)
        c.text(x0 + 0.3, y + 0.15, cw - 0.6, h - 0.95, g["about"], 13.5, C["ink"], 600, ls=1.12)
        badge = g.get("badge")
        line_w = cw - 0.6 - (2.2 if badge else 0)
        c.text(x0 + 0.3, y + h - 0.8, line_w, 0.66, g["sponsor_line"], 10.5, C["muted_on_light"], 500,
               valign="middle", ls=1.05)
        if badge:
            c.image(P(badge["logo"]), x0 + cw - 2.2, y + h - 0.66, 1.9, 0.38, align="right")
        footer(c, d, total)

        # ------------------------------------------------------------ LEADERSHIP TEAM
        c.new_slide()
        t = g["team"]
        title_pill(c, [Run(g["name"], 700, C["cyan"]), Run(" " + t["title"], 700, C["white"])])
        people = t["people"]
        n, gap = len(people), 0.22
        pw = (SLIDE_W - 2 * MARGIN - gap * (n - 1)) / n
        for i, p in enumerate(people):
            x, y, h = MARGIN + i * (pw + gap), 1.8, 3.3
            c.card(x, y, pw, h, framed=True)
            dia = 1.4
            avatar(c, P(p["photo"]), x + (pw - dia) / 2, y + 0.3, dia)
            c.text(x + 0.12, y + 1.85, pw - 0.24, 0.72, p["name"], 15, C["black"], 700, "center", "middle", ls=1.0)
            if p.get("hosting"):
                c.pill(x + pw / 2, y + 2.72, 0.36, t["host_badge"], C["blue"], C["white"], 10, 700, align="center")
        cs = g["call_for_speakers"]
        y, h = 5.5, 1.3
        c.card(MARGIN, y, SLIDE_W - 2 * MARGIN, h, framed=True)
        q = 1.1
        c.qr(cs["qr_name"], cs["url"], q, MARGIN + 0.12, y + (h - q) / 2)
        c.text(MARGIN + 1.45, y + 0.14, 9.0, 0.45, cs["title"], 17, C["black"], 700, valign="middle")
        c.text(MARGIN + 1.45, y + 0.58, 10.4, 0.6, cs["text"], 12, C["muted_on_light"], 500, ls=1.08)
        footer(c, d, total)

    # ------------------------------------------------------------------- SPONSORS & PARTNERS
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
            logo_fit(c, P(it["logo"]), x + 0.3, 2.45, w - 0.6, 1.75, area=1.25 * it.get("scale", 1.0))
            c.rect(x + 0.3, 4.4, w - 0.6, 0.012, C["grey_line"], radius=0)
            c.text(x + 0.2, 4.5, w - 0.4, 0.62, it["name"], 15, C["black"], 700, "center", "middle")
    c.rect(MARGIN + sec_w + sec_gap / 2 - 0.01, 1.8, 0.02, 3.5, C["black"], radius=0)
    c.text(MARGIN + 0.4, 5.6, SLIDE_W - 2 * MARGIN - 0.8, 1.15, sp["thanks"], 16, C["black"], 600, "center", "middle", ls=1.12)
    footer(c, d, total)

    # ----------------------------------------------------------- KCD (KCD Porto colours)
    k = d["kcd"]
    K = k["colors"]
    c.new_slide(bg=K["bg"])
    kicker_pill(c, MARGIN, 0.55, k["title"], K["title"], C["white"], size=17)
    c.text(MARGIN + c.m.width_in(k["title"], 17, 700) + 0.85, 0.55, 6.5, 0.6, k["warmup_note"], 14, K["ink"], 600,
           valign="middle")
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
    footer(c, d, total, color=K["ink"])

    # ---------------------------------------------------------------------------- AGENDA
    c.new_slide()
    ag = d["agenda"]
    title_pill(c, [Run(ag["title"], 700, C["white"])], size=24)
    items = ag["items"]
    short_h, talk_h = 0.58, 1.42
    body_h = sum(talk_h if it.get("speaker") else short_h for it in items)
    top, pad = 1.6, 0.22
    framed(c, MARGIN, top, SLIDE_W - 2 * MARGIN, body_h + 2 * pad, C["white"], C["black"], C["black"])
    y = top + pad
    tx = 2.55
    for i, it in enumerate(items):
        sp_ = it.get("speaker")
        h = talk_h if sp_ else short_h
        c.pill(MARGIN + 0.3, y + (0.13 if sp_ else (h - 0.44) / 2), 0.44, it["time"], C["blue"], C["white"], 13, 700, w=1.3)
        if sp_:
            dia = 1.05
            avatar(c, P(sp_["photo"]), tx, y + (h - dia) / 2, dia)
            c.text(tx + 1.3, y + 0.06, 12.73 - tx - 1.6, 0.8, it["what"], 16, C["black"], 700, ls=1.02, valign="middle")
            nw = c.pill(tx + 1.3, y + 0.92, 0.38, sp_["name"], C["blue"], C["white"], 12, 700)
            c.pill(tx + 1.3 + nw + 0.15, y + 0.92, 0.38, sp_["role"], C["grey"], C["black"], 10.5, 500)
        else:
            c.text(tx, y, 12.73 - tx - 0.3, h, it["what"], 16, C["black"], 700, valign="middle")
        if i < len(items) - 1:
            c.rect(MARGIN + 0.3, y + h - 0.006, SLIDE_W - 2 * MARGIN - 0.6, 0.012, C["grey_line"], radius=0)
        y += h
    footer(c, d, total)

    # ---------------------------------------------- SPONSOR SLIDES (last one = hand-over to the host)
    for ss in d["sponsor_slides"]:
        c.new_slide()
        c.text(MARGIN, 0.75, SLIDE_W - 2 * MARGIN, 2.0, ss["text"], 38, C["black"], 700, "center", "middle", ls=1.08)
        cw_, ch_ = 6.6, 2.75
        c.card((SLIDE_W - cw_) / 2, 3.05, cw_, ch_, framed=True)
        logo_fit(c, P(ss["logo"]), (SLIDE_W - cw_) / 2 + 0.6, 3.4, cw_ - 1.2, ch_ - 0.7, area=5.2 * ss.get("scale", 1.0))
        if ss.get("pill"):
            c.pill(SLIDE_W / 2, 6.2, 0.5, ss["pill"], C["black"], C["white"], 14, 700, align="center")
        footer(c, d, total)

    assert c.n == total, f"expected {total} slides, drew {c.n}"


# ----------------------------------------------------------------------------- pipeline

def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    mdir = Path(sys.argv[1]).resolve()
    repo = mdir.parent.parent
    deck = json.loads((mdir / "deck.json").read_text(encoding="utf-8"))
    slug = deck["slug"]
    ensure_background(deck, mdir)
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
