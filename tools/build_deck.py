#!/usr/bin/env python3
"""
Meetup deck builder: one layout definition -> PPTX + PDF + self-contained web deck.

Usage:  python3 tools/build_deck.py "<Community> Presentations/<YYYY-MM-DD - Meetup Name>"

Reads <meetup dir>/deck.json, generates QR codes into assets/qr/, builds
  <meetup dir>/<slug>.pptx   (python-pptx, 16:9, real text boxes)
  <meetup dir>/<slug>.pdf    (LibreOffice headless; page count verified)
  <meetup dir>/web/index.html (single self-contained file, keyboard navigation)
  <meetup dir>/preview/*.png  (rendered PDF pages + optional Chrome screenshots of the web deck)
and publishes the web deck to <repo>/docs/<slug>/ (GitHub Pages) with an index page.

Every drawing call goes through a Canvas that fans out to a PPTX backend and an
HTML backend, so geometry is computed once and both deliverables match.
"""
from __future__ import annotations
import base64, json, hashlib, os, re, shutil, subprocess, sys, textwrap, html as htmlmod
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import qrcode
from qrcode.constants import ERROR_CORRECT_Q
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

SLIDE_W, SLIDE_H = 13.333, 7.5          # inches (16:9)
PX = 120                                 # web design px per inch -> 1600 x 900 stage
TOOLS_DIR = Path(__file__).resolve().parent
FONT_DIRS = [TOOLS_DIR / "fonts", Path.home() / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts")]

# ----------------------------------------------------------------------------- helpers

def find_font_file(family: str, weight: int) -> Path | None:
    names = {400: "Regular", 500: "Medium", 600: "SemiBold", 700: "Bold", 800: "ExtraBold"}
    fam = family.replace(" ", "")
    for d in FONT_DIRS:
        for sub in (d / fam.lower(), d / fam, d):
            for ext in ("ttf", "otf"):
                p = sub / f"{fam}-{names[weight]}.{ext}"
                if p.exists():
                    return p
    return None


class Measurer:
    """Text measurement with the real font files (falls back to a heuristic)."""
    def __init__(self, family):
        self.family = family
        self.cache = {}

    def font(self, size_pt, weight):
        key = (round(size_pt * 4), weight)
        if key not in self.cache:
            f = find_font_file(self.family, weight)
            self.cache[key] = ImageFont.truetype(str(f), max(1, int(round(size_pt * 4)))) if f else None
        return self.cache[key]

    def width_in(self, text, size_pt, weight=400) -> float:
        f = self.font(size_pt, weight)
        if f is None:
            return len(text) * size_pt * 0.55 / 72
        return f.getlength(text) / 4 / 72          # PIL size is in px @72dpi -> inches

    def wrap(self, text, size_pt, weight, width_in) -> list[str]:
        lines = []
        for para in text.split("\n"):
            words, cur = para.split(" "), ""
            for w in words:
                trial = (cur + " " + w).strip()
                if self.width_in(trial, size_pt, weight) <= width_in or not cur:
                    cur = trial
                else:
                    lines.append(cur); cur = w
            lines.append(cur)
        return lines


def rgb(hexstr):
    h = hexstr.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def data_uri(path: Path) -> str:
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "svg": "image/svg+xml"}[path.suffix.lower()[1:]]
    return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode()


def contain_box(img_w, img_h, x, y, w, h, align="center"):
    s = min(w / img_w, h / img_h)
    dw, dh = img_w * s, img_h * s
    dx = x + (w - dw) / 2 if align in ("center", "top", "bottom") else (x if align == "left" else x + w - dw)
    dy = y + (h - dh) / 2 if align in ("center", "left", "right") else (y if align == "top" else y + h - dh)
    return dx, dy, dw, dh


# ----------------------------------------------------------------------------- run/paragraph model

class Run:
    def __init__(self, text, weight=None, color=None, size=None, italic=False):
        self.text, self.weight, self.color, self.size, self.italic = text, weight, color, size, italic


def norm_paras(paras):
    """str | list[str | list[Run|str]] -> list[list[Run]]"""
    if isinstance(paras, str):
        paras = paras.split("\n")
    out = []
    for p in paras:
        if isinstance(p, str):
            out.append([Run(p)])
        elif isinstance(p, Run):
            out.append([p])
        else:
            out.append([r if isinstance(r, Run) else Run(r) for r in p])
    return out


# ----------------------------------------------------------------------------- backends

class PptxBackend:
    def __init__(self, brand):
        self.brand = brand
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
        self.slide = None

    def new_slide(self):
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def _font_for(self, weight):
        fam = self.brand["font"]
        return {400: (fam, False), 500: (f"{fam} Medium", False), 600: (f"{fam} SemiBold", False), 700: (fam, True)}[weight]

    def rect(self, x, y, w, h, fill, radius=0, line=None, line_w=0, link=None):
        shp = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                          Inches(x), Inches(y), Inches(w), Inches(h))
        if radius:
            shp.adjustments[0] = min(0.5, radius / min(w, h))
        shp.shadow.inherit = False
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor(*rgb(fill))
        if line:
            shp.line.color.rgb = RGBColor(*rgb(line)); shp.line.width = Pt(line_w)
        else:
            shp.line.fill.background()
        if link:
            shp.click_action.hyperlink.address = link
        shp.text_frame.text = ""  # keep empty
        return shp

    def ellipse(self, x, y, w, h, fill):
        shp = self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.shadow.inherit = False
        shp.fill.solid(); shp.fill.fore_color.rgb = RGBColor(*rgb(fill)); shp.line.fill.background()

    def text(self, x, y, w, h, paras, size, color, weight, align, valign, ls, para_space, link, letter_spacing=0):
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[valign]
        for i, runs in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
            p.line_spacing = ls
            if para_space:
                p.space_after = Pt(para_space)
            for run in runs:
                r = p.add_run()
                r.text = run.text
                fam, bold = self._font_for(run.weight or weight)
                r.font.name = fam
                r.font.bold = bold
                r.font.italic = run.italic
                r.font.size = Pt(run.size or size)
                r.font.color.rgb = RGBColor(*rgb(run.color or color))
                if letter_spacing:
                    r.font._element.set("spc", str(int(letter_spacing * 100)))
        if link:
            tb.click_action.hyperlink.address = link

    def image(self, path, x, y, w, h, link=None):
        pic = self.slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
        if link:
            pic.click_action.hyperlink.address = link

    def save(self, path):
        self.prs.save(str(path))


class HtmlBackend:
    def __init__(self, brand, meta):
        self.brand, self.meta = brand, meta
        self.slides: list[list[str]] = []
        self.uri_cache = {}
        self.bg_uri = None

    def _uri(self, path):
        key = str(path)
        if key not in self.uri_cache:
            self.uri_cache[key] = data_uri(Path(path))
        return self.uri_cache[key]

    def new_slide(self):
        self.slides.append([])

    @staticmethod
    def px(v):
        return f"{v * PX:.1f}px"

    def _pos(self, x, y, w, h):
        return f"left:{self.px(x)};top:{self.px(y)};width:{self.px(w)};height:{self.px(h)};"

    def rect(self, x, y, w, h, fill, radius=0, line=None, line_w=0, link=None):
        st = self._pos(x, y, w, h) + f"border-radius:{self.px(radius)};"
        st += f"background:{fill};" if fill else ""
        if line:
            st += f"border:{line_w * PX / 72:.1f}px solid {line};box-sizing:border-box;"
        el = f'<div class="r" style="{st}"></div>'
        self.slides[-1].append(self._wrap_link(el, link))

    def ellipse(self, x, y, w, h, fill):
        self.slides[-1].append(f'<div class="r" style="{self._pos(x, y, w, h)}border-radius:50%;background:{fill};"></div>')

    def text(self, x, y, w, h, paras, size, color, weight, align, valign, ls, para_space, link, letter_spacing=0):
        jc = {"top": "flex-start", "middle": "center", "bottom": "flex-end"}[valign]
        st = self._pos(x, y, w, h) + f"font-size:{size * PX / 72:.1f}px;color:{color};font-weight:{weight};text-align:{align};" \
             f"justify-content:{jc};line-height:{ls * 1.4:.2f};"
        if letter_spacing:
            st += f"letter-spacing:{letter_spacing * PX / 72:.1f}px;"
        ps = []
        for runs in paras:
            spans = []
            for r in runs:
                rs = ""
                if r.weight: rs += f"font-weight:{r.weight};"
                if r.color: rs += f"color:{r.color};"
                if r.size: rs += f"font-size:{r.size * PX / 72:.1f}px;"
                if r.italic: rs += "font-style:italic;"
                t = htmlmod.escape(r.text)
                spans.append(f'<span style="{rs}">{t}</span>' if rs else t)
            pst = f' style="margin-bottom:{para_space * PX / 72:.1f}px"' if para_space else ""
            ps.append(f"<p{pst}>{''.join(spans) or '&nbsp;'}</p>")
        el = f'<div class="t" style="{st}">{"".join(ps)}</div>'
        self.slides[-1].append(self._wrap_link(el, link))

    def image(self, path, x, y, w, h, link=None):
        el = f'<img class="i" style="{self._pos(x, y, w, h)}" src="{self._uri(path)}" alt="">'
        self.slides[-1].append(self._wrap_link(el, link))

    def _wrap_link(self, el, link):
        return f'<a href="{htmlmod.escape(link)}" target="_blank" rel="noopener">{el}</a>' if link else el

    def save(self, path, bg_path):
        b = self.brand
        font = b["font"]
        gf = font.replace(" ", "+")
        n = len(self.slides)
        slides_html = "\n".join(
            f'<section class="slide" data-n="{i + 1}" aria-label="Slide {i + 1}">{"".join(els)}</section>'
            for i, els in enumerate(self.slides))
        doc = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{htmlmod.escape(self.meta["title"])}</title>
<meta name="description" content="{htmlmod.escape(self.meta.get("description", ""))}">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family={gf}:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{{--purple:{b["colors"]["purple"]};--navy:{b["colors"]["navy"]};}}
*{{box-sizing:border-box}}
html,body{{margin:0;height:100%;background:var(--navy);overflow:hidden;font-family:"{font}","Segoe UI",Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased}}
#viewport{{position:fixed;inset:0;display:flex;align-items:center;justify-content:center}}
#stage{{position:relative;width:1600px;height:900px;transform-origin:center center;overflow:hidden;border-radius:6px;box-shadow:0 20px 60px rgba(0,0,0,.45)}}
.slide{{position:absolute;inset:0;background:url("{self._uri(bg_path)}") center/cover no-repeat;opacity:0;visibility:hidden;transition:opacity .35s ease}}
.slide.active{{opacity:1;visibility:visible}}
.r,.t,.i{{position:absolute;display:block}}
.t{{display:flex;flex-direction:column;overflow:hidden;white-space:pre-wrap;overflow-wrap:break-word}}
.t p{{margin:0}}
.i{{object-fit:fill}}
a{{color:inherit;text-decoration:none}}
a .r,a .t,a .i{{cursor:pointer}}
#bar{{position:fixed;left:0;top:0;height:4px;background:#fff;opacity:.85;transition:width .3s;z-index:9}}
#nav{{position:fixed;right:14px;top:12px;color:#fff;font-size:13px;opacity:.75;z-index:9;display:flex;gap:10px;align-items:center;user-select:none}}
#nav button{{all:unset;cursor:pointer;padding:4px 10px;border:1px solid rgba(255,255,255,.5);border-radius:999px;font:inherit;color:#fff}}
#nav button:hover{{background:rgba(255,255,255,.15)}}
#hint{{position:fixed;left:50%;bottom:18px;transform:translateX(-50%);color:#fff;background:rgba(0,0,0,.45);padding:8px 14px;border-radius:999px;font-size:13px;z-index:9;transition:opacity .6s}}
@media print{{@page{{size:1600px 900px;margin:0}}html,body{{overflow:visible;background:#fff;width:1600px}}#viewport{{position:static;display:block}}#stage{{transform:none!important;box-shadow:none;border-radius:0;height:auto;overflow:visible}}.slide{{position:relative;opacity:1;visibility:visible;height:900px;width:1600px;break-after:page;page-break-after:always;transition:none}}#bar,#nav,#hint{{display:none}}}}
</style></head>
<body>
<div id="bar"></div>
<div id="viewport"><div id="stage">
{slides_html}
</div></div>
<div id="nav"><button id="prev" aria-label="Previous">&larr;</button><span id="count">1 / {n}</span><button id="next" aria-label="Next">&rarr;</button><button id="fs" aria-label="Fullscreen">&#x26F6;</button></div>
<div id="hint">Use &larr; &rarr; keys, swipe, or click the arrows</div>
<script>
(function(){{
  const slides=[...document.querySelectorAll('.slide')],N=slides.length,stage=document.getElementById('stage');
  let i=0;
  function fit(){{const s=Math.min(innerWidth/1600,innerHeight/900);stage.style.transform='scale('+s+')';}}
  function show(k,push=true){{i=Math.max(0,Math.min(N-1,k));slides.forEach((s,j)=>s.classList.toggle('active',j===i));
    document.getElementById('count').textContent=(i+1)+' / '+N;document.getElementById('bar').style.width=((i+1)/N*100)+'%';
    if(push)history.replaceState(null,'','#'+(i+1));}}
  function fromHash(){{const h=parseInt(location.hash.slice(1),10);show(isNaN(h)?0:h-1,false);}}
  addEventListener('resize',fit);addEventListener('hashchange',fromHash);
  addEventListener('keydown',e=>{{if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(e.key)){{e.preventDefault();show(i+1);}}
    else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(e.key)){{e.preventDefault();show(i-1);}}
    else if(e.key==='Home')show(0);else if(e.key==='End')show(N-1);else if(e.key==='f')toggleFs();}});
  let tx=null;addEventListener('touchstart',e=>tx=e.touches[0].clientX,{{passive:true}});
  addEventListener('touchend',e=>{{if(tx===null)return;const dx=e.changedTouches[0].clientX-tx;if(Math.abs(dx)>50)show(i+(dx<0?1:-1));tx=null;}});
  stage.addEventListener('click',e=>{{if(e.target.closest('a'))return;const r=stage.getBoundingClientRect();show(i+((e.clientX-r.left)/r.width<0.3?-1:1));}});
  document.getElementById('prev').onclick=()=>show(i-1);document.getElementById('next').onclick=()=>show(i+1);
  function toggleFs(){{document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen();}}
  document.getElementById('fs').onclick=toggleFs;
  setTimeout(()=>document.getElementById('hint').style.opacity=0,4000);
  fit();fromHash();
}})();
</script>
</body></html>"""
        Path(path).write_text(doc, encoding="utf-8")


# ----------------------------------------------------------------------------- canvas (fans out)

class Canvas:
    def __init__(self, deck, mdir: Path):
        self.deck, self.mdir = deck, mdir
        self.brand = deck["brand"]
        self.C = self.brand["colors"]
        self.m = Measurer(self.brand["font"])
        self.pptx = PptxBackend(self.brand)
        self.html = HtmlBackend(self.brand, {"title": deck["title"], "description": deck.get("description", "")})
        self.backends = [self.pptx, self.html]
        self.derived = mdir / "assets" / "derived"
        self.derived.mkdir(parents=True, exist_ok=True)
        self.bg_path = mdir / self.brand["background"]
        self.n = 0
        self.warnings = []

    # -- primitives
    def new_slide(self, bg=True):
        self.n += 1
        for b in self.backends:
            b.new_slide()
        if bg:
            self.image(self.bg_path, 0, 0, SLIDE_W, SLIDE_H, fit="fill")

    def rect(self, x, y, w, h, fill="#FFFFFF", radius=0.2, line=None, line_w=0, shadow=False, link=None):
        if shadow:
            d = self.brand.get("shadow_offset", 0.09)
            for b in self.backends:
                b.rect(x + d, y + d, w, h, self.C["black"], radius)
        for b in self.backends:
            b.rect(x, y, w, h, fill, radius, line, line_w, link)

    def card(self, x, y, w, h, framed=False, link=None):
        """White card. framed=True -> black outline + hard shadow (banner's hero card style)."""
        self.rect(x, y, w, h, self.C["white"], 0.2, self.C["black"] if framed else None, 2.5 if framed else 0, shadow=framed, link=link)

    def ellipse(self, x, y, w, h, fill):
        for b in self.backends:
            b.ellipse(x, y, w, h, fill)

    def text(self, x, y, w, h, paras, size=14, color=None, weight=400, align="left", valign="top", ls=1.0,
             para_space=0, link=None, letter_spacing=0, check=True):
        color = color or self.C["black"]
        P = norm_paras(paras)
        if check:
            self._check_overflow(P, size, weight, w, h, ls, para_space, x, y)
        for b in self.backends:
            b.text(x, y, w, h, P, size, color, weight, align, valign, ls, para_space, link, letter_spacing)

    def _check_overflow(self, P, size, weight, w, h, ls, para_space, x, y):
        lines = 0
        for runs in P:
            t = "".join(r.text for r in runs)
            wt = max([r.weight or weight for r in runs] + [weight])
            sz = max([r.size or size for r in runs] + [size])
            lines += len(self.m.wrap(t, sz, wt, w))
        need = lines * size * ls * 1.4 / 72 + (len(P) - 1) * para_space / 72
        if need > h + 0.02:
            snippet = "".join(r.text for r in P[0])[:40]
            self.warnings.append(f"slide {self.n}: text may overflow ({need:.2f}in > {h:.2f}in) at ({x:.2f},{y:.2f}) '{snippet}'")
        for runs in P:
            t = "".join(r.text for r in runs)
            for word in t.split():
                if self.m.width_in(word, size, weight) > w:
                    self.warnings.append(f"slide {self.n}: word '{word}' wider than box {w:.2f}in")

    def image(self, path, x, y, w, h, fit="contain", radius=0, align="center", link=None):
        path = Path(path)
        if not path.is_absolute():
            path = self.mdir / path
        im = Image.open(path)
        if fit == "contain":
            x, y, w, h = contain_box(im.width, im.height, x, y, w, h, align)
            src = path
        elif fit == "cover":
            src = self._derived_cover(path, im, w, h, radius)
        else:
            src = path
        for b in self.backends:
            b.image(src, x, y, w, h, link)
        return x, y, w, h

    def _derived_cover(self, path, im, w, h, radius):
        """Crop to the box aspect ratio (cover) and round corners; cached under assets/derived."""
        dpi = 220
        tw, th = int(w * dpi), int(h * dpi)
        key = hashlib.md5(f"{path.name}{tw}{th}{radius}{path.stat().st_size}".encode()).hexdigest()[:10]
        out = self.derived / f"{path.stem}-{key}.png"
        if not out.exists():
            im = im.convert("RGB")
            s = max(tw / im.width, th / im.height)
            im = im.resize((max(tw, int(im.width * s + 0.5)), max(th, int(im.height * s + 0.5))), Image.LANCZOS)
            l, t = (im.width - tw) // 2, (im.height - th) // 2
            im = im.crop((l, t, l + tw, t + th)).convert("RGBA")
            if radius:
                r = int(radius * dpi)
                mask = Image.new("L", (tw, th), 0)
                ImageDraw.Draw(mask).rounded_rectangle((0, 0, tw - 1, th - 1), r, fill=255)
                im.putalpha(mask)
            im.save(out)
        return out

    def pill(self, x, y, h, text, bg, fg, size=12, weight=600, pad=None, w=None, align="left", link=None):
        pad = pad if pad is not None else h * 0.55
        tw = self.m.width_in(text, size, weight)
        w = w or (tw + 2 * pad)
        if align == "center":
            x = x - w / 2
        elif align == "right":
            x = x - w
        self.rect(x, y, w, h, bg, radius=h / 2, link=link)
        self.text(x, y, w, h, text, size, fg, weight, "center", "middle", 1.0, check=False)
        return w

    def label(self, x, y, w, text, color=None):
        """Small letter-spaced section label."""
        self.text(x, y, w, 0.3, text.upper(), 11, color or self.C["white"], 700, letter_spacing=1.5, valign="middle")

    def title_card(self, title, size=30):
        w = self.m.width_in(title, size, 700) + 0.8
        self.card(0.6, 0.45, w, 0.95, framed=True)
        self.text(0.95, 0.45, w - 0.6, 0.95, title, size, self.C["black"], 700, valign="middle")

    def footer(self, total):
        f = self.deck.get("footer", "")
        self.text(0.6, 7.05, 9, 0.3, f, 10.5, self.C["muted_on_dark"], 500, valign="middle")
        self.text(11.6, 7.05, 1.13, 0.3, f"{self.n} / {total}", 10.5, self.C["muted_on_dark"], 500, "right", "middle")

    def initials_avatar(self, cx, cy, d, name):
        self.ellipse(cx - d / 2, cy - d / 2, d, d, self.C["deep"])
        parts = [p for p in name.split() if p]
        ini = (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()
        self.text(cx - d / 2, cy - d / 2, d, d, ini, d * 22, self.C["white"], 700, "center", "middle", check=False)

    def qr(self, name, url, size_in, x, y, link=True):
        """Generate (once) and place a QR code. Files live in assets/qr so they can be regenerated."""
        qdir = self.mdir / "assets" / "qr"
        qdir.mkdir(parents=True, exist_ok=True)
        out = qdir / f"{name}.png"
        q = qrcode.QRCode(error_correction=ERROR_CORRECT_Q, box_size=16, border=4)   # border = quiet zone (4 modules)
        q.add_data(url)
        q.make(fit=True)
        q.make_image(fill_color="black", back_color="white").save(out)
        self.image(out, x, y, size_in, size_in, fit="fill", link=url if link else None)
        return out

    # -- output
    def save(self, pptx_path, html_path):
        self.pptx.save(pptx_path)
        self.html.save(html_path, self.bg_path)


# ----------------------------------------------------------------------------- slides

def draw_deck(c: Canvas, d: dict):
    C = c.C
    total = 9

    # 1. COVER (banner composition: framed hero card left, speaker cards right)
    c.new_slide()
    c.card(0.55, 0.5, 4.95, 5.2, framed=True)
    logos = d["cover"]["logos"]
    slot_w = (4.95 - 0.5) / len(logos)
    for i, lg in enumerate(logos):
        c.image(lg["file"], 0.8 + i * slot_w + 0.08, 0.72, slot_w - 0.16, 0.66)
    c.rect(0.55, 1.58, 4.95, 0.035, C["black"], radius=0)
    c.image(d["community_logo"], 0.9, 1.85, 4.25, 2.25)
    c.text(0.85, 4.2, 4.4, 1.4, [
        [Run(d["edition_label"], 700)],
        [Run(d["cover"]["title_line_1"], 700)],
        [Run(d["cover"]["title_line_2"], 700)],
    ], 25, C["black"], 700, ls=0.95)
    c.card(0.55, 5.95, 4.95, 1.05, framed=True)
    c.text(0.7, 6.03, 4.65, 0.5, d["cover"]["date_line"], 19, C["black"], 700, "center", "middle")
    c.text(0.7, 6.48, 4.65, 0.47, d["cover"]["venue_line"], 11, C["black"], 500, "center", "middle", ls=0.95)
    spk = d["speakers"]
    for i, s in enumerate(spk[:2]):
        y = 0.5 + i * 3.4
        c.card(5.85, y, 6.95, 3.05)
        c.image(s["photo"], 6.1, y + 0.25, 2.05, 2.55, fit="cover", radius=0.12)
        c.text(8.4, y + 0.28, 4.2, 1.55, s["talk"], 21, C["black"], 700, ls=1.05)
        nw = c.pill(8.4, y + 1.9, 0.42, s["name"], C["deep"], C["white"], 13, 700)
        c.pill(8.4, y + 2.4, 0.4, s["role"], C["grey"], C["black"], 10.5, 500)

    # 2. ABOUT
    c.new_slide()
    a = d["about"]
    c.title_card(a["title"])
    c.card(0.6, 1.75, 5.2, 4.95)
    c.image(d["community_logo"], 0.85, 2.0, 4.7, 2.5)
    c.text(0.95, 4.65, 4.5, 1.95, [[Run(a["tagline"], 600, C["black"], 15)], [Run(a["quote"], 400, C["muted_on_light"], 12.5, italic=True)]],
           15, C["black"], 400, ls=1.05, para_space=8)
    c.text(6.3, 1.8, 4.55, 4.9, [f"•  {b}" for b in a["bullets"]], 18, C["white"], 500, ls=1.05, para_space=10, valign="middle")
    c.card(11.05, 1.75, 1.7, 2.25)
    c.qr("meetup-group", a["join_url"], 1.4, 11.2, 1.85)
    c.text(11.05, 3.3, 1.7, 0.6, a["join_label"], 10.5, C["black"], 700, "center", "middle", ls=1.0)
    c.footer(total)

    # 3. ORGANIZERS
    c.new_slide()
    o = d["organizers"]
    c.title_card(o["title"])
    people = o["people"]
    n = len(people)
    gap = 0.16
    cw = (12.13 - gap * (n - 1)) / n
    for i, p in enumerate(people):
        x = 0.6 + i * (cw + gap)
        y, h = 2.15, 3.55
        c.card(x, y, cw, h)
        c.initials_avatar(x + cw / 2, y + 0.95, 1.25, p["name"])
        c.text(x + 0.12, y + 1.75, cw - 0.24, 0.75, p["name"], 14.5, C["black"], 700, "center", "middle", ls=0.95)
        c.pill(x + cw / 2, y + 2.6, 0.36, p["role"], C["grey"], C["black"], 9.5, 500, align="center")
        if p.get("hosting"):
            c.pill(x + cw / 2, y + h - 0.5, 0.32, o["host_badge"], C["purple"], C["white"], 8.5, 700, align="center")
    c.text(0.6, 6.0, 12.13, 0.7, o.get("note", ""), 13, C["white"], 500, "center", "middle")
    c.footer(total)

    # 4. SPEAKERS & TALKS
    c.new_slide()
    c.title_card(d["speakers_title"])
    for i, s in enumerate(spk[:2]):
        y = 1.65 + i * 2.7
        c.card(0.6, y, 12.13, 2.5)
        c.image(s["photo"], 0.85, y + 0.25, 1.65, 2.0, fit="cover", radius=0.12)
        c.pill(2.8, y + 0.28, 0.32, f"TALK {i + 1}", C["purple"], C["white"], 9.5, 700)
        c.text(2.8, y + 0.68, 6.5, 0.75, s["talk"], 17.5, C["black"], 700, ls=1.0)
        c.text(2.8, y + 1.42, 6.6, 1.0, s["abstract"], 10, C["muted_on_light"], 400, ls=1.0)
        c.pill(9.7, y + 0.5, 0.42, s["name"], C["deep"], C["white"], 13, 700)
        c.text(9.7, y + 1.05, 2.85, 0.9, s["role"], 10.5, C["black"], 500, ls=1.05)
    c.footer(total)

    # 5. AGENDA
    c.new_slide()
    ag = d["agenda"]
    c.title_card(ag["title"])
    items = ag["items"]
    c.card(0.6, 1.65, 12.13, 5.15)
    row_h = (5.15 - 0.5) / len(items)
    for i, it in enumerate(items):
        y = 1.9 + i * row_h
        c.pill(0.95, y + 0.1, 0.44, it["time"], C["deep"], C["white"], 12.5, 700, w=1.55)
        c.text(2.8, y, 9.7, 0.5, it["what"], 16.5, C["black"], 700, valign="middle")
        if it.get("detail"):
            c.text(2.8, y + 0.48, 9.7, 0.4, it["detail"], 11, C["muted_on_light"], 400, valign="top")
        if i < len(items) - 1:
            c.rect(2.8, y + row_h - 0.06, 9.7, 0.012, C["grey"], radius=0)
    c.footer(total)

    # 6. SPONSORS + MEETUP PARTNER (distinct sections)
    c.new_slide()
    sp = d["sponsors_slide"]
    c.title_card(sp["title"])
    c.label(0.6, 1.62, 6, sp["sponsors_label"])
    sponsors = sp["sponsors"]
    sw = (7.85 - 0.25 * (len(sponsors) - 1)) / len(sponsors)
    for i, s in enumerate(sponsors):
        x = 0.6 + i * (sw + 0.25)
        c.card(x, 2.0, sw, 3.15)
        c.image(s["logo"], x + 0.35, 2.3, sw - 0.7, 1.45)
        c.pill(x + sw / 2, 3.95, 0.36, s["kind"], C["deep"], C["white"], 10, 700, align="center")
        c.text(x + 0.2, 4.4, sw - 0.4, 0.65, s["name"], 12, C["black"], 500, "center", "middle", ls=1.0)
    c.rect(8.72, 1.62, 0.014, 3.53, C["muted_on_dark"], radius=0)
    c.label(9.0, 1.62, 3.8, sp["partner_label"])
    pt = sp["partner"]
    c.card(9.0, 2.0, 3.73, 3.15)
    c.image(pt["logo"], 9.35, 2.3, 3.03, 1.45)
    c.pill(9.0 + 3.73 / 2, 3.95, 0.36, pt["kind"], C["black"], C["white"], 10, 700, align="center")
    c.text(9.2, 4.4, 3.33, 0.65, pt["name"], 12, C["black"], 500, "center", "middle", ls=1.0)
    c.text(0.6, 5.45, 12.13, 1.3, sp["thanks"], 14.5, C["white"], 500, "center", "middle", ls=1.1)
    c.footer(total)

    # 7. KCD x DEVOPS DAYS
    c.new_slide()
    k = d["kcd"]
    c.title_card(k["title"])
    c.card(0.6, 1.62, 7.95, 5.2)
    c.image(k["logo"], 0.85, 1.85, 1.85, 0.95, align="left")
    c.text(2.9, 1.82, 5.45, 0.62, k["event_name"], 15, C["black"], 700, ls=1.0, valign="middle")
    c.text(2.9, 2.46, 5.45, 0.7, [[Run(k["when"], 600, C["black"])], [Run(k["where"], 400, C["muted_on_light"])]],
           11.5, C["black"], 400, ls=1.0, para_space=1)
    c.text(0.85, 3.05, 7.45, 0.6, k["blurb"], 10.5, C["muted_on_light"], 400, ls=1.0, valign="middle")
    c.text(0.85, 3.7, 4, 0.3, k["keynotes_label"].upper(), 10, C["deep"], 700, letter_spacing=1.2, valign="middle")
    keys = k["keynotes"]
    kw = (7.45 - 0.2 * (len(keys) - 1)) / len(keys)
    for i, kn in enumerate(keys):
        x = 0.85 + i * (kw + 0.2)
        c.image(kn["photo"], x, 4.02, kw, 1.35, fit="cover", radius=0.1)
        c.text(x, 5.4, kw, 0.3, kn["name"], 11, C["black"], 700, valign="middle")
        c.text(x, 5.68, kw, 0.52, kn["role"], 8.5, C["muted_on_light"], 400, ls=1.0)
    c.pill(0.85, 6.3, 0.38, k["agenda_note"], C["purple"], C["white"], 11, 700)
    # QR + discount code
    c.card(8.8, 1.62, 3.93, 5.2, framed=True)
    c.qr("kcd-site", k["url"], 2.75, 8.8 + (3.93 - 2.75) / 2, 1.85)
    c.text(8.95, 4.62, 3.63, 0.35, k["qr_label"], 11.5, C["black"], 700, "center", "middle")
    c.text(8.95, 5.08, 3.63, 0.3, k["code_label"].upper(), 9.5, C["muted_on_light"], 700, "center", "middle", letter_spacing=1.2)
    c.pill(8.8 + 3.93 / 2, 5.42, 0.6, k["discount_code"], C["deep"], C["white"], 17, 700, align="center", w=3.3)
    c.text(8.95, 6.15, 3.63, 0.45, k["code_hint"], 9, C["muted_on_light"], 400, "center", "middle", ls=1.0)
    c.footer(total)

    # 8. RAFFLE
    c.new_slide()
    r = d["raffle"]
    c.title_card(r["title"])
    c.text(0.6, 1.9, 6.0, 2.5, r["headline"], 27, C["white"], 700, ls=1.05, valign="middle")
    c.text(0.6, 4.55, 6.0, 1.0, r["instructions"], 15.5, C["white"], 500, ls=1.1)
    c.text(0.6, 5.7, 6.0, 0.5, r["url_label"], 12, C["muted_on_dark"], 500, link=r["url"])
    c.card(7.05, 1.55, 5.68, 5.3, framed=True)
    qs = r.get("qr_size_in", 4.15)   # >= 1.97in (5 cm) required
    c.qr("raffle-form", r["url"], qs, 7.05 + (5.68 - qs) / 2, 1.55 + 0.3)
    c.text(7.2, 1.55 + 0.3 + qs + 0.12, 5.38, 0.5, r["qr_label"], 14, C["black"], 700, "center", "middle")
    c.footer(total)

    # 9. CLOSING
    c.new_slide()
    cl = d["closing"]
    c.text(0.6, 0.75, 8, 1.4, cl["title"], 58, C["white"], 700, valign="middle", ls=0.95)
    c.text(0.6, 2.15, 8.5, 1.05, cl["subtitle"], 17, C["white"], 500, ls=1.1)
    c.label(0.6, 3.4, 6, cl["follow_label"])
    links = cl["links"]
    lw = (12.13 - 0.24 * (len(links) - 1)) / len(links)
    for i, ln in enumerate(links):
        x = 0.6 + i * (lw + 0.24)
        y, h = 3.8, 2.98
        c.card(x, y, lw, h, link=ln["url"])
        c.image(ln["icon"], x + 0.22, y + 0.2, 0.55, 0.55)
        c.text(x + 0.9, y + 0.2, lw - 1.05, 0.55, ln["label"], 11.5, C["black"], 700, valign="middle", ls=0.95)
        qsz = 1.7
        c.qr(ln["qr_name"], ln["url"], qsz, x + (lw - qsz) / 2, y + 0.85)
        c.text(x + 0.12, y + 2.58, lw - 0.24, 0.3, ln["short"], 7.5, C["muted_on_light"], 500, "center", "middle", link=ln["url"])
    c.footer(total)
    assert c.n == total, f"expected {total} slides, drew {c.n}"


# ----------------------------------------------------------------------------- pipeline

CHROME = Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")


def convert_pdf_libreoffice(pptx_path: Path) -> Path:
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pptx_path.parent), str(pptx_path)],
                   check=True, capture_output=True, timeout=300)
    pdf = pptx_path.with_suffix(".pdf")
    assert pdf.exists(), "LibreOffice PDF conversion failed"
    return pdf


def print_pdf_chrome(html: Path, out: Path):
    """Print the web deck (same geometry as the PPTX) to PDF with headless Chrome; fonts come from Google Fonts."""
    assert CHROME.exists(), "Google Chrome not found for the PDF fallback"
    subprocess.run([str(CHROME), "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--virtual-time-budget=8000",
                    f"--print-to-pdf={out}", html.as_uri()], check=True, capture_output=True, timeout=300)
    assert out.exists(), "Chrome PDF print failed"


def pdf_info(pdf: Path):
    import pymupdf
    doc = pymupdf.open(pdf)
    fonts = set()
    for page in doc:
        fonts.update(f[3] for f in page.get_fonts())
    return len(doc), sorted(fonts)


def render_previews(pdf: Path, out_dir: Path, prefix="slide", dpi=110):
    import pymupdf
    out_dir.mkdir(exist_ok=True)
    for old in out_dir.glob(f"{prefix}-*.png"):
        old.unlink()
    for i, page in enumerate(pymupdf.open(pdf), 1):
        page.get_pixmap(dpi=dpi).save(out_dir / f"{prefix}-{i:02d}.png")


def screenshot_web(html: Path, out_dir: Path, n: int):
    if not CHROME.exists():
        return False
    for old in out_dir.glob("web-*.png"):
        old.unlink()
    for i in range(1, n + 1):
        subprocess.run([str(CHROME), "--headless=new", "--disable-gpu", "--hide-scrollbars", "--window-size=1600,900",
                        "--virtual-time-budget=4000", f"--screenshot={out_dir / f'web-{i:02d}.png'}", f"{html.as_uri()}#{i}"],
                       capture_output=True, timeout=120)
    return True


def publish_docs(repo: Path, mdir: Path, deck: dict, html: Path, pdf: Path):
    docs = repo / "docs"
    slug = deck["slug"]
    dest = docs / slug
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copy2(html, dest / "index.html")
    shutil.copy2(pdf, dest / f"{slug}.pdf")
    (docs / ".nojekyll").touch()
    reg_path = docs / "decks.json"
    reg = json.loads(reg_path.read_text()) if reg_path.exists() else []
    reg = [r for r in reg if r["slug"] != slug]
    reg.append({"slug": slug, "title": deck["title"], "date": deck["date_iso"], "community": deck["community"],
                "folder": str(mdir.relative_to(repo))})
    reg.sort(key=lambda r: r["date"], reverse=True)
    reg_path.write_text(json.dumps(reg, indent=2, ensure_ascii=False) + "\n")
    rows = "\n".join(
        f'<li><span class="d">{r["date"]}</span> <a href="{r["slug"]}/">{htmlmod.escape(r["title"])}</a>'
        f' <span class="c">{htmlmod.escape(r["community"])}</span> · <a class="pdf" href="{r["slug"]}/{r["slug"]}.pdf">PDF</a></li>'
        for r in reg)
    (docs / "index.html").write_text(f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Meetup Presentations</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap" rel="stylesheet">
<style>body{{margin:0;min-height:100vh;font-family:Poppins,Helvetica,Arial,sans-serif;background:linear-gradient(110deg,#201E37 0%,#4D3788 45%,#7B5AC2 100%);color:#fff;display:flex;align-items:center;justify-content:center;padding:24px;box-sizing:border-box}}
.card{{background:#fff;color:#000;border-radius:16px;border:3px solid #000;box-shadow:12px 12px 0 #000;padding:32px 36px;max-width:820px;width:100%}}
h1{{margin:0 0 6px;font-size:28px}}p{{margin:0 0 18px;color:#3A3A4A}}ul{{list-style:none;padding:0;margin:0}}li{{padding:12px 0;border-top:1px solid #E5EAEB;display:flex;gap:12px;align-items:baseline;flex-wrap:wrap}}
.d{{font-weight:600;color:#45307F;font-variant-numeric:tabular-nums}}a{{color:#000;font-weight:700;text-decoration:none}}a:hover{{text-decoration:underline}}.c{{color:#3A3A4A;font-size:13px}}.pdf{{font-weight:600;color:#7B5AC2}}</style></head>
<body><main class="card"><h1>Meetup Presentations</h1><p>Decks for the community meetups I organise. Open a deck and use the arrow keys.</p><ul>{rows}</ul></main></body></html>""")
    return dest


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    mdir = Path(sys.argv[1]).resolve()
    repo = mdir.parent.parent
    deck = json.loads((mdir / "deck.json").read_text(encoding="utf-8"))
    slug = deck["slug"]
    c = Canvas(deck, mdir)
    draw_deck(c, deck)
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
    # 1) LibreOffice export of the PPTX: structural check (page count) + PDF candidate
    pdf_path = convert_pdf_libreoffice(pptx_path)
    lo_pages, lo_fonts = pdf_info(pdf_path)
    assert lo_pages == c.n, f"LibreOffice PDF has {lo_pages} pages, expected {c.n}"
    lo_check = preview / "libreoffice-check.pdf"
    shutil.copy2(pdf_path, lo_check)
    render_previews(lo_check, preview, prefix="pptx")
    engine = "libreoffice"
    if not any(brand_font.lower() in f.lower() for f in lo_fonts):
        # LibreOffice could not see the brand font (macOS user-font service): print the web deck instead.
        engine = "chrome"
        print_pdf_chrome(html_path, pdf_path)
    pages, fonts = pdf_info(pdf_path)
    assert pages == c.n, f"PDF has {pages} pages, expected {c.n}"
    render_previews(pdf_path, preview, prefix="slide")
    print(f"PDF   {pdf_path}  pages={pages} (slides={c.n})  engine={engine}  fonts={fonts}")
    if engine == "chrome":
        print(f"NOTE  LibreOffice embedded {lo_fonts} instead of {brand_font}; the PDF was printed from the web deck (same geometry). "
              f"LibreOffice render kept for the PPTX geometry check: {lo_check}")
    if screenshot_web(html_path, preview, c.n):
        print(f"WEB screenshots -> {preview}/web-XX.png")
    if "--no-publish" not in sys.argv:
        dest = publish_docs(repo, mdir, deck, html_path, pdf_path)
        print(f"DOCS  {dest}")


if __name__ == "__main__":
    main()
