#!/usr/bin/env python3
"""Prepare the image assets for the generic synvert xgeeks company deck from Sources/.

Re-runnable: everything under assets/ (logos, icons, photos, derived) is pulled
from Sources/xgeeks presentation/ so it can be regenerated from scratch.
"""
from pathlib import Path
import io, subprocess, sys, zipfile
from PIL import Image, ImageChops

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
SRC = ROOT / "Sources" / "xgeeks presentation"
A = HERE / "assets"
for d in ("logos", "icons", "photos", "derived"):
    (A / d).mkdir(parents=True, exist_ok=True)

VTT = zipfile.ZipFile(SRC / "previous-deck" / "#9 VTT PTT.pptx")
ALLHANDS = zipfile.ZipFile(SRC / "previous-deck" / "2026 38 xgeeks all hands (leads and clients).pptx")


def media(zf, name):
    return Image.open(io.BytesIO(zf.read(f"ppt/media/{name}")))


def trim(im, pad=0):
    bbox = im.getchannel("A").getbbox() if im.mode == "RGBA" else im.getbbox()
    if not bbox:
        return im
    l, t, r, b = bbox
    return im.crop((max(0, l - pad), max(0, t - pad), min(im.width, r + pad), min(im.height, b + pad)))


def whiten_to_alpha(im, thresh=235):
    im = im.convert("RGBA")
    px = im.load()
    for y in range(im.height):
        for x in range(im.width):
            r, g, b, a = px[x, y]
            if r > thresh and g > thresh and b > thresh:
                px[x, y] = (r, g, b, 0)
    return im


def save_rgb(im, path, quality=92):
    im.convert("RGB").save(path, quality=quality)


# --- Brand marks --------------------------------------------------------------
media(VTT, "image1.png").save(A / "logos" / "synvert-xgeeks-navy.png")     # for use on light backgrounds
media(VTT, "image14.png").save(A / "logos" / "synvert-xgeeks-white.png")   # for use on dark backgrounds
Image.open(SRC / "logos" / "GlobalLogic-logo.jpg").convert("RGB").save(A / "logos" / "globallogic.png")

# --- Decorative line-art icons reused from the previous deck -----------------
media(VTT, "image15.png").save(A / "icons" / "rocket-laptop.png")
media(VTT, "image16.png").save(A / "icons" / "lightbulb-planet.png")
media(VTT, "image25.png").save(A / "icons" / "hand-planet.png")
media(VTT, "image31.png").save(A / "icons" / "astronaut-helmet.png")
media(VTT, "image9.png").save(A / "icons" / "binoculars.png")

# --- Team photo for About xgeeks ----------------------------------------------
save_rgb(media(VTT, "image30.jpeg"), A / "photos" / "team.jpg")

# --- Client logos (slide 7): existing clients.yaml + the leads/clients funnel -
CLIENT_MAP = [
    (VTT, "image17.png", "mercedes-benz.png"),
    (VTT, "image18.jpeg", "lichtblick.png"),
    (VTT, "image19.jpeg", "telekom.png"),
    (VTT, "image20.png", "putzmeister.png"),
    (VTT, "image21.jpeg", "tblx.png"),
    (VTT, "image22.png", "coverflex.png"),
    (VTT, "image23.png", "datacamp.png"),
    (VTT, "image24.png", "mbition.png"),
    (VTT, "image26.png", "american-tower.png"),
    (VTT, "image28.jpeg", "bmw.png"),
    (ALLHANDS, "image13.jpeg", "salsify.png"),
    (ALLHANDS, "image14.png", "siemens.png"),
    (ALLHANDS, "image24.png", "die-continentale.png"),
    (ALLHANDS, "image28.jpeg", "omip.png"),
    (ALLHANDS, "image16.png", "hitachi.png"),
    (ALLHANDS, "image30.png", "helvetia.png"),
]
for zf, src_name, out_name in CLIENT_MAP:
    im = media(zf, src_name)
    if im.mode != "RGBA":
        im = whiten_to_alpha(im.convert("RGB"), 245)
    trim(im, pad=6).save(A / "logos" / out_name)

# CJR Renewables is a vector group in the "leads" slide (not a plain embedded
# picture) -> render that slide at high DPI with LibreOffice and crop it.
cjr_pdf = A / "derived" / "_all_hands_leads.pdf"
soffice = "/opt/homebrew/bin/soffice"
subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(A / "derived"),
                str(SRC / "previous-deck" / "2026 38 xgeeks all hands (leads and clients).pptx")],
               check=True, capture_output=True, timeout=180)
made_pdf = A / "derived" / "2026 38 xgeeks all hands (leads and clients).pdf"
if made_pdf.exists():
    made_pdf.rename(cjr_pdf)
import pymupdf
doc = pymupdf.open(cjr_pdf)
pix = doc[0].get_pixmap(dpi=600)
pix.save(A / "derived" / "_all_hands_leads_hires.png")
hires = Image.open(A / "derived" / "_all_hands_leads_hires.png").convert("RGB")
sc = hires.width / 26.664931102362203  # px per inch at this render
x, y, w, h = 20.68 * sc, 7.25 * sc, 2.20 * sc, 0.75 * sc
pad = 0.05 * sc
hires.crop((int(x - pad), int(y - pad), int(x + w + pad), int(y + h + pad))).save(A / "logos" / "cjr-renewables.png")
cjr_pdf.unlink(missing_ok=True)
(A / "derived" / "_all_hands_leads_hires.png").unlink(missing_ok=True)

# --- Community logos (slide 10): 6 from the deck + 3 cropped from the screenshots
COMMUNITY_MAP = [
    (VTT, "image35.png", "viseu-tech-talks.png"),
    (VTT, "image34.png", "leiria-tech-talks.png"),
    (VTT, "image33.png", "lisbon-genai-community.png"),
    (VTT, "image37.png", "aws-user-group-porto.png"),
    (VTT, "image39.png", "cloud-native-lisbon.png"),
    (VTT, "image38.png", "cloud-native-porto.png"),
]
for zf, src_name, out_name in COMMUNITY_MAP:
    im = media(zf, src_name)
    if im.mode != "RGBA":
        im = whiten_to_alpha(im.convert("RGB"), 245)
    trim(im, pad=6).save(A / "logos" / out_name)

# The 3 newer communities only exist as small dashboard thumbnails in the
# "communities-list-2" screenshot -> crop + upscale them.
shot2 = Image.open(SRC / "communities" / "communities-list-2.png").convert("RGB")
SHOT2_BOXES = {
    "devops-braga.png": (108, 105, 272, 205),
    "leaddev-meetup-lisbon.png": (797, 116, 965, 200),
    "devops-lisbon.png": (1462, 118, 1655, 222),
}
for out_name, box in SHOT2_BOXES.items():
    crop = shot2.crop(box)
    crop = crop.resize((crop.width * 5, crop.height * 5), Image.LANCZOS)
    crop.save(A / "logos" / out_name)

# --- Tech Talks format logo (VTT has one; PTT has none in Sources -> left out,
# the deck builder draws an explicit placeholder slot for it instead) ---------
media(VTT, "image35.png").save(A / "logos" / "viseu-tech-talks-format.png")

print("assets prepared ->", A)
