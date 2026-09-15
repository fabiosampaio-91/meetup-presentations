---
name: meetup-presentation
description: Build a community meetup deck (PPTX + PDF + web) from the files in Sources/ and publish it to GitHub Pages. Use whenever the user asks for a new meetup presentation, deck or slides for any community (AWS UG, KCD, DevOps Days, OutSystems UG, ...), OR for the generic synvert xgeeks company deck reused across every meetup (sponsors & partners, about xgeeks, GlobalLogic, what we do, clients, communities). Encodes the house procedure - sources first, banner-driven branding, fixed slide order, sponsors vs meetup partner, folder conventions, QR codes, render check, commit + push + Pages + live link.
---

# Meetup presentation

Repeatable procedure for every meetup deck in this repository. Follow it in order; do not skip the
verification or publishing steps. The deliverables are always **PPTX + PDF + web deck**.

Two deck kinds share this repo:
- **Event deck** (per meetup edition: AWS UG, KCD, OutSystems UG, ...) - built by `tools/build_deck.py`
  from a per-meetup `deck.json`. §1-9 below.
- **Company deck** (the generic, reusable synvert xgeeks deck presented before the talks at *any*
  meetup) - built by `tools/build_company_deck.py` from its own `deck.json`. See §10.

## 1. Read the sources first, derive everything from them

`Sources/` is the inbox. Before writing any code, open **every** file that is there for the new meetup:

- the previous deck (`.pptx`): dump slide texts/positions with python-pptx, convert to PDF with
  `soffice --headless --convert-to pdf`, render pages with PyMuPDF and look at them; unzip `ppt/media/`
  to recover logos, icons and QR codes (decode QR images with OpenCV `QRCodeDetector` to get the URLs);
- the banner image(s) for this edition: the branding reference (see §2);
- documents (`.docx`, notes): links to the event page, KCD site, discount code, raffle form, etc.
  Extract text with python-docx and also read `word/_rels/document.xml.rels` for hyperlinks;
- logo files and photos.

If a document links to the event page (meetup.com), fetch it: `curl -A "Mozilla/5.0" <url>` and parse the
`<script id="__NEXT_DATA__">` JSON. It carries the exact title, `dateTime`, venue, hosts, the full
description (agenda, talks, speakers, sponsors thanks) and the `featuredEventPhoto` (the real event
banner). The group page gives member count, founding date and the leadership list. Fetch the KCD/CNCF
page for dates, venue and keynotes.

Organizer photos: the group page JSON carries `memberPhoto` ids for (some) leaders; the same photo is served at
`https://secure-content.meetupstatic.com/images/classic-member/<photoId>/300x300.webp` (save it under
`Sources/<meetup>/organizers/`). The leaders page and member profiles require a login, so a screenshot of the
leadership list dropped in Sources is the fallback: crop the 80px avatar circles (see `prepare_assets.py`).

Facts to extract: meetup name and edition number, date, time, venue + address, agenda, speakers with
role and talk title, organizers and roles, community description and numbers, social links, sponsors,
partner(s), KCD event (dates, venue, keynotes, site URL, discount code), raffle form URL.

**Never invent a value.** If something is not in the sources or the linked pages, leave it out of the
deck and list it under "Open items" in the final report. Check that the banner really belongs to this
meetup (city, date, speakers); a banner from another chapter is still a valid *branding* reference but
never a *content* source.

## 2. Branding: current banner + previous deck

- Sample the palette from the banner with Pillow (quantize + point samples) and put the hex values in
  `deck.json > brand.colors`. Reference for the AWS UG design system: purple `#7B5AC2`, navy `#201E37`,
  deep purple pill `#45307F`, light grey pill `#E5EAEB`, white cards with a 2.5pt black outline and a hard
  black shadow, `Poppins` typography (bold titles, regular/medium body).
- Reuse the previous deck's visual motifs (background gradient/diagonal bands, footer, icon set) and the
  banner's composition for the cover (framed hero card left, speaker cards right).
- Brand font files live in `tools/fonts/<family>/` (TTF + licence, e.g. Poppins from Google Fonts, OFL).
  The builder measures text with them and the web deck loads the same family from Google Fonts.
  On this Mac, apps started from the CLI cannot see user-installed fonts (`~/Library/Fonts`), so
  LibreOffice falls back to Arial/Liberation. The builder detects that (PDF font list without the brand
  font) and prints the PDF from the web deck with headless Chrome instead (same geometry, Poppins
  embedded); the LibreOffice export is still produced and checked for page count. The line
  `PDF ... engine=libreoffice|chrome` tells you which path was used.
- Logos: use the files from Sources, proportional (`fit="contain"`), on a white card. Turn near-white
  logo backgrounds transparent, trim, and upscale tiny screenshots (see `prepare_assets.py`).

## 3. Slide order (fixed)

1. Cover: meetup name, edition, date, venue, community/sponsor/partner logos, speaker cards
2. About the meetup / community intro (+ "join us" QR to the Meetup group)
3. Organizers (roles as listed on Meetup; mark tonight's hosts)
4. Speakers + talks (name, role, title, short abstract from the event description)
5. Agenda / running order (times exactly as published; no invented slot times)
6. Sponsors **and** meetup partner in two clearly separated, labelled sections
7. KCD x DevOps Days: dates, venue, confirmed keynotes, "full agenda released this week", QR to the
   KCD site, discount code in a large pill next to the QR
8. Raffle: one KCD ticket, QR to the raffle form (>= 5 cm; the builder uses ~10.5 cm), one line of
   instructions (scan, fill the form, winner drawn at the end)
9. Closing: thank you, community social networks + how to join (QR per link)

## 4. Sponsors vs meetup partner

Sponsors and partners come from the **current** banner and event page (plus explicit user
instructions), never carried over from the previous edition. A meetup partner (co-hosting community)
is a different role from a sponsor: it gets its own labelled section on slide 6 and is never placed in
the sponsor lockup. Label the sponsor kind (venue sponsor, main sponsor, ...) as stated in the sources.

## 5. Folder conventions

```
<Community> Presentations/YYYY-MM-DD - <Meetup Name>/     # deliverables
  deck.json, prepare_assets.py, <slug>.pptx, <slug>.pdf, web/index.html,
  assets/{logos,photos,icons,qr,derived}/, preview/
Sources/YYYY-MM-DD - <Meetup Name>/                        # the mirrored, organised sources
  banner/  previous-deck/  logos/  docs/  <anything else>/
docs/<slug>/index.html + <slug>.pdf, docs/index.html        # GitHub Pages (generated)
```

Create both folders first and **move** every source file for the meetup out of the `Sources/` root
into its mirrored folder so `Sources/` stays clean. Save downloaded references (e.g. the real event
banner from meetup.com) there too. `slug` = `YYYY-MM-DD-<kebab-name>`.

## 6. Build

1. Copy `prepare_assets.py` from the latest meetup folder, adapt paths/crops, run it (logos, photo
   crops, background). Check the result visually (contact sheet).
2. Write `deck.json` (schema and field list in `tools/README.md`). Every string must trace back to a
   source.
3. `python3 tools/build_deck.py "<Community> Presentations/<YYYY-MM-DD - Meetup Name>"`
   - draws each slide once and emits PPTX (python-pptx, 16:9, real text boxes) **and** the
     self-contained web deck (one HTML file, inlined assets, arrow-key/swipe navigation, links on QRs);
   - generates all QR codes with `qrcode` + Pillow into `assets/qr/` (error correction Q, 4-module
     quiet zone, black on white);
   - converts to PDF with LibreOffice headless and asserts page count == slide count; if LibreOffice
     could not embed the brand font, prints the PDF from the web deck with headless Chrome instead
     (`engine=chrome` in the output) and keeps the LibreOffice render as `preview/pptx-XX.png`;
   - renders `preview/slide-XX.png` from the PDF, `preview/web-XX.png` (each slide at 1600x900) and
     `preview/web-fit-<size>.png` (slide 1 at laptop, portrait and phone window sizes; headless Chrome cannot go below 500px wide) with headless Chrome;
   - copies the web deck + PDF into `docs/<slug>/` and regenerates `docs/index.html`.
   Dependencies: `python-pptx qrcode pillow python-docx pymupdf opencv-python-headless`, LibreOffice
   (`soffice`), Google Chrome (optional screenshots).

## 7. Render check (mandatory before delivering)

- Read every `WARN` line from the builder (estimated text overflow / over-wide words) and fix them.
- Open **all** `preview/slide-XX.png` (final PDF), `preview/web-XX.png` (interactive page) and
  `preview/pptx-XX.png` (LibreOffice render of the PPTX, geometry only if fonts fell back) and check: no overflowing or clipped
  text, no overlapping elements, logos fully visible and proportional, readable contrast, QR codes
  crisp with a quiet zone, page number/footer not covered. In `web-fit-*.png` the whole slide must be
  visible and centred with the background around it at every window size (never clipped).
- Confirm the PDF font list contains the brand font and the page count equals the slide count.
- Fix `deck.json` / the layout, rebuild, re-check. Only report done when the check is clean.

## 8. Publish

- `git add -A && git commit -m "<Community> #<n>: <meetup name> deck (pptx, pdf, web)"` and push.
- GitHub Pages serves `main:/docs`. Live deck URL: `https://<owner>.github.io/<repo>/<slug>/`.
  Enable Pages once with `gh api -X POST repos/<owner>/<repo>/pages -f build_type=legacy -f "source[branch]=main" -f "source[path]=/docs"`.
- Wait for the Pages build (`gh api repos/<owner>/<repo>/pages/builds/latest`) and `curl -I` the URL
  until it returns 200. Give the live link at the very end of the reply.

## 9. Report

Finish with: what was built, where each file lives, what was moved in `Sources/`, the commit hash,
the live site link, and the open items that could not be resolved from the sources.

## 10. Variant: the generic company deck

The user reuses ONE synvert xgeeks deck across every meetup they run, tweaking a few slides in
PowerPoint per event. When asked to build/refresh *that* deck (not a specific meetup edition), it
is a different shape from §1-9 above - same house rules (sources first, never invent, render check,
publish), different slide set and a different builder:

- **Branding is exclusive to the reference deck.** Pull colour palette, typography, spacing, card/pill
  language and icon set only from the current xgeeks company deck in `Sources/` (the one the user
  already presents - look for a deck with "WE'RE synvert xgeeks", "$ kubectl apply -f ..." slides).
  Never take colours from a banner or another community's deck; a differently-branded deck (e.g. an
  AWS UG deck) can still be used as a *layout* reference (see sponsors & partners below) but never as
  a colour/font source. This xgeeks deck's own title convention is a small square tab overlapping a
  rounded pill (`$ kubectl apply -f ./x.yaml`, "What About ...") - not the AWS deck's black-bordered
  `title_card()`; keep using plain `card(framed=True)` for content tiles (both decks already share
  that hard-shadow black-border language) but build titles with the tab+pill pattern instead
  (`kicker_pill()` in `tools/build_company_deck.py`).
- **Audience/tone**: presented live to an engineer audience right before the talks start, ~3-4 minutes
  total, 5-second-from-the-back-of-the-room readability. Headline + a few words per slide, no bullet
  walls, no marketing filler. Positioning is AI-native (built *with* AI, specific about AI/Cloud Native
  engineering) - this has to read in the ACCELERATE and WHAT WE DO slides especially, without hype.
- **Folder conventions**: organise every source file for this deck under `Sources/xgeeks presentation/`
  (previous-deck/, logos/, communities/, ...) so `Sources/` stays clean, mirroring §5. Deliverables go
  in `xgeeks Presentations/YYYY-MM-DD - synvert xgeeks Company Deck/`, same internal layout as an event
  deck (`deck.json`, `prepare_assets.py`, pptx/pdf/web, `assets/`, `preview/`).
- **Fixed 10-slide order**: 1 Cover (banner placeholder, nothing else) · 2 Sponsors & Partners
  (generic, editable logo slots - not the old "intro" slide; use a prior AWS-style sponsors/meetup-partner
  slide purely as the *layout* reference: labelled SPONSORS/PARTNERS sections, one card per logo) ·
  3 About xgeeks · 4 We're synvert xgeeks + GlobalLogic lockup · 5 Accelerate (must carry the
  GlobalLogic-backed, AI-native positioning, condensed) · 6 What We Do (Cloud Native Applications,
  Platform Engineering, AI Engineering, Data Engineering) · 7 Clients (logo grid only) · 8 What About
  \<Community\> (reusable placeholder) · 9 Tech Talks Formats (two-block, e.g. Viseu/Porto Tech Talks)
  · 10 `$ kubectl apply -f ./community.yaml` (every community logo, the hero "show the room" slide).
- **synvert xgeeks + GlobalLogic lockup**: since the GlobalLogic acquisition, slide 4 keeps the existing
  big wordmark treatment and adds the official GlobalLogic logo (`Sources/.../logos/GlobalLogic-logo.jpg`
  or similar) in its own clean white plaque - GlobalLogic keeps its own official black-on-white colours,
  never recoloured/stretched, with clear space respected. Reinforce with an "A GlobalLogic Company"
  pill/badge on the Accelerate slide too.
- **Reusable placeholder slides are real, editable PPTX objects, not screenshots.** Cover banner slot,
  sponsor/partner logo slots, the community name + stats on "What About \<Community\>", and any missing
  logo (e.g. a sibling Tech Talks chapter with no logo yet in `Sources/`) are all dashed-border boxes
  with clear labels the user fills in PowerPoint - never bake them into a background image. Use
  `Canvas.placeholder(x, y, w, h, label, sublabel)` (added to the shared `tools/build_deck.py`, dashed
  outline via `Canvas.rect(..., dashed=True)`) for every one of these. If a required logo is genuinely
  missing from `Sources/` (e.g. a community with no distinct logo yet), don't invent one - draw a
  placeholder slot and list it as an open item, exactly like a missing fact in the event-deck flow.
  Content that *is* in the sources (client logos, community logos, stats) must trace back to a file -
  never invented, matching §1's rule.
- **Build**: `python3 tools/build_company_deck.py "xgeeks Presentations/<dated folder>"`. This script
  imports the shared `Canvas`/backends/PDF pipeline from `tools/build_deck.py` (same dual PPTX+web
  geometry, same LibreOffice+Chrome PDF fallback, same preview renders and GitHub Pages publish step)
  but defines its own 10-slide `draw_deck()`, since the slide types don't overlap with the fixed
  9-slide event template. `deck.json` here has its own schema (see the file itself, or
  `tools/build_company_deck.py`'s field access) - not the event-deck schema in `tools/README.md`.
- **Fonts**: self-hosting the brand font (`tools/fonts/<family>/*.ttf`, base64-embedded via
  `embedded_font_css()`) is now the default in the shared web-deck HTML, not just a Poppins-specific
  workaround - some Google-Fonts-served families (Montserrat included) get silently rasterised as
  nameless Type3 fonts by headless Chrome's PDF printer instead of embedding cleanly, which fails the
  "PDF font list contains the brand font" check in §7; self-hosting sidesteps it and works offline.
  If the brand font isn't already vendored under `tools/fonts/`, get static weights (Regular/Medium/
  SemiBold/Bold, + Italic/BoldItalic if used) - if only a variable font is available (e.g. via
  `brew install --cask font-<family>`), instance it with `fontTools.varLib.instancer` the way
  `tools/fonts/montserrat/` was generated, and carry over the font's own OFL/licence text.
- Everything else - render check (§7, now checking 10 slides), commit + push + Pages + live link (§8),
  and the final report (§9) - applies unchanged.
