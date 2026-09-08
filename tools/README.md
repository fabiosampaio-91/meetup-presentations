# tools/build_deck.py

Builds a meetup deck from `<meetup dir>/deck.json`:

```
python3 tools/build_deck.py "<Community> Presentations/<YYYY-MM-DD - Meetup Name>" [--no-publish]
```

One layout definition (`draw_deck`) is rendered by two backends at once, so the PPTX and the web deck
share identical geometry. Outputs: `<slug>.pptx`, `<slug>.pdf`, `web/index.html` (self-contained),
`assets/qr/*.png`, `preview/*.png`, and the GitHub Pages copy under `docs/<slug>/`.

PDF: LibreOffice headless exports the PPTX first (page count asserted, render kept as
`preview/pptx-XX.png` + `preview/libreoffice-check.pdf`). If that PDF does not embed the brand font
(macOS user fonts invisible to CLI-launched apps), the final PDF is printed from the web deck with
headless Chrome (`engine=chrome`), which embeds the Google Fonts version of the brand font.

Fonts: `tools/fonts/<family>/` holds the TTFs used for text measurement (Poppins, SIL OFL).

## deck.json fields

| key | content |
|---|---|
| `slug`, `community`, `title`, `description`, `date_iso`, `edition_label`, `footer` | identity + footer line |
| `community_logo` | path to the community logo (shown on the cover and About slide) |
| `brand.font`, `brand.background`, `brand.shadow_offset`, `brand.colors{purple,navy,deep,grey,white,black,teal,muted_on_light,muted_on_dark}` | branding sampled from the banner |
| `cover.logos[] {name,file}`, `cover.title_line_1/2`, `cover.date_line`, `cover.venue_line` | slide 1 |
| `about {title,tagline,quote,bullets[],join_url,join_label}` | slide 2 |
| `organizers {title,host_badge,note,people[] {name,role,hosting}}` | slide 3 |
| `speakers_title`, `speakers[] {name,role,talk,abstract,photo}` | slides 1 and 4 |
| `agenda {title,items[] {time,what,detail}}` | slide 5 |
| `sponsors_slide {title,sponsors_label,sponsors[] {name,kind,logo},partner_label,partner{name,kind,logo},thanks}` | slide 6 |
| `kcd {title,logo,event_name,when,where,blurb,keynotes_label,keynotes[] {name,role,photo},agenda_note,url,qr_label,code_label,discount_code,code_hint}` | slide 7 |
| `raffle {title,headline,instructions,url,url_label,qr_label,qr_size_in}` | slide 8 |
| `closing {title,subtitle,follow_label,links[] {label,short,icon,qr_name,url}}` | slide 9 |

All paths are relative to the meetup folder. Fonts are measured with the real font files
(`~/Library/Fonts`) to size pills and to warn about text that will not fit.
