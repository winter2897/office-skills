---
name: docx-report
description: "Write technical reports as .docx on an organisation's own template – logo, cover page, document-control table, auto-numbered headings, table of contents, list of figures and tables, auto-numbered figure/table captions, abbreviations, references, styled tables, page footer. Branding (company name, logo, cover colour) comes from assets/brand.json, so the same skill serves any company. Use whenever the user asks for a technical, integration, feasibility or project report as a Word document, wants a report in their company format, or hands over a Template.docx and asks for a report in that style. Triggers: 'viết báo cáo', 'báo cáo theo template', 'technical report', 'write a report docx', 'docx-report'."
---

# Report template

Two modes, one script – a written report and the blank template come out of the
same renderer, so output is always on-template.

```bash
S=<skill>/scripts/report.py

python3 "$S" template [out.docx]             # blank template with {{PLACEHOLDER}}
python3 "$S" write content.json [out.docx]   # finished report
python3 "$S" selfcheck                       # verify the script still works
```

Needs `python-docx`. If it is missing the script stops with the install
command instead of a traceback - run `pip3 install python-docx` and retry.

## Branding

Everything company-specific lives in **`assets/brand.json` at the root of this
repository**, one directory shared by every skill under it – there is no logo
copy per skill. To make the template yours: drop your logo in `assets/`, change
`company`, done. No code edit.

```json
{
  "company": "RTROBOTICS",
  "logo": "logo.png",
  "mark": "logo-mark.png",
  "logo_box_cm": [2.9, 1.2],
  "mark_box_cm": [1.6, 1.6],
  "cover_bg": "F4F4F4",
  "ink": "1A1A1A",
  "lang": "en",
  "font": "Times New Roman",
  "doc_no_example": "RTR-REP-2026-014",
  "email_example": "quan.tran@rtrobotics.me"
}
```

| Key | What it does |
|---|---|
| `company` | Cover top left, page footer, document properties |
| `logo` | Header mark, every page. Relative to `assets/`, or an absolute path |
| `mark` | Cover mark, bottom right – usually the logo without the wordmark |
| `logo_box_cm` / `mark_box_cm` | `[width, height]` the mark is fitted inside |
| `cover_bg` | Cover ground, hex. The PNG is generated, there is no file to edit |
| `ink` | Body text colour, hex |
| `lang` | Furniture language when `content.json` does not say (`vi` / `en`) |
| `font` | Default typeface; a report can still override it |
| `doc_no_example` / `email_example` | Shown on the template's instruction page |

**The shipped `brand.json` is RTROBOTICS as a working sample.** Before rendering
for anyone else, check it still names the right company – otherwise the report
goes out with someone else's logo on it. Ask the user once, then leave it alone.

A single report can override any of these without touching the file:

```json
{"brand": {"company": "ACME Corp", "logo": "/abs/path/acme.png"}, "title": "..."}
```

Logos are **fitted inside their box, keeping the aspect ratio** – a square logo
does not become 2.9 cm tall and inflate the header on every page. Give a PNG or
JPEG; an SVG cannot be embedded in a .docx and falls back to the company name in
text, with a warning on stderr. Also warned about: an image over 2000 px on its
long side (the .docx carries every pixel it is given) and an aspect ratio outside
8:1 … 1:3 (widen `logo_box_cm`). **Read stderr** – a report can render perfectly
and still have lost its logo.

No `brand.json` found at all – the skill was copied out on its own – means brand
defaults and the company name set in text where the logo would be. Nothing
crashes.

## Language

**Write the report in the language the user prompted in**, unless they ask for
another one. Vietnamese prompt → Vietnamese report. Vietnamese prompt that says
"write it in English" → English report. Nothing to go on → the `lang` in
`brand.json`.

Set the `lang` key in `content.json` to match: `"vi"` or `"en"`. It drives the
document furniture (cover labels, `MỤC LỤC` / `TABLE OF CONTENTS`, `Trang` /
`Page`, `Hình` / `Figure`, `Bảng` / `Table`) and it **must agree with the prose
you write** – English furniture around Vietnamese paragraphs is a bug.

Writing Vietnamese, keep a term in English only when there is no clear
Vietnamese equivalent – `payload`, `gimbal`, `firmware`, `companion computer`,
`binding`, protocol and part names. Do not invent awkward translations for
those, and do not leave ordinary prose in English.

## Punctuation

Use the en dash `–` (U+2013) for ranges, asides and caption separators.
**Never the longer em dash (U+2014).** The selfcheck fails if one reaches the
document.

## Workflow

1. Collect the content from the user (or from the source docs they point at).
   Missing metadata → ask once for doc number / date / author, or use `TBD`.
2. Check `assets/brand.json` names the user's organisation, not the sample.
3. Write `content.json` (schema below). **This is where the writing happens** –
   compose real prose, not placeholders.
4. Run `python3 "$S" write content.json <DOC_NO>.docx`, and read stderr.
5. Tell the user the path.

Do **not** hand-edit the .docx afterwards – change the JSON and re-run.

## content.json

```json
{
  "lang": "vi",
  "font": "Times New Roman",
  "document_type": "Báo cáo",
  "title": "Tích hợp Phase One P3",
  "object": "Phase One P3",
  "doc_no": "REP-2026-014",
  "version": "1.0",
  "date": "17/08/2026",
  "prepared_by": "Tên – Chức danh",
  "email": "name@company.com",
  "tested_by": "",
  "reviewed_by": "", "approved_by": "",
  "classification": "Nội bộ",
  "revisions": [["1.0", "17/08/2026", "Phát hành lần đầu", "HQT"]],
  "toc": true,
  "list_of_figures": true,
  "list_of_tables": true,
  "abbreviations": [["GCS", "Ground Control Station", "Trạm điều khiển mặt đất"]],
  "references": ["AirSim Team, AirSim API Documentation, Microsoft Research, 2024."],
  "sections": [
    {"heading": "Tổng quan",
     "blocks": [{"type": "para", "text": "..."}],
     "sections": [{"heading": "Mục con", "blocks": []}]},
    {"heading": "Kết luận", "numbered": false, "blocks": []}
  ]
}
```

`font` sets the typeface for the whole document – styles, theme and tables. It
only accepts fonts Office installs on both Windows and macOS:

`Times New Roman`, `Arial`, `Calibri`, `Cambria`, `Georgia`, `Tahoma`,
`Trebuchet MS`, `Verdana`

Anything else raises. .docx stores the font *name*, not the font, so Inter or
Roboto on a machine that lacks them becomes a silent substitute and a layout
nobody proof-read. Keep Times New Roman for anything a Vietnamese authority
reads (TT 01/2011 asks for Times New Roman 13-14 pt). Code blocks stay on
Consolas.

`document_type`, `title` and `classification` fall back to the language default
(`Báo cáo` / `Report`, `Nội bộ` / `Confidential`) when left empty. The cover
shows only `title`, `version`, the month of `date`, and the author block
(`prepared_by` + `email`); every other field prints on page i. Full worked
example: `example.json`.

### Layout

`"layout": "formal"` (the default) is the document described here: cover page,
document-control page, roman front matter, a contents list, and a new page for
every top-level section.

`"layout": "simple"` drops all of it. The document opens with the title, the
author, `tested_by` if set, and the revision table, then goes straight into
section 1. One section, page numbers from 1, sections flow without a page break,
level-1 headings keep their sentence case, and the contents and figure/table
lists default to off — pass `"toc": true` to bring one back. The styles,
branding, header, footer, auto-numbered headings and `SEQ` captions are
unchanged, so a flat document is the same template, not a different one.

Use it for anything read at a bench rather than filed: see the
[`docx-test`](../test/SKILL.md) skill, which is this renderer with the simple
layout and a fixed seven-section skeleton.

`tested_by` names whoever ran a test, which is often not whoever wrote the
document up. Left empty it prints nowhere.

### Sections

`sections` nest recursively → Heading 1 / 2 / 3, numbered automatically.

`"numbered": false` gives a heading with **no number that still appears in the
table of contents** – for `Kết luận`, `Phụ lục`, and anything else outside the
numbered body. The whole subtree under it stays unnumbered.

`references` renders its own unnumbered references section at the end,
numbered `[1]`, `[2]`… Cite them in prose as `[1]`.

`abbreviations` renders the abbreviations list in the front matter.
Rows of 2 (short form, full term) or 3 (plus a gloss).

`revisions` gets its own page straight after the cover. It is kept off the cover
on purpose: the cover has a fixed job (what the document is, who approved it)
while the revision table grows with every issue and would overflow the page.

### Block types

| `type` | Fields | Renders as |
|---|---|---|
| `para` | `text` | Body Text, justified |
| `bullets` | `items` – a nested list becomes level-2 bullets | List Bullet / List Bullet 2 |
| `kv` | `rows: [[label, value]]`, `label_width` (twips, default 2800) | 2-column spec table, labels shaded |
| `table` | `header`, `rows`, `widths` (twips, sum 9638), `caption` | Report Table, header row repeats |
| `figure` | `path`, `width_cm` (default 14), `caption`, `placeholder` | centred image + caption |
| `code` | `text` (newlines kept), `caption`, `language` | Consolas 10 pt listing in a shaded box, caption above it |
| `equation` | `text` | centred, numbered `(1)` at the right margin |
| `quote` | `text` | Quote style |
| `pagebreak` | – | page break |

Inline code: wrap a command, flag, path or file name in backticks and that span
alone comes out in Consolas, one point below the body – `"text": "chạy `sudo
systemctl restart ros` trước khi đo"`. Backticks are the only markup the text
fields understand, and they work in `para`, `bullets`, `quote` and every table
cell. A `code` block without `caption` renders the box alone, and `language`
only ever shows up inside the caption, so it needs one to appear at all.

`path` is resolved from the current directory, so give an **absolute path**.
A path that does not resolve prints a warning on stderr and falls back to the
placeholder box – check the console before sending a report that should have
figures in it. Omit `path` entirely for a deliberate empty figure slot.

`content.json` is validated before anything renders: an unknown or misspelt key
(top level, `brand`, section or block), a missing required field, a `widths` list
that does not add up, a table row wider than its header – all raise with the
location (`section 3.1, block 2 (table): ...`). Nothing is silently ignored.

## Rules

1. **Never type section numbers** into a heading – the styles number themselves.
   Typing them is what broke the original document (`3.3` jumped to `3.5`).
2. **Never type figure or table numbers either.** Captions are built from `SEQ`
   fields, so they number themselves (`Hình 1`, `Bảng 2`) and renumber when
   something is inserted. Give only the caption *text*.
   Numbering is flat, not chapter-scoped: `STYLEREF 1 \s` can only read a
   heading number stored on the paragraph, and ours lives on the *style*, so
   Word printed the heading text instead (`Bảng Câu hỏi mở cần chốt.9`).
3. **Table and listing captions sit above, figure captions below the figure.**
   The renderer enforces this; do not fight it. Listings number themselves too,
   so cite them as "xem Đoạn mã 1" / "see Listing 1".
4. **Every top-level section starts on a new page.** Automatic – do not add a
   `pagebreak` block before a Heading 1. Keep sections substantial enough to
   fill a page, or the report turns into a stack of half-empty sheets.
5. Prose must cite figures and tables by number – "xem Hình 2", "see Table 4" –
   never "hình dưới đây" / "the figure below".
6. Keep `widths` summing to 9638 twips (A4 minus 2 cm margins) or omit them –
   a list that does not add up is refused, not silently overflowed.
7. The header carries the logo, doc no and version, the footer carries the
   company name and the page number – all driven by the brand and the top-level
   metadata, so never write them into a section. `classification` prints on page
   i only.
8. Only `write` mode produces a sendable document; `template` mode adds a red
   "delete this page" instruction page at the end.

## Layout produced

```
Cover      one borderless table: company + month top left, version chip top
           right, rubric and title in the middle band, author block bottom
           left, logo bottom right
           no header/footer, not counted in the page numbering
Front      document information + revision history (same page), contents,
matter     list of figures, list of tables, abbreviations
           one page each, numbered i, ii, iii …
Body       auto-numbered Heading 1/2/3, page numbers restart at 1
           each Heading 1 opens a new page
           references last, unnumbered, also on its own page
```

Every generated list (contents, figures, tables) is **written out in full** –
heading text, dot leader, `PAGEREF` page number, clickable hyperlink – and then
wrapped in its Word field (`TOC \o "1-3"`, `TOC \c "Hình"`, `TOC \c "Bảng"`), so
Word rebuilds it on update while every other viewer still shows a finished list.
The *entries* are right everywhere; the cached page numbers all read `1` until
something paginates the document and refreshes the fields. Only an editor can:
Word does it on open (`updateFields`), ONLYOFFICE when you refresh the list by
hand, Google Docs never – it drops fields on import and keeps the cached text.
So refresh once in an editor before sending the .docx or uploading it to Drive,
and export the PDF from there – never from a viewer that leaves fields alone.

Heading 1 prints upper case as *text*, not through the `w:caps` effect: Google
Docs has no all-caps run format and imports `w:caps` as small caps. Word rebuilds
the contents list from that text, so level-1 entries shout there too.

Body 13 pt per the Vietnamese report convention, language tag set from `lang` so
spell-check behaves. Rules and table borders `#BFBFBF`, header-row fill
`#F2F2F2`, code background `#F7F7F7`. The cover adds only greys: version chip
`#DDDDDD`, title rule `#8C8C8C`; its ground comes from `cover_bg`.

Every element of the cover sits in a cell of one borderless table
(`build_cover`), because a cover built out of floating boxes falls apart the
first time somebody edits it in Word – and this template exists to be edited.
Row heights are `atLeast`, so a title of one, two or three lines leaves the
author block and the logo exactly where they are. The only floating object is
the cover ground – a flat PNG generated from `cover_bg` at render time, anchored
to the page behind the text and `locked` so it cannot be dragged.

To make a `mark` from a full logo SVG, strip the wordmark and rasterise it:

```bash
rsvg-convert -w 1200 assets/logo-mark.svg -o assets/logo-mark.png
```

## Changing the design

Brand, typeface and cover colour are all config, not edits. For anything deeper,
edit `scripts/report.py`: `BRAND_DEFAULTS` / `DEFAULT_FONT` / `BODY_PT` / `DASH`
at the top, `install_styles` (type scale, spacing, colours), `install_table_style`,
`build_header` / `build_footer`, `placeholder_spec` (default skeleton). Then run
`selfcheck` – it renders `example.json` alongside both modes and checks the
fonts, the language tag, the caption fields, both generated lists, the
roman-numeral front matter, the per-section page breaks, the bullet glyphs, that
no two tables end up adjacent (Word merges those into one), that the spec
validator rejects malformed JSON, that a brand override reaches the cover and the
footer, that logos of any aspect ratio stay inside their box, that a missing or
unreadable logo warns and falls back to text, that the cover ground carries the
requested colour, that STYLEREF stays out, and that no em dash slipped in.
