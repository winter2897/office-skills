# Office Skills

> **Beta** — This project is under active development. Skills, schemas, and the branding format may change without notice. Feedback and contributions welcome.

Document skills for AI coding agents. Plug into your favorite AI coding tool and get finished Word documents on your own company template — you supply the content, the skill supplies the cover page, the numbering, the contents list and the styles. Point it at your logo once and every document that follows is branded.

## Skills

| Skill | Description | Status |
|-------|-------------|--------|
| `docx/report` | Technical, integration and feasibility reports as `.docx`. Cover page, document-control table, revision history, auto-numbered headings, table of contents, list of figures and tables, `SEQ`-numbered figure / table / listing captions, abbreviations, references, styled tables, running header and footer. Writes Vietnamese or English from the same schema. | Available |
| `docx/test` | Test reports, qualification reports and test records. The seven-section engineering structure — purpose, system configuration, items tested / not tested, approach, PASS/FAIL criteria, results, conclusions. Flat layout: no cover, no contents list, straight into section 1, the way a bench document is read. | Available |
| `docx/sop` | Standard operating procedures. | Planned |
| `docx/manual` | User and maintenance manuals. | Planned |
| `docx/plan` | Project and test plans. | Planned |
| `docx/record` | Test records and inspection sheets. | Planned |
| `docx/specs` | Requirement and interface specifications. | Planned |

Every skill reads the same branding, so a company sets its logo once and each
document type comes out matching. They share one renderer too: `docx/test` is
the report renderer with `"layout": "simple"` and a different section skeleton,
not a second copy of the code.

## Branding

Branding is organisation-wide, not per skill, so it lives once at the root in
`assets/`. Making the template yours takes two steps and no code edit:

1. Put your logo in `assets/` as PNG or JPEG — an SVG cannot be embedded in a
   `.docx`. `logo.png` is the header mark, `logo-mark.png` the cover mark.
2. Change `company` in `assets/brand.json`.

```json
{
  "company": "Your Company",
  "logo": "logo.png",
  "mark": "logo-mark.png",
  "logo_box_cm": [2.9, 1.2],
  "mark_box_cm": [1.6, 1.6],
  "cover_bg": "F4F4F4",
  "ink": "1A1A1A",
  "lang": "en",
  "font": "Times New Roman"
}
```

| Key | What it does |
|-----|--------------|
| `company` | Cover top left, page footer, document properties |
| `logo` / `mark` | Header mark and cover mark. Relative to `assets/`, or an absolute path |
| `logo_box_cm` / `mark_box_cm` | `[width, height]` in cm the mark is fitted inside |
| `cover_bg` | Cover ground colour, hex. The image is generated, there is no file to edit |
| `ink` | Body text colour, hex |
| `lang` | Furniture language when a document does not say — `vi` or `en` |
| `font` | Default typeface; a single document can still override it |

Logos are **fitted inside their box keeping the aspect ratio**, so a square logo
does not become as tall as it is wide and inflate the header on every page.
Anything unusable — missing file, SVG, an image over 2000 px, an aspect ratio
outside 8:1 … 1:3 — warns on stderr, and an unreadable logo falls back to setting
the company name in text rather than failing the render. Read stderr: a document
can come out perfect and still have lost its logo.

A single document can override any brand key without touching the file:

```json
{"brand": {"company": "ACME Corp", "logo": "/abs/path/acme.png"}, "title": "..."}
```

**This repository ships RTROBOTICS as a working sample.** Replace it before you
send anything.

## Installation

### Claude Code

```bash
claude plugin marketplace add https://github.com/winter2897/office-skills
claude plugin install office-skills
```

Two lines install the **whole set** — every skill in the repository, together
with the shared `assets/`, so branding works out of the box. Adding a skill later
means `claude plugin update office-skills`, not a second install. Restart Claude
Code to discover them.

The skills need `python-docx`. You do not have to install it up front: the first
render without it stops with the exact command to run, and the agent runs it.

### Claude.ai

Skills upload one directory at a time and cannot reach the shared `assets/`, so
copy the branding in before zipping:

```bash
git clone https://github.com/winter2897/office-skills.git
cp -R office-skills/assets office-skills/docx/report/assets
cd office-skills/docx && zip -r docx-report.zip report
```

`docx/test` renders through the report script, so it cannot be uploaded on its
own — zip `report/` and use the test skill's `example.json` as the shape.

Settings → Capabilities → Skills → Upload skill. Code execution must be enabled.

### Manual

No agent needed — the skills are plain Python scripts:

```bash
git clone https://github.com/winter2897/office-skills.git
pip3 install python-docx
```

## Usage

Ask in plain language — *"write a report about the payload binding test"* — and
the agent composes the JSON and renders it. The report skill writes in the
language you prompt in, so a Vietnamese prompt gives a Vietnamese report unless
you ask for another one.

By hand:

```bash
S=~/.office-skills/docx/report/scripts/report.py

python3 "$S" template out.docx             # blank template full of {{PLACEHOLDER}}
python3 "$S" write content.json out.docx   # finished document
python3 "$S" selfcheck                     # verify the script still works
```

Schema and writing rules: [`docx/report/SKILL.md`](docx/report/SKILL.md).
A full worked example: [`docx/report/example.json`](docx/report/example.json).

## How it works

The blank template and a real document go through the **same renderer** — the
template is just a document whose content is `{{PLACEHOLDER}}`. There is no path
by which the output drifts off-template.

Everything that carries a number is a **real Word field**, not dead text:

| Element | Mechanism |
|---------|-----------|
| Section numbers (1., 1.1, 1.1.1) | numbering bound to the Heading styles |
| Figure, table and listing numbers | `SEQ` fields, renumbered on insert |
| Page numbers | `PAGE` fields, roman front matter and arabic body |
| Contents, list of figures, list of tables | written out in full with `PAGEREF`, wrapped in a `TOC` field |

The contents list is written out rather than left as an empty field, so Word
rebuilds it on open while every other reader still shows a finished list.

Content is validated before anything renders: a misspelt key, a missing required
field, a column-width list that does not add up — all raise with the location
(`section 3.1, block 2 (table): ...`). Nothing is silently ignored.

## Contributing

Every skill ships a self-check. Run it before opening a pull request:

```bash
python3 docx/report/scripts/report.py selfcheck
claude plugin validate . --strict
```

A new skill is a directory with a `SKILL.md` under its format group, plus one
line in the `skills` array of [`.claude-plugin/plugin.json`](.claude-plugin/plugin.json)
so it ships with the bundle. Prefer reusing an existing renderer over adding a
second one, the way `docx/test` does — a skill can be nothing but a `SKILL.md`
and a worked example. The path is listed explicitly rather than scanning
the group, because that is what makes the `name:` in the skill's frontmatter the
name it is invoked by.

It renders both modes and the worked example, then asserts the fonts, the
language tag, the caption fields, the generated lists, the page breaks, the brand
overrides, logo fitting at three aspect ratios, and that no em dash reached the
document.

## Requirements

Python 3 and `python-docx`. Microsoft Word is not needed to generate the files —
only to refresh the field values before you export a PDF.

## License

[MIT](LICENSE)
