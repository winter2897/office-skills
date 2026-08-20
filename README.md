# office-skills

Skills for Claude Code / Claude.ai that generate professional office documents.
You write the content as JSON, the script does every bit of formatting – no one
has to remember which style to use, and no one types a section, figure or table
number by hand.

## Layout

```
office-skills/
├── assets/              shared branding – one logo for every skill below
│   ├── brand.json
│   ├── logo.png
│   ├── logo-mark.png
│   └── logo-mark.svg
└── docx/
    ├── report/          technical / integration / feasibility reports
    ├── manual/
    ├── plan/
    ├── record/
    ├── sop/
    └── specs/
```

Branding is organisation-wide, not per skill, so it lives once at the root.
Each skill walks up from its own directory to find `assets/brand.json`.

## Make it yours

Two steps, no code edit:

1. Put your logo in `assets/` (PNG or JPEG – an SVG cannot be embedded in a
   .docx). `logo.png` is the header mark, `logo-mark.png` the cover mark.
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

Logos are fitted inside their box keeping the aspect ratio, so a square logo
does not become as tall as it is wide and inflate the header on every page.
Anything unusable – missing file, SVG, extreme proportions – warns on stderr and
falls back to setting the company name in text.

**The repository ships RTROBOTICS as a working sample.** Replace it before you
send anything.

## Install

```bash
git clone https://github.com/winter2897/office-skills.git ~/.claude/skills/office-skills
pip3 install python-docx
```

## Use

Ask for the document in plain language – *"write a report about …"* – and the
skill composes the JSON and renders it. Or drive it by hand:

```bash
S=~/.claude/skills/office-skills/docx/report/scripts/report.py

python3 "$S" template out.docx             # blank template full of {{PLACEHOLDER}}
python3 "$S" write content.json out.docx   # finished report
python3 "$S" selfcheck                     # verify the script still works
```

The report skill writes in the language you prompt in – a Vietnamese prompt
gives a Vietnamese report, unless you ask for another language. Schema and
writing rules: [`docx/report/SKILL.md`](docx/report/SKILL.md); a full worked
example: [`docx/report/example.json`](docx/report/example.json).

## Design principle

The blank template and a real report go through the **same renderer** – the
template is just a report whose content is `{{PLACEHOLDER}}`. There is no path
by which the output drifts off-template.

Everything that carries a number is a **real Word field**, not dead text:

| Element | Mechanism |
|---|---|
| Section numbers (1., 1.1, 1.1.1) | numbering bound to the Heading styles |
| Figure and table numbers | `SEQ` fields, renumbered on insert |
| Page numbers | `PAGE` fields, roman front matter / arabic body |
| Contents, list of figures, list of tables | written out in full with `PAGEREF`, wrapped in a `TOC` field |

The contents list is written out rather than left as an empty field: Word
rebuilds it on open, and every other reader still shows a finished list.

## Requirements

Python 3 and `python-docx`. Microsoft Word is not needed to generate the files.
