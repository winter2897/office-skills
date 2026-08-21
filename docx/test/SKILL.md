---
name: docx-test
description: "Write test reports, qualification reports and test records as .docx – the seven-section engineering test structure (purpose, system configuration, items tested / not tested, approach, PASS/FAIL criteria, results, conclusions), on your organisation's branding. Flat layout: no cover page, no table of contents, straight into section 1, the way a bench document is read. Use whenever the user asks for a test report, qualification report, test record, verification or validation report, a DVT/EVT/PVT write-up, or hands over test data and asks for it written up. Triggers: 'báo cáo test', 'báo cáo thử nghiệm', 'test report', 'qualification report', 'test record', 'docx-test'."
---

# Test report

A test report is the report renderer with a different shape: the same styles,
branding, auto-numbered headings and `SEQ` captions, but flat — no cover page, no
document-control page, no contents list. It opens with the title and who ran the
test, then goes straight into section 1.

```bash
S=<office-skills>/docx/report/scripts/report.py

python3 "$S" write content.json <DOC_NO>.docx     # finished test report
python3 "$S" write <this-skill>/skeleton.json blank.docx   # blank, all placeholders
```

There is no script here on purpose — one renderer, one set of styles, nothing to
drift. **Read [`../report/SKILL.md`](../report/SKILL.md) for the full schema**:
block types, branding, fonts, captions, validation. This file only covers what a
test document does differently.

## What is different

```json
{
  "layout": "simple",
  "title": "DJI Motor/ESC Qualification",
  "doc_no": "EV1-26",
  "prepared_by": "Raul R.",
  "tested_by": "Raul R.",
  "revisions": [["1.0", "14/05/2019", "Initial", "Raul R."]],
  "sections": [ ... the seven below ... ]
}
```

| Key | Why |
|---|---|
| `"layout": "simple"` | **Required.** Drops the cover, the control page and the roman front matter; sections no longer each start a new page; headings keep their sentence case; contents and figure/table lists default off |
| `tested_by` | Who ran the test, which is often not who wrote it up |

Everything else — `lang`, `font`, `brand`, `doc_no`, `date`, `revisions`,
`references`, every block type — behaves exactly as in the report skill. Turn a
contents list back on with `"toc": true` if a test report grows past ~15 pages.

## The seven sections

Use these headings, in this order. They are the structure a reviewer looks for;
do not rename or reorder them, and do not number them by hand — the styles do it.

| # | Heading | What goes in it |
|---|---|---|
| 1 | Purpose of Testing | Why this test exists, in 2-4 sentences: what is being qualified, what question the test answers, what the deliverable is (a limit, a procedure, a go/no-go). |
| 2 | System Configuration | The exact thing under test — part numbers, firmware and hardware revisions, the rig around it. Someone must be able to rebuild the setup from this section alone. A `kv` block suits it. |
| 3 | Items to be Tested / Not Tested | A `table`: `Item to Test / Test Overview / Responsibility`. **Say what is *not* tested too**, and why — that is half the value of the section. |
| 4 | Test Approach | How each test is run. Bullets for the quantities measured, prose for the procedure, a `table` for the methods or waveforms (`Name / Purpose / Example`). Figures of the rig or the stimulus go here. |
| 5 | Test PASS / FAIL Criteria | The threshold each test is judged against. If the criteria are stated per test in section 4, this section says so in one line and does not repeat them. |
| 6 | Test Results | One Heading 2 per test, in the order of section 3. Data, figures, observations. Close with a `table`: `Test Name / Test Description / Pass Criteria / Result`. Failures get their own Heading 2 with what failed, the evidence and the root cause. |
| 7 | Conclusions/Recommendations | The verdict in the first sentence — qualified, qualified with limits, or not qualified. Then the recommendations, and what is still open. |

## Rules

1. **Write the result, not the hope.** A test that failed is written up as
   failed, with the evidence. A test not run is `Not tested` with the reason, not
   a blank cell. Never round a marginal result into a pass.
2. **Every number carries its unit and its condition** — `4031 W at 1950 µs PWM,
   25 °C ambient`. A number without its condition cannot be reproduced.
3. **Section 6 mirrors section 3.** Every item listed as to-be-tested has a
   result; anything tested that was not in the list is added to section 3 first.
4. **Cite figures and tables by number** — "see Figure 2", "xem Bảng 4". The
   captions number themselves; never type a number into one.
5. Failures are worth more than passes. Give a failure its own Heading 2 under
   section 6: what happened, when, the measurement that shows it, the
   disassembly or log evidence, and the root cause if known.
6. **Do not add a cover page or a contents list** by switching `layout` back to
   `formal`. If someone wants the formal furniture they want the report skill.

## Language

The same rule as the report skill: **write in the language the user prompted
in**, unless they ask otherwise, and set `lang` to match. Test documents are
often read by a supplier or a customer — if the user says the audience is
foreign, `"lang": "en"` and English prose.

Keep instrument, part and protocol names in their original form in both
languages — `PWM`, `FET`, `ESC`, `duty cycle`, `burn-in`.

## Files

| File | What it is |
|---|---|
| `skeleton.json` | The seven sections as `{{PLACEHOLDER}}`, ready to fill |
| `example.json` | A complete worked test report, modelled on a real qualification document |
