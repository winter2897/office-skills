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
  "font": "Arial",
  "body_pt": 11,
  "document_type": "Test Document",
  "title": "E7000 Motor/ESC Qualification",
  "doc_no": "TD-EV1-26",
  "project": "ALTA X",
  "prepared_by": "Quan Tran",
  "tested_by": "Quan Tran, Minh Le",
  "revisions": [["1.0", "14/05/2026", "Initial", "QT"]],
  "sections": [ ... the seven below ... ]
}
```

| Key | Why |
|---|---|
| `"layout": "simple"` | **Required.** No cover, no control page, no roman front matter; sections do not each start a page; headings keep sentence case and lose the rule under them; the title is set at heading weight with a `TITLE:` prefix; contents and figure/table lists default off |
| `"font": "Arial"`, `"body_pt": 11` | **The defaults for this skill.** A test document is Arial 11 pt, not the Times New Roman 13 pt a formal report uses |
| `document_type` | Prints in the header as `TEST DOCUMENT: <doc_no>`. Keep it `Test Document` |
| `project` | Prints in the header as `Project Reference: <project>` |
| `tested_by` | Who ran the test, which is often not who wrote it up |

Everything else — `lang`, `brand`, `date`, `revisions`, `references`, every block
type — behaves exactly as in the report skill. Turn a contents list back on with
`"toc": true` if a test report grows past ~15 pages.

## The seven sections

Use these headings, in this order. They are the structure a reviewer looks for;
do not rename or reorder them, and do not number them by hand — the styles do it.

| # | Heading | What goes in it |
|---|---|---|
| 1 | Purpose of Testing | What is being qualified and what question the test answers. One sentence, plus one for the deliverable if there is one. |
| 2 | System Configuration | The exact thing under test — part numbers, firmware and hardware revisions, the rig. A `kv` block and nothing else: someone must be able to rebuild the setup from it, without reading a paragraph. |
| 3 | Items to be Tested / Not Tested | A `table`: `Item to Test / Test Overview / Responsibility`. **Say what is *not* tested too**, and why — that is half the value of the section. |
| 4 | Test Approach | Bullets for the quantities measured, a `table` for the methods or waveforms (`Name / Purpose / Definition`), one line for the conditions. Figures of the rig or the stimulus go here. |
| 5 | Test PASS / FAIL Criteria | The threshold each test is judged against. If the criteria are stated per test in section 4, this section says so in one line and does not repeat them. |
| 6 | Test Results | One Heading 2 per test, in the order of section 3, each two or three sentences of measurement and meaning. A failure gets its own Heading 2. Close with a `Summary` subsection holding the `Test Name / Test Description / Pass Criteria / Result` table. |
| 7 | Conclusions/Recommendations | The verdict in the first sentence — qualified, qualified with limits, or not qualified. Then the recommendations, and what is still open. |

## How to write it

**Short.** A test document is scanned by someone deciding whether to ship, not
read for pleasure. The reference document this skill is modelled on puts one
sentence in section 2, one in section 5 and one in section 7.

1. **One to three sentences per section**, except section 6. If a section runs
   longer, the detail belongs in a table.
2. **Tables carry the detail, prose carries the judgement.** Do not narrate a
   table in the paragraph above it.
3. **No process narration.** Not "we then proceeded to disassemble the unit in
   order to determine" — "Dissection showed a failed FET."
4. **A failure is three lines: what happened with the measurement, the root
   cause, the action.** In that order, labelled `Root cause:` and `Action:`. What
   the reader wants is why it broke and what you are doing about it.
5. **Section 7 opens with the verdict in one sentence.** Recommendations follow
   as bullets, one action each.
6. Cut every sentence that only restates the heading, and every hedge that does
   not carry a number.

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
5. **Do not add a cover page or a contents list** by switching `layout` back to
   `formal`. If someone wants the formal furniture they want the report skill.
6. **Do not drop `font` or `body_pt`.** Losing them silently reverts the document
   to Times New Roman 13 pt, which no longer matches the format.

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
