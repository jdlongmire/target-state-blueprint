---
name: target-state-blueprint
description: Build a target-state blueprint "field guide" in a fixed form: a transit-line document in which every content page is a numbered stop that answers one question its audience keeps asking, with a cover, letter, index, charter, destination, CONOPS, decision rights, RACI, roadmap, measures, assumptions, glossary, and closing ask. Use when the user asks for a blueprint, field guide, or target-state document for an organization function that must answer its stakeholders' recurring questions in one governed document.
---

# Target State Blueprint

The skill makes one artifact: a **blueprint field guide** for an
organization function (a strategy and architecture team, an
operations function, any group whose stakeholders keep asking the
same questions). The form is fixed: the document is a transit
line, every content page is a numbered **stop**, and each stop
answers exactly one question, named in the header pill:
`Answers: "<question>"`. A sustained metaphor carries the whole
argument (in the worked example the function runs a manufacturing
works: the review board is the inspection gate, the roadmap is the
route map, the closing ask is the start of production).

**You write a JSON spec; the builder renders it.** You never lay out
pages yourself. Everything you need is in this folder:

- `reference/design-system.md` is the visual system: palette,
  typography, page furniture, and block treatments, with the
  parameters the builder implements.
- `reference/stop-catalog.md` is the canonical stop sequence with
  include/cut guidance and the drafting rules.
- `reference/content-spec.md` is the full spec field reference,
  including the scale behavior at 20+ stops.
- `templates/blueprint-spec-template.json` is the fill-in skeleton
  with the full stop sequence stubbed.
- `scripts/build_blueprint.py` is the PDF builder: spec JSON in,
  PDF and cover PNG out.
- `scripts/build_blueprint_pptx.py` is the native PPTX builder: the
  same spec JSON in, an editable PowerPoint out (every text element
  a real text frame, tables native, cards and bands shapes with
  text).
- `examples/` holds one finished worked example on a fictional
  organization (Meridian Works): spec, built PDF, and cover PNG.
  It is the reference build: the form, fully exercised.

Requirements: Python 3 with WeasyPrint (see `requirements.txt`; on
this machine the docs venv `~/workspace/.venvs/docs/bin/python`
already has it), plus `pdftoppm` (poppler-utils) on PATH for the
cover PNG and QA renders. The PPTX builder needs python-pptx and
Pillow (on this machine the deck venv
`~/workspace/.venvs/deck/bin/python` has both).

---

## The workflow

### Step 1: Study the framework

Read `reference/stop-catalog.md` and `reference/design-system.md`
end to end, and open the worked example's built PDF
(`examples/fictional-example-blueprint.pdf`). Internalize the
furniture before writing anything: cover
(kicker, giant display title, three summary cards, dark plan band,
transit-line graphics top and bottom, FIRST STOP note), then per
stop a header (wordmark left, `STOP NN / TOTAL` center, dark
Answers pill right), a 4-segment color bar, a large assertion
title, a lede, content blocks, and a footer (document id left,
station-progress strip right with the current stop amber).

### Step 2: Intake

Establish, from the user and their source material:

- **The function and its audience.** Whose blueprint is it, and who
  keeps asking it questions?
- **The questions.** Collect the real recurring questions ("Where's
  the RACI?", "Who decides what?", "When will I get it?"). These
  become the stops. A question nobody asks is not a stop.
- **The metaphor mapping.** Map the metaphor BEFORE drafting: the
  function, its disciplines, its governance bodies, its rhythm, its
  roadmap, and its closing ask each get one metaphor term, and the
  mapping is used consistently in every stop, title, and band. A
  metaphor that appears on the cover and vanishes by stop 5 is a
  defect.
- **Source material.** Charters, org charts, process docs, prior
  plans: the facts the stops will state. Do not invent numbers,
  names, dates, or commitments the sources do not carry.

### Step 3: Draft the stop plan

Work from `reference/stop-catalog.md`. Start from the canonical
sequence, apply the include/cut guidance, and record for each stop:
its question, its assertion title, and its block types. The floor
for a full blueprint is the catalog's required set: letter, index,
charter, destination, CONOPS, decision rights, RACI, deliverables,
roadmap, measures, assumptions, glossary, closing ask. **Measures,
assumptions, glossary, and the closing ask are not optional in a
full blueprint.** A focused blueprint for a smaller function runs
18-24 stops; a full blueprint runs to about 33.

### Step 4: Write the content spec

Copy `templates/blueprint-spec-template.json` and fill it in per
`reference/content-spec.md`. Numbering, the `STOP NN / TOTAL`
header, and the footer progress strip are generated; never write
them in the spec. The index stop's numbers are typed by hand: fill
them in after the stop plan settles, and re-check after any reorder.
Delete template stops you have no content for; never leave `<<...>>`
in the file. Spec files may carry full-line `//` comments.

### Step 5: Build

```
python3 scripts/build_blueprint.py spec.json --out DIR
```

Writes `<name>.pdf` and `<name>.png` (150 dpi cover render) into
DIR; `<name>` is the spec's `name` field or the spec file name. The
script prints `RESULT: PASS` with the file paths and the page
count, or `RESULT: INVALID` with the exact field at fault, building
nothing. A NOTE prints when the page count differs from the stop
count, so an overflow is never silent.

Build the native PPTX from the same spec:

```
python3 scripts/build_blueprint_pptx.py spec.json --out DIR
```

Writes `<name>.pptx`: US Letter portrait slides, one per stop,
fully editable. **The PPTX is the editable master; the PDF is the
fixed rendering of record.** Build both for every delivery. The
PPTX builder prints the same PASS/INVALID/NOTE contract (slides
vs stops).

### Step 6: Visual QA (mandatory)

Render the built PDF's pages to PNG (`pdftoppm`) and view at
least: the cover, a card stop, the RACI stop, and the roadmap stop.
Do the same for the PPTX: convert it with LibreOffice
(`soffice --headless --convert-to pdf`) and view the same stops;
host line heights differ from CSS, so text that fits the PDF can
overflow a native shape. Check against the design system and the
worked example:

1. Furniture on every stop: wordmark, stop counter, Answers pill,
   4-segment color bar, footer document id and progress strip with
   the current stop amber.
2. Density: stops fill their page the way the worked example's do; no
   near-empty stops, no orphan stops, no stop furniture split or
   stranded on a page of its own.
3. Tables: dark header band present, repeated on every page a long
   table spans; RACI letter chips render as chips, rotated actor
   headers legible.
4. The metaphor reads consistently from cover to closing.

Fix defects in the spec or the builders, never by hand-editing
the outputs. Do not deliver a blueprint whose pages you have not
viewed.

### Step 7: Deliver

Deliver the PPTX (the editable master) and the PDF (and the cover
PNG if asked). Say in one or two
sentences what the blueprint covers, how many stops it runs, and
any stop where the build falls short of the design system. Do not
paste the JSON spec into your answer unless the user asks for it.

---

## Drafting rules

- **One question per stop.** If a stop answers two questions, it is
  two stops. If two stops answer the same question, one is cut.
- **Assertion titles.** The title states the answer's headline
  ("The destination: one blueprint"), never the topic alone
  ("Destination").
- **The lede answers first.** Each stop's lede paragraph answers its
  question in brief before the blocks argue it.
- **One page per stop.** Size blocks to the page; split a stop
  (RACI 1 of 2, processes 1 of 2) rather than letting furniture
  break. Overflow is handled cleanly when it happens (see the
  scale behavior in `reference/content-spec.md`), but a planned
  split reads better than an accidental one.
- **Order is load-bearing.** Charter before destination, CONOPS
  before handoffs, decision rights before the RACI, roadmap after
  deliverables, assumptions and glossary just before the ask.
- **The index is written last and checked twice**, against the built
  PDF, not from memory.

## Scale behavior

Verified against a synthetic 24-stop blueprint with a 30-row,
10-actor RACI and a deliberately overflowing prose stop: stop
numbering and progress strips stay correct at 24 stops (dots shrink
as the count grows; two-digit numbers fit the amber disc); tables
that span pages repeat their dark header band, rotated RACI headers
included; the footer is a CSS running element, so it prints at the
bottom of every page a stop touches and can never split across a
break; an overflowing stop takes a second page under its own stop
number and the next stop still starts fresh with the right number.
Full detail in `reference/content-spec.md`.

## Worked example

`examples/fictional-example-blueprint.json` builds the 12-stop
"Plant Plan" for Meridian Works, a fictional mid-size manufacturer:
cover, letter, index, charter, destination (card grid plus dark
band), team structure, RACI, roadmap grid with timeline, measures
in three groups, assumptions table, glossary table, and closing
ask. Its built PDF and cover PNG sit beside the spec. The example
is invented end to end: no real organization, people, or numbers.
