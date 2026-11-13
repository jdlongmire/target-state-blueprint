# Content Spec Reference

Two builders take the same JSON spec file and an output
directory:

```
python3 scripts/build_blueprint.py spec.json --out DIR
python3 scripts/build_blueprint_pptx.py spec.json --out DIR
```

The first writes `<name>.pdf` and `<name>.png` (a 150 dpi render of
the cover) into DIR. The second writes `<name>.pptx`: the same
blueprint as native, fully editable PowerPoint (US Letter portrait
slides, one per stop; text frames, native tables, shapes). The
PPTX is the editable master; the PDF is the fixed rendering.
`<name>` is the spec's `name` field, or the spec
file name when `name` is absent. Spec files may carry full-line `//`
comments (the template uses them); the loader drops lines that start
with `//` before parsing. A spec that fails validation prints
`RESULT: INVALID` with the exact field at fault and builds nothing.

All spec text is plain text; the builder escapes it. Write the
middle dot separator as `\u00b7` inside JSON strings so spec files
stay ASCII. The builder itself supplies the typographic quotes
around the question in the Answers pill.

## Document fields

| Field | Type | Required | Meaning |
|---|---|---|---|
| `name` | string | no | Output file base name. Default: the spec file name. |
| `document.wordmark` | string | yes | The word printed top-left on every stop and used in the cover furniture (for example `BLUEPRINT`). |
| `document.kicker` | array of strings | yes | Cover kicker lines, mono caps: organization and function, then edition line (`FIELD GUIDE \u00b7 V1.0 \u00b7 MONTH YEAR`). |
| `document.title_lines` | array of strings | yes | The giant cover title, one entry per line. |
| `document.line_name` | string | no | Badge text across the cover's line roundel (the worked example uses `TS&A LINE`). Omit to drop the badge. |
| `document.subtitle` | string | no | Bold line under the cover title. |
| `document.deck` | string | no | The cover's selling sentence under the subtitle. |
| `document.doc_id` | string | yes | The footer line on every stop (the worked example uses `PLANT PLAN \u00b7 The Meridian works plan \u00b7 v1.0 \u00b7 Mar 2027`). Keep it short; long ids wrap to two lines. |
| `document.author_line` | string | no | Cover credit line above the bottom line graphic. |
| `document.cover_cards` | array | no | Exactly three cards work best: `{label, text}`. Labels are mono caps (scope, route, readers). |
| `document.cover_band` | object | no | `{label, items}`: the dark "what's in the plan" band. Items render in a 4-column grid. |
| `document.line_top` | object | no | `{tag: [line1, line2], stations: [...]}`. The cover's top transit line. Stations are `{label, color?}` or plain strings. |
| `document.line_bottom` | object | no | Same shape; the tag draws below the right end. |

Station colors cycle blue, orange, teal, purple when `color` is
absent. A station label longer than 9 characters with a space in it
wraps to two lines under its dot. The first station's label is
left-anchored and the last station's right-anchored so neither
clips the page edge. Keep to 7-9 stations per line.

## Stops

`stops` is an array; the cover is stop 1 and the first entry in
`stops` becomes stop 2. Numbering, the `STOP NN / TOTAL` header and
the footer progress strip are generated, never written in the spec.

| Field | Type | Required | Meaning |
|---|---|---|---|
| `title` | string | yes | The stop's assertion title, large. For `kind: "letter"` the title is not printed (the letter speaks for itself) but is still required as the stop's name. |
| `question` | string | no | The question the stop answers. The pill renders `Answers: "<question>"`. Omit the surrounding quotes; the builder adds them. |
| `pill` | string | no | Overrides the pill text outright (the letter stop uses `Start here`, the index stop `Answers: all of them`). |
| `lede` | string | no | One paragraph under the title that answers the question in brief before the blocks argue it. |
| `kind` | string | no | `content` (default) or `letter`. Letter stops print no title. |
| `blocks` | array | yes | The stop's content, in order. Any block may carry a `label`, printed as a mono caps section label above it. |

One stop should fill one page. If a stop's blocks overflow, the
build still succeeds: content flows to a second page, tables repeat
their header band, and the footer prints on every page (see Scale
behavior). The builder prints a NOTE when the page count differs
from the stop count so an overflow is never silent.

## Blocks

### prose
`{type: "prose", paragraphs: [...]}`. Each paragraph is a string, or
`{lead: "...", text: "..."}` where `lead` prints bold at the start.

### cards
`{type: "cards", columns: 2|3|4, cards: [{glyph?, color?, title, text}]}`.
Cards render in rows of `columns`; a short final row is padded with
invisible cards so widths hold. `glyph` is one short ASCII character
or letter shown in a rounded-square chip in `color` (a palette name:
blue, orange, teal, purple, ink, amber). Omit `glyph` for a
title-only card. Titles up to about 30 characters and text up to
about 240 characters read well at 2 columns; halve the text at 4.

### band
`{type: "band", label?, items: [...]}`. The dark ink band. Items are
strings or `{lead, text}` and render as equal columns (1 to 4).
With `grid: true`, string items render in a 4-column grid instead
(the cover band uses this form).

### steps
`{type: "steps", steps: [{title, owner?, text, tag?, color?}]}`. A
numbered horizontal flow with chevron separators. Numbers are
generated and colored by the line-color cycle unless `color` is
set. `owner` prints as a tiny mono line under the title; `tag` as a
mono amber line under the text (dates and cadences live there).
Three to five steps read well.

### list
`{type: "list", color?, items: [...]}`. Hairline-separated rows with
a colored dot. Items are strings or `{lead, text}`.

### groups
`{type: "groups", groups: [{label, color, items: [{title, text}]}]}`.
Two or three side-by-side groups under colored mono header bars;
the measures stop uses it (signals / proofs / health checks). Keep
item text to one or two lines.

### timeline
`{type: "timeline", items: [{when, text}]}`. Two-column rows: a mono
date column and a text column, hairline separated.

### index
`{type: "index", entries: [{question, headline?, stops}]}`. The
find-your-stop grid, three entries per row: the question in bold, the
stop's `headline` (when given) as a subordinate mono line under it,
then `STOP NN` in mono amber. `stops` is a string so ranges work
(`"10-14"`). The
numbers are the author's responsibility: fill them in after the stop
plan settles, and re-check after any reorder.

### table
`{type: "table", mode?, columns: [...], widths?, rows: [...], note?, d_label?}`.

- `widths` are percentages and should sum to 100; omit for equal
  columns.
- A row may be `{"section": "STRATEGY", "color": "blue"}` instead of
  an array: a full-width sub-band row, tinted by the named pillar
  (blue, orange, teal, purple, amber; anything else gets the warm
  neutral tint). Zebra striping restarts after each section row.
- `note` prints under the table in small muted text.
- The header band repeats on every page the table spans; rows never
  split across pages.

Mode `raci` changes two things: every column after the first is
centered and its header is drawn rotated (keep actor labels at or
under 24 characters), and cells in those columns are parsed as RACI
values:

- `A`, `R`, `D`, `C`, `I` render as letter chips: A ink-filled, R
  outlined, D teal-filled, C and I warm-tint filled.
- `-` or empty renders as a muted dash.
- `A::Plant OT` renders the A chip with the tag `Plant OT` in small
  type under it.
- A legend prints above the table. The D line reads `d_label`
  (default `Decides by vote`); set it to the body's actual decision
  rule.

### roadmap
`{type: "roadmap", phases: [...], subs?, milestones?, rows: [...]}`.

- `phases`: `[{label, span, band, tint}]`. `span` is how many grid
  columns the phase covers; the spans sum to the grid width (4 is
  typical). `band` is the header treatment: `light` (warm neutral),
  `amber`, or `ink`. `tint` is the background its columns' cells get
  unless a cell overrides it.
- `subs`: one short label per grid column (quarters or years),
  printed under the phase bands.
- `milestones`: `[{col, lead, text}]`, `col` a 0-based grid column.
  Each renders with an amber diamond.
- `rows`: `[{label, color, cells: [{span, text, tint?}]}]`. The
  label cell takes the row's palette color. Each row's cell spans
  must sum to the grid width.

## Scale behavior

Verified against a synthetic 24-stop blueprint with a 30-row,
10-actor RACI and a deliberately overflowing prose stop:

- Stop numbering and progress strips are correct at 24 stops; the
  strip's dots shrink as the count grows, and two-digit numbers fit
  the amber disc at any count.
- A table that spans pages repeats its dark header band, rotated
  RACI headers included, on every page; zebra striping and section
  bands continue correctly.
- The footer is a CSS running element (`position: running`), so it
  prints at the bottom of every page a stop touches and can never
  split across a break or strand on a page of its own. The cover
  has no footer.
- An overflowing stop simply takes a second page under its own
  header's stop; the NEXT stop still starts on a fresh page with
  the right number.
