#!/usr/bin/env python3
"""Build a target-state blueprint PDF (plus a page-1 PNG) from a JSON spec.

The spec describes the document and its stops; this script renders the
field-guide form defined in reference/design-system.md: a cover with
transit-line graphics, then one numbered stop per page, each stop
answering one question, with a station-progress footer.

Usage:
    python3 scripts/build_blueprint.py SPEC.json --out DIR

Writes <name>.pdf and <name>.png into DIR. <name> is the spec's "name"
field, or the spec file name when "name" is absent. Prints
"RESULT: PASS" with the file paths, or "RESULT: INVALID" with the
exact field at fault and builds nothing.

Scale behavior: stops may overflow one page. Tables repeat their
header band on every page they span. The stop footer (document id
and progress strip) is a CSS running element: it prints at the
bottom of every page the stop touches, the cover carries none, and
it can never split or strand across a page break. Stop numbering
and the progress strip are generated from the spec, so any stop
count works; the strip's dots shrink as the count grows.
"""

import argparse
import html
import json
import subprocess
import sys
import tempfile
from pathlib import Path

# Palette tokens (see reference/design-system.md for the full
# token list and their uses).
PAL = {
    "ink": "#1a1d29",
    "body": "#3d4149",
    "muted": "#6a6e78",
    "blue": "#2d5ba8",
    "orange": "#b8501a",
    "teal": "#12806f",
    "purple": "#6a4c9c",
    "amber": "#f4c542",
    "amber_dark": "#9a7410",
    "hairline": "#dcd5c4",
    "row_tint": "#f3efe6",
    "section_tint": "#e6e0d2",
}
TINTS = {
    "blue": "#e3eaf6",
    "orange": "#f7e4d8",
    "teal": "#ddf0ec",
    "purple": "#e9e3f3",
    "amber": "#fff6d9",
    "": "#e6e0d2",
}
CYCLE = ["blue", "orange", "teal", "purple"]
DOT = "\u00b7"          # middle dot separator used by the furniture
LQ = "\u201c"           # left double quotation mark (Answers pill)
RQ = "\u201d"           # right double quotation mark
DIA = "\u25c6"          # milestone diamond on the roadmap

# SVG text is styled with presentation attributes, not CSS classes:
# WeasyPrint does not apply document CSS font sizes inside inline SVG.
SVG_MONO = 'font-family="Liberation Mono, DejaVu Sans Mono, monospace"'
SVG_SANS = 'font-family="Liberation Sans, DejaVu Sans, sans-serif"'

CSS = """
@page { size: Letter; margin: 44pt 54pt 62pt 54pt;
  @bottom-center { content: element(stopfoot); vertical-align: top; } }
* { box-sizing: border-box; }
body { font-family: "Liberation Sans", "DejaVu Sans", sans-serif;
       color: #3d4149; margin: 0; font-size: 10pt; }
section { break-before: page; }
section:first-of-type { break-before: auto; }
section.cover { display: flex; flex-direction: column; min-height: 686pt; }
.spacer { flex: 1 1 auto; }

/* ---- stop furniture ---- */
.shead { display: flex; align-items: center; justify-content: space-between; }
.wm { font-weight: bold; font-size: 11pt; letter-spacing: 2.2pt; color: #1a1d29; }
.stopno { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 8.5pt; letter-spacing: 1.6pt; color: #1a1d29; }
.pill { background: #1a1d29; color: #ffffff; border-radius: 9pt;
  padding: 3.5pt 9pt; font-size: 8.5pt; font-weight: bold;
  max-width: 62%; text-align: right; }
.colorbar { display: flex; height: 4.5pt; margin-top: 7pt; }
.colorbar .cb1 { flex: 2 0 0; background: #2d5ba8; }
.colorbar .cb2 { flex: 1 0 0; background: #b8501a; }
.colorbar .cb3 { flex: 1 0 0; background: #12806f; }
.colorbar .cb4 { flex: 2 0 0; background: #6a4c9c; }
h1.stitle { font-size: 26pt; line-height: 1.1; color: #1a1d29;
  margin: 13pt 0 7pt 0; font-weight: bold; letter-spacing: -0.2pt; }
.lede { font-size: 11.5pt; line-height: 1.45; margin: 0 0 12pt 0; }
.runfoot { position: running(stopfoot); display: table; width: 504pt;
  border-top: 0.75pt solid #1a1d29; padding-top: 5pt; }
.runfoot .dcell { display: table-cell; width: 200pt;
  font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.3pt; letter-spacing: 0.6pt; color: #1a1d29;
  vertical-align: bottom; padding-bottom: 3pt; }
.runfoot .scell { display: table-cell; text-align: right;
  vertical-align: bottom; }
.stlab { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 6.6pt; fill: #3d4149; text-anchor: middle; letter-spacing: 0.4pt; }
.taglab { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 6.2pt; fill: #f4c542; text-anchor: middle; letter-spacing: 0.5pt; }
.stripnum { font-family: "Liberation Sans", sans-serif; font-weight: bold;
  fill: #1a1d29; text-anchor: middle; }
.rothead { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.1pt; fill: #ffffff; letter-spacing: 0.5pt; }

/* ---- cover ---- */
.kickrow { display: flex; justify-content: space-between;
  align-items: flex-start; margin-top: 14pt; }
.kicker { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 9.5pt; letter-spacing: 1.8pt; color: #1a1d29; line-height: 1.8; }
.badge { position: relative; width: 56pt; height: 56pt; flex: 0 0 auto;
  border: 3.5pt solid #2d5ba8; border-radius: 50%; }
.badge .bpill { position: absolute; top: 50%; left: 50%;
  transform: translate(-50%, -50%); background: #1a1d29; color: #f4c542;
  font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 6.8pt; letter-spacing: 0.6pt; padding: 2.5pt 5pt;
  white-space: nowrap; }
h1.ctitle { font-size: 63pt; line-height: 0.98; color: #1a1d29;
  margin: 16pt 0 10pt 0; letter-spacing: -0.5pt; font-weight: bold; }
.csub { font-size: 19pt; font-weight: bold; color: #1a1d29; margin: 0 0 6pt 0; }
.cdeck { font-size: 12pt; line-height: 1.45; width: 84%; margin: 0 0 16pt 0; }
.ccards { display: flex; gap: 9pt; margin-bottom: 10pt; }
.ccard { flex: 1 1 0; border: 1.3pt solid #1a1d29; border-radius: 9pt;
  padding: 9pt 11pt 10pt 11pt; }
.ccard .lab { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.5pt; letter-spacing: 1.4pt; color: #6a6e78; }
.ccard .txt { font-size: 10.5pt; font-weight: bold; color: #1a1d29;
  line-height: 1.35; margin-top: 3pt; }
.authorrow { display: flex; justify-content: space-between;
  align-items: flex-end; margin: 10pt 0 8pt 0; }
.authorline { font-size: 9pt; line-height: 1.45; width: 68%; }
.firststop { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 8pt; letter-spacing: 1.2pt; color: #1a1d29; }

/* ---- shared blocks ---- */
.blabel { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 8.4pt; letter-spacing: 1.6pt; color: #1a1d29;
  margin: 12pt 0 6pt 0; }
.prose p { font-size: 10.5pt; line-height: 1.5; margin: 0 0 8pt 0; }
.prose .lead { font-weight: bold; color: #1a1d29; }
.cards { display: flex; flex-direction: column; gap: 9pt; margin-bottom: 10pt; }
.crow { display: flex; gap: 9pt; }
.card { flex: 1 1 0; border: 1.3pt solid #1a1d29; border-radius: 9pt;
  padding: 9pt 11pt 10pt 11pt; break-inside: avoid; background: #ffffff; }
.card.ghost { border-color: transparent; }
.chead { display: flex; align-items: center; gap: 7pt; }
.card .chip { flex: 0 0 auto; width: 19pt; height: 19pt; border-radius: 5pt;
  color: #ffffff; font-weight: bold; font-size: 11pt; text-align: center;
  line-height: 19pt; }
.card h3 { font-size: 11.5pt; color: #1a1d29; margin: 0; }
.card p { font-size: 9.3pt; line-height: 1.42; margin: 6pt 0 0 0; }
.band { background: #1a1d29; border-radius: 9pt; padding: 11pt 13pt 12pt 13pt;
  margin-bottom: 10pt; break-inside: avoid; }
.band .blab { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  color: #f4c542; font-size: 8.3pt; letter-spacing: 1.8pt; margin-bottom: 7pt; }
.band .bcols { display: flex; gap: 14pt; }
.band .bcol { flex: 1 1 0; color: #eef0f4; font-size: 9.3pt; line-height: 1.45; }
.band .bcol .lead { color: #ffffff; font-weight: bold; }
.band .bgrid { display: flex; flex-wrap: wrap; gap: 5pt 0; }
.band .bitem { flex: 0 0 25%; color: #ffffff; font-size: 9.5pt; }
.steps { display: flex; align-items: flex-start; margin-bottom: 10pt; }
.step { flex: 1 1 0; padding: 2pt 3pt; break-inside: avoid; }
.step .num { font-size: 16pt; font-weight: bold; line-height: 1; }
.step h3 { font-size: 10.5pt; color: #1a1d29; margin: 4pt 0 2pt 0; }
.step .owner { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 6.8pt; letter-spacing: 0.8pt; color: #6a6e78; }
.step p { font-size: 8.8pt; line-height: 1.4; margin: 3pt 0 0 0; }
.step .tag { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.3pt; color: #9a7410; margin-top: 5pt; letter-spacing: 0.8pt; }
.chev { flex: 0 0 13pt; display: flex; justify-content: center;
  padding-top: 26pt; }
.chev span { display: block; width: 7pt; height: 7pt;
  border-top: 2pt solid #1a1d29; border-right: 2pt solid #1a1d29;
  transform: rotate(45deg); }
.blist { margin: 0 0 10pt 0; padding: 0; list-style: none; }
.blist li { font-size: 10pt; line-height: 1.45; padding: 3.5pt 0 3.5pt 14pt;
  position: relative; border-bottom: 0.75pt solid #efe9dc; }
.blist li .bdot { position: absolute; left: 1pt; top: 8.5pt; width: 5pt;
  height: 5pt; border-radius: 50%; }
.blist .lead { font-weight: bold; color: #1a1d29; }
.groups { display: flex; gap: 9pt; margin-bottom: 10pt; }
.group { flex: 1 1 0; }
.ghead { border-radius: 6pt; padding: 4.5pt 8pt; color: #ffffff;
  font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 8pt; letter-spacing: 1.4pt; margin-bottom: 7pt; }
.gitem { margin-bottom: 8pt; break-inside: avoid; }
.gitem h4 { font-size: 9.8pt; color: #1a1d29; margin: 0 0 1.5pt 0; }
.gitem p { font-size: 8.8pt; line-height: 1.4; margin: 0; }
.tl { margin-bottom: 10pt; }
.tlrow { display: flex; border-bottom: 0.75pt solid #dcd5c4;
  padding: 4.5pt 0; break-inside: avoid; }
.tlwhen { flex: 0 0 92pt; font-family: "Liberation Mono", "DejaVu Sans Mono",
  monospace; font-size: 8.3pt; color: #1a1d29; letter-spacing: 0.4pt;
  padding-top: 1.5pt; }
.tltext { flex: 1 1 0; font-size: 9.5pt; line-height: 1.42; }
.indexrows { display: flex; flex-direction: column; gap: 12pt;
  margin-bottom: 10pt; }
.irow { display: flex; gap: 14pt; }
.ientry { flex: 1 1 0; break-inside: avoid; }
.ientry .q { font-size: 10pt; font-weight: bold; color: #1a1d29;
  line-height: 1.3; }
.ientry .s { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.8pt; color: #9a7410; letter-spacing: 1.2pt; margin-top: 2.5pt; }

/* ---- tables ---- */
.twrap { border: 1.3pt solid #1a1d29; border-radius: 9pt; overflow: hidden;
  margin-bottom: 6pt; }
.twrap table { width: 100%; border-collapse: collapse; }
.twrap thead { display: table-header-group; }
.twrap th { background: #1a1d29; color: #ffffff;
  font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7.3pt; letter-spacing: 0.7pt; text-transform: uppercase;
  text-align: left; padding: 5pt 6pt; }
.twrap th.rot { padding: 0; }
.twrap td { font-size: 9pt; line-height: 1.35; padding: 4.5pt 6pt;
  border-top: 0.75pt solid #dcd5c4; vertical-align: top; color: #2c2f38; }
.twrap tr { break-inside: avoid; }
.twrap tr.alt td { background: #f3efe6; }
.twrap tr.section td { font-family: "Liberation Mono", "DejaVu Sans Mono",
  monospace; font-size: 8pt; letter-spacing: 1.4pt; color: #1a1d29;
  font-weight: bold; padding: 4pt 6pt; }
.twrap table.raci td.rcell { text-align: center; }
.legend { font-size: 8.5pt; margin-bottom: 7pt; color: #1a1d29; }
.rchip { display: inline-block; min-width: 12pt; padding: 1.2pt 3pt;
  border-radius: 3.5pt; font-size: 8pt; font-weight: bold; text-align: center; }
.rchip.a { background: #1a1d29; color: #ffffff; }
.rchip.r { background: #ffffff; border: 0.9pt solid #1a1d29; color: #1a1d29; }
.rchip.d { background: #12806f; color: #ffffff; }
.rchip.c, .rchip.i { background: #e6e0d2; color: #1a1d29; }
.celltag { display: block; font-size: 6.3pt; color: #6a6e78; margin-top: 1pt; }
.dash { color: #9a958a; }
.tnote { font-size: 8.3pt; color: #6a6e78; margin: 2pt 0 10pt 0;
  line-height: 1.4; }

/* ---- roadmap ---- */
.rmwrap { border: 1.3pt solid #1a1d29; border-radius: 9pt; padding: 7pt;
  margin-bottom: 10pt; }
table.rm { width: 100%; border-collapse: separate; border-spacing: 0; }
table.rm thead { display: table-header-group; }
table.rm th, table.rm td { vertical-align: top; }
.rmband { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 8pt; letter-spacing: 1.2pt; text-align: center;
  padding: 4.5pt 4pt; border-radius: 5pt; }
.rmb-ink { background: #1a1d29; color: #ffffff; }
.rmb-amber { background: #f4c542; color: #1a1d29; }
.rmb-light { background: #e6e0d2; color: #1a1d29; }
.rm-sub { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 7pt; color: #6a6e78; text-align: center; padding: 3pt 2pt; }
.rm-ms-h { font-family: "Liberation Mono", "DejaVu Sans Mono", monospace;
  font-size: 6.8pt; letter-spacing: 0.8pt; color: #6a6e78; text-align: left;
  padding: 4pt 4pt 4pt 0; }
.rm-ms { font-size: 7.8pt; line-height: 1.3; padding: 3pt 4pt;
  color: #1a1d29; }
.dia { color: #dd9d0d; font-family: "DejaVu Sans", sans-serif; font-size: 9pt; }
.rm-label { color: #ffffff; font-size: 9.5pt; font-weight: bold;
  text-align: left; padding: 5pt 7pt; border: 2pt solid #ffffff;
  border-radius: 6pt; width: 12.5%; }
.rm-cell { border: 0.75pt solid #dcd5c4; font-size: 8.6pt; line-height: 1.38;
  padding: 5pt 6pt; color: #2c2f38; }
"""


def esc(s):
    return html.escape(str(s), quote=False)


def pal(name, default="blue"):
    return PAL.get(name, PAL[default])


# ---------------------------------------------------------------- SVG ---

def line_svg(stations, tag=None, tag_at="start", width=504):
    """A transit line: colored segments, station dots, labels below.

    tag is {"lines": ["DEPART", "OCT 26"]}; tag_at "start" draws the
    tag above the left end (cover top line), "end" below the right
    end (cover bottom line).
    """
    n = len(stations)
    top_pad = 26 if (tag and tag_at == "start") else 6
    cy = top_pad + 8
    bottom_extra = 48 if (tag and tag_at == "end") else 4
    height = cy + 8 + 26 + bottom_extra
    if n == 1:
        xs = [width / 2]
    else:
        step = (width - 24) / (n - 1)
        xs = [12 + i * step for i in range(n)]

    def st_color(i):
        st = stations[i]
        if isinstance(st, dict):
            return pal(st.get("color", CYCLE[i % 4]))
        return pal(CYCLE[i % 4])

    parts = [f'<svg width="{width}pt" height="{height}pt" '
             f'viewBox="0 0 {width} {height}">']
    for i in range(n - 1):
        parts.append(
            f'<line x1="{xs[i]:.1f}" y1="{cy}" x2="{xs[i + 1]:.1f}" '
            f'y2="{cy}" stroke="{st_color(i)}" stroke-width="3.6"/>')
    for i, st in enumerate(stations):
        label = st["label"] if isinstance(st, dict) else st
        fill = st_color(i)
        if i == 0 and tag and tag_at == "start":
            fill = PAL["amber"]
        if i == n - 1 and tag and tag_at == "end":
            fill = PAL["amber"]
        parts.append(
            f'<circle cx="{xs[i]:.1f}" cy="{cy}" r="6" fill="{fill}" '
            f'stroke="{PAL["ink"]}" stroke-width="2"/>')
        if i == 0:
            anchor, ax = "start", xs[i] - 6
        elif i == n - 1:
            anchor, ax = "end", xs[i] + 6
        else:
            anchor, ax = "middle", xs[i]
        words = label.split(" ")
        if len(words) > 1 and len(label) > 9:
            parts.append(
                f'<text x="{ax:.1f}" y="{cy + 16}" {SVG_MONO} '
                f'font-size="6.6" fill="{PAL["body"]}" '
                f'text-anchor="{anchor}" letter-spacing="0.4">'
                f'{esc(words[0])}</text>')
            parts.append(
                f'<text x="{ax:.1f}" y="{cy + 24.5}" {SVG_MONO} '
                f'font-size="6.6" fill="{PAL["body"]}" '
                f'text-anchor="{anchor}" letter-spacing="0.4">'
                f'{esc(" ".join(words[1:]))}</text>')
        else:
            parts.append(
                f'<text x="{ax:.1f}" y="{cy + 16}" {SVG_MONO} '
                f'font-size="6.6" fill="{PAL["body"]}" '
                f'text-anchor="{anchor}" letter-spacing="0.4">'
                f'{esc(label)}</text>')
    if tag:
        lines = tag["lines"]
        tw = max(len(s) for s in lines) * 4.4 + 12
        if tag_at == "start":
            tx, ty = 0, 0
        else:
            tx, ty = width - tw, cy + 38
        parts.append(f'<rect x="{tx:.1f}" y="{ty}" width="{tw:.1f}" '
                     f'height="19" rx="2.5" fill="{PAL["ink"]}"/>')
        parts.append(
            f'<text x="{tx + tw / 2:.1f}" y="{ty + 8}" {SVG_MONO} '
            f'font-size="6.2" fill="{PAL["amber"]}" '
            f'text-anchor="middle" letter-spacing="0.5">'
            f'{esc(lines[0])}</text>')
        if len(lines) > 1:
            parts.append(
                f'<text x="{tx + tw / 2:.1f}" y="{ty + 15.5}" {SVG_MONO} '
                f'font-size="6.2" fill="{PAL["amber"]}" '
                f'text-anchor="middle" letter-spacing="0.5">'
                f'{esc(lines[1])}</text>')
    parts.append("</svg>")
    return "".join(parts)


def strip_svg(total, current, width=296):
    """The footer station-progress strip: one dot per stop, the
    current stop an amber disc carrying its number."""
    h, cy = 15, 7.5
    if total <= 1:
        xs = [width / 2]
    else:
        step = (width - 16) / (total - 1)
        xs = [8 + i * step for i in range(total)]
    r = 2.5 if total <= 26 else 2.1
    parts = [f'<svg width="{width}pt" height="{h}pt" '
             f'viewBox="0 0 {width} {h}">']
    parts.append(f'<line x1="{xs[0]:.1f}" y1="{cy}" x2="{xs[-1]:.1f}" '
                 f'y2="{cy}" stroke="#b9b2a2" stroke-width="1"/>')
    for i, x in enumerate(xs, start=1):
        if i == current:
            fs = 6.6 if total > 20 else 7.2
            parts.append(
                f'<circle cx="{x:.1f}" cy="{cy}" r="6.8" '
                f'fill="{PAL["amber"]}" stroke="{PAL["ink"]}" '
                f'stroke-width="1.1"/>')
            parts.append(
                f'<text x="{x:.1f}" y="{cy + 2.4}" {SVG_SANS} '
                f'font-size="{fs}" font-weight="bold" '
                f'fill="{PAL["ink"]}" text-anchor="middle">'
                f'{i:02d}</text>')
        else:
            parts.append(
                f'<circle cx="{x:.1f}" cy="{cy}" r="{r}" fill="#ffffff" '
                f'stroke="{PAL["ink"]}" stroke-width="0.9"/>')
    parts.append("</svg>")
    return "".join(parts)


def rot_svg(label, col_pt):
    """A rotated actor-column header for RACI tables, drawn as SVG so
    the rotation is exact and repeats with the table header band."""
    return (f'<svg width="{col_pt:.0f}pt" height="106pt" '
            f'viewBox="0 0 {col_pt:.0f} 106">'
            f'<text x="{col_pt / 2 + 2.6:.1f}" y="102" fill="#ffffff" '
            f'font-size="7.1" letter-spacing="0.5" {SVG_MONO} '
            f'transform="rotate(-90 {col_pt / 2 + 2.6:.1f} 102)">'
            f'{esc(label)}</text></svg>')


# ------------------------------------------------------------- blocks ---

def with_label(b, inner):
    if b.get("label"):
        return f'<div class="blabel">{esc(b["label"])}</div>' + inner
    return inner


def render_prose(b):
    out = []
    for p in b["paragraphs"]:
        if isinstance(p, dict):
            lead = f'<span class="lead">{esc(p["lead"])}</span> ' \
                if p.get("lead") else ""
            out.append(f"<p>{lead}{esc(p.get('text', ''))}</p>")
        else:
            out.append(f"<p>{esc(p)}</p>")
    return with_label(b, '<div class="prose">' + "".join(out) + "</div>")


def render_cards(b):
    cols = int(b.get("columns", 2))
    cards = b["cards"]
    rows = []
    for i in range(0, len(cards), cols):
        cells = []
        for c in cards[i:i + cols]:
            chip = ""
            if c.get("glyph"):
                chip = (f'<span class="chip" style="background:'
                        f'{pal(c.get("color", "blue"))}">'
                        f'{esc(c["glyph"])}</span>')
            cells.append(
                f'<div class="card"><div class="chead">{chip}'
                f'<h3>{esc(c["title"])}</h3></div>'
                f'<p>{esc(c.get("text", ""))}</p></div>')
        while len(cells) < cols:
            cells.append('<div class="card ghost"></div>')
        rows.append('<div class="crow">' + "".join(cells) + "</div>")
    return with_label(b, '<div class="cards">' + "".join(rows) + "</div>")


def render_band(b):
    label = f'<div class="blab">{esc(b["label"])}</div>' \
        if b.get("label") else ""
    items = b.get("items", [])
    if b.get("grid"):
        cells = "".join(f'<div class="bitem">{esc(it)}</div>'
                        for it in items)
        return (f'<div class="band">{label}'
                f'<div class="bgrid">{cells}</div></div>')
    cols = []
    for it in items:
        if isinstance(it, dict):
            lead = f'<span class="lead">{esc(it["lead"])}</span> ' \
                if it.get("lead") else ""
            cols.append(f'<div class="bcol">{lead}'
                        f'{esc(it.get("text", ""))}</div>')
        else:
            cols.append(f'<div class="bcol">{esc(it)}</div>')
    return (f'<div class="band">{label}'
            f'<div class="bcols">' + "".join(cols) + "</div></div>")


def render_steps(b):
    steps = b["steps"]
    parts = []
    for i, s in enumerate(steps):
        color = pal(s.get("color", CYCLE[i % 4]))
        owner = f'<div class="owner">{esc(s["owner"])}</div>' \
            if s.get("owner") else ""
        tag = f'<div class="tag">{esc(s["tag"])}</div>' \
            if s.get("tag") else ""
        parts.append(
            f'<div class="step"><div class="num" style="color:{color}">'
            f'{i + 1}</div><h3>{esc(s["title"])}</h3>{owner}'
            f'<p>{esc(s.get("text", ""))}</p>{tag}</div>')
        if i < len(steps) - 1:
            parts.append('<div class="chev"><span></span></div>')
    return with_label(b, '<div class="steps">' + "".join(parts) + "</div>")


def render_list(b):
    color = pal(b.get("color", "blue"))
    items = []
    for it in b["items"]:
        if isinstance(it, dict):
            lead = f'<span class="lead">{esc(it["lead"])}</span> ' \
                if it.get("lead") else ""
            items.append(f'<li><span class="bdot" style="background:'
                         f'{color}"></span>{lead}'
                         f'{esc(it.get("text", ""))}</li>')
        else:
            items.append(f'<li><span class="bdot" style="background:'
                         f'{color}"></span>{esc(it)}</li>')
    return with_label(b, '<ul class="blist">' + "".join(items) + "</ul>")


def render_groups(b):
    groups = []
    for g in b["groups"]:
        items = "".join(
            f'<div class="gitem"><h4>{esc(it["title"])}</h4>'
            f'<p>{esc(it.get("text", ""))}</p></div>'
            for it in g["items"])
        groups.append(
            f'<div class="group"><div class="ghead" style="background:'
            f'{pal(g.get("color", "blue"))}">{esc(g["label"])}</div>'
            f'{items}</div>')
    return with_label(b, '<div class="groups">' + "".join(groups) + "</div>")


def render_timeline(b):
    rows = "".join(
        f'<div class="tlrow"><div class="tlwhen">{esc(it["when"])}</div>'
        f'<div class="tltext">{esc(it["text"])}</div></div>'
        for it in b["items"])
    return with_label(b, f'<div class="tl">{rows}</div>')


def render_index(b):
    entries = b["entries"]
    rows = []
    for i in range(0, len(entries), 3):
        cells = []
        for e in entries[i:i + 3]:
            cells.append(
                f'<div class="ientry"><div class="q">'
                f'{esc(e["question"])}</div>'
                f'<div class="s">STOP {esc(e["stops"])}</div></div>')
        while len(cells) < 3:
            cells.append('<div class="ientry"></div>')
        rows.append('<div class="irow">' + "".join(cells) + "</div>")
    return with_label(b, '<div class="indexrows">' + "".join(rows) + "</div>")


def raci_cell(cell):
    s = str(cell).strip()
    if s in ("-", ""):
        return '<span class="dash">-</span>'
    tag = None
    if "::" in s:
        s, tag = s.split("::", 1)
    letter = s.strip().upper()
    cls = letter.lower() if letter in "ARDCI" else "c"
    out = f'<span class="rchip {cls}">{esc(letter)}</span>'
    if tag:
        out += f'<span class="celltag">{esc(tag.strip())}</span>'
    return out


def render_table(b):
    mode = b.get("mode", "plain")
    cols = b["columns"]
    widths = b.get("widths")
    colgroup = ""
    if widths:
        colgroup = "<colgroup>" + "".join(
            f'<col style="width:{w}%"/>' for w in widths) + "</colgroup>"
    if mode == "raci":
        ths = [f'<th>{esc(cols[0])}</th>']
        for j, c in enumerate(cols[1:], start=1):
            if widths and len(widths) == len(cols):
                col_pt = widths[j] / 100 * 500
            else:
                col_pt = 500 / len(cols)
            ths.append(f'<th class="rot">{rot_svg(c, col_pt)}</th>')
        head = "".join(ths)
    else:
        head = "".join(f"<th>{esc(c)}</th>" for c in cols)
    rows_html = []
    alt = False
    for row in b["rows"]:
        if isinstance(row, dict) and "section" in row:
            bg = TINTS.get(row.get("color", ""), TINTS[""])
            rows_html.append(
                f'<tr class="section"><td colspan="{len(cols)}" '
                f'style="background:{bg}">{esc(row["section"])}</td></tr>')
            alt = False
            continue
        tds = []
        for j, cell in enumerate(row):
            if mode == "raci" and j > 0:
                tds.append(f'<td class="rcell">{raci_cell(cell)}</td>')
            else:
                tds.append(f"<td>{esc(cell)}</td>")
        cls = ' class="alt"' if alt else ""
        alt = not alt
        rows_html.append(f"<tr{cls}>" + "".join(tds) + "</tr>")
    legend = ""
    if mode == "raci":
        d_label = esc(b.get("d_label", "Decides by vote"))
        legend = (
            '<div class="legend"><span class="rchip a">A</span> '
            'Accountable, one per row &nbsp; '
            '<span class="rchip r">R</span> Responsible &nbsp; '
            f'<span class="rchip d">D</span> {d_label} &nbsp; '
            '<span class="rchip c">C</span> Consulted &nbsp; '
            '<span class="rchip i">I</span> Informed</div>')
    note = f'<p class="tnote">{esc(b["note"])}</p>' if b.get("note") else ""
    inner = (f'<div class="twrap"><table class="{mode}">{colgroup}'
             f"<thead><tr>{head}</tr></thead><tbody>"
             + "".join(rows_html) + "</tbody></table></div>")
    return with_label(b, legend + inner) + note


def render_roadmap(b):
    phases = b["phases"]
    grid = sum(p["span"] for p in phases)
    col_tint = []
    for p in phases:
        col_tint += [p.get("tint", "#ffffff")] * p["span"]
    bandcls = {"ink": "rmb-ink", "amber": "rmb-amber", "light": "rmb-light"}
    head1 = ['<th class="rm-label-h"></th>']
    for p in phases:
        head1.append(
            f'<th class="rmband {bandcls.get(p.get("band", "light"))}" '
            f'colspan="{p["span"]}">{esc(p["label"])}</th>')
    head2 = ['<td class="rm-sub-h"></td>'] + [
        f'<td class="rm-sub">{esc(s)}</td>' for s in b.get("subs", [])]
    ms = {m["col"]: m for m in b.get("milestones", [])}
    msrow = ['<th class="rm-ms-h">MILESTONES</th>']
    for c in range(grid):
        if c in ms:
            m = ms[c]
            lead = f'<b>{esc(m["lead"])}</b> ' if m.get("lead") else ""
            msrow.append(f'<td class="rm-ms"><span class="dia">{DIA}</span>'
                         f'<br/>{lead}{esc(m.get("text", ""))}</td>')
        else:
            msrow.append('<td class="rm-ms"></td>')
    body = []
    for row in b["rows"]:
        cells = [f'<th class="rm-label" style="background:'
                 f'{pal(row.get("color", "blue"))}">'
                 f'{esc(row["label"])}</th>']
        col = 0
        for cell in row["cells"]:
            span = int(cell.get("span", 1))
            tint = cell.get("tint") or (
                col_tint[col] if col < len(col_tint) else "#ffffff")
            cells.append(f'<td class="rm-cell" colspan="{span}" '
                         f'style="background:{tint}">'
                         f'{esc(cell["text"])}</td>')
            col += span
        body.append("<tr>" + "".join(cells) + "</tr>")
    inner = ('<div class="rmwrap"><table class="rm"><thead><tr>'
             + "".join(head1) + "</tr><tr>" + "".join(head2)
             + '</tr><tr class="rm-ms-row">' + "".join(msrow)
             + "</tr></thead><tbody>" + "".join(body)
             + "</tbody></table></div>")
    return with_label(b, inner)


RENDERERS = {
    "prose": render_prose,
    "cards": render_cards,
    "band": render_band,
    "steps": render_steps,
    "list": render_list,
    "groups": render_groups,
    "timeline": render_timeline,
    "index": render_index,
    "table": render_table,
    "roadmap": render_roadmap,
}


# ----------------------------------------------------------- document ---

def render_cover(doc):
    parts = ['<section class="cover">']
    if doc.get("line_top"):
        lt = doc["line_top"]
        tag = {"lines": lt["tag"]} if lt.get("tag") else None
        parts.append(line_svg(lt["stations"], tag=tag, tag_at="start"))
    kicker = "<br/>".join(esc(x) for x in doc["kicker"])
    badge = ""
    if doc.get("line_name"):
        badge = (f'<div class="badge"><span class="bpill">'
                 f'{esc(doc["line_name"])}</span></div>')
    parts.append(f'<div class="kickrow"><div class="kicker">{kicker}'
                 f'</div>{badge}</div>')
    title = "<br/>".join(esc(x) for x in doc["title_lines"])
    parts.append(f'<h1 class="ctitle">{title}</h1>')
    if doc.get("subtitle"):
        parts.append(f'<p class="csub">{esc(doc["subtitle"])}</p>')
    if doc.get("deck"):
        parts.append(f'<p class="cdeck">{esc(doc["deck"])}</p>')
    if doc.get("cover_cards"):
        cells = "".join(
            f'<div class="ccard"><div class="lab">{esc(c["label"])}</div>'
            f'<div class="txt">{esc(c["text"])}</div></div>'
            for c in doc["cover_cards"])
        parts.append(f'<div class="ccards">{cells}</div>')
    if doc.get("cover_band"):
        band = doc["cover_band"]
        label = f'<div class="blab">{esc(band["label"])}</div>' \
            if band.get("label") else ""
        cells = "".join(f'<div class="bitem">{esc(it)}</div>'
                        for it in band["items"])
        parts.append(f'<div class="band">{label}'
                     f'<div class="bgrid">{cells}</div></div>')
    parts.append('<div class="spacer"></div>')
    author = esc(doc.get("author_line", ""))
    first = f"FIRST STOP: PAGE 02"
    parts.append(f'<div class="authorrow"><div class="authorline">'
                 f'{author}</div><div class="firststop">{first}</div>'
                 f'</div>')
    if doc.get("line_bottom"):
        lb = doc["line_bottom"]
        tag = {"lines": lb["tag"]} if lb.get("tag") else None
        parts.append(line_svg(lb["stations"], tag=tag, tag_at="end"))
    parts.append("</section>")
    return "".join(parts)


def render_stop(stop, number, total, doc):
    parts = ['<section class="stop">']
    parts.append(
        f'<div class="runfoot"><div class="dcell">'
        f'{esc(doc["doc_id"])}</div><div class="scell">'
        f'{strip_svg(total, number)}</div></div>')
    pill = ""
    if stop.get("pill"):
        pill = esc(stop["pill"])
    elif stop.get("question"):
        pill = f"Answers: {LQ}{esc(stop['question'])}{RQ}"
    parts.append(
        f'<div class="shead"><div class="wm">{esc(doc["wordmark"])}</div>'
        f'<div class="stopno">STOP {number:02d} / {total:02d}</div>'
        f'<div class="pill">{pill}</div></div>')
    parts.append('<div class="colorbar"><span class="cb1"></span>'
                 '<span class="cb2"></span><span class="cb3"></span>'
                 '<span class="cb4"></span></div>')
    if stop.get("kind") != "letter":
        parts.append(f'<h1 class="stitle">{esc(stop["title"])}</h1>')
    if stop.get("lede"):
        parts.append(f'<p class="lede">{esc(stop["lede"])}</p>')
    for b in stop.get("blocks", []):
        parts.append(RENDERERS[b["type"]](b))
    parts.append("</section>")
    return "".join(parts)


def build_html(spec):
    doc = spec["document"]
    total = 1 + len(spec["stops"])
    body = [render_cover(doc)]
    for i, stop in enumerate(spec["stops"]):
        body.append(render_stop(stop, i + 2, total, doc))
    return ("<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
            f"<title>{esc(doc.get('wordmark', 'Blueprint'))}</title>"
            f"<style>{CSS}</style></head><body>"
            + "".join(body) + "</body></html>")


# --------------------------------------------------------- validation ---

BLOCK_TYPES = set(RENDERERS)


def validate(spec):
    errs = []
    if not isinstance(spec, dict):
        return ["spec must be a JSON object"]
    doc = spec.get("document")
    if not isinstance(doc, dict):
        errs.append("document is required")
        doc = {}
    for f in ("wordmark", "kicker", "title_lines", "doc_id"):
        if not doc.get(f):
            errs.append(f"document.{f} is required")
    stops = spec.get("stops")
    if not isinstance(stops, list) or not stops:
        errs.append("stops must be a non-empty array")
        return errs
    for i, stop in enumerate(stops):
        p = f"stops[{i}]"
        if not stop.get("title"):
            errs.append(f"{p}.title is required")
        for j, b in enumerate(stop.get("blocks", [])):
            bp = f"{p}.blocks[{j}]"
            t = b.get("type")
            if t not in BLOCK_TYPES:
                errs.append(f"{bp}.type {t!r} is not a known block type")
                continue
            if t == "prose" and not b.get("paragraphs"):
                errs.append(f"{bp}.paragraphs is required")
            if t == "cards" and not b.get("cards"):
                errs.append(f"{bp}.cards is required")
            if t == "table" and not (b.get("columns") and b.get("rows")):
                errs.append(f"{bp}.columns and .rows are required")
            if t == "roadmap" and not (b.get("phases") and b.get("rows")):
                errs.append(f"{bp}.phases and .rows are required")
            if t in ("steps", "list", "timeline", "index") \
                    and not (b.get("steps") or b.get("items")
                             or b.get("entries")):
                errs.append(f"{bp} has no content items")
            if t == "groups" and not b.get("groups"):
                errs.append(f"{bp}.groups is required")
    return errs


def load_spec(path):
    raw = Path(path).read_text(encoding="utf-8")
    lines = [l for l in raw.splitlines()
             if not l.lstrip().startswith("//")]
    return json.loads("\n".join(lines))


def page_count(pdf_path):
    try:
        import pymupdf
        with pymupdf.open(pdf_path) as d:
            return d.page_count
    except Exception:
        pass
    try:
        out = subprocess.run(["pdfinfo", str(pdf_path)],
                             capture_output=True, text=True).stdout
        for line in out.splitlines():
            if line.startswith("Pages:"):
                return int(line.split(":")[1])
    except Exception:
        pass
    return None


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("spec", help="Blueprint spec JSON file")
    ap.add_argument("--out", required=True, help="Output directory")
    args = ap.parse_args(argv)

    try:
        spec = load_spec(args.spec)
    except (OSError, json.JSONDecodeError) as e:
        print(f"RESULT: INVALID - cannot load spec: {e}")
        return 1
    errs = validate(spec)
    if errs:
        for e in errs:
            print(f"RESULT: INVALID - {e}")
        return 1

    name = spec.get("name") or Path(args.spec).stem
    outdir = Path(args.out)
    outdir.mkdir(parents=True, exist_ok=True)
    pdf_path = outdir / f"{name}.pdf"
    png_path = outdir / f"{name}.png"

    html_doc = build_html(spec)
    try:
        from weasyprint import HTML
    except ImportError:
        print("RESULT: INVALID - weasyprint is not installed "
              "(see requirements.txt)")
        return 1
    with tempfile.NamedTemporaryFile("w", suffix=".html",
                                     delete=False,
                                     encoding="utf-8") as f:
        f.write(html_doc)
        tmp_html = f.name
    try:
        HTML(filename=tmp_html).write_pdf(str(pdf_path))
    finally:
        Path(tmp_html).unlink(missing_ok=True)

    total_stops = 1 + len(spec["stops"])
    pages = page_count(pdf_path)
    subprocess.run(["pdftoppm", "-png", "-r", "150", "-f", "1", "-l", "1",
                    str(pdf_path), str(outdir / name)],
                   check=True, capture_output=True)
    produced = sorted(outdir.glob(f"{name}-*.png"))
    if produced:
        produced[0].replace(png_path)
        for extra in produced[1:]:
            extra.unlink()

    print(f"RESULT: PASS")
    print(f"  pdf: {pdf_path}")
    print(f"  png: {png_path}")
    if pages is not None:
        print(f"  pages: {pages} (stops: {total_stops})")
        if pages != total_stops:
            print("  NOTE: page count differs from stop count; at "
                  "least one stop overflowed its page (content is "
                  "intact; tables repeat their header band).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
