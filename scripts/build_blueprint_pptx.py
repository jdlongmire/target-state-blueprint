#!/usr/bin/env python3
"""Build a NATIVE, editable PowerPoint blueprint from a JSON spec.

Companion to scripts/build_blueprint.py (the PDF builder). It takes
the same spec (see reference/content-spec.md) and renders the same
field-guide form -- cover with transit-line graphics, then one
numbered stop per slide, each answering one question -- as native
PowerPoint objects on US Letter portrait slides (8.5 x 11 in):
every text element is a real text frame, tables are native PPTX
tables, cards and bands are shapes carrying text. Nothing is a
pasted page image, so the deck stays fully editable. The PDF from
build_blueprint.py remains the fixed rendering of record; this PPTX
is the editable master.

Usage:
    python3 scripts/build_blueprint_pptx.py SPEC.json --out DIR

Writes <name>.pptx into DIR. <name> is the spec's "name" field, or
the spec file name when "name" is absent. Prints "RESULT: PASS"
with the file path and slide count, or "RESULT: INVALID" with the
exact field at fault and builds nothing. A NOTE prints when the
slide count differs from the stop count, so an overflow is never
silent.

Requires python-pptx and Pillow (Pillow is used only to measure text
with the Liberation fonts, which are metric-identical to Arial and
Courier New, so the layout math holds on the reader's machine).
Fonts named in the file: Arial for sans, Courier New for mono, bold
weight for display -- the same stacks the design system names, in
their metric-safe form.

Fidelity notes (where native constructs differ from the PDF):
- RACI letter chips are small rounded-rectangle shapes with text,
  overlaid on the native table at the exact cell positions; the
  letters are real text and move with manual edits only if the
  table geometry is left alone.
- RACI actor headers are rotated text boxes (rotation = 270)
  overlaid on the table's header band at the exact column
  positions: host support for table-cell text rotation is too
  uneven to trust, and a rotated text box is editable everywhere.
- If a stop overflows one slide, its splittable blocks (prose,
  list, timeline, index, table) continue on a following slide that
  repeats the stop's header furniture and footer; the PDF instead
  flows within one numbered stop. Non-splittable blocks move whole.
"""

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import ImageFont

# ---------------------------------------------------------------- palette
HEX = {
    "ink": "1A1D29", "body": "3D4149", "muted": "6A6E78",
    "blue": "2D5BA8", "orange": "B8501A", "teal": "12806F",
    "purple": "6A4C9C", "amber": "F4C542", "amber_dark": "9A7410",
    "amber_glyph": "DD9D0D", "hairline": "DCD5C4", "row_tint": "F3EFE6",
    "section_tint": "E6E0D2", "offwhite": "EEF0F4", "white": "FFFFFF",
}
TINTS = {"blue": "E3EAF6", "orange": "F7E4D8", "teal": "DDF0EC",
         "purple": "E9E3F3", "amber": "FFF6D9", "": "E6E0D2"}
CYCLE = ["blue", "orange", "teal", "purple"]
RGB = {k: RGBColor.from_string(v) for k, v in HEX.items()}

SANS = "Arial"
MONO = "Courier New"
_FONT_DIR = "/usr/share/fonts/truetype/liberation"
_FONT_FILES = {
    (SANS, False): f"{_FONT_DIR}/LiberationSans-Regular.ttf",
    (SANS, True): f"{_FONT_DIR}/LiberationSans-Bold.ttf",
    (MONO, False): f"{_FONT_DIR}/LiberationMono-Regular.ttf",
    (MONO, True): f"{_FONT_DIR}/LiberationMono-Bold.ttf",
}
_fonts = {}


def _font(font, bold):
    key = (font, bold)
    if key not in _fonts:
        path = _FONT_FILES.get(key)
        _fonts[key] = ImageFont.truetype(path, 200) if path else None
    return _fonts[key]


LH_BASE = 1.20   # host line height: PowerPoint/LibreOffice set a line
# at size x spacing x ~1.20 (the font's natural height), where CSS
# uses exactly size x spacing. All height math below uses this.
BODY_SCALE = 0.94  # body text (<=13pt spec size) renders at 94% so a
# stop sized for the CSS page still fits the host's taller lines.


def _sz(size):
    return size * BODY_SCALE if size <= 13 else size


def text_w(text, font, size, bold=False, spc=0.0):
    """Width of a string in inches. spc is letter spacing in points."""
    size = _sz(size)
    f = _font(font, bold)
    if f is None:
        w = len(text) * size * (0.6 if font == MONO else 0.52)
    else:
        w = f.getlength(text) / 200.0 * size
    return w / 72.0 + len(text) * spc / 72.0


def wrap(text, font, size, width, bold=False, spc=0.0):
    """Greedy word wrap; returns the list of lines for the width (inches)."""
    words = str(text).split()
    if not words:
        return [""]
    lines, cur = [], ""
    for word in words:
        trial = word if not cur else cur + " " + word
        if cur and text_w(trial, font, size, bold, spc) > width:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines


def lines_h(n, size, lh):
    return n * _sz(size) * lh * LH_BASE / 72.0


def _space_w(font, size, bold):
    return text_w(" ", font, size, bold)


def wrap_mixed(segments, font, size, width, spc=0.0):
    """Wrap a run of (text, bold) segments; returns the line count.

    Bold leads (band items, prose leads, list leads) are measurably
    wider than the same words in regular weight; measuring the whole
    string as regular undercounts lines and overflows fixed shapes.
    """
    words = []
    for text, bold in segments:
        for w in str(text).split():
            words.append((w, bold))
    if not words:
        return 1
    lines, cur = 1, 0.0
    for w, bold in words:
        ww = text_w(w, font, size, bold, spc)
        add = ww if cur == 0.0 else cur + _space_w(font, size, False) + ww
        if cur > 0.0 and add > width:
            lines += 1
            cur = ww
        else:
            cur = add
    return lines


# ---------------------------------------------------------------- geometry
PW, PH = 8.5, 11.0
MX = 0.75                    # side margin (54pt)
CW = PW - 2 * MX             # content width 7.0
TOP = 0.61                   # top margin (44pt)
CBOT = 10.12                 # content bottom limit
FOOT_Y = 10.30               # footer rule
LQ, RQ = "\u201c", "\u201d"
DOT = "\u00b7"
DIA_OK = True


# ---------------------------------------------------------------- helpers
def _noshadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass
    return shape


def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    _noshadow(sp)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = max(0.0, min(0.5, radius / min(w, h)))
        except Exception:
            pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGB[fill] if isinstance(fill, str) else fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGB[line] if isinstance(line, str) else line
        sp.line.width = Pt(line_w)
    sp.text_frame.word_wrap = True
    return sp


def tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    return tf


def para(tf, runs, align=PP_ALIGN.LEFT, lh=1.0, before=0.0, after=0.0, first=False):
    """runs: list of (text, font, size, bold, color, spc_pt)."""
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = lh
    if before:
        p.space_before = Pt(before)
    if after:
        p.space_after = Pt(after)
    for text, font, size, bold, color, spc in runs:
        r = p.add_run()
        r.text = str(text)
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = RGB[color] if isinstance(color, str) else color
        if spc:
            r.font._rPr.set("spc", str(int(spc * 100)))
    return p


def R(text, font=SANS, size=10.5, bold=False, color="body", spc=0.0):
    return (text, font, _sz(size), bold, color, spc)


def block_label(slide, x, y, label):
    tf = tb(slide, x, y, CW, 0.2)
    para(tf, [R(label.upper(), MONO, 7.5, True, "muted", 1.2)], first=True)
    return 0.24


def chip(slide, cx, cy, letter, kind):
    """A RACI letter chip centered at (cx, cy); returns nothing."""
    w, h = 0.34, 0.25
    fill, tcol, ln = {
        "A": ("ink", "white", None),
        "R": ("white", "ink", "ink"),
        "D": ("teal", "white", None),
        "C": ("section_tint", "ink", None),
        "I": ("section_tint", "ink", None),
    }[kind]
    sp = rect(slide, cx - w / 2, cy - h / 2, w, h, fill=fill, line=ln, line_w=1.0,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.055)
    tf = sp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    para(tf, [R(letter, SANS, 9.5, True, tcol)], align=PP_ALIGN.CENTER, first=True)


# ---------------------------------------------------------------- furniture
def draw_header(slide, doc, stop_no, total, stop):
    tf = tb(slide, MX, TOP, 2.2, 0.3)
    para(tf, [R(doc["wordmark"], SANS, 11, True, "ink", 2.2)], first=True)
    tf = tb(slide, MX + 2.2, TOP + 0.02, CW - 4.4, 0.3)
    para(tf, [R(f"STOP {stop_no:02d} / {total:02d}", MONO, 8.5, False, "ink", 1.6)],
         align=PP_ALIGN.CENTER, first=True)
    pill_text = stop.get("pill") or f"Answers: {LQ}{stop.get('question', '')}{RQ}"
    pw = min(CW * 0.62, text_w(pill_text, SANS, 8.5, True) + 0.28)
    sp = rect(slide, MX + CW - pw, TOP - 0.03, pw, 0.27, fill="ink",
              shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.135)
    tf = sp.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    para(tf, [R(pill_text, SANS, 8.5, True, "white")], align=PP_ALIGN.CENTER, first=True)
    # color bar: four segments at 2:1:1:2
    y = TOP + 0.36
    x = MX
    for name, frac in zip(CYCLE, (2, 1, 1, 2)):
        w = CW * frac / 6.0
        rect(slide, x, y, w + 0.004, 0.0625, fill=name)
        x += w
    return y + 0.0625


def draw_footer(slide, doc, stop_no, total):
    rect(slide, MX, FOOT_Y, CW, 0.011, fill="ink")
    doc_id = doc["doc_id"]
    id_lines = wrap(doc_id, MONO, 7.3, 2.6, spc=0.6)
    tf = tb(slide, MX, FOOT_Y + 0.07, 2.7, 0.3)
    for i, ln in enumerate(id_lines[:2]):
        para(tf, [R(ln, MONO, 7.3, False, "ink", 0.6)], lh=1.25, first=(i == 0))
    # progress strip
    strip_w = min(3.9, max(1.4, total * 0.16))
    x1 = MX + CW
    x0 = x1 - strip_w
    cy = FOOT_Y + 0.20
    rect(slide, x0, cy - 0.006, strip_w, 0.012, fill="muted")
    for i in range(1, total + 1):
        cx = x0 + strip_w * (i - 1) / max(1, total - 1) if total > 1 else x0
        if i == stop_no:
            d = 0.21
            sp = rect(slide, cx - d / 2, cy - d / 2, d, d, fill="amber", line="ink",
                      line_w=1.0, shape=MSO_SHAPE.OVAL)
            tf = sp.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                setattr(tf, m, 0)
            para(tf, [R(str(i), SANS, 7.5, True, "ink")], align=PP_ALIGN.CENTER, first=True)
        else:
            d = 0.075
            rect(slide, cx - d / 2, cy - d / 2, d, d, fill=CYCLE[(i - 1) % 4],
                 line="ink", line_w=0.75, shape=MSO_SHAPE.OVAL)


def new_stop_slide(prs, doc, stop_no, total, stop, with_title=True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar_bottom = draw_header(slide, doc, stop_no, total, stop)
    draw_footer(slide, doc, stop_no, total)
    y = bar_bottom + 0.17
    if with_title and stop.get("kind") != "letter":
        title = stop["title"]
        n = len(wrap(title, SANS, 26, CW, bold=True))
        tf = tb(slide, MX, y, CW, lines_h(n, 26, 1.1) + 0.05)
        para(tf, [R(title, SANS, 26, True, "ink")], lh=1.1, first=True)
        y += lines_h(n, 26, 1.1) + 0.10
    if with_title and stop.get("lede"):
        lede = stop["lede"]
        n = len(wrap(lede, SANS, 11.5, CW))
        tf = tb(slide, MX, y, CW, lines_h(n, 11.5, 1.45) + 0.03)
        para(tf, [R(lede, SANS, 11.5, False, "body")], lh=1.45, first=True)
        y += lines_h(n, 11.5, 1.45) + 0.14
    return slide, y


# ---------------------------------------------------------------- segments
class FixedSeg:
    def __init__(self, height, fn):
        self.height = height
        self.fn = fn
        self.done = False

    def draw(self, slide, y, max_h):
        self.fn(slide, y)
        self.done = True
        return self.height


def label_seg(label):
    def fn(slide, y):
        block_label(slide, MX, y, label)
    return FixedSeg(0.24, fn)


def prose_segs(block):
    segs = []
    paras = block["paragraphs"]
    for i, p in enumerate(paras):
        if isinstance(p, dict):
            runs_spec = [(p.get("lead", ""), True), (p.get("text", ""), False)]
        else:
            runs_spec = [(str(p), False)]
        n = wrap_mixed(runs_spec, SANS, 10.5, CW)
        h = lines_h(n, 10.5, 1.5) + (0.02 if i == len(paras) - 1 else 0.09)

        def make(runs_spec=runs_spec):
            def fn(slide, y):
                tf = tb(slide, MX, y, CW, 1.0)
                runs = [R(t, SANS, 10.5, b, "ink" if b else "body") for t, b in runs_spec if t]
                para(tf, runs, lh=1.5, first=True)
            return fn
        segs.append(FixedSeg(h, make()))
    return segs


def cards_seg(block):
    cols = int(block.get("columns", 2))
    gap_x, gap_y = 0.14, 0.12
    cardw = (CW - (cols - 1) * gap_x) / cols
    cards = block["cards"]
    rows = [cards[i:i + cols] for i in range(0, len(cards), cols)]

    def card_h(card):
        glyph = card.get("glyph")
        tw = cardw - 0.26 - (0.38 if glyph else 0)
        tl = len(wrap(card["title"], SANS, 11.5, tw, bold=True))
        title_area = max(0.27 if glyph else 0.0, lines_h(tl, 11.5, 1.22))
        xl = len(wrap(card.get("text", ""), SANS, 9.3, cardw - 0.26))
        return 0.13 + title_area + 0.05 + lines_h(xl, 9.3, 1.42) + 0.12

    row_hs = [max(card_h(c) for c in row) for row in rows]
    total = sum(row_hs) + gap_y * (len(rows) - 1)

    def fn(slide, y):
        cy = y
        for row, rh in zip(rows, row_hs):
            for j, card in enumerate(row):
                cx = MX + j * (cardw + gap_x)
                rect(slide, cx, cy, cardw, rh, fill="white", line="ink", line_w=1.3,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
                glyph = card.get("glyph")
                tx = cx + 0.13
                if glyph:
                    g = rect(slide, tx, cy + 0.13, 0.27, 0.27, fill=card.get("color", "blue"),
                             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
                    gtf = g.text_frame
                    gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
                    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                        setattr(gtf, m, 0)
                    para(gtf, [R(glyph, SANS, 10.5, True, "white")],
                         align=PP_ALIGN.CENTER, first=True)
                    tx += 0.38
                tw = cx + cardw - 0.13 - tx
                tl = len(wrap(card["title"], SANS, 11.5, tw, bold=True))
                title_h = lines_h(tl, 11.5, 1.22)
                title_area = max(0.27 if glyph else 0.0, title_h)
                ty = cy + 0.13 + (max(0.0, (title_area - title_h) / 2) if glyph else 0)
                tf = tb(slide, tx, ty, tw, title_h + 0.04)
                para(tf, [R(card["title"], SANS, 11.5, True, "ink")], lh=1.22, first=True)
                text = card.get("text", "")
                if text:
                    tf = tb(slide, cx + 0.13, cy + 0.13 + title_area + 0.05,
                            cardw - 0.26, rh)
                    para(tf, [R(text, SANS, 9.3, False, "body")], lh=1.42, first=True)
            cy += rh + gap_y
    return [FixedSeg(total, fn)]


def band_seg(block):
    items = block["items"]
    label = block.get("label")
    grid = bool(block.get("grid"))
    pad = 0.15
    if grid:
        cols = 4
        colw = (CW - 2 * pad - (cols - 1) * 0.2) / cols
        rows = [items[i:i + cols] for i in range(0, len(items), cols)]
        content_h = len(rows) * 0.235
    else:
        cols = max(1, len(items))
        colw = (CW - 2 * pad - (cols - 1) * 0.22) / cols
        content_h = 0.0
        for it in items:
            if isinstance(it, str):
                segs_in = [(it, False)]
            else:
                segs_in = ([(it.get("lead", ""), True)] if it.get("lead") else []) + \
                          [(it.get("text", ""), False)]
            content_h = max(content_h, lines_h(wrap_mixed(segs_in, SANS, 10, colw), 10, 1.45))
    total = pad + (0.26 if label else 0.0) + content_h + pad + 0.02

    def fn(slide, y):
        rect(slide, MX, y, CW, total, fill="ink", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             radius=0.09)
        iy = y + pad - 0.02
        if label:
            tf = tb(slide, MX + pad, iy, CW - 2 * pad, 0.22)
            para(tf, [R(label.upper(), MONO, 8, True, "amber", 1.4)], first=True)
            iy += 0.26
        if grid:
            for r, row in enumerate(rows):
                for j, it in enumerate(row):
                    tf = tb(slide, MX + pad + j * (colw + 0.2), iy + r * 0.235, colw, 0.22)
                    para(tf, [R(str(it), SANS, 10.5, False, "white")], first=True)
        else:
            for j, it in enumerate(items):
                x = MX + pad + j * (colw + 0.22)
                tf = tb(slide, x, iy, colw, content_h + 0.05)
                if isinstance(it, str):
                    para(tf, [R(it, SANS, 10, False, "offwhite")], lh=1.45, first=True)
                else:
                    runs = []
                    if it.get("lead"):
                        runs.append(R(it["lead"] + " ", SANS, 10, True, "white"))
                    if it.get("text"):
                        runs.append(R(it["text"], SANS, 10, False, "offwhite"))
                    para(tf, runs, lh=1.45, first=True)
    return [FixedSeg(total, fn)]


def steps_seg(block):
    steps = block["steps"]
    n = len(steps)
    gap = 0.26
    colw = (CW - (n - 1) * gap) / n

    def step_h(s):
        h = 0.44
        h += lines_h(len(wrap(s["title"], SANS, 11, colw, bold=True)), 11, 1.2)
        if s.get("owner"):
            h += 0.16
        h += lines_h(len(wrap(s.get("text", ""), SANS, 9.5, colw)), 9.5, 1.4)
        if s.get("tag"):
            h += 0.19
        return h

    total = max(step_h(s) for s in steps)

    def fn(slide, y):
        for j, s in enumerate(steps):
            x = MX + j * (colw + gap)
            color = s.get("color") or CYCLE[j % 4]
            tf = tb(slide, x, y, colw, 0.44)
            para(tf, [R(str(j + 1), SANS, 23, True, color)], first=True)
            ty = y + 0.44
            tf = tb(slide, x, ty, colw, 0.6)
            para(tf, [R(s["title"], SANS, 11, True, "ink")], lh=1.2, first=True)
            ty += lines_h(len(wrap(s["title"], SANS, 11, colw, bold=True)), 11, 1.2) + 0.03
            if s.get("owner"):
                tf = tb(slide, x, ty, colw, 0.16)
                para(tf, [R(s["owner"].upper(), MONO, 7, False, "muted", 0.8)], first=True)
                ty += 0.16
            if s.get("text"):
                tf = tb(slide, x, ty, colw, 1.0)
                para(tf, [R(s["text"], SANS, 9.5, False, "body")], lh=1.4, first=True)
                ty += lines_h(len(wrap(s["text"], SANS, 9.5, colw)), 9.5, 1.4) + 0.04
            if s.get("tag"):
                tf = tb(slide, x, ty, colw, 0.17)
                para(tf, [R(s["tag"].upper(), MONO, 7.5, True, "amber_dark", 0.8)], first=True)
            if j < n - 1:
                rect(slide, x + colw + 0.05, y + 0.10, 0.15, 0.24, fill="hairline",
                     shape=MSO_SHAPE.CHEVRON)
    return [FixedSeg(total, fn)]


def list_segs(block):
    color = block.get("color", "blue")
    segs = []
    items = block["items"]
    for i, it in enumerate(items):
        if isinstance(it, dict):
            lead = it.get("lead", "")
            rest = it.get("text", "")
        else:
            lead, rest = "", str(it)
        segs_in = ([(lead, True)] if lead else []) + [(rest, False)]
        n = wrap_mixed(segs_in, SANS, 10.5, CW - 0.28)
        h = lines_h(n, 10.5, 1.45) + 0.13

        def make(lead=lead, rest=rest, first=(i == 0)):
            def fn(slide, y):
                if not first:
                    rect(slide, MX, y - 0.065, CW, 0.01, fill="hairline")
                rect(slide, MX + 0.02, y + 0.075, 0.09, 0.09, fill=color, shape=MSO_SHAPE.OVAL)
                tf = tb(slide, MX + 0.28, y, CW - 0.28, 1.0)
                runs = []
                if lead:
                    runs.append(R(lead + " ", SANS, 10.5, True, "ink"))
                if rest:
                    runs.append(R(rest, SANS, 10.5, False, "body"))
                para(tf, runs, lh=1.45, first=True)
            return fn
        segs.append(FixedSeg(h, make()))
    return segs


def timeline_segs(block):
    segs = []
    items = block["items"]
    for i, it in enumerate(items):
        wn = len(wrap(it["when"], MONO, 8.5, 1.24, spc=0.3))
        tn = len(wrap(it["text"], SANS, 10, CW - 1.42))
        h = max(lines_h(wn, 8.5, 1.35), lines_h(tn, 10, 1.4)) + 0.13

        def make(it=it, first=(i == 0)):
            def fn(slide, y):
                if not first:
                    rect(slide, MX, y - 0.065, CW, 0.01, fill="hairline")
                tf = tb(slide, MX, y, 1.28, 0.5)
                para(tf, [R(it["when"], MONO, 8.5, False, "ink", 0.3)], lh=1.35, first=True)
                tf = tb(slide, MX + 1.42, y, CW - 1.42, 0.8)
                para(tf, [R(it["text"], SANS, 10, False, "body")], lh=1.4, first=True)
            return fn
        segs.append(FixedSeg(h, make()))
    return segs


def index_segs(block):
    entries = block["entries"]
    cols = 3
    gap = 0.24
    cellw = (CW - (cols - 1) * gap) / cols
    rows = [entries[i:i + cols] for i in range(0, len(entries), cols)]
    segs = []
    for row in rows:
        rh = 0.0
        for e in row:
            q = len(wrap(e["question"], SANS, 10.5, cellw, bold=True))
            rh = max(rh, lines_h(q, 10.5, 1.3) + 0.24)
        rh += 0.10

        def make(row=row, rh=rh):
            def fn(slide, y):
                for j, e in enumerate(row):
                    x = MX + j * (cellw + gap)
                    tf = tb(slide, x, y, cellw, rh)
                    para(tf, [R(e["question"], SANS, 10.5, True, "ink")], lh=1.3, first=True)
                    q = len(wrap(e["question"], SANS, 10.5, cellw, bold=True))
                    tf = tb(slide, x, y + lines_h(q, 10.5, 1.3) + 0.03, cellw, 0.18)
                    para(tf, [R("STOP " + str(e["stops"]), MONO, 8, True, "amber_dark", 0.8)],
                         first=True)
            return fn
        segs.append(FixedSeg(rh, make()))
    return segs


def groups_seg(block):
    groups = block["groups"]
    cols = len(groups)
    gap = 0.16
    colw = (CW - (cols - 1) * gap) / cols

    def col_h(g):
        h = 0.30 + 0.12
        for it in g["items"]:
            h += lines_h(len(wrap(it["title"], SANS, 10.5, colw, bold=True)), 10.5, 1.25)
            h += lines_h(len(wrap(it["text"], SANS, 9.3, colw)), 9.3, 1.4) + 0.10
        return h

    total = max(col_h(g) for g in groups)

    def fn(slide, y):
        for j, g in enumerate(groups):
            x = MX + j * (colw + gap)
            bar = rect(slide, x, y, colw, 0.30, fill=g["color"])
            tf = bar.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.1)
            para(tf, [R(g["label"].upper(), MONO, 8.5, True, "white", 1.0)], first=True)
            iy = y + 0.42
            for it in g["items"]:
                tf = tb(slide, x, iy, colw, 0.3)
                para(tf, [R(it["title"], SANS, 10.5, True, "ink")], lh=1.25, first=True)
                iy += lines_h(len(wrap(it["title"], SANS, 10.5, colw, bold=True)), 10.5, 1.25) + 0.02
                tf = tb(slide, x, iy, colw, 0.6)
                para(tf, [R(it["text"], SANS, 9.3, False, "body")], lh=1.4, first=True)
                iy += lines_h(len(wrap(it["text"], SANS, 9.3, colw)), 9.3, 1.4) + 0.10
    return [FixedSeg(total, fn)]


# ---------------------------------------------------------------- tables
def _set_cell_border(cell, color="DCD5C4", width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    tag_order = ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]
    for tag in reversed(tag_order):
        ln = tcPr.makeelement(qn(tag), {"w": str(int(width_pt * 12700)),
                                        "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = ln.makeelement(qn("a:srgbClr"), {"val": color})
        fill.append(clr)
        ln.append(fill)
        tcPr.insert(0, ln)


def _cell_text(cell, runs, align=PP_ALIGN.LEFT, lh=1.35, vert=None):
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.03)
    cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if vert:
        cell._tc.get_or_add_tcPr().set("vert", vert)
    tf = cell.text_frame
    tf.word_wrap = True
    para(tf, runs, align=align, lh=lh, first=True)


class TableSeg:
    def __init__(self, block):
        self.block = block
        self.mode = block.get("mode", "std")
        self.cols = block["columns"]
        n = len(self.cols)
        widths = block.get("widths") or [100.0 / n] * n
        self.colw = [CW * w / 100.0 for w in widths]
        self.rows = block["rows"]
        self.next_row = 0
        self.zebra = 0
        self.header_h = self._header_h()
        self.row_hs = [self._row_h(r) for r in self.rows]
        self.height = self.header_h + sum(self.row_hs)

    def _header_h(self):
        if self.mode == "raci":
            # Actor headers are drawn as rotated overlay text boxes, so
            # the band height is the longest label's vertical extent.
            h = 0.95
            for label in self.cols[1:]:
                h = max(h, text_w(label, MONO, 7.1, True, 0.5) + 0.16)
            return min(h, 1.50)
        h = 0.30
        for j, label in enumerate(self.cols):
            n = len(wrap(label, MONO, 7.3, self.colw[j] - 0.16, bold=True, spc=0.5))
            h = max(h, lines_h(n, 7.3, 1.3) + 0.07)
        return h

    def _row_h(self, row):
        if isinstance(row, dict):
            return 0.26
        if self.mode == "raci":
            tagged = any("::" in str(v) for v in row[1:])
            n = len(wrap(str(row[0]), SANS, 9, self.colw[0] - 0.16))
            return 0.44 if tagged else max(0.34, lines_h(n, 9, 1.35) + 0.07)
        h = 0.0
        for j, val in enumerate(row):
            n = len(wrap(str(val), SANS, 9, self.colw[j] - 0.16))
            h = max(h, lines_h(n, 9, 1.35) + 0.07)
        return h

    @property
    def done(self):
        return self.next_row >= len(self.rows)

    def draw(self, slide, y, max_h):
        if self.done:
            return 0
        avail = max_h
        if self.header_h + self.row_hs[self.next_row] > avail and self.next_row > 0:
            return None
        chunk = []
        used = self.header_h
        i = self.next_row
        while i < len(self.rows) and (used + self.row_hs[i] <= avail or not chunk):
            chunk.append(i)
            used += self.row_hs[i]
            i += 1
        self._draw_chunk(slide, y, chunk)
        self.next_row = i
        self.height = self.header_h + sum(self.row_hs[i:]) if not self.done else 0
        return used

    def _draw_chunk(self, slide, y, chunk):
        block = self.block
        n = len(self.cols)
        total_h = self.header_h + sum(self.row_hs[i] for i in chunk)
        gf = slide.shapes.add_table(len(chunk) + 1, n, Inches(MX), Inches(y),
                                    Inches(CW), Inches(total_h))
        table = gf.table
        table.first_row = False
        table.horz_banding = False
        for j, w in enumerate(self.colw):
            table.columns[j].width = Inches(w)
        table.rows[0].height = Inches(self.header_h)
        # header
        for j, label in enumerate(self.cols):
            cell = table.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGB["ink"]
            _set_cell_border(cell)
            if self.mode == "raci" and j > 0:
                # The label is drawn below as a rotated overlay text
                # box (host table-cell text rotation is unreliable);
                # the cell itself carries only a sized empty run so
                # the host does not price the row at its default
                # font size.
                _cell_text(cell, [R(" ", MONO, 7.1, False, "ink")],
                           align=PP_ALIGN.CENTER)
            else:
                _cell_text(cell, [R(label.upper(), MONO, 7.3, True, "white", 0.5)])
        if self.mode == "raci":
            for j, label in enumerate(self.cols[1:], start=1):
                cx = MX + sum(self.colw[:j]) + self.colw[j] / 2
                cy = y + self.header_h / 2
                w = self.header_h - 0.14
                shape = slide.shapes.add_textbox(Inches(cx - w / 2), Inches(cy - 0.12),
                                                 Inches(w), Inches(0.24))
                shape.rotation = 270
                tf = shape.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                    setattr(tf, m, 0)
                para(tf, [R(label, MONO, 7.1, True, "white", 0.5)],
                     align=PP_ALIGN.CENTER, first=True)
        overlays = []
        ry = y + self.header_h
        for k, ri in enumerate(chunk):
            row = self.rows[ri]
            table.rows[k + 1].height = Inches(self.row_hs[ri])
            if isinstance(row, dict):
                tint = TINTS.get(row.get("color", ""), TINTS[""])
                a = table.cell(k + 1, 0)
                if n > 1:
                    a.merge(table.cell(k + 1, n - 1))
                a.fill.solid()
                a.fill.fore_color.rgb = RGBColor.from_string(tint)
                _set_cell_border(a)
                _cell_text(a, [R(str(row.get("section", "")).upper(), MONO, 7.5, True, "ink", 1.0)])
                self.zebra = 0
            else:
                fill = "row_tint" if self.zebra % 2 == 1 else "white"
                self.zebra += 1
                for j, val in enumerate(row):
                    cell = table.cell(k + 1, j)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGB[fill]
                    _set_cell_border(cell)
                    text = str(val)
                    if self.mode == "raci" and j > 0:
                        letter, _, tag = text.partition("::")
                        letter = letter.strip()
                        if letter in ("A", "R", "D", "C", "I"):
                            # A sized space run keeps the host from
                            # pricing this row at its default font
                            # size; the chip itself is an overlay.
                            _cell_text(cell, [R(" ", SANS, 9, False, fill)],
                                       align=PP_ALIGN.CENTER)
                            cx = MX + sum(self.colw[:j]) + self.colw[j] / 2
                            cy = ry + self.row_hs[ri] / 2 - (0.09 if tag else 0.0)
                            overlays.append((cx, cy, letter, tag, ry, self.row_hs[ri], j))
                        else:
                            _cell_text(cell, [R("-", SANS, 9, False, "muted")],
                                       align=PP_ALIGN.CENTER)
                    elif j == 0:
                        _cell_text(cell, [R(text, SANS, 9, False, "ink")])
                    else:
                        align = PP_ALIGN.CENTER if self.mode == "raci" else PP_ALIGN.LEFT
                        _cell_text(cell, [R(text, SANS, 9, False, "body")], align=align)
            ry += self.row_hs[ri]
        for cx, cy, letter, tag, ry0, rh, j in overlays:
            chip(slide, cx, cy, letter, letter)
            if tag:
                tf = tb(slide, cx - self.colw[j] / 2, ry0 + rh / 2 + 0.045,
                        self.colw[j], 0.13)
                para(tf, [R(tag, SANS, 6.3, False, "muted")], align=PP_ALIGN.CENTER,
                     first=True)
        note = block.get("note")
        if note and self.done:
            tf = tb(slide, MX, y + total_h + 0.06, CW, 0.3)
            para(tf, [R(note, SANS, 8, False, "muted")], lh=1.35, first=True)


def raci_legend_seg(block):
    d_label = block.get("d_label", "Decides by vote")
    items = [("A", "Accountable, one per row"), ("R", "Responsible"),
             ("D", d_label), ("C", "Consulted"), ("I", "Informed")]

    def fn(slide, y):
        x = MX
        for letter, label in items:
            chip(slide, x + 0.17, y + 0.125, letter, letter)
            w = text_w(label, SANS, 8.5)
            tf = tb(slide, x + 0.42, y + 0.045, w + 0.1, 0.2)
            para(tf, [R(label, SANS, 8.5, False, "body")], first=True)
            x += 0.42 + w + 0.22
    return FixedSeg(0.27, fn)


def table_segs(block):
    segs = []
    if block.get("mode") == "raci":
        segs.append(raci_legend_seg(block))
    seg = TableSeg(block)
    segs.append(seg)
    if block.get("note"):
        seg.height += 0.30
    return segs


# ---------------------------------------------------------------- roadmap
def roadmap_seg(block):
    phases = block["phases"]
    span_total = sum(p["span"] for p in phases)
    pad = 0.14
    labelw = 1.02
    gridw = CW - 2 * pad - labelw
    colw = gridw / span_total
    col_phase = []
    for pi, p in enumerate(phases):
        col_phase += [pi] * p["span"]
    subs = block.get("subs") or []
    milestones = block.get("milestones") or []
    rows = block["rows"]

    def cell_lines(text, span):
        return len(wrap(text, SANS, 8.8, colw * span - 0.14))

    ms_h = 0.0
    for m in milestones:
        segs_in = ([(m.get("lead", ""), True)] if m.get("lead") else []) + \
                  [(m.get("text", ""), False)]
        ms_h = max(ms_h, lines_h(wrap_mixed(segs_in, SANS, 8.5, colw - 0.24), 8.5, 1.3) + 0.08)
    if milestones:
        ms_h = max(ms_h, 0.40)
    row_hs = []
    for row in rows:
        h = 0.34
        for cell in row["cells"]:
            h = max(h, lines_h(cell_lines(cell["text"], cell["span"]), 8.8, 1.32) + 0.12)
        lab = len(wrap(row["label"], SANS, 9.5, labelw - 0.16, bold=True))
        h = max(h, lines_h(lab, 9.5, 1.25) + 0.10)
        row_hs.append(h)
    total = pad + 0.30 + (0.20 if subs else 0.0) + ms_h + sum(row_hs) + pad + 0.06

    def fn(slide, y):
        rect(slide, MX, y, CW, total, fill="white", line="ink", line_w=1.3,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
        gx = MX + pad + labelw
        iy = y + pad
        # phase bands
        x = gx
        for p in phases:
            w = colw * p["span"]
            band = p.get("band", "light")
            fill = {"light": "section_tint", "amber": "amber", "ink": "ink"}[band]
            tcol = "white" if band == "ink" else "ink"
            sp = rect(slide, x, iy, w - 0.015, 0.30, fill=fill)
            tf = sp.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                setattr(tf, m, 0)
            para(tf, [R(p["label"].upper(), MONO, 8, True, tcol, 0.6)],
                 align=PP_ALIGN.CENTER, first=True)
            x += w
        iy += 0.30
        if subs:
            for c, sub in enumerate(subs):
                tf = tb(slide, gx + c * colw, iy + 0.025, colw, 0.17)
                para(tf, [R(sub, MONO, 7, False, "muted", 0.4)], align=PP_ALIGN.CENTER,
                     first=True)
            iy += 0.20
        if milestones:
            tf = tb(slide, MX + pad, iy + 0.02, labelw - 0.06, 0.4)
            para(tf, [R("MILESTONES", MONO, 7, False, "muted", 0.6)], lh=1.3, first=True)
            for m in milestones:
                cx = gx + m["col"] * colw
                rect(slide, cx + 0.01, iy + 0.035, 0.13, 0.13, fill="amber_glyph",
                     shape=MSO_SHAPE.DIAMOND)
                tf = tb(slide, cx + 0.20, iy, colw - 0.22, ms_h)
                runs = []
                if m.get("lead"):
                    runs.append(R(m["lead"] + " ", SANS, 8.5, True, "ink"))
                runs.append(R(m.get("text", ""), SANS, 8.5, False, "body"))
                para(tf, runs, lh=1.3, first=True)
            iy += ms_h
        for row, rh in zip(rows, row_hs):
            sp = rect(slide, MX + pad, iy, labelw, rh, fill=row["color"])
            tf = sp.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.08)
            tf.margin_right = Inches(0.04)
            para(tf, [R(row["label"], SANS, 9.5, True, "white")], lh=1.25, first=True)
            x = gx
            for cell in row["cells"]:
                w = colw * cell["span"]
                tint = cell.get("tint")
                if tint:
                    fill = tint.lstrip("#")
                    fill_rgb = RGBColor.from_string(fill)
                else:
                    ph = phases[col_phase[int(round((x - gx) / colw))]]
                    fill_rgb = RGBColor.from_string(ph.get("tint", "#FFFFFF").lstrip("#"))
                sp = rect(slide, x, iy, w - 0.015, rh, fill=fill_rgb, line="hairline",
                          line_w=0.75)
                tf = sp.text_frame
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                for m in ("margin_left", "margin_right"):
                    setattr(tf, m, Inches(0.06))
                for m in ("margin_top", "margin_bottom"):
                    setattr(tf, m, Inches(0.02))
                para(tf, [R(cell["text"], SANS, 8.8, False, "body")], lh=1.32, first=True)
                x += w
            iy += rh
    return [FixedSeg(total, fn)]


BLOCK_BUILDERS = {
    "prose": prose_segs,
    "cards": cards_seg,
    "band": band_seg,
    "steps": steps_seg,
    "list": list_segs,
    "timeline": timeline_segs,
    "index": index_segs,
    "groups": groups_seg,
    "table": table_segs,
    "roadmap": roadmap_seg,
}


def block_segments(block):
    segs = []
    label = block.get("label")
    if label and block["type"] != "band":
        segs.append(label_seg(label))
    segs.extend(BLOCK_BUILDERS[block["type"]](block))
    return segs


# ---------------------------------------------------------------- cover
def _station_label(slide, cx, y, label, mode):
    words = label.split()
    if len(label) > 9 and len(words) > 1:
        mid = (len(words) + 1) // 2
        text = " ".join(words[:mid]) + "\n" + " ".join(words[mid:])
    else:
        text = label
    if mode == "left":
        x, align = MX, PP_ALIGN.LEFT
    elif mode == "right":
        x, align = MX + CW - 1.1, PP_ALIGN.RIGHT
    else:
        x, align = cx - 0.55, PP_ALIGN.CENTER
    tf = tb(slide, x, y, 1.1, 0.3)
    for i, ln in enumerate(text.split("\n")):
        para(tf, [R(ln, MONO, 6.6, False, "body", 0.4)], align=align, lh=1.3,
             first=(i == 0))


def transit_line(slide, line_spec, y, tag_below=False):
    stations = line_spec.get("stations", [])
    n = len(stations)
    if n == 0:
        return
    x0, x1 = MX + 0.30, MX + CW - 0.30
    centers = [x0 + (x1 - x0) * i / (n - 1) for i in range(n)] if n > 1 else [x0]
    colors = []
    for i, st in enumerate(stations):
        c = st.get("color") if isinstance(st, dict) else None
        colors.append(c or CYCLE[i % 4])
    for i in range(n - 1):
        rect(slide, centers[i], y - 0.028, centers[i + 1] - centers[i] + 0.01, 0.056,
             fill=colors[i])
    tag = line_spec.get("tag")
    if tag and not tag_below:
        tw = 0.80
        sp = rect(slide, centers[0] - tw / 2, y - 0.62, tw, 0.42, fill="ink",
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
        tf = sp.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        for i, ln in enumerate(tag[:2]):
            para(tf, [R(ln, MONO, 7.5, True, "amber", 0.8)], align=PP_ALIGN.CENTER,
                 lh=1.25, first=(i == 0))
    for i, cx in enumerate(centers):
        first, last = i == 0, i == n - 1
        fill = "amber" if (first or last) else colors[i]
        d = 0.21
        rect(slide, cx - d / 2, y - d / 2, d, d, fill=fill, line="ink", line_w=2.0,
             shape=MSO_SHAPE.OVAL)
        label = stations[i]["label"] if isinstance(stations[i], dict) else str(stations[i])
        mode = "left" if first else ("right" if last else "center")
        _station_label(slide, cx, y + 0.17, label, mode)
    if tag and tag_below:
        tw = 0.80
        sp = rect(slide, centers[-1] - tw / 2, y + 0.50, tw, 0.42, fill="ink",
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
        tf = sp.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        for i, ln in enumerate(tag[:2]):
            para(tf, [R(ln, MONO, 7.5, True, "amber", 0.8)], align=PP_ALIGN.CENTER,
                 lh=1.25, first=(i == 0))


def render_cover(prs, doc):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if doc.get("line_top"):
        transit_line(slide, doc["line_top"], 1.06)
    y = 1.62
    kicker = doc.get("kicker", [])
    tf = tb(slide, MX, y, CW - 1.3, 0.85)
    for i, ln in enumerate(kicker):
        para(tf, [R(ln, MONO, 9.5, False, "ink", 1.8)], lh=1.8, first=(i == 0))
    if doc.get("line_name"):
        bx = MX + CW - 0.82
        rect(slide, bx, y - 0.06, 0.82, 0.82, fill=None, line="blue", line_w=3.5,
             shape=MSO_SHAPE.OVAL)
        name = doc["line_name"]
        pw = text_w(name, MONO, 8, True, 0.6) + 0.22
        sp = rect(slide, bx + 0.41 - pw / 2, y + 0.24, pw, 0.26, fill="ink")
        tf = sp.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, 0)
        para(tf, [R(name, MONO, 8, True, "amber", 0.6)], align=PP_ALIGN.CENTER, first=True)
    y += 0.48 + 0.34 * len(kicker)
    titles = doc.get("title_lines", [])
    tf = tb(slide, MX, y, CW, lines_h(len(titles), 63, 1.02) + 0.1)
    for i, ln in enumerate(titles):
        para(tf, [R(ln, SANS, 63, True, "ink")], lh=1.02, first=(i == 0))
    y += lines_h(len(titles), 63, 1.02) + 0.18
    if doc.get("subtitle"):
        tf = tb(slide, MX, y, CW, 0.4)
        para(tf, [R(doc["subtitle"], SANS, 16.5, True, "ink")], first=True)
        y += 0.36
    if doc.get("deck"):
        n = len(wrap(doc["deck"], SANS, 12.5, CW - 0.7))
        tf = tb(slide, MX, y, CW - 0.7, 0.35 * n)
        para(tf, [R(doc["deck"], SANS, 12.5, False, "body")], lh=1.35, first=True)
        y += n * 0.235 + 0.22
    cards = doc.get("cover_cards") or []
    if cards:
        gap = 0.16
        cardw = (CW - (len(cards) - 1) * gap) / len(cards)
        card_h = 0.0
        for c in cards:
            n = len(wrap(c["text"], SANS, 12, cardw - 0.3, bold=True))
            card_h = max(card_h, 0.16 + 0.24 + lines_h(n, 12, 1.3) + 0.16)
        for j, c in enumerate(cards):
            cx = MX + j * (cardw + gap)
            rect(slide, cx, y, cardw, card_h, fill="white", line="ink", line_w=1.3,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
            tf = tb(slide, cx + 0.15, y + 0.14, cardw - 0.3, 0.22)
            para(tf, [R(c["label"].upper(), MONO, 8, False, "muted", 1.5)], first=True)
            tf = tb(slide, cx + 0.15, y + 0.38, cardw - 0.3, card_h - 0.45)
            para(tf, [R(c["text"], SANS, 12, True, "ink")], lh=1.3, first=True)
        y += card_h + 0.20
    band = doc.get("cover_band")
    if band:
        items = band.get("items", [])
        label = band.get("label", "")
        cols = 4
        rows = [items[i:i + cols] for i in range(0, len(items), cols)]
        bh = 0.17 + 0.28 + len(rows) * 0.25 + 0.13
        rect(slide, MX, y, CW, bh, fill="ink", shape=MSO_SHAPE.ROUNDED_RECTANGLE,
             radius=0.09)
        tf = tb(slide, MX + 0.17, y + 0.13, CW - 0.34, 0.22)
        para(tf, [R(label.upper(), MONO, 8, True, "amber", 1.4)], first=True)
        colw = (CW - 0.34) / cols
        for r, row in enumerate(rows):
            for j, it in enumerate(row):
                tf = tb(slide, MX + 0.17 + j * colw, y + 0.42 + r * 0.25, colw, 0.22)
                para(tf, [R(str(it), SANS, 10.5, False, "white")], first=True)
        y += bh + 0.26
    if doc.get("author_line"):
        tf = tb(slide, MX, y, 4.9, 0.55)
        para(tf, [R(doc["author_line"], SANS, 10, False, "body")], lh=1.4, first=True)
    tf = tb(slide, MX + CW - 2.2, y + 0.06, 2.2, 0.22)
    para(tf, [R("FIRST STOP: PAGE 02", MONO, 8.5, False, "ink", 1.2)],
         align=PP_ALIGN.RIGHT, first=True)
    if doc.get("line_bottom"):
        transit_line(slide, doc["line_bottom"], 9.92, tag_below=True)
    return slide


# ---------------------------------------------------------------- packer
def render_stop(prs, doc, stop, stop_no, total):
    segs = []
    for block in stop["blocks"]:
        bsegs = block_segments(block)
        for k, s in enumerate(bsegs):
            s.block_start = (k == 0)
        segs.extend(bsegs)
    slide, y = new_stop_slide(prs, doc, stop_no, total, stop, with_title=True)
    top_y = y
    pages = 1
    i = 0
    while i < len(segs):
        seg = segs[i]
        gap = 0.11 if (getattr(seg, "block_start", False) and y > top_y + 0.001) else 0.0
        avail = CBOT - (y + gap)
        if seg.height <= avail:
            y += gap
            seg.draw(slide, y, avail)
            y += seg.height
            i += 1
        elif isinstance(seg, TableSeg):
            drawn = seg.draw(slide, y + gap, avail)
            if drawn:
                y += gap + drawn
            else:
                seg = seg  # fall through to a fresh page below
            if not seg.done:
                slide, y = new_stop_slide(prs, doc, stop_no, total, stop, with_title=False)
                top_y = y
                pages += 1
                if seg.done:
                    i += 1
            else:
                i += 1
        else:
            if y <= top_y + 0.001:
                # A single block taller than a fresh page: draw it anyway
                # (overflow) rather than loop forever; the count NOTE in
                # main() will not fire, so warn here.
                print(f"NOTE: stop {stop_no}: a block is taller than one slide "
                      f"and will overflow its slide", file=sys.stderr)
                seg.draw(slide, y, avail)
                y += seg.height
                i += 1
            else:
                slide, y = new_stop_slide(prs, doc, stop_no, total, stop, with_title=False)
                top_y = y
                pages += 1
    return pages


# ---------------------------------------------------------------- main
def load_spec(path):
    raw = Path(path).read_text(encoding="utf-8")
    lines = [l for l in raw.splitlines() if not l.lstrip().startswith("//")]
    return json.loads("\n".join(lines))


def fail(field):
    print(f"RESULT: INVALID {field}")
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    try:
        spec = load_spec(args.spec)
    except Exception as exc:
        fail(f"spec: cannot parse ({exc})")
    doc = spec.get("document") or fail("document")
    for field in ("wordmark", "kicker", "title_lines", "doc_id"):
        if not doc.get(field):
            fail(f"document.{field}")
    stops = spec.get("stops")
    if not stops:
        fail("stops")
    for i, s in enumerate(stops):
        if not s.get("title"):
            fail(f"stops[{i}].title")
        if "blocks" not in s:
            fail(f"stops[{i}].blocks")
        for j, b in enumerate(s["blocks"]):
            if b.get("type") not in BLOCK_BUILDERS:
                fail(f"stops[{i}].blocks[{j}].type")
    total = len(stops) + 1
    prs = Presentation()
    prs.slide_width = Inches(PW)
    prs.slide_height = Inches(PH)
    render_cover(prs, doc)
    slides = 1
    for i, stop in enumerate(stops):
        slides += render_stop(prs, doc, stop, i + 2, total)
    name = spec.get("name") or Path(args.spec).stem
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"{name}.pptx"
    prs.save(str(path))
    print(f"RESULT: PASS pptx={path} slides={slides} stops={total}")
    if slides != total:
        print(f"NOTE: slide count {slides} differs from stop count {total} "
              f"(a stop overflowed onto a continuation slide)")


if __name__ == "__main__":
    main()
