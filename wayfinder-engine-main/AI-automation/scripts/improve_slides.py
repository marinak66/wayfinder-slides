"""
Improve specific slides in the Session 8 deck:
  - Slide 9  (index 8):  Query hygiene — two beautiful code blocks + footer
  - Slide 11 (index 10): Reveal table — dark-themed table with pair highlight
"""

from __future__ import annotations
from pathlib import Path
from copy import deepcopy
import lxml.etree as etree

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Palette — matched to the dark-background deck
# ---------------------------------------------------------------------------
DARK_BG     = RGBColor(0x04, 0x0D, 0x1A)   # near-black slide bg
CODE_BG     = RGBColor(0x07, 0x13, 0x22)   # slightly lifted code panel
CODE_BORDER = RGBColor(0x1A, 0x35, 0x55)   # subtle border for code blocks
TEAL        = RGBColor(0x57, 0xE5, 0xE3)   # primary teal accent
TEAL_DIM    = RGBColor(0x2A, 0x7F, 0x7D)   # dimmed teal for strips
CYAN        = RGBColor(0x7F, 0xE7, 0xE0)   # lighter teal
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
MUTED       = RGBColor(0xA9, 0xB6, 0xC4)   # secondary body text
ORANGE_KW   = RGBColor(0xFF, 0xA0, 0x47)   # DQL keyword colour
PURPLE_KW   = RGBColor(0xC0, 0x92, 0xFF)   # DQL field name colour
GREEN_KW    = RGBColor(0x7E, 0xD3, 0x21)   # DQL function
DIM_FILL    = RGBColor(0x0A, 0x1E, 0x35)   # table normal row
PAIR_FILL   = RGBColor(0x0D, 0x2A, 0x45)   # highlighted pair rows
PAIR_BORDER = RGBColor(0x57, 0xE5, 0xE3)   # teal border on pair rows
HDR_FILL    = RGBColor(0x08, 0x33, 0x66)   # table header
STAGE_FILL  = RGBColor(0x0B, 0x42, 0x7A)   # stage number column
NO_SHADOW_XML = "<a:effectLst/>"

FONT_HEAD = "DT Flow Extrabold"
FONT_BODY = "DT Flow"
FONT_MONO = "Consolas"          # code blocks

SW = Emu(12192000)              # 13.333"
SH = Emu(6858000)               #  7.500"


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------
def _no_shadow(shp):
    sp = shp._element
    sp_pr = sp.find(qn("p:spPr"))
    if sp_pr is None:
        return
    eff = sp_pr.find(qn("a:effectLst"))
    if eff is None:
        sp_pr.append(etree.fromstring(f'<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))


def _rect(slide, l, t, w, h, *, fill, border=None, border_w=1.0, rounding=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounding else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(kind, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if border:
        shp.line.color.rgb = border
        shp.line.width = Pt(border_w)
    else:
        shp.line.fill.background()
    _no_shadow(shp)
    return shp


def _textbox(slide, l, t, w, h, *, wrap=True, anchor=None):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left  = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top   = Inches(0)
    tf.margin_bottom= Inches(0)
    if anchor:
        tf.vertical_anchor = anchor
    return tf


def _run(para, text, *, font=FONT_BODY, size=None, bold=None, italic=None, color=WHITE):
    r = para.add_run()
    r.text = text
    r.font.name = font
    if size:  r.font.size = Pt(size)
    if bold is not None:  r.font.bold = bold
    if italic is not None: r.font.italic = italic
    r.font.color.rgb = color
    return r


def _para(tf, *, align=PP_ALIGN.LEFT, space_before=0, space_after=0, first=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after  = Pt(space_after)
    return p


def _remove_content_shapes(slide):
    """Remove every non-placeholder shape (textboxes, auto-shapes, tables)."""
    KEEP_IDX = {0, 1, 2, 3, 10, 11, 12}   # title + eyebrow + slide-num placeholders
    to_drop = []
    for shp in slide.shapes:
        try:
            ph = shp.placeholder_format
        except Exception:
            ph = None
        if ph is not None and ph.idx in KEEP_IDX:
            continue
        to_drop.append(shp._element)
    for el in to_drop:
        el.getparent().remove(el)


# ---------------------------------------------------------------------------
# Slide 9 — Query hygiene: two code blocks + closing line
# ---------------------------------------------------------------------------
RULE1 = {
    "label": "Rule 1",
    "headline": "Always filter out duplicates",
    "lines": [
        [("| fetch", ORANGE_KW), (" dt.davis.problems", WHITE)],
        [("| filter not(", ORANGE_KW), ("dt.davis.is_duplicate", PURPLE_KW), (")", ORANGE_KW)],
    ],
    "note": "Omitting this filter inflates every metric — counts, rates and trends.",
}
RULE2 = {
    "label": "Rule 2",
    "headline": "Confirm the root-cause field",
    "lines": [
        [("root_cause_entity_id", PURPLE_KW), ("       // classic tenant", MUTED)],
        [("root_cause.smartscape_entity", PURPLE_KW), ("  // 3rd-gen tenant", MUTED)],
        [("// include both in the dashboard", MUTED)],
    ],
    "note": "Field name differs by tenant generation — use both if your estate is mixed.",
}

FOOTER_TEXT = (
    "Cut-points in the tolerance bands (red / amber / green) are CoE starting values, "
    "not product SLAs.  Say this to customers explicitly."
)


def _code_block(slide, rule, left, top, width, height):
    gap = Inches(0.0)

    # outer card
    _rect(slide, left, top, width, height, fill=CODE_BG, border=CODE_BORDER, border_w=1.5, rounding=True)

    # teal header strip
    strip_h = Inches(0.56)
    _rect(slide, left, top, width, strip_h, fill=TEAL_DIM, rounding=True)
    # label pill
    pill_w = Inches(1.1)
    _rect(slide, left + Inches(0.22), top + Inches(0.1),
          pill_w, Inches(0.36), fill=TEAL, rounding=True)
    label_tf = _textbox(slide, left + Inches(0.22), top + Inches(0.1),
                        pill_w, Inches(0.36), anchor=MSO_ANCHOR.MIDDLE)
    lp = _para(label_tf, align=PP_ALIGN.CENTER, first=True)
    _run(lp, rule["label"], font=FONT_HEAD, size=13, bold=True, color=DARK_BG)

    # headline
    hl_tf = _textbox(slide, left + Inches(1.5), top + Inches(0.1),
                     width - Inches(1.72), Inches(0.36), anchor=MSO_ANCHOR.MIDDLE)
    hlp = _para(hl_tf, first=True)
    _run(hlp, rule["headline"], font=FONT_HEAD, size=14, bold=True, color=WHITE)

    # code body
    code_top = top + strip_h + Inches(0.18)
    code_h   = height - strip_h - Inches(0.52) - Inches(0.18)
    code_tf  = _textbox(slide, left + Inches(0.3), code_top,
                        width - Inches(0.6), code_h)
    code_tf.margin_top = Inches(0.05)
    first = True
    for token_line in rule["lines"]:
        lp2 = _para(code_tf, space_after=2, first=first)
        first = False
        for text, color in token_line:
            _run(lp2, text, font=FONT_MONO, size=13, color=color)

    # note line at bottom of card
    note_top = top + height - Inches(0.44)
    note_tf  = _textbox(slide, left + Inches(0.22), note_top,
                        width - Inches(0.44), Inches(0.38))
    np_ = _para(note_tf, first=True)
    _run(np_, rule["note"], font=FONT_BODY, size=11, italic=True, color=MUTED)


def rebuild_slide9(slide):
    _remove_content_shapes(slide)

    # layout constants
    content_top    = Inches(1.78)
    content_bottom = Inches(6.65)
    content_h      = content_bottom - content_top
    footer_h       = Inches(0.76)
    blocks_h       = content_h - footer_h - Inches(0.22)

    margin_l = Inches(0.5)
    total_w  = SW - margin_l * 2
    gap      = Inches(0.35)
    block_w  = (total_w - gap) / 2

    # Rule 1 (left)
    _code_block(slide, RULE1, margin_l, content_top, Emu(int(block_w)), Emu(int(blocks_h)))
    # Rule 2 (right)
    _code_block(slide, RULE2, margin_l + Emu(int(block_w)) + gap,
                content_top, Emu(int(block_w)), Emu(int(blocks_h)))

    # Full-width footer callout
    footer_top = Emu(int(content_top) + int(blocks_h) + int(Inches(0.22)))
    _rect(slide, margin_l, footer_top, total_w, Emu(int(footer_h)),
          fill=RGBColor(0x0B, 0x2F, 0x55), border=TEAL, border_w=1.5, rounding=True)
    # teal left accent bar
    _rect(slide, margin_l, footer_top, Inches(0.1), Emu(int(footer_h)),
          fill=TEAL, rounding=False)
    ftf = _textbox(slide, margin_l + Inches(0.25), footer_top,
                   total_w - Inches(0.45), Emu(int(footer_h)),
                   anchor=MSO_ANCHOR.MIDDLE)
    fp = _para(ftf, first=True)
    _run(fp, "⚠  ", font=FONT_HEAD, size=15, bold=True, color=TEAL)
    _run(fp, FOOTER_TEXT, font=FONT_BODY, size=14, color=WHITE)


# ---------------------------------------------------------------------------
# Slide 11 — Reveal table with pair highlight
# ---------------------------------------------------------------------------
REVEAL_ROWS = [
    # (symptom, stage, explanation, is_pair)
    ("40 auto-closing response time alerts",    "3",
     "Auto-close threshold too low — Stage 3 raises it",                   False),
    ("One event per region, 12 problems",       "4",
     "event.name contains region; fix: stabilise the name field",           False),
    ("No merge, topology present",              "5",
     "Correlation rule not configured — topology exists but rule is absent", True),
    ("No merge, no topology",                   "2",
     "Topology edge missing — fix topology before the rule works",           True),
    ("80% empty RCA",                           "6",
     "Root-cause field unpopulated; check RCA configuration",                False),
    ("Noise returned after six months",         "9",
     "Detection settings drifted — periodic review skipped",                 False),
]

PAIR_NOTE = (
    "Rows 3 & 4: identical surface symptom — the discriminator is one topology question."
)


def rebuild_slide11(slide):
    _remove_content_shapes(slide)

    # geometry
    tbl_left   = Inches(0.5)
    tbl_top    = Inches(1.78)
    tbl_w      = SW - Inches(1.0)
    tbl_bottom = Inches(6.50)
    note_h     = Inches(0.55)
    tbl_h      = tbl_bottom - tbl_top - note_h - Inches(0.18)

    n_rows  = 1 + len(REVEAL_ROWS)   # header + 6 data rows
    row_h   = Emu(int(tbl_h) // n_rows)

    # column widths: symptom 55%, stage 10%, explanation 35%
    col_w = [
        Emu(int(tbl_w) * 55 // 100),
        Emu(int(tbl_w) * 10 // 100),
        Emu(int(tbl_w) * 35 // 100),
    ]

    shape = slide.shapes.add_table(n_rows, 3, tbl_left, tbl_top, tbl_w, Emu(int(row_h) * n_rows))
    tbl   = shape.table
    for ci, cw in enumerate(col_w):
        tbl.columns[ci].width = cw
    tbl.first_row = True

    def cell_style(r, c, text, *, fill, text_color=WHITE, size=13,
                   bold=False, align=PP_ALIGN.LEFT, font=FONT_BODY,
                   border_color=None, border_w=1.0):
        cell = tbl.cell(r, c)
        cell.margin_left   = Inches(0.16)
        cell.margin_right  = Inches(0.10)
        cell.margin_top    = Inches(0.06)
        cell.margin_bottom = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = align
        rn = p.add_run()
        rn.text = text
        rn.font.name  = font
        rn.font.size  = Pt(size)
        rn.font.bold  = bold
        rn.font.color.rgb = text_color
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
        if border_color:
            _set_cell_border(cell, border_color, border_w)

    def _set_cell_border(cell, color, w_pt):
        """Apply a uniform border to a table cell via lxml."""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        w_emu = int(Pt(w_pt) * 12700)
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            existing = tcPr.find(qn(tag))
            if existing is not None:
                tcPr.remove(existing)
            ln = etree.SubElement(tcPr, qn(tag),
                                  attrib={"w": str(w_emu), "cap": "flat", "cmpd": "sng"})
            solidFill = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(solidFill, qn("a:srgbClr"),
                                    attrib={"val": f"{color}"})

    # Header row
    headers = ["Symptom", "Stage", "Explanation"]
    for c, h in enumerate(headers):
        cell_style(0, c, h, fill=HDR_FILL, text_color=TEAL,
                   size=12, bold=True, font=FONT_HEAD,
                   align=PP_ALIGN.CENTER if c == 1 else PP_ALIGN.LEFT)

    # Data rows
    for ri, (symptom, stage, explanation, is_pair) in enumerate(REVEAL_ROWS):
        r = ri + 1
        row_fill    = PAIR_FILL  if is_pair else DIM_FILL
        bdr_color   = PAIR_BORDER if is_pair else None
        bdr_w       = 1.5 if is_pair else 1.0

        # symptom — bold white (pair) or muted
        txt_color   = WHITE  if is_pair else RGBColor(0xD8, 0xE6, 0xF0)
        cell_style(r, 0, symptom, fill=row_fill, text_color=txt_color,
                   size=13, bold=is_pair, border_color=bdr_color, border_w=bdr_w)

        # stage number — teal badge column
        cell_style(r, 1, stage, fill=STAGE_FILL if not is_pair else TEAL_DIM,
                   text_color=TEAL if not is_pair else WHITE,
                   size=16, bold=True, font=FONT_HEAD, align=PP_ALIGN.CENTER,
                   border_color=bdr_color, border_w=bdr_w)

        # explanation
        cell_style(r, 2, explanation, fill=row_fill, text_color=MUTED,
                   size=12, border_color=bdr_color, border_w=bdr_w)

    # Pair note banner below the table
    note_top = Emu(int(tbl_top) + int(row_h) * n_rows + int(Inches(0.1)))
    _rect(slide, tbl_left, note_top, tbl_w, Emu(int(note_h)),
          fill=RGBColor(0x0A, 0x22, 0x3A), border=PAIR_BORDER, border_w=1.0, rounding=True)
    _rect(slide, tbl_left, note_top, Inches(0.09), Emu(int(note_h)),
          fill=PAIR_BORDER, rounding=False)
    ntf = _textbox(slide, tbl_left + Inches(0.24), note_top,
                   tbl_w - Inches(0.44), Emu(int(note_h)), anchor=MSO_ANCHOR.MIDDLE)
    np_ = _para(ntf, first=True)
    _run(np_, "🔗  ", font=FONT_HEAD, size=13, bold=True, color=PAIR_BORDER)
    _run(np_, PAIR_NOTE, font=FONT_BODY, size=12, italic=True, color=MUTED)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    src = Path("/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/inputs/"
               "60% validated - 8 - From Noise to Signal - Problem Quality & Root Cause Analysis - Draft v1 (1).pptx")
    out = Path("/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/"
               "projects/session-8/session-8-improved.pptx")
    out.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation(str(src))

    rebuild_slide9(prs.slides[8])
    print("Rebuilt slide 9  (Query hygiene — code blocks)")

    rebuild_slide11(prs.slides[10])
    print("Rebuilt slide 11 (Reveal table with pair highlight)")

    prs.save(str(out))
    print(f"\nSaved → {out}")


if __name__ == "__main__":
    main()
