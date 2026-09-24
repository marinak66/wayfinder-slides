#!/usr/bin/env python3
"""Session 8 slide improvement script — v3.
Key fixes over v2:
- Preserve divider line (idx=15) and eyebrow (idx=14) on all slides
- Much larger font sizes throughout
- DQL panels respect picture left edges on slides 24-28
- Less cluttered layouts, more breathing room
- Follow reference deck design patterns
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from lxml import etree

INPUT  = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/inputs/60% validated - 8 - From Noise to Signal - Problem Quality & Root Cause Analysis - Draft v1 (1).pptx"
OUTPUT = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-v3.pptx"

# ---------- Palette ----------
TEAL        = RGBColor(0x57, 0xE5, 0xE3)
TEAL_DIM    = RGBColor(0x2A, 0x7F, 0x7D)
CYAN        = RGBColor(0x7F, 0xE7, 0xE0)
MUTED       = RGBColor(0xA9, 0xB6, 0xC4)
CODE_BG     = RGBColor(0x07, 0x13, 0x2A)
CODE_BORDER = RGBColor(0x1A, 0x35, 0x55)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DIM_FILL    = RGBColor(0x0A, 0x1E, 0x35)
HDR_FILL    = RGBColor(0x08, 0x33, 0x66)
ORANGE_KW   = RGBColor(0xFF, 0xA0, 0x47)
PURPLE_KW   = RGBColor(0xC0, 0x92, 0xFF)
GREEN_KW    = RGBColor(0x7E, 0xD3, 0x21)
WARN        = RGBColor(0xFF, 0xB5, 0x47)
GOOD        = RGBColor(0x4C, 0xAF, 0x50)
ERR         = RGBColor(0xF4, 0x43, 0x36)
NAVY        = RGBColor(0x0B, 0x1E, 0x3A)
PAIR_FILL   = RGBColor(0x0D, 0x2A, 0x45)
PAIR_BORDER = RGBColor(0x57, 0xE5, 0xE3)
AMBER       = RGBColor(0xFF, 0xB3, 0x00)
PURPLE_LIGHT= RGBColor(0xAB, 0x7E, 0xFF)
DARK_RED    = RGBColor(0x2A, 0x05, 0x05)

# ---------- Geometry ----------
SW = 12192000   # 13.333"
SH = 6858000    # 7.5"
LEFT_MARGIN = Inches(0.62)
RIGHT_MARGIN = Inches(0.62)
CONTENT_TOP = Inches(1.78)   # below divider line
CONTENT_W   = SW - LEFT_MARGIN - RIGHT_MARGIN   # 12.09"

# ---------- Fonts ----------
# DT Flow Extrabold — card/section headers, emphasis labels inside content
# DT Flow Heavy     — large metric/stat labels
# DT Flow           — body text, eyebrow, captions, bullets
# Consolas          — code
F_HEAD  = "DT Flow Extrabold"
F_HEAVY = "DT Flow Heavy"
F_BODY  = "DT Flow"
F_CODE  = "Consolas"

# ---------- Keep these placeholder indices (NEVER remove them) ----------
KEEP_PH_IDX = {0, 1, 2, 10, 11, 12, 14, 15, 17}


# ===================================================================
# Core helpers
# ===================================================================

def _find_layout(prs, name):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name.lower() == name.lower():
                return layout
    return prs.slide_masters[0].slide_layouts[11]


def _remove_content_shapes(slide, keep_images=True):
    """Remove added shapes but keep title (0), eyebrow (14), divider (15), etc."""
    to_drop = []
    for shp in slide.shapes:
        try:
            ph = shp.placeholder_format
            if ph is not None and ph.idx in KEEP_PH_IDX:
                continue
        except Exception:
            pass
        if keep_images:
            try:
                if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    continue
            except Exception:
                pass
        to_drop.append(shp._element)
    for el in to_drop:
        el.getparent().remove(el)


def _set_title(slide, text):
    """Update title placeholder (idx=0). Never set explicit font — let theme apply."""
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.idx == 0:
                tf = ph.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                r = p.add_run()
                r.text = text
                return
        except Exception:
            pass


def _set_eyebrow(slide, text, color=MUTED):
    """Update eyebrow placeholder (idx=14 then idx=1 fallback). Uses DT Flow body font."""
    for try_idx in (14, 1):
        for ph in slide.placeholders:
            try:
                if ph.placeholder_format.idx == try_idx:
                    tf = ph.text_frame
                    tf.clear()
                    p = tf.paragraphs[0]
                    r = p.add_run()
                    r.text = text
                    r.font.name = F_BODY
                    r.font.color.rgb = color
                    return True
            except Exception:
                pass
    return False


def _tb(slide, text, left, top, width, height,
        font=F_BODY, size=14, bold=False, color=WHITE,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.italic = italic
    return txBox


def _rect(slide, left, top, width, height,
          fill=None, line_color=None, line_pt=1.5):
    shp = slide.shapes.add_shape(1, left, top, width, height)
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line_color:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(line_pt)
    else:
        shp.line.fill.background()
    return shp


def _band(slide, left, top, width, height, fill, border_color=None, border_pt=1.5):
    """Draw a flat colored band/row."""
    return _rect(slide, left, top, width, height, fill=fill,
                 line_color=border_color, line_pt=border_pt)


def _footer(slide, text, fill=TEAL_DIM, text_color=WHITE, size=11, top=None):
    if top is None:
        top = Inches(6.45)
    h = Inches(0.4)
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, h, fill=fill)
    _tb(slide, text, LEFT_MARGIN + Inches(0.2), top + Inches(0.05),
        CONTENT_W - Inches(0.4), h - Inches(0.1),
        font=F_BODY, size=size, color=text_color, wrap=True)


def _section_tag(slide, text, left, top, width, color=TEAL):
    """Small uppercase label/tag."""
    _tb(slide, text.upper(), left, top, width, Inches(0.28),
        font=F_HEAD, size=10, bold=True, color=color)


def _code_box(slide, code_text, left, top, width, height, title=None):
    """Dark code panel with optional title strip."""
    _rect(slide, left, top, width, height, fill=CODE_BG, line_color=CODE_BORDER, line_pt=1.5)
    if title:
        _tb(slide, title, left + Inches(0.12), top + Inches(0.06), width - Inches(0.24), Inches(0.28),
            font=F_HEAD, size=10, bold=True, color=TEAL)
        code_top = top + Inches(0.38)
        code_h = height - Inches(0.44)
    else:
        code_top = top + Inches(0.1)
        code_h = height - Inches(0.2)
    tb = slide.shapes.add_textbox(left + Inches(0.15), code_top, width - Inches(0.3), code_h)
    tf = tb.text_frame
    tf.word_wrap = False
    first = True
    for line in code_text.split('\n'):
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.name = F_CODE
        r.font.size = Pt(10)
        r.font.color.rgb = WHITE


def _card(slide, left, top, width, height,
          header, body, border_color=TEAL, fill=DIM_FILL,
          h_size=13, b_size=11, h_color=None):
    _rect(slide, left, top, width, height, fill=fill, line_color=border_color, line_pt=2.0)
    hc = h_color if h_color else border_color
    _tb(slide, header, left + Inches(0.15), top + Inches(0.1), width - Inches(0.3), Inches(0.35),
        font=F_HEAD, size=h_size, bold=True, color=hc)
    _tb(slide, body, left + Inches(0.15), top + Inches(0.5), width - Inches(0.3), height - Inches(0.6),
        font=F_BODY, size=b_size, color=WHITE, wrap=True)


def _warn_bar(slide, text, top=None, fill=DARK_RED, border=ERR, size=12):
    if top is None:
        top = CONTENT_TOP
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, Inches(0.5), fill=fill, line_color=border, line_pt=2)
    _tb(slide, text, LEFT_MARGIN + Inches(0.2), top + Inches(0.08),
        CONTENT_W - Inches(0.4), Inches(0.38),
        font=F_BODY, size=size, bold=True, color=WARN, wrap=True)
    return top + Inches(0.6)


def _info_bar(slide, text, top, fill=DIM_FILL, border=TEAL, size=12):
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, Inches(0.55), fill=fill, line_color=border, line_pt=2)
    _tb(slide, text, LEFT_MARGIN + Inches(0.2), top + Inches(0.08),
        CONTENT_W - Inches(0.4), Inches(0.42),
        font=F_BODY, size=size, color=WHITE, wrap=True)
    return top + Inches(0.65)


def _set_cell_fill(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for el in tcPr.findall(qn('a:solidFill')):
        tcPr.remove(el)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    srgb = etree.SubElement(sf, qn('a:srgbClr'))
    srgb.set('val', f"{color}")


def _set_cell_border(cell, color, w_pt=1.0):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ('lnL', 'lnR', 'lnT', 'lnB'):
        ln = etree.SubElement(tcPr, qn(f'a:{side}'))
        ln.set('w', str(int(w_pt * 12700)))
        sf = etree.SubElement(ln, qn('a:solidFill'))
        srgb = etree.SubElement(sf, qn('a:srgbClr'))
        srgb.set('val', f"{color}")


def _cell_text(cell, text, font=F_BODY, size=11, bold=False,
               color=WHITE, align=PP_ALIGN.LEFT):
    tf = cell.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


# ===================================================================
# Horizontal band helpers (reference-style layouts)
# ===================================================================

BAND_FILLS = [
    RGBColor(0x08, 0x30, 0x65),
    RGBColor(0x09, 0x47, 0x6D),
    RGBColor(0x0B, 0x5E, 0x72),
    RGBColor(0x0D, 0x75, 0x75),
    RGBColor(0x0F, 0x8A, 0x78),
]

def _horiz_band_row(slide, left, top, width, height, fill, icon, label, desc,
                    label_size=16, desc_size=13, icon_size=22):
    _rect(slide, left, top, width, height, fill=fill)
    # icon circle
    circle_r = Inches(0.35)
    cx = left + Inches(0.2)
    cy = top + (height - circle_r) // 2
    circ = slide.shapes.add_shape(9, cx, cy, circle_r, circle_r)  # oval
    circ.fill.solid()
    circ.fill.fore_color.rgb = NAVY
    circ.line.fill.background()
    _tb(slide, icon, cx, cy - Inches(0.04), circle_r, circle_r + Inches(0.05),
        font=F_BODY, size=icon_size, align=PP_ALIGN.CENTER, color=WHITE)
    # label (monospace)
    label_left = left + Inches(0.75)
    label_w = Inches(3.0)
    _tb(slide, label, label_left, top + Inches(0.12), label_w, height - Inches(0.2),
        font=F_CODE, size=label_size, bold=True, color=WHITE)
    # vertical divider
    div = slide.shapes.add_shape(1, label_left + label_w + Inches(0.1), top + Inches(0.15),
                                  Pt(1.5), height - Inches(0.3))
    div.fill.solid()
    div.fill.fore_color.rgb = RGBColor(0x30, 0x60, 0x80)
    div.line.fill.background()
    # description
    desc_left = label_left + label_w + Inches(0.25)
    _tb(slide, desc, desc_left, top + Inches(0.12), width - (desc_left - left) - Inches(0.1),
        height - Inches(0.2), font=F_BODY, size=desc_size, color=WHITE, wrap=True)


# ===================================================================
# Per-slide rebuild functions
# ===================================================================

def rebuild_slide11(slide):
    """pptx 12 — How reports become events: identity tuple"""
    _set_title(slide, "How reports become events: identity tuple")
    _set_eyebrow(slide, "One report, one event — or many")
    _remove_content_shapes(slide, keep_images=False)

    half = Inches(5.65)
    gap = Inches(0.3)
    panel_h = Inches(2.0)

    # Left panel
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, panel_h, fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "Same tuple → event refreshes", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.4), font=F_HEAD, size=14, bold=True, color=TEAL)
    items = [("✓  event.name", "matches", TEAL),
             ("✓  event.type", "matches", TEAL),
             ("✓  source.entity.id", "matches", TEAL)]
    for i, (field, val, col) in enumerate(items):
        _tb(slide, f"{field}  —  {val}", LEFT_MARGIN + Inches(0.2),
            CONTENT_TOP + Inches(0.6) + i * Inches(0.38), half - Inches(0.4), Inches(0.35),
            font=F_BODY, size=12, color=col)
    _tb(slide, "→  Refreshes existing event", LEFT_MARGIN + Inches(0.2),
        CONTENT_TOP + Inches(1.55), half - Inches(0.4), Inches(0.35),
        font=F_HEAD, size=12, bold=True, color=WHITE)

    # Right panel
    rp = LEFT_MARGIN + half + gap
    _rect(slide, rp, CONTENT_TOP, half, panel_h, fill=DIM_FILL, line_color=ORANGE_KW, line_pt=2)
    _tb(slide, "Change any field → new event", rp + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.4), font=F_HEAD, size=14, bold=True, color=ORANGE_KW)
    items2 = [("✗  event.name", "CHANGED", ERR),
              ("✓  event.type", "matches", MUTED),
              ("✓  source.entity.id", "matches", MUTED)]
    for i, (field, val, col) in enumerate(items2):
        _tb(slide, f"{field}  —  {val}", rp + Inches(0.2),
            CONTENT_TOP + Inches(0.6) + i * Inches(0.38), half - Inches(0.4), Inches(0.35),
            font=F_BODY, size=12, color=col)
    _tb(slide, "→  Creates NEW separate event", rp + Inches(0.2),
        CONTENT_TOP + Inches(1.55), half - Inches(0.4), Inches(0.35),
        font=F_HEAD, size=12, bold=True, color=ORANGE_KW)

    # Identity tuple table
    tbl_top = CONTENT_TOP + panel_h + Inches(0.25)
    headers = ["Field", "Purpose", "Key rule"]
    col_ws = [Inches(2.8), Inches(3.0), Inches(5.6)]
    rows = [
        ["event.name", "Identifies the event", "Volatile values = new event every time"],
        ["event.type", "Semantic signal", "Carries platform meaning — choose carefully"],
        ["source.entity.id", "Entity binding", "Must bind to a real entity"],
        ["dt.event.correlation_id", "Grail identity", "Grail exposes this as the correlation key"],
    ]
    tbl = slide.shapes.add_table(5, 3, LEFT_MARGIN, tbl_top, CONTENT_W, Inches(2.1)).table
    for j, w in enumerate(col_ws): tbl.columns[j].width = w
    for j, h in enumerate(headers):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=11, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, row in enumerate(rows):
        hl = i == 1
        for j, val in enumerate(row):
            c = tbl.cell(i+1, j)
            _cell_text(c, val, font=F_CODE if j == 0 else F_BODY, size=11,
                       color=TEAL if hl else WHITE)
            _set_cell_fill(c, PAIR_FILL if hl else DIM_FILL)
            if hl: _set_cell_border(c, PAIR_BORDER)

    _footer(slide, "Same tuple → report refreshes the event.  Change any field → new, separate event.  Grail exposes the identity as dt.event.correlation_id.")


def rebuild_slide12(slide):
    """pptx 13 — Match category to condition — horizontal bands"""
    _set_title(slide, "Match the category to the condition")
    _set_eyebrow(slide, "event.type is a semantic signal, not a formality")
    _remove_content_shapes(slide, keep_images=False)

    # Warning bar
    next_top = _warn_bar(slide, "Do not default to CUSTOM_ALERT for error, slowdown or availability signals.", fill=DARK_RED, border=AMBER)

    # Tier 1 label
    _section_tag(slide, "Opens a Problem", LEFT_MARGIN, next_top, Inches(3), color=TEAL)
    next_top += Inches(0.32)

    # 4 horizontal bands — Tier 1
    tier1 = [
        ("AVAILABILITY_EVENT",          "A service or entity is unavailable or unreachable."),
        ("ERROR_EVENT",                 "An error rate or error condition is detected."),
        ("PERFORMANCE_EVENT",           "Performance degradation — latency, throughput."),
        ("RESOURCE_CONTENTION_EVENT",   "CPU, memory, disk or network saturation."),
    ]
    band_h = Inches(0.5)
    for i, (name, desc) in enumerate(tier1):
        fill = BAND_FILLS[i % len(BAND_FILLS)]
        _horiz_band_row(slide, LEFT_MARGIN, next_top + i * (band_h + Inches(0.04)),
                        CONTENT_W, band_h, fill, "⬤", name, desc, label_size=13, desc_size=11)
    next_top += len(tier1) * (band_h + Inches(0.04)) + Inches(0.15)

    # CUSTOM_ALERT special row
    _rect(slide, LEFT_MARGIN, next_top, CONTENT_W, Inches(0.48), fill=RGBColor(0x2A, 0x18, 0x00), line_color=AMBER, line_pt=1.5)
    _tb(slide, "CUSTOM_ALERT", LEFT_MARGIN + Inches(0.2), next_top + Inches(0.08), Inches(3.5), Inches(0.35),
        font=F_CODE, size=13, bold=True, color=AMBER)
    _tb(slide, "Last resort — use only when no built-in category fits.",
        LEFT_MARGIN + Inches(3.8), next_top + Inches(0.08), CONTENT_W - Inches(4.0), Inches(0.35),
        font=F_BODY, size=11, color=WHITE)
    next_top += Inches(0.58)

    # Tier 2
    _section_tag(slide, "Does NOT open a Problem", LEFT_MARGIN, next_top, Inches(4), color=MUTED)
    next_top += Inches(0.3)
    tier2 = [
        ("CUSTOM_INFO",           "Informational — no problem triggered."),
        ("CUSTOM_ANNOTATION",     "Deployment or change marker."),
        ("MARKED_FOR_TERMINATION","Entity marked for removal."),
    ]
    tier2_w = (CONTENT_W - Inches(0.4)) / 3
    for i, (name, desc) in enumerate(tier2):
        tl = LEFT_MARGIN + i * (tier2_w + Inches(0.2))
        _rect(slide, tl, next_top, tier2_w, Inches(0.65), fill=CODE_BG, line_color=CODE_BORDER, line_pt=1)
        _tb(slide, name, tl + Inches(0.1), next_top + Inches(0.06), tier2_w - Inches(0.2), Inches(0.3),
            font=F_CODE, size=11, color=MUTED)
        _tb(slide, desc, tl + Inches(0.1), next_top + Inches(0.36), tier2_w - Inches(0.2), Inches(0.25),
            font=F_BODY, size=9, color=MUTED)

    _footer(slide, "Match the category to the actual condition. A specific category gives Davis AI more context. CUSTOM_ALERT is a last resort.")


def rebuild_slide13(slide):
    """pptx 14 — Fragmentation patterns — keep images, add footer"""
    _remove_content_shapes(slide, keep_images=True)
    _footer(slide, "Rule of thumb: If one anomaly manifests across many dimensions, that dimension is data on one event — not a reason to emit one event per value. Keep high-cardinality context in event.description.", fill=TEAL_DIM)


def rebuild_slide14(slide):
    """pptx 15 — Keep images, add left-side bullet panel"""
    _remove_content_shapes(slide, keep_images=True)
    # Left bullet panel (images are on right ~5.5"+)
    panel_w = Inches(4.5)
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, panel_w, Inches(3.0), fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _section_tag(slide, "Shared fields enable merge", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.12), panel_w - Inches(0.3))
    items = [
        "Process, disk and network-interface events carry dt.smartscape.host.",
        "Kubernetes pod events carry k8s.namespace.name, k8s.workload.name and cluster fields.",
        "With the shared field present, one rule merges the whole vertical stack — no topology walk required.",
    ]
    for i, item in enumerate(items):
        _tb(slide, f"•  {item}", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.5) + i * Inches(0.78),
            panel_w - Inches(0.3), Inches(0.7), font=F_BODY, size=12, color=WHITE, wrap=True)


def rebuild_slide15(slide):
    """pptx 16 — Correlation rules — horizontal bands like reference"""
    _set_title(slide, "Correlation rules — definition and structure")
    _set_eyebrow(slide, "Definition and structure")
    _remove_content_shapes(slide, keep_images=False)

    # Subtitle italic
    _tb(slide, "Correlation rules merge only. They do not produce a Root Cause or Visual Resolution Path.",
        LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(0.35),
        font=F_BODY, size=12, italic=True, color=TEAL)

    # 5 horizontal bands
    bands = [
        ("🏷", "displayName",           "A unique, human-readable name you assign to the rule"),
        ("🔽", "matchingCondition",     "Which events the rule evaluates — a high-performance matcher query"),
        ("👥", "groupingFields",        "Events with identical values across all fields merge into one problem"),
        ("⏱", "timeWindow",            "How far back — compares event start to problem start. Default 15 min, max 1 day"),
        ("🔗", "correlationNamespace",  "Optional — lets several rules share one merge scope"),
    ]
    band_h = Inches(0.72)
    band_top = CONTENT_TOP + Inches(0.45)
    for i, (icon, label, desc) in enumerate(bands):
        fill = BAND_FILLS[i % len(BAND_FILLS)]
        _horiz_band_row(slide, LEFT_MARGIN, band_top + i * (band_h + Inches(0.05)),
                        CONTENT_W, band_h, fill, icon, label, desc, label_size=15, desc_size=12, icon_size=18)

    _footer(slide, "Correlation rules merge only. They do not produce a Root Cause or Visual Resolution Path.", fill=TEAL_DIM)


def rebuild_slide16(slide):
    """pptx 17 — Exercise 1 — 3 scenario cards"""
    _remove_content_shapes(slide, keep_images=False)

    # Instructions bar
    _info_bar(slide, "3 scenario cards per team  ·  Everyone does all cards  ·  7 minutes  ·  Debrief together",
              CONTENT_TOP, fill=TEAL_DIM, border=TEAL, size=13)
    card_top = CONTENT_TOP + Inches(0.7)
    card_w = (CONTENT_W - Inches(0.4)) / 3
    cards = [
        (TEAL,        "Card A · Audit",
         "A host's memory, disk and process problems — check the built-in rule first.\n\nKey question: does a shipped rule already cover this?"),
        (AMBER,       "Card B · Trap",
         "Per-region error-rate events with region in both the name and entity — write the rule.\n\nKey question: why is this impossible?"),
        (PURPLE_LIGHT,"Card C · Build",
         "A release across three services — deployment markers, warnings and errors under one namespace.\n\nKey question: what does the namespace actually solve?"),
    ]
    for i, (col, hdr, body) in enumerate(cards):
        cl = LEFT_MARGIN + i * (card_w + Inches(0.2))
        _card(slide, cl, card_top, card_w, Inches(3.8), hdr, body,
              border_color=col, fill=DIM_FILL, h_size=15, b_size=12)

    _tb(slide, "⏱  7 min", LEFT_MARGIN, card_top + Inches(4.0), CONTENT_W, Inches(0.9),
        font=F_HEAVY, size=40, bold=True, color=TEAL, align=PP_ALIGN.CENTER)


def rebuild_slide17(slide):
    """pptx 18 — Card A lesson"""
    _set_title(slide, "Card A lesson — the audit reflex")
    _set_eyebrow(slide, "Check what ships before you build")
    _remove_content_shapes(slide, keep_images=False)

    half = Inches(5.65)
    gap = Inches(0.3)

    _section_tag(slide, "The discipline", LEFT_MARGIN, CONTENT_TOP, half)
    bullets = [
        "The platform ships rules for same-entity, vertical stack, Kubernetes, and host-cloud grouping.",
        "Best practice: never edit a shipped rule directly.",
        "Disable it and create a new one with narrower scope — keeps a one-click revert path.",
    ]
    for i, b in enumerate(bullets):
        _tb(slide, f"•  {b}", LEFT_MARGIN, CONTENT_TOP + Inches(0.35) + i * Inches(0.85),
            half - Inches(0.1), Inches(0.75), font=F_BODY, size=13, color=WHITE, wrap=True)

    # Code panel right side
    json_code = """{
  "displayName": "Correlate host infrastructure by host",
  "enabled": true,
  "timeWindow": "15m",
  "groupingFields": ["dt.smartscape.host"],
  "matchingCondition": "matchesValue(
    dt.smartscape_source.type,
    {\\"CONTAINER\\", \\"DISK\\",
     \\"NETWORK_INTERFACE\\",
     \\"PROCESS\\", \\"HOST\\",
     \\"OS_SERVICE\\"})",
  "correlationNamespace":
    "default-dt.smartscape.host"
}"""
    _code_box(slide, json_code, LEFT_MARGIN + half + gap, CONTENT_TOP,
              half, Inches(4.0), title="Built-in rule — do not edit directly")


def rebuild_slide18(slide):
    """pptx 19 — Card B lesson"""
    _set_title(slide, "Card B lesson — the rule that cannot exist")
    _set_eyebrow(slide, "You cannot fix improper event design with a correlation rule")
    _remove_content_shapes(slide, keep_images=False)

    half = Inches(5.65)
    gap = Inches(0.3)
    ph = Inches(3.2)

    # Left — volatile identity
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=AMBER, line_pt=2)
    _tb(slide, "VOLATILE IDENTITY", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35), font=F_HEAD, size=13, bold=True, color=AMBER)
    _tb(slide, "event.name contains region", LEFT_MARGIN + Inches(0.15),
        CONTENT_TOP + Inches(0.55), half - Inches(0.3), Inches(0.4),
        font=F_CODE, size=14, bold=True, color=ORANGE_KW)
    _tb(slide, "source entity differs per region", LEFT_MARGIN + Inches(0.15),
        CONTENT_TOP + Inches(1.05), half - Inches(0.3), Inches(0.35),
        font=F_BODY, size=13, color=WHITE)
    _tb(slide, "Transferable question: Which field in the identity tuple carries a value that changes?",
        LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(1.55), half - Inches(0.3), Inches(0.8),
        font=F_BODY, size=12, color=MUTED, italic=True, wrap=True)

    # Right — fix upstream
    rp = LEFT_MARGIN + half + gap
    _rect(slide, rp, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "FIX UPSTREAM", rp + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35), font=F_HEAD, size=13, bold=True, color=TEAL)
    fixes = [
        "Fix at Stage 4: remove region from event.name.",
        "Bind to the real entity (not the per-region variant).",
        "Carry regions in event.description.",
        "Problem title inherits the first event's name — volatile values always mislead.",
    ]
    for i, f in enumerate(fixes):
        _tb(slide, f"•  {f}", rp + Inches(0.15), CONTENT_TOP + Inches(0.55) + i * Inches(0.62),
            half - Inches(0.3), Inches(0.55), font=F_BODY, size=12, color=WHITE, wrap=True)

    _footer(slide, "Even with a rule, you'd get a problem titled with the region name. The fix is always upstream.")


def rebuild_slide19(slide):
    """pptx 20 — Card C lesson"""
    _set_title(slide, "Card C lesson — what the namespace solves")
    _set_eyebrow(slide, "The namespace pulls context into the problem")
    _remove_content_shapes(slide, keep_images=False)

    _info_bar(slide, "One shared namespace (release-impact-correlation) ties deployment markers + warnings + errors into one problem.",
              CONTENT_TOP, border=TEAL, size=13)

    c_top = CONTENT_TOP + Inches(0.7)
    c_w = (CONTENT_W - Inches(0.4)) / 3
    rules = [
        ("Rule 1 — Deployment marker",
         'timeWindow: "1h"\nmatchingCondition:\n  event.type ==\n  CUSTOM_INFO\ngroupingFields:\n  [release_version]'),
        ("Rule 2 — Warning",
         'timeWindow: "30m"\nmatchingCondition:\n  event.type ==\n  WARNING\ngroupingFields:\n  [release_version]'),
        ("Rule 3 — Error / Perf",
         'timeWindow: "15m"\nmatchingCondition:\n  event.type in\n  [ERROR_EVENT,\n   PERFORMANCE_EVENT]\ngroupingFields:\n  [release_version]'),
    ]
    for i, (title, code) in enumerate(rules):
        cl = LEFT_MARGIN + i * (c_w + Inches(0.2))
        _code_box(slide, code, cl, c_top, c_w, Inches(2.8), title=title)

    bullets = [
        "The error rule merges across services on release_version alone.",
        "The namespace adds deployment markers and warnings into the same problem.",
        "Without it, different group keys — the explanatory context is lost.",
    ]
    b_top = c_top + Inches(3.0)
    for i, b in enumerate(bullets):
        _tb(slide, f"•  {b}", LEFT_MARGIN, b_top + i * Inches(0.38), CONTENT_W, Inches(0.35),
            font=F_BODY, size=12, color=WHITE, wrap=True)


def rebuild_slide20(slide):
    """pptx 21 — Before/After + reflection question + DQL"""
    _remove_content_shapes(slide, keep_images=True)

    q_top = Inches(4.5)
    _rect(slide, LEFT_MARGIN, q_top, CONTENT_W, Inches(0.58), fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "Which field on the events was the grouping key, and how did you know it existed before writing the rule?",
        LEFT_MARGIN + Inches(0.2), q_top + Inches(0.08), CONTENT_W - Inches(0.4), Inches(0.44),
        font=F_BODY, size=13, bold=True, color=WHITE, wrap=True)

    _code_box(slide, "fetch dt.davis.events\n| filter isNotNull(release_version)\n| summarize count(), by: {release_version, event.name}",
              LEFT_MARGIN, q_top + Inches(0.7), Inches(7.5), Inches(1.2),
              title="DQL — verify the grouping field exists")


def rebuild_slide22(slide):
    """pptx 23 — When RCA is legitimately empty — 5 tiles"""
    _set_title(slide, "When RCA is legitimately empty")
    _set_eyebrow(slide, "'No root cause' in a Problem could be expected behaviour")
    _remove_content_shapes(slide, keep_images=False)

    _info_bar(slide, "Causal RCA exists only where one event explains another. No chain to build → the panel correctly shows 'No root cause.'",
              CONTENT_TOP, border=TEAL, size=13)

    tiles = [
        ("🔍", "Single-event problem",
         "One event, one entity. It is both fault and cause.\nNothing upstream to trace."),
        ("🔗", "Same-entity merge",
         "Multiple events, one entity.\nNo second entity for a chain."),
        ("🏗", "Pure infrastructure stack",
         "Host, disk, process — no service layer,\nno topology walk needed."),
        ("⚙", "Custom alert — review",
         "Two checks: right event.category?\nEntity reference present?"),
        ("🌐", "Non-standard entity types",
         "Network, custom, extension entities —\noutside deterministic traversal."),
    ]
    tile_top = CONTENT_TOP + Inches(0.7)
    tile_w = (CONTENT_W - Inches(0.8)) / 5
    tile_h = Inches(2.8)
    for i, (icon, name, desc) in enumerate(tiles):
        tl = LEFT_MARGIN + i * (tile_w + Inches(0.2))
        _rect(slide, tl, tile_top, tile_w, tile_h, fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
        _tb(slide, icon, tl, tile_top + Inches(0.15), tile_w, Inches(0.65),
            font=F_BODY, size=28, align=PP_ALIGN.CENTER, color=WHITE)
        _tb(slide, name, tl + Inches(0.1), tile_top + Inches(0.85), tile_w - Inches(0.2), Inches(0.5),
            font=F_HEAD, size=11, bold=True, color=TEAL, wrap=True, align=PP_ALIGN.CENTER)
        _tb(slide, desc, tl + Inches(0.1), tile_top + Inches(1.4), tile_w - Inches(0.2), Inches(1.2),
            font=F_BODY, size=10, color=WHITE, wrap=True, align=PP_ALIGN.CENTER)

    _footer(slide, "Do not tune detection to force a root cause. Confirm the empty is expected before treating it as a defect.",
            fill=RGBColor(0x2A, 0x10, 0x00))


def rebuild_slide23(slide):
    """pptx 24 — Three dimensions of RCA quality"""
    _set_title(slide, "Three dimensions of RCA quality")
    _set_eyebrow(slide, "Populated RCA is not the same as correct")
    _remove_content_shapes(slide, keep_images=False)

    bars_data = [
        (TEAL,        1.0,  "1 · Attached",   "Is the root cause field populated? A query answers this.",           "QUERY"),
        (PURPLE_LIGHT,0.75, "2 · Correct",    "The named entity is the right one? Needs service-owner knowledge.", "DOMAIN"),
        (MUTED,       0.5,  "3 · Actionable", "Given a correct root cause, could someone act? Depends on #2.",    "PEOPLE"),
    ]
    bar_h = Inches(1.3)
    bar_gap = Inches(0.3)
    b_top = CONTENT_TOP + Inches(0.2)
    for i, (color, pct, label, desc, badge) in enumerate(bars_data):
        top = b_top + i * (bar_h + bar_gap)
        bw = int(CONTENT_W * pct)
        _rect(slide, LEFT_MARGIN, top, bw, bar_h, fill=DIM_FILL, line_color=color, line_pt=2)
        # Color accent strip
        _rect(slide, LEFT_MARGIN, top, Inches(0.28), bar_h, fill=color)
        _tb(slide, label, LEFT_MARGIN + Inches(0.4), top + Inches(0.12), Inches(3.0), Inches(0.45),
            font=F_HEAD, size=18, bold=True, color=color)
        _tb(slide, desc, LEFT_MARGIN + Inches(0.4), top + Inches(0.65), bw - Inches(2.0), Inches(0.5),
            font=F_BODY, size=13, color=WHITE, wrap=True)
        # Badge pill
        _rect(slide, LEFT_MARGIN + bw - Inches(1.15), top + Inches(0.4), Inches(1.05), Inches(0.38), fill=color)
        _tb(slide, badge, LEFT_MARGIN + bw - Inches(1.1), top + Inches(0.45), Inches(1.0), Inches(0.3),
            font=F_HEAD, size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    _footer(slide, "A populated field only clears the first of three bars.  |  CoE framework — not a native Dynatrace concept.", fill=RGBColor(0x1A, 0x10, 0x00))


# Slides 24-28: pictures are on RIGHT side; DQL panel on LEFT within available space
# Slide 24: pic starts at 5.50" → DQL left=0.62", width=4.5"
# Slide 25: pic starts at 5.15" → width=4.2"
# Slide 26: pic starts at 5.15" → width=4.2"
# Slide 27: pic starts at 3.94" → width=3.0" (tight — use smaller font)
# Slide 28: pic starts at 2.93" → place DQL below images, full width bottom strip

def rebuild_slide24_dql(slide):
    _remove_content_shapes(slide, keep_images=True)
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
    and isNull(root_cause.smartscape_entity.id)
| fieldsAdd
    affected_count = arraySize(smartscape.affected_entities),
    event_count = arraySize(dt.davis.event_ids),
    affected_str = toString(smartscape.affected_entities)
| fieldsAdd has_classic =
    matchesPhrase(affected_str, "SERVICE-") or
    matchesPhrase(affected_str, "HOST-") or
    matchesPhrase(affected_str, "PROCESS-")
| fieldsAdd bucket =
    if(event_count <= 1, "expected_single_event",
    else: if(affected_count == 1, "expected_same_entity",
    else: if(in("Infrastructure", dt.davis.impact_level)
        and arraySize(dt.davis.impact_level)==1,
        "expected_infra_stack",
    else: if(event.category == "CUSTOM_ALERT",
        "review_custom_alert",
    else: if(not(has_classic),
        "expected_non_classic_entity",
    else: "should_have_had")))))
| summarize problems=count(), by:{bucket}
| sort problems desc"""
    _code_box(slide, dql, LEFT_MARGIN, CONTENT_TOP, Inches(4.5), Inches(5.0), title="Master classification query")


def rebuild_slide25_dql(slide):
    _remove_content_shapes(slide, keep_images=True)
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
    and isNull(root_cause.smartscape_entity.id)
| fieldsAdd
    event_count = arraySize(dt.davis.event_ids)
| filter event_count <= 1"""
    _code_box(slide, dql, LEFT_MARGIN, CONTENT_TOP, Inches(4.2), Inches(2.2), title="Case 1 — expected_single_event")


def rebuild_slide26_dql(slide):
    _remove_content_shapes(slide, keep_images=True)
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
    and isNull(root_cause.smartscape_entity.id)
| fieldsAdd
    affected_count = arraySize(smartscape.affected_entities),
    event_count = arraySize(dt.davis.event_ids)
| filter event_count <= 1
| filter affected_count == 1"""
    _code_box(slide, dql, LEFT_MARGIN, CONTENT_TOP, Inches(4.2), Inches(2.4), title="Case 2 — expected_same_entity")


def rebuild_slide27_dql(slide):
    _remove_content_shapes(slide, keep_images=True)
    # Picture starts at 3.94" — only ~3" available on left
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
| fieldsAdd
    affected_count = arraySize(
      smartscape.affected_entities)
| filter in("Infrastructure",
    dt.davis.impact_level)
    and arraySize(
      dt.davis.impact_level) == 1"""
    _code_box(slide, dql, LEFT_MARGIN, CONTENT_TOP, Inches(3.0), Inches(2.5), title="Case 3 — expected_infra_stack")


def rebuild_slide28_dql(slide):
    _remove_content_shapes(slide, keep_images=True)
    # Picture starts at 2.93" — put DQL panel at bottom spanning full width
    dql = """| filter bucket == "should_have_had"
| filter event_count > 1
| filter arraySize(dt.davis.impact_level) > 1
| filter in("Services", dt.davis.impact_level)
    and in("Infrastructure", dt.davis.impact_level)"""
    _code_box(slide, dql, LEFT_MARGIN, Inches(5.2), CONTENT_W, Inches(1.4), title="Case 4 — should_have_had key filters")


def rebuild_slide29(slide):
    """pptx 30 — Exercise 2"""
    _remove_content_shapes(slide, keep_images=False)

    _warn_bar(slide, "The tests are ordered deliberately. A problem matches one bucket only — the first one it hits. Do not skip rows.", top=CONTENT_TOP)
    task_top = CONTENT_TOP + Inches(0.65)
    half = (CONTENT_W - Inches(0.3)) / 2
    _card(slide, LEFT_MARGIN, task_top, half, Inches(2.5),
          "Task 1 · Classify",
          "Assign every problem card to the correct bucket.\nWork top to bottom — stop at the first match.\n5 problem cards.",
          border_color=TEAL, fill=DIM_FILL, h_size=16, b_size=13)
    _card(slide, LEFT_MARGIN + half + Inches(0.3), task_top, half, Inches(2.5),
          "Task 2 · Investigate",
          "For the problem you classify as should_have_had (the genuine gap), write down two diagnostic steps.",
          border_color=PURPLE_LIGHT, fill=DIM_FILL, h_size=16, b_size=13)
    _tb(slide, "⏱  6 min", LEFT_MARGIN, task_top + Inches(2.7), CONTENT_W, Inches(1.0),
        font=F_HEAVY, size=48, bold=True, color=TEAL, align=PP_ALIGN.CENTER)


def rebuild_slide30(slide):
    """pptx 31 — Classification reference + 5 records"""
    _set_title(slide, "Classification reference + five records")
    _set_eyebrow(slide, "The five buckets and five problem records")
    _remove_content_shapes(slide, keep_images=False)

    ref_rows = [
        ("1", "expected_single_event",       "event_count <= 1"),
        ("2", "expected_same_entity",         "affected_count == 1"),
        ("3", "expected_infra_stack",          'in("Infrastructure", impact_level) AND arraySize == 1'),
        ("4", "review_custom_alert",           'event.category == "CUSTOM_ALERT"'),
        ("5", "expected_non_classic_entity",  "affected_count > 0 AND not(has_classic)"),
        ("6", "should_have_had",               "Falls through all above — genuine gap"),
    ]
    tbl = slide.shapes.add_table(7, 3, LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(1.8)).table
    tbl.columns[0].width = Inches(0.5)
    tbl.columns[1].width = Inches(4.2)
    tbl.columns[2].width = CONTENT_W - Inches(4.7)
    for j, h in enumerate(["#", "Bucket", "Rule"]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=11, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, (num, bucket, rule) in enumerate(ref_rows):
        hl = bucket == "should_have_had"
        fill = RGBColor(0x2A, 0x10, 0x00) if hl else DIM_FILL
        for j, val in enumerate([num, bucket, rule]):
            c = tbl.cell(i+1, j)
            _cell_text(c, val, font=F_CODE if j == 1 else F_BODY, size=10,
                       color=AMBER if hl else WHITE)
            _set_cell_fill(c, fill)

    # Five problem records
    p_rows = [
        ("P-1", "1", "1", "[Services]",              "CUSTOM_ALERT",        "review_custom_alert"),
        ("P-2", "3", "2", "[Services, Infra]",        "ERROR_EVENT",         "should_have_had"),
        ("P-3", "4", "1", "[Services]",               "ERROR_EVENT",         "expected_same_entity"),
        ("P-4", "2", "2", "[Infrastructure]",         "RESOURCE_CONTENTION", "expected_infra_stack"),
        ("P-5", "1", "1", "[Infrastructure]",         "AVAILABILITY",        "expected_single_event"),
    ]
    p_top = CONTENT_TOP + Inches(2.0)
    ptbl = slide.shapes.add_table(6, 6, LEFT_MARGIN, p_top, CONTENT_W, Inches(2.2)).table
    col_ws = [Inches(0.7), Inches(0.7), Inches(0.8), Inches(2.8), Inches(2.4), Inches(2.2)]
    for k, w in enumerate(col_ws): ptbl.columns[k].width = w
    for j, h in enumerate(["#", "Events", "Entities", "Impact level", "Category", "Bucket"]):
        _cell_text(ptbl.cell(0, j), h, font=F_HEAD, size=10, bold=True, color=TEAL)
        _set_cell_fill(ptbl.cell(0, j), HDR_FILL)
    for i, row in enumerate(p_rows):
        hl = row[0] == "P-2"
        fill = RGBColor(0x2A, 0x10, 0x00) if hl else DIM_FILL
        for j, val in enumerate(row):
            c = ptbl.cell(i+1, j)
            _cell_text(c, val, font=F_BODY, size=10, color=AMBER if hl else WHITE)
            _set_cell_fill(c, fill)


def rebuild_slide31(slide):
    """pptx 32 — Answer reveal"""
    _set_title(slide, "Answer reveal — which problem goes where")
    _set_eyebrow(slide, "The answers explained")
    _remove_content_shapes(slide, keep_images=False)

    answers = [
        ("P-1", "review_custom_alert",  "Custom alert — apply the two-check test."),
        ("P-2", "should_have_had",      "Multiple entities, service impact, no root cause — genuine gap."),
        ("P-3", "expected_same_entity", "Four events, one service — merging is working."),
        ("P-4", "expected_infra_stack", "Infrastructure-only, arraySize == 1."),
        ("P-5", "expected_single_event","One event, one entity."),
    ]
    tbl = slide.shapes.add_table(6, 3, LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(4.0)).table
    tbl.columns[0].width = Inches(0.9)
    tbl.columns[1].width = Inches(4.0)
    tbl.columns[2].width = CONTENT_W - Inches(4.9)
    for j, h in enumerate(["Problem", "Bucket", "Explanation"]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=12, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, (prob, bucket, expl) in enumerate(answers):
        hl = prob == "P-2"
        fill = PAIR_FILL if hl else DIM_FILL
        for j, val in enumerate([prob, bucket, expl]):
            c = tbl.cell(i+1, j)
            _cell_text(c, val, font=F_CODE if j == 1 else F_BODY, size=12,
                       color=TEAL if hl else WHITE)
            _set_cell_fill(c, fill)
            if hl: _set_cell_border(c, PAIR_BORDER)


def rebuild_slide32(slide):
    """pptx 33 — P-4 guard"""
    _set_title(slide, "P-4 — the guard")
    _set_eyebrow(slide, "What the array-size guard prevents")
    _remove_content_shapes(slide, keep_images=False)

    half = Inches(5.65)
    gap = Inches(0.3)
    ph = Inches(2.8)

    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "✓  PASSES", LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(0.12),
        half - Inches(0.4), Inches(0.4), font=F_HEAD, size=16, bold=True, color=TEAL)
    _tb(slide, 'impact_level = ["Infrastructure"]', LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(0.62),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=13, color=WHITE)
    _tb(slide, 'arraySize = 1', LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(1.1),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=13, color=WHITE)
    _tb(slide, '→  expected_infra_stack', LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(1.62),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=14, bold=True, color=GOOD)

    rp = LEFT_MARGIN + half + gap
    _rect(slide, rp, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=AMBER, line_pt=2)
    _tb(slide, "✗  CORRECTLY FAILS", rp + Inches(0.2), CONTENT_TOP + Inches(0.12),
        half - Inches(0.4), Inches(0.4), font=F_HEAD, size=16, bold=True, color=AMBER)
    _tb(slide, 'impact_level = ["Services", "Infrastructure"]', rp + Inches(0.2), CONTENT_TOP + Inches(0.62),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=13, color=WHITE)
    _tb(slide, 'arraySize = 2', rp + Inches(0.2), CONTENT_TOP + Inches(1.1),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=13, color=WHITE)
    _tb(slide, '→  does NOT match expected_infra_stack', rp + Inches(0.2), CONTENT_TOP + Inches(1.62),
        half - Inches(0.4), Inches(0.4), font=F_CODE, size=13, bold=True, color=ERR)

    q_top = CONTENT_TOP + ph + Inches(0.3)
    _rect(slide, LEFT_MARGIN, q_top, CONTENT_W, Inches(0.55), fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _tb(slide, "What does arraySize(impact_level) == 1 actually prevent?",
        LEFT_MARGIN + Inches(0.2), q_top + Inches(0.1), CONTENT_W - Inches(0.4), Inches(0.4),
        font=F_HEAD, size=14, bold=True, color=TEAL)
    _tb(slide, "Without it, [Services, Infrastructure] would also match in('Infrastructure', ...) and get labelled expected_infra_stack — masking a genuine gap.",
        LEFT_MARGIN, q_top + Inches(0.65), CONTENT_W, Inches(0.65),
        font=F_BODY, size=13, color=WHITE, wrap=True)


def rebuild_slide33(slide):
    """pptx 34 — P-1 two checks"""
    _set_title(slide, "P-1 — the two checks")
    _set_eyebrow(slide, "Custom alert: check the category first, then the reference")
    _remove_content_shapes(slide, keep_images=False)

    half = Inches(5.65)
    gap = Inches(0.3)
    ph = Inches(3.2)

    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "Step 1 — Was CUSTOM_ALERT the right category?",
        LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.1), half - Inches(0.3), Inches(0.5),
        font=F_HEAD, size=14, bold=True, color=TEAL, wrap=True)
    _tb(slide, "✓  Yes: A computed business ratio has no built-in category — CUSTOM_ALERT is correct.",
        LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.7), half - Inches(0.3), Inches(0.6),
        font=F_BODY, size=12, color=GOOD, wrap=True)
    _tb(slide, "✗  No: A service's traffic dropping to zero is AVAILABILITY — using CUSTOM_ALERT is a governance miss.",
        LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(1.42), half - Inches(0.3), Inches(0.65),
        font=F_BODY, size=12, color=ERR, wrap=True)

    rp = LEFT_MARGIN + half + gap
    _rect(slide, rp, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=PURPLE_LIGHT, line_pt=2)
    _tb(slide, "Step 2 — Does the template carry an entity reference?",
        rp + Inches(0.15), CONTENT_TOP + Inches(0.1), half - Inches(0.3), Inches(0.5),
        font=F_HEAD, size=14, bold=True, color=PURPLE_LIGHT, wrap=True)
    _tb(slide, "✓  Empty is correct when nothing can be bound.",
        rp + Inches(0.15), CONTENT_TOP + Inches(0.7), half - Inches(0.3), Inches(0.45),
        font=F_BODY, size=12, color=GOOD, wrap=True)
    _tb(slide, "⚠  If a Smartscape entity can be mapped, the template must not miss it.",
        rp + Inches(0.15), CONTENT_TOP + Inches(1.25), half - Inches(0.3), Inches(0.6),
        font=F_BODY, size=12, color=AMBER, wrap=True)

    result_top = CONTENT_TOP + ph + Inches(0.25)
    _rect(slide, LEFT_MARGIN, result_top, CONTENT_W, Inches(0.55), fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _tb(slide, "P-1 outcome: review_custom_alert — apply both checks before marking it resolved.",
        LEFT_MARGIN + Inches(0.2), result_top + Inches(0.1), CONTENT_W - Inches(0.4), Inches(0.38),
        font=F_BODY, size=14, bold=True, color=WHITE)


def rebuild_slide34(slide):
    """pptx 35 — P-2 investigation: keep image, add DQL"""
    _remove_content_shapes(slide, keep_images=True)
    dql = """| filter bucket == "should_have_had"
| filter event_count > 1
| filter arraySize(dt.davis.impact_level) > 1
| filter in("Services", dt.davis.impact_level)
    and in("Infrastructure", dt.davis.impact_level)"""
    _code_box(slide, dql, LEFT_MARGIN, CONTENT_TOP, Inches(5.5), Inches(2.0), title="Key filter — P-2 pattern")
    _rect(slide, LEFT_MARGIN, CONTENT_TOP + Inches(2.15), Inches(5.5), Inches(0.55),
          fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _tb(slide, "Finding: a relationship exists — this is not a missing-topology gap.",
        LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(2.25), Inches(5.3), Inches(0.38),
        font=F_BODY, size=13, bold=True, color=WHITE)


def insert_tcca_slide(prs, after_idx):
    layout = _find_layout(prs, "Title+eyebrow_left")
    new_slide = prs.slides.add_slide(layout)
    xml_slides = prs.slides._sldIdLst
    last_el = xml_slides[-1]
    xml_slides.remove(last_el)
    xml_slides.insert(after_idx + 1, last_el)

    _set_title(new_slide, "P-2 — Answer is TCCA")
    _set_eyebrow(new_slide, "How TopologyCausalChainAnalyzer builds a causal chain")

    _info_bar(new_slide,
              "TCCA turns five separate alerts into one problem with a root cause and a Visual Resolution Path. It only draws a causal link when two things are both connected by topology AND there is a causal signal.",
              CONTENT_TOP, border=TEAL, size=13)

    conditions = [
        (TEAL,        "🔗", "Topology link",  "A directed relationship exists between the affected entities in Smartscape."),
        (PURPLE_LIGHT,"📊", "Causal signal",  "One entity's anomaly temporally precedes the other's in a way consistent with causation."),
        (CYAN,        "✓",  "Traversal",      "TCCA walks the topology graph to find the root cause entity."),
    ]
    c_top = CONTENT_TOP + Inches(0.7)
    c_w = (CONTENT_W - Inches(0.4)) / 3
    for i, (col, icon, hdr, body) in enumerate(conditions):
        cl = LEFT_MARGIN + i * (c_w + Inches(0.2))
        _rect(new_slide, cl, c_top, c_w, Inches(2.2), fill=DIM_FILL, line_color=col, line_pt=2)
        _tb(new_slide, icon, cl, c_top + Inches(0.1), c_w, Inches(0.55),
            font=F_BODY, size=26, align=PP_ALIGN.CENTER, color=WHITE)
        _tb(new_slide, hdr, cl + Inches(0.15), c_top + Inches(0.7), c_w - Inches(0.3), Inches(0.4),
            font=F_HEAD, size=14, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(new_slide, body, cl + Inches(0.15), c_top + Inches(1.2), c_w - Inches(0.3), Inches(0.8),
            font=F_BODY, size=12, color=WHITE, wrap=True, align=PP_ALIGN.CENTER)

    _rect(new_slide, LEFT_MARGIN, c_top + Inches(2.4), CONTENT_W, Inches(0.65),
          fill=RGBColor(0x1A, 0x10, 0x00), line_color=AMBER, line_pt=1.5)
    _tb(new_slide, "A relationship exists but TCCA didn't draw the chain → the edge type may not be treated as causal. This is the interesting finding — not a missing-topology gap.",
        LEFT_MARGIN + Inches(0.2), c_top + Inches(2.5), CONTENT_W - Inches(0.4), Inches(0.5),
        font=F_BODY, size=12, color=AMBER, wrap=True)

    try:
        notes_slide = new_slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.clear()
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = ("TCCA is what turns five separate alerts into one problem with a root cause and a Visual Resolution Path. "
                  "It only draws a causal link when two things are both connected by topology AND there is a causal signal. "
                  "Key insight: in P-2 the topology link exists but the causal signal may not satisfy TCCA requirements for that edge type.")
    except Exception:
        pass
    return new_slide


def rebuild_slide35(slide):
    """pptx 36 — Honest conversation / Commit to the trend"""
    _set_title(slide, "Customer conversation and the honest position")
    _set_eyebrow(slide, "Commit to the trend, not the number")
    _remove_content_shapes(slide, keep_images=False)

    left_w = Inches(3.8)
    right_w = CONTENT_W - left_w - Inches(0.3)
    ph = Inches(4.0)

    # Left — headline alone
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, left_w, ph, fill=DIM_FILL, line_color=CODE_BORDER, line_pt=1.5)
    _tb(slide, "20%?", LEFT_MARGIN + Inches(0.1), CONTENT_TOP + Inches(0.25), left_w - Inches(0.2), Inches(1.3),
        font=F_HEAVY, size=56, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    _tb(slide, "RCA attachment", LEFT_MARGIN + Inches(0.1), CONTENT_TOP + Inches(1.55), left_w - Inches(0.2), Inches(0.35),
        font=F_BODY, size=13, color=MUTED, align=PP_ALIGN.CENTER)
    _tb(slide, '"Headline alone"', LEFT_MARGIN + Inches(0.1), CONTENT_TOP + Inches(2.0), left_w - Inches(0.2), Inches(0.35),
        font=F_BODY, size=12, color=AMBER, italic=True, align=PP_ALIGN.CENTER)
    _tb(slide, '"We are at 20% RCA attachment — the product is broken."',
        LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(2.5), left_w - Inches(0.3), Inches(0.9),
        font=F_BODY, size=12, color=WHITE, italic=True, wrap=True)

    # Right — breakdown
    rp = LEFT_MARGIN + left_w + Inches(0.3)
    _rect(slide, rp, CONTENT_TOP, right_w, ph, fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _tb(slide, "20% broken down:", rp + Inches(0.15), CONTENT_TOP + Inches(0.15), right_w - Inches(0.3), Inches(0.35),
        font=F_HEAD, size=14, bold=True, color=TEAL)
    buckets = [
        (TEAL,        "40%", "expected_single_event"),
        (CYAN,        "30%", "expected_same_entity"),
        (PURPLE_LIGHT,"15%", "expected_infra_stack"),
        (AMBER,       "10%", "review_custom_alert"),
        (ERR,         " 5%", "should_have_had  ← genuine work"),
    ]
    for i, (col, pct, label) in enumerate(buckets):
        top = CONTENT_TOP + Inches(0.6) + i * Inches(0.65)
        _rect(slide, rp + Inches(0.15), top + Inches(0.05), Inches(0.45), Inches(0.38), fill=col)
        _tb(slide, f"{pct}  {label}", rp + Inches(0.72), top + Inches(0.05), right_w - Inches(0.85), Inches(0.38),
            font=F_BODY, size=12, color=col if "genuine" in label else WHITE)

    _tb(slide, "📈  Trending down — the metric that stays true",
        rp + Inches(0.15), CONTENT_TOP + Inches(3.55), right_w - Inches(0.3), Inches(0.35),
        font=F_BODY, size=12, bold=True, color=TEAL)

    points = [
        "First move: run the bucket query before agreeing — headline alone cannot distinguish healthy from broken.",
        "A fixed percentage gets disproven the moment the environment changes.",
        "A trend stays true.",
    ]
    for i, pt in enumerate(points):
        _tb(slide, f"•  {pt}", LEFT_MARGIN, CONTENT_TOP + ph + Inches(0.15) + i * Inches(0.38),
            CONTENT_W, Inches(0.35), font=F_BODY, size=12, color=WHITE, wrap=True)


def rebuild_slide36(slide):
    """pptx 37 — Agentic comparison — proper table like reference"""
    _set_title(slide, "Agentic — two surfaces, shared reasoning, hard rule")
    _set_eyebrow(slide, "SRE Agent vs. Assist")
    _remove_content_shapes(slide, keep_images=False)

    # Rule bar
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(0.5),
          fill=RGBColor(0x1A, 0x08, 0x00), line_color=AMBER, line_pt=2)
    _tb(slide, "The Smartscape root cause panel and Visual Resolution Path are written exclusively by deterministic analysis. Agentic output never writes there.",
        LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(0.08), CONTENT_W - Inches(0.4), Inches(0.38),
        font=F_BODY, size=12, bold=True, color=AMBER, wrap=True)

    # Comparison table — 3 columns, proper sizing
    tbl_top = CONTENT_TOP + Inches(0.62)
    tbl = slide.shapes.add_table(6, 3, LEFT_MARGIN, tbl_top, CONTENT_W, Inches(3.2)).table
    tbl.columns[0].width = Inches(2.5)
    tbl.columns[1].width = Inches(4.5)
    tbl.columns[2].width = CONTENT_W - Inches(7.0)
    for j, (h, col) in enumerate([("FEATURE", WHITE), ("DYNATRACE ASSIST", TEAL), ("SRE AGENT (WORKFLOW)", PURPLE_LIGHT)]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=11, bold=True,
                   color=col, align=PP_ALIGN.CENTER)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    rows = [
        ("INVOCATION",  "On-demand, conversational",    "Autonomous, workflow-triggered"),
        ("LICENSE",     "Included — free",              "Paid (AI units + Grail + execution)"),
        ("OUTPUT",      "Explanation in chat",           "Comment on problem card"),
        ("BEST FOR",    '"Explain this to me"',           "Automatic first-response"),
        ("WHEN TO USE", "Ad-hoc investigation on a problem you're viewing",
                        "Systematic coverage on a scoped subset"),
    ]
    for i, (feat, assist, sre) in enumerate(rows):
        fill = CODE_BG if i % 2 == 0 else DIM_FILL
        _cell_text(tbl.cell(i+1, 0), feat, font=F_HEAD, size=11, bold=True, color=MUTED)
        _set_cell_fill(tbl.cell(i+1, 0), fill)
        _cell_text(tbl.cell(i+1, 1), assist, font=F_BODY, size=11, color=WHITE)
        _set_cell_fill(tbl.cell(i+1, 1), fill)
        _cell_text(tbl.cell(i+1, 2), sre, font=F_BODY, size=11, color=WHITE)
        _set_cell_fill(tbl.cell(i+1, 2), fill)

    # What they share
    share_top = tbl_top + Inches(3.35)
    _rect(slide, LEFT_MARGIN, share_top, CONTENT_W, Inches(0.5), fill=DIM_FILL)
    _tb(slide, "What they share:  Multi-step reasoning over logs, traces, metrics, events, deployments — narrative with evidence.",
        LEFT_MARGIN + Inches(0.2), share_top + Inches(0.08), CONTENT_W - Inches(0.4), Inches(0.38),
        font=F_BODY, size=12, color=WHITE)


def rebuild_slide37(slide):
    """pptx 38 — SRE Agent deployment"""
    _set_title(slide, "SRE Agent deployment")
    _set_eyebrow(slide, "Never run the template as-is")
    _remove_content_shapes(slide, keep_images=True)

    _warn_bar(slide, "The SRE Agent ships as a workflow template. Customize before running in production.", top=CONTENT_TOP)
    c_top = CONTENT_TOP + Inches(0.65)
    c_w = (CONTENT_W - Inches(0.4)) / 3
    concerns = [
        (ERR,         "⚠  Trigger scope",
         "Unscoped = budget gone fast.\nPreview invocation count with 'Query past events' first."),
        (AMBER,       "⚙  Prompt and tools",
         "Tight prompt + few tools completes in time.\nBroad prompt hits timeout."),
        (TEAL,        "✓  Duplicates excluded",
         "Mandatory filter.\nWithout it, you pay twice for the same incident."),
    ]
    for i, (col, hdr, body) in enumerate(concerns):
        cl = LEFT_MARGIN + i * (c_w + Inches(0.2))
        _card(slide, cl, c_top, c_w, Inches(2.2), hdr, body, border_color=col, fill=DIM_FILL, h_size=14, b_size=12)


def rebuild_slide38(slide):
    """pptx 39 — Assist on populated root cause"""
    _remove_content_shapes(slide, keep_images=True)
    _tb(slide, "Assist interprets and explains what deterministic RCA already named.",
        LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(0.4),
        font=F_BODY, size=14, color=TEAL, italic=True)
    _footer(slide, "A problem where deterministic RCA already named the root cause entity — Assist adds the narrative layer on top.")


def rebuild_slide39(slide):
    """pptx 40 — SRE Agent on empty RCA"""
    _set_title(slide, "SRE Agent on an empty root cause")
    _set_eyebrow(slide, "Agentic narrative where deterministic correctly has nothing")
    _remove_content_shapes(slide, keep_images=True)

    panel_w = Inches(5.0)
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, panel_w, Inches(2.4), fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
    _section_tag(slide, "Agent output includes:", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.12), panel_w - Inches(0.3))
    items = [
        "Verdict — recurring structural problem, not a transient spike.",
        "Recurrence history — same entity failing days earlier.",
        "Recommendations — in priority order, including the governance call.",
    ]
    for i, item in enumerate(items):
        _tb(slide, f"•  {item}", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.5) + i * Inches(0.6),
            panel_w - Inches(0.3), Inches(0.5), font=F_BODY, size=13, color=WHITE, wrap=True)


def rebuild_slide40(slide):
    """pptx 41 — Agentic analysis / auditability"""
    _set_title(slide, "Agentic Analysis")
    _set_eyebrow(slide, "The auditability beat")
    _remove_content_shapes(slide, keep_images=False)

    _rect(slide, LEFT_MARGIN, CONTENT_TOP, CONTENT_W, Inches(1.8), fill=NAVY, line_color=TEAL, line_pt=3)
    _tb(slide,
        '"No application-level logs were returned — the pod is crash-looping before it can emit logs. No distributed traces — the workload does not expose them."',
        LEFT_MARGIN + Inches(0.35), CONTENT_TOP + Inches(0.3), CONTENT_W - Inches(0.7), Inches(1.2),
        font=F_BODY, size=15, italic=True, color=WHITE, wrap=True)

    half = (CONTENT_W - Inches(0.3)) / 2
    b_top = CONTENT_TOP + Inches(2.0)
    _card(slide, LEFT_MARGIN, b_top, half, Inches(1.6),
          "The rule",
          "An agent that says 'no logs, because crash-looping' is auditable. An agent that stays silent is not.",
          border_color=TEAL, fill=DIM_FILL, h_size=14, b_size=12)
    _card(slide, LEFT_MARGIN + half + Inches(0.3), b_top, half, Inches(1.6),
          "The action",
          "One prompt instruction to SRE Agent: if evidence is unavailable, state so explicitly and explain why.",
          border_color=PURPLE_LIGHT, fill=DIM_FILL, h_size=14, b_size=12)

    _footer(slide, "Deterministic tells you which entity. Agentic tells you the story — and tells you what it couldn't find. Assist is free. The workflow requires a license.")


def rebuild_slide41(slide):
    """pptx 42 — Routing"""
    _set_title(slide, "Routing — the handoff")
    _set_eyebrow(slide, "Where a clean problem goes next")
    _remove_content_shapes(slide, keep_images=False)

    # Schema row
    boxes = [("Stages 2-6", DIM_FILL, TEAL),
             ("Clean problem", TEAL_DIM, TEAL),
             ("ITSM ticket", DIM_FILL, TEAL),
             ("On-call alert", DIM_FILL, AMBER),
             ("Chat notification", DIM_FILL, PURPLE_LIGHT)]
    bw = Inches(2.1)
    for i, (label, fill, border) in enumerate(boxes):
        bl = LEFT_MARGIN + i * (bw + Inches(0.28))
        _rect(slide, bl, CONTENT_TOP, bw, Inches(0.7), fill=fill, line_color=border, line_pt=2)
        _tb(slide, label, bl + Inches(0.1), CONTENT_TOP + Inches(0.18), bw - Inches(0.2), Inches(0.38),
            font=F_HEAD, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i < len(boxes) - 1:
            _tb(slide, "→", bl + bw + Inches(0.05), CONTENT_TOP + Inches(0.18), Inches(0.2), Inches(0.38),
                font=F_HEAD, size=16, color=TEAL)

    bullets = [
        "Pattern: problem trigger → filter on alert group + severity → open/close pairing.",
        "Clean, merged, deduplicated problems are what make downstream consumption work.",
        "Routing noise faster only scales the noise — Stage 7 sits after Stages 2–6.",
    ]
    for i, b in enumerate(bullets):
        _tb(slide, f"•  {b}", LEFT_MARGIN, CONTENT_TOP + Inches(0.9) + i * Inches(0.55),
            CONTENT_W, Inches(0.5), font=F_BODY, size=13, color=WHITE, wrap=True)


def rebuild_slide42(slide):
    """pptx 43 — Honest ITSM position + DQL"""
    _set_title(slide, "The honest ITSM position + DQL")
    _set_eyebrow(slide, "What downstream systems can consume today")
    _remove_content_shapes(slide, keep_images=False)

    half = (CONTENT_W - Inches(0.3)) / 2
    ph = Inches(2.3)

    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _tb(slide, "CONSUMABLE TODAY", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35), font=F_HEAD, size=13, bold=True, color=TEAL)
    items_l = [
        "Deterministic root cause — queryable field on the problem record.",
        "Consumable directly — no parsing needed.",
        "Direct ITSM consumption.",
    ]
    for i, item in enumerate(items_l):
        _tb(slide, f"•  {item}", LEFT_MARGIN + Inches(0.15), CONTENT_TOP + Inches(0.55) + i * Inches(0.55),
            half - Inches(0.3), Inches(0.5), font=F_BODY, size=12, color=WHITE, wrap=True)

    rp = LEFT_MARGIN + half + Inches(0.3)
    _rect(slide, rp, CONTENT_TOP, half, ph, fill=DIM_FILL, line_color=AMBER, line_pt=2)
    _tb(slide, "NEEDS A BRIDGE", rp + Inches(0.15), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35), font=F_HEAD, size=13, bold=True, color=AMBER)
    items_r = [
        "Agentic output: today a comment on problem (CUSTOM_ANNOTATION).",
        "Consuming it means a parsing job via standard workflow.",
        "Native field targeted for December Rally release.",
    ]
    for i, item in enumerate(items_r):
        _tb(slide, f"•  {item}", rp + Inches(0.15), CONTENT_TOP + Inches(0.55) + i * Inches(0.55),
            half - Inches(0.3), Inches(0.5), font=F_BODY, size=12, color=WHITE, wrap=True)

    _code_box(slide, 'fetch dt.davis.events.snapshots, from: "DATE", to: now()\n| filter event.type == "CUSTOM_ANNOTATION"\n| filter in(annotation.problem_ids, "YOUR_PROBLEM_ID")',
              LEFT_MARGIN, CONTENT_TOP + ph + Inches(0.2), CONTENT_W, Inches(1.5),
              title="Query agentic annotations on a problem")


def rebuild_slide43(slide):
    """pptx 44 — Governance cycle"""
    _set_title(slide, "Governance — Wayfinder Stage 9")
    _set_eyebrow(slide, "One-off tuning delivers temporary improvement only")
    _remove_content_shapes(slide, keep_images=False)

    c_w = (CONTENT_W - Inches(0.4)) / 3
    cards = [
        (TEAL,        "🔁  Quarterly", "Detector justification", "Review all active detectors.\nAsk: is this still relevant?"),
        (PURPLE_LIGHT,"📅  Monthly",   "Noise review",           "Check problem volume trend.\nIdentify new noisy sources."),
        (AMBER,       "🔄  On change", "Reassessment",           "New service, topology change,\ndetector added — reassess."),
    ]
    c_top = CONTENT_TOP + Inches(0.2)
    for i, (col, freq, title, desc) in enumerate(cards):
        cl = LEFT_MARGIN + i * (c_w + Inches(0.2))
        _rect(slide, cl, c_top, c_w, Inches(3.2), fill=DIM_FILL, line_color=col, line_pt=2)
        _tb(slide, freq, cl + Inches(0.15), c_top + Inches(0.15), c_w - Inches(0.3), Inches(0.55),
            font=F_HEAD, size=20, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(slide, title, cl + Inches(0.15), c_top + Inches(0.82), c_w - Inches(0.3), Inches(0.45),
            font=F_HEAD, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _tb(slide, desc, cl + Inches(0.15), c_top + Inches(1.38), c_w - Inches(0.3), Inches(0.9),
            font=F_BODY, size=13, color=MUTED, wrap=True, align=PP_ALIGN.CENTER)

    _footer(slide, "Configs decay as the environment changes. Wayfinder Stage 9 is mandatory, not optional.", fill=TEAL_DIM)


def rebuild_slide44(slide):
    """pptx 45 — Reinforcement Q&A"""
    _set_title(slide, "Reinforcement")
    _set_eyebrow(slide, "Quick recall")
    _remove_content_shapes(slide, keep_images=False)

    qa_pairs = [
        ("Q1 — A problem has four events, all on one service, and no deterministic root cause. Expected, or a gap?",
         "A:  Expected (same-entity merge). Multiple symptoms, one entity, merging is working. No second entity for a causal chain."),
        ("Q2 — A detector fires one event per Kubernetes pod. Which stage owns the fix?",
         "A:  Stage 4 — event design. The pod name is a volatile value in the identity tuple."),
    ]
    q_h = Inches(0.6)
    a_h = Inches(0.6)
    gap = Inches(0.2)
    q_top = CONTENT_TOP + Inches(0.2)
    pair_h = q_h + a_h + gap + Inches(0.35)
    for i, (q, a) in enumerate(qa_pairs):
        qt = q_top + i * pair_h
        _rect(slide, LEFT_MARGIN, qt, CONTENT_W, q_h, fill=TEAL_DIM)
        _tb(slide, q, LEFT_MARGIN + Inches(0.2), qt + Inches(0.1),
            CONTENT_W - Inches(0.4), q_h - Inches(0.15),
            font=F_BODY, size=14, bold=True, color=WHITE, wrap=True)
        _rect(slide, LEFT_MARGIN, qt + q_h + Inches(0.05), CONTENT_W, a_h,
              fill=DIM_FILL, line_color=GOOD, line_pt=2)
        _tb(slide, a, LEFT_MARGIN + Inches(0.2), qt + q_h + Inches(0.13),
            CONTENT_W - Inches(0.4), a_h - Inches(0.18),
            font=F_BODY, size=13, color=WHITE, wrap=True)


def rebuild_slide45(slide):
    """pptx 46 — Closing question"""
    _remove_content_shapes(slide, keep_images=False)

    ph = Inches(4.2)
    top = CONTENT_TOP + Inches(0.4)
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, ph, fill=NAVY, line_color=TEAL, line_pt=3)
    _tb(slide,
        "The next time a customer says RCA doesn't work —\nwhat's your first move before you agree with them?",
        LEFT_MARGIN + Inches(0.6), top + Inches(0.9), CONTENT_W - Inches(1.2), Inches(2.5),
        font=F_HEAVY, size=24, bold=True, color=TEAL, align=PP_ALIGN.CENTER, wrap=True)


# ===================================================================
# Slide deletion
# ===================================================================

def delete_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    rId = xml_slides[idx].get(qn('r:id'))
    prs.part.drop_rel(rId)
    del xml_slides[idx]
    print(f"  Deleted slide at original idx={idx}")


# ===================================================================
# Main
# ===================================================================

def main():
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    prs = Presentation(INPUT)
    print(f"Loaded: {len(prs.slides)} slides")

    ops = [
        (11, rebuild_slide11),
        (12, rebuild_slide12),
        (13, rebuild_slide13),
        (14, rebuild_slide14),
        (15, rebuild_slide15),
        (16, rebuild_slide16),
        (17, rebuild_slide17),
        (18, rebuild_slide18),
        (19, rebuild_slide19),
        (20, rebuild_slide20),
        (22, rebuild_slide22),
        (23, rebuild_slide23),
        (24, rebuild_slide24_dql),
        (25, rebuild_slide25_dql),
        (26, rebuild_slide26_dql),
        (27, rebuild_slide27_dql),
        (28, rebuild_slide28_dql),
        (29, rebuild_slide29),
        (30, rebuild_slide30),
        (31, rebuild_slide31),
        (32, rebuild_slide32),
        (33, rebuild_slide33),
        (34, rebuild_slide34),
        (35, rebuild_slide35),
        (36, rebuild_slide36),
        (37, rebuild_slide37),
        (38, rebuild_slide38),
        (39, rebuild_slide39),
        (40, rebuild_slide40),
        (41, rebuild_slide41),
        (42, rebuild_slide42),
        (43, rebuild_slide43),
        (44, rebuild_slide44),
        (45, rebuild_slide45),
    ]

    for idx, fn in ops:
        try:
            fn(prs.slides[idx])
            print(f"  Rebuilt idx={idx}")
        except Exception as e:
            print(f"  ERROR idx={idx}: {e}")

    insert_tcca_slide(prs, after_idx=34)
    print("  Inserted TCCA slide after idx=34")

    delete_slide(prs, 21)

    prs.save(OUTPUT)
    print(f"\nSaved → {OUTPUT}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
