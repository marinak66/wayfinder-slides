#!/usr/bin/env python3
"""Session 8 comprehensive slide improvement script — v2.
Processes all 34+ slides in one pass and saves a new PPTX.
"""
import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from lxml import etree

INPUT = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/inputs/60% validated - 8 - From Noise to Signal - Problem Quality & Root Cause Analysis - Draft v1 (1).pptx"
OUTPUT = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-v2.pptx"

# ---------- Palette ----------
TEAL       = RGBColor(0x57, 0xE5, 0xE3)
TEAL_DIM   = RGBColor(0x2A, 0x7F, 0x7D)
CYAN       = RGBColor(0x7F, 0xE7, 0xE0)
MUTED      = RGBColor(0xA9, 0xB6, 0xC4)
CODE_BG    = RGBColor(0x07, 0x13, 0x2A)
CODE_BORDER= RGBColor(0x1A, 0x35, 0x55)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DIM_FILL   = RGBColor(0x0A, 0x1E, 0x35)
HDR_FILL   = RGBColor(0x08, 0x33, 0x66)
ORANGE_KW  = RGBColor(0xFF, 0xA0, 0x47)
PURPLE_KW  = RGBColor(0xC0, 0x92, 0xFF)
GREEN_KW   = RGBColor(0x7E, 0xD3, 0x21)
WARN       = RGBColor(0xFF, 0xB5, 0x47)
GOOD       = RGBColor(0x4C, 0xAF, 0x50)
ERR        = RGBColor(0xF4, 0x43, 0x36)
NAVY       = RGBColor(0x0B, 0x1E, 0x3A)
PAIR_FILL  = RGBColor(0x0D, 0x2A, 0x45)
PAIR_BORDER= RGBColor(0x57, 0xE5, 0xE3)
DARK_BG    = RGBColor(0x06, 0x0F, 0x1E)
AMBER      = RGBColor(0xFF, 0xB3, 0x00)
PURPLE     = RGBColor(0x9C, 0x27, 0xB0)
PURPLE_LIGHT = RGBColor(0xAB, 0x7E, 0xFF)

# ---------- Slide geometry ----------
SW = 12192000   # slide width  EMU
SH = 6858000    # slide height EMU
ML = Inches(0.5)
MR = Inches(0.5)
CONTENT_TOP = Inches(1.78)
CONTENT_BOT = Inches(6.55)

# ---------- Fonts ----------
# DT Flow Extrabold — content card headers, labels, callout emphasis
# DT Flow Heavy     — prominent metric-style labels inside slides
# DT Flow           — all body text, eyebrow/subheader, captions
# Consolas          — code blocks
F_HEAD  = "DT Flow Extrabold"   # card/section headers inside content
F_HEAVY = "DT Flow Heavy"       # large emphasis labels
F_BODY  = "DT Flow"             # body text, eyebrow placeholder
F_CODE  = "Consolas"            # code
# Title placeholder inherits theme major font — never override with explicit font


# ============================================================
# Helper utilities
# ============================================================

def _find_layout(prs, name):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name.lower() == name.lower():
                return layout
    return prs.slide_masters[0].slide_layouts[11]  # Title+eyebrow_left fallback


def _remove_content_shapes(slide, keep_images=True):
    KEEP_PH_IDX = {0, 1, 2, 10, 11, 12}
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


def _update_placeholder(slide, idx, text, size=None, bold=None, color=None, font=None, italic=None):
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.idx == idx:
                tf = ph.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                r = p.add_run()
                r.text = text
                # Never set explicit font on title (idx=0) — let theme major font apply
                if font and idx != 0:
                    r.font.name = font
                if size:
                    r.font.size = Pt(size)
                if bold is not None and idx != 0:
                    r.font.bold = bold
                if color:
                    r.font.color.rgb = color
                if italic is not None:
                    r.font.italic = italic
                return
        except Exception:
            pass


def _update_eyebrow(slide, text, color=None, size=None):
    """Update eyebrow/subheader placeholder — tries idx=14 then idx=1."""
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
                    if size:
                        r.font.size = Pt(size)
                    if color:
                        r.font.color.rgb = color
                    return True
            except Exception:
                pass
    return False


def _add_textbox(slide, text, left, top, width, height,
                 font=F_BODY, size=11, bold=False, color=WHITE,
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


def _add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width_pt=1.5):
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE_TYPE.RECTANGLE=1
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def _set_cell_fill(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
    srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgb.set('val', f"{color}")


def _set_cell_border(cell, color, w_pt=1.0):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ('lnL', 'lnR', 'lnT', 'lnB'):
        ln = etree.SubElement(tcPr, qn(f'a:{side}'))
        ln.set('w', str(int(w_pt * 12700)))
        solidFill = etree.SubElement(ln, qn('a:solidFill'))
        srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
        srgb.set('val', f"{color}")


def _cell_text(cell, text, font=F_BODY, size=10, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
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


def _footer_bar(slide, text, top=None, fill=TEAL_DIM, text_color=WHITE, font_size=10):
    if top is None:
        top = Inches(6.45)
    bar = _add_rect(slide, ML, top, SW - ML - MR, Inches(0.35), fill_color=fill)
    txBox = slide.shapes.add_textbox(ML + Inches(0.15), top + Inches(0.05), SW - ML - MR - Inches(0.3), Inches(0.3))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.name = F_BODY
    r.font.size = Pt(font_size)
    r.font.color.rgb = text_color
    return bar


def _code_panel(slide, lines_tokens, left, top, width, height, title=None, title_color=TEAL):
    """Draw a code panel. lines_tokens is a list of lists of (text, color) tuples per line."""
    panel = _add_rect(slide, left, top, width, height, fill_color=CODE_BG, line_color=CODE_BORDER, line_width_pt=1.5)
    if title:
        _add_textbox(slide, title, left + Inches(0.1), top + Inches(0.05), width - Inches(0.2), Inches(0.25),
                     font=F_HEAD, size=9, bold=True, color=title_color)
        code_top = top + Inches(0.32)
    else:
        code_top = top + Inches(0.1)

    line_h = Inches(0.22)
    for i, tokens in enumerate(lines_tokens):
        x = left + Inches(0.15)
        for token_text, token_color in tokens:
            if not token_text:
                continue
            tb = slide.shapes.add_textbox(x, code_top + i * line_h, Inches(0.01) + Pt(len(token_text) * 6.5), line_h)
            tf = tb.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = token_text
            r.font.name = F_CODE
            r.font.size = Pt(9)
            r.font.color.rgb = token_color
            # rough width estimation — textboxes auto-size, so just stack them
            x += Pt(len(token_text) * 5.5)
    return panel


def _simple_code_box(slide, code_text, left, top, width, height, title=None):
    """Simpler code panel using a single multiline textbox."""
    panel = _add_rect(slide, left, top, width, height, fill_color=CODE_BG, line_color=CODE_BORDER, line_width_pt=1.5)
    if title:
        tb_t = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.05), width - Inches(0.2), Inches(0.25))
        tf_t = tb_t.text_frame
        p_t = tf_t.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.name = F_HEAD
        r_t.font.size = Pt(9)
        r_t.font.bold = True
        r_t.font.color.rgb = TEAL
        code_top = top + Inches(0.32)
    else:
        code_top = top + Inches(0.1)

    tb = slide.shapes.add_textbox(left + Inches(0.15), code_top, width - Inches(0.3), height - Inches(0.15))
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
        r.font.size = Pt(9)
        r.font.color.rgb = WHITE
    return panel


def _section_label(slide, text, left, top, width, color=TEAL):
    tb = slide.shapes.add_textbox(left, top, width, Inches(0.3))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text.upper()
    r.font.name = F_HEAD
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = color


def _card(slide, left, top, width, height, header, body, border_color=TEAL, fill=DIM_FILL, header_size=10, body_size=9):
    _add_rect(slide, left, top, width, height, fill_color=fill, line_color=border_color, line_width_pt=1.5)
    _add_textbox(slide, header, left + Inches(0.1), top + Inches(0.08), width - Inches(0.2), Inches(0.25),
                 font=F_HEAD, size=header_size, bold=True, color=border_color)
    _add_textbox(slide, body, left + Inches(0.1), top + Inches(0.35), width - Inches(0.2), height - Inches(0.45),
                 font=F_BODY, size=body_size, color=WHITE, wrap=True)


# ============================================================
# Per-slide rebuild functions
# ============================================================

def rebuild_slide11(slide):
    """pptx slide 12 — How reports become events: identity tuple"""
    _update_placeholder(slide, 0, "How reports become events: identity tuple")
    _update_eyebrow(slide, "One report, one event — or many", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    panel_w = Inches(5.5)
    panel_h = Inches(1.8)
    panel_top = CONTENT_TOP

    # Left panel — same tuple
    lp = _add_rect(slide, ML, panel_top, panel_w, panel_h, fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "Same tuple → event refreshes", ML + Inches(0.1), panel_top + Inches(0.08), panel_w - Inches(0.2), Inches(0.3),
                 font=F_HEAD, size=11, bold=True, color=TEAL)
    items = ["✓  event.name  — matches", "✓  event.type  — matches", "✓  source.entity.id  — matches"]
    for i, item in enumerate(items):
        _add_textbox(slide, item, ML + Inches(0.15), panel_top + Inches(0.45) + i * Inches(0.32), panel_w - Inches(0.3), Inches(0.3),
                     font=F_BODY, size=10, color=TEAL)
    _add_textbox(slide, "→ Refreshes existing event", ML + Inches(0.15), panel_top + Inches(1.45), panel_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=10, bold=True, color=WHITE)

    # Right panel — changed field
    rp_left = ML + panel_w + Inches(0.4)
    rp = _add_rect(slide, rp_left, panel_top, panel_w, panel_h, fill_color=DIM_FILL, line_color=ORANGE_KW, line_width_pt=2)
    _add_textbox(slide, "Change any field → new event", rp_left + Inches(0.1), panel_top + Inches(0.08), panel_w - Inches(0.2), Inches(0.3),
                 font=F_HEAD, size=11, bold=True, color=ORANGE_KW)
    items2 = ["✗  event.name  — CHANGED", "✓  event.type  — matches", "✓  source.entity.id  — matches"]
    colors2 = [ERR, MUTED, MUTED]
    for i, (item, col) in enumerate(zip(items2, colors2)):
        _add_textbox(slide, item, rp_left + Inches(0.15), panel_top + Inches(0.45) + i * Inches(0.32), panel_w - Inches(0.3), Inches(0.3),
                     font=F_BODY, size=10, color=col)
    _add_textbox(slide, "→ Creates NEW separate event", rp_left + Inches(0.15), panel_top + Inches(1.45), panel_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=10, bold=True, color=ORANGE_KW)

    # Identity tuple table
    tbl_top = panel_top + panel_h + Inches(0.3)
    tbl_h = Inches(2.0)
    tbl_w = SW - ML - MR
    headers = ["Field", "Purpose", "Key rule"]
    col_ws = [Inches(2.8), Inches(3.0), Inches(5.6)]
    rows_data = [
        ["event.name", "Identifies the event", "Volatile values = new event every time"],
        ["event.type", "Semantic signal", "Carries platform meaning — choose carefully"],
        ["source.entity.id", "Entity binding", "Must bind to a real entity"],
        ["dt.event.correlation_id", "Grail identity", "Grail exposes this as the correlation key"],
    ]
    tbl = slide.shapes.add_table(len(rows_data) + 1, 3, ML, tbl_top, tbl_w, tbl_h).table
    tbl.columns[0].width = col_ws[0]
    tbl.columns[1].width = col_ws[1]
    tbl.columns[2].width = col_ws[2]
    for j, h in enumerate(headers):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=10, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, row in enumerate(rows_data):
        fill = PAIR_FILL if i == 1 else DIM_FILL
        bdr = PAIR_BORDER if i == 1 else CODE_BORDER
        for j, val in enumerate(row):
            c = tbl.cell(i + 1, j)
            col_c = TEAL if j == 0 else WHITE
            _cell_text(c, val, font=F_BODY if j > 0 else F_CODE, size=9, color=col_c)
            _set_cell_fill(c, fill)
            if i == 1:
                _set_cell_border(c, bdr, 1.0)

    _footer_bar(slide, "Same tuple → report refreshes the event. Change any field → new, separate event. Grail exposes the identity as dt.event.correlation_id.")


def rebuild_slide12(slide):
    """pptx slide 13 — Match category to condition"""
    _update_placeholder(slide, 0, "Match the category to the condition")
    _update_eyebrow(slide, "event.type is a semantic signal, not a formality", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Warning bar
    warn = _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.35), fill_color=WARN, line_color=None)
    _add_textbox(slide, "Do not default to CUSTOM_ALERT for error, slowdown or availability signals.", ML + Inches(0.15), CONTENT_TOP + Inches(0.06),
                 SW - ML - MR - Inches(0.3), Inches(0.28), font=F_BODY, size=10, bold=True, color=NAVY)

    # Tier 1 label
    tier1_top = CONTENT_TOP + Inches(0.45)
    _section_label(slide, "Opens a Problem", ML, tier1_top, Inches(3), color=TEAL)

    # 4 cards: Tier 1
    cards1 = [
        ("AVAILABILITY_EVENT", "A service or entity is unavailable or unreachable."),
        ("ERROR_EVENT", "An error rate or error condition is detected."),
        ("PERFORMANCE_EVENT", "Performance degradation — latency, throughput."),
        ("RESOURCE_CONTENTION_EVENT", "CPU, memory, disk or network saturation."),
    ]
    c_top = tier1_top + Inches(0.32)
    c_w = Inches(2.85)
    c_h = Inches(1.3)
    gap = Inches(0.08)
    for i, (name, desc) in enumerate(cards1):
        c_left = ML + i * (c_w + gap)
        _card(slide, c_left, c_top, c_w, c_h, name, desc, border_color=TEAL, fill=DIM_FILL, header_size=8)

    # CUSTOM_ALERT special card
    ca_top = c_top + c_h + Inches(0.08)
    _card(slide, ML, ca_top, Inches(5.5), Inches(0.8), "CUSTOM_ALERT", "Last resort — use only when no built-in category fits.", border_color=AMBER, fill=RGBColor(0x1A, 0x10, 0x00), header_size=9)

    # Tier 2 label
    tier2_top = ca_top + Inches(0.88)
    _section_label(slide, "Does NOT open a Problem", ML, tier2_top, Inches(4), color=MUTED)

    # 3 dimmed cards
    cards2 = [
        ("CUSTOM_INFO", "Informational — no problem triggered."),
        ("CUSTOM_ANNOTATION", "Deployment or change marker."),
        ("MARKED_FOR_TERMINATION", "Entity marked for removal."),
    ]
    c2_top = tier2_top + Inches(0.3)
    c2_w = Inches(3.8)
    for i, (name, desc) in enumerate(cards2):
        c2_left = ML + i * (c2_w + Inches(0.1))
        _card(slide, c2_left, c2_top, c2_w, Inches(0.75), name, desc, border_color=CODE_BORDER, fill=CODE_BG, header_size=9)

    # Bottom note
    _footer_bar(slide, "Match the category to the actual condition. A specific category gives Dynatrace Intelligence more to work with. CUSTOM_ALERT is a last resort.")


def rebuild_slide13(slide):
    """pptx slide 14 — Fragmentation patterns — keep images, add footer"""
    _remove_content_shapes(slide, keep_images=True)
    _footer_bar(slide,
        "Rule of thumb: If one anomaly manifests across many dimensions — cities, partitions, pods, regions — that dimension is data on one event, not a reason to emit one event per value. Keep high-cardinality context in event.description.",
        fill=TEAL_DIM)


def rebuild_slide14(slide):
    """pptx slide 15 — Keep images, add bullet panel"""
    _remove_content_shapes(slide, keep_images=True)
    bullets = [
        "Process, disk and network-interface events carry dt.smartscape.host.",
        "Kubernetes pod events carry k8s.namespace.name, k8s.workload.name and cluster fields.",
        "With the shared field present, one rule merges the whole vertical stack — no topology walk required.",
    ]
    panel_left = SW - MR - Inches(4.5)
    panel_top = CONTENT_TOP
    _add_rect(slide, panel_left, panel_top, Inches(4.4), Inches(2.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _section_label(slide, "Shared fields enable merge", panel_left + Inches(0.1), panel_top + Inches(0.08), Inches(4.0), color=TEAL)
    for i, b in enumerate(bullets):
        _add_textbox(slide, f"• {b}", panel_left + Inches(0.1), panel_top + Inches(0.42) + i * Inches(0.65), Inches(4.2), Inches(0.6),
                     font=F_BODY, size=9, color=WHITE, wrap=True)


def rebuild_slide15(slide):
    """pptx slide 16 — Correlation rules definition"""
    _update_placeholder(slide, 0, "Correlation rules — definition and structure")
    _update_eyebrow(slide, "Correlation rules explained", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Definition bar
    def_top = CONTENT_TOP
    def_rect = _add_rect(slide, ML, def_top, SW - ML - MR, Inches(0.45), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "Definition: DQL-based matching definitions that tell the platform which arriving events belong in the same problem — based on a shared field, within a time window.",
                 ML + Inches(0.15), def_top + Inches(0.05), SW - ML - MR - Inches(0.3), Inches(0.38),
                 font=F_BODY, size=10, color=WHITE, wrap=True)

    # Constraint bar
    cst_top = def_top + Inches(0.55)
    _add_rect(slide, ML, cst_top, SW - ML - MR, Inches(0.35), fill_color=RGBColor(0x1A, 0x10, 0x00), line_color=WARN, line_width_pt=1.5)
    _add_textbox(slide, "Correlation rules produce merging only: no root cause, no Visual Resolution Path.",
                 ML + Inches(0.15), cst_top + Inches(0.06), SW - ML - MR - Inches(0.3), Inches(0.28),
                 font=F_BODY, size=10, color=WARN, wrap=True)

    # 4 part cards
    cards = [
        (TEAL,        "timeWindow",            "How far back to compare", '"15m"'),
        (PURPLE_LIGHT,"groupingFields",         "Values that must match",   '"dt.smartscape.host"'),
        (CYAN,        "matchingCondition",      "Which events qualify",     "DQL expression"),
        (MUTED,       "correlationNamespace",   "Correlation group name",   '"default-dt.smartscape.host"'),
    ]
    c_top = cst_top + Inches(0.45)
    c_w = Inches(2.85)
    c_h = Inches(1.2)
    for i, (col, name, desc, ex) in enumerate(cards):
        c_left = ML + i * (c_w + Inches(0.08))
        _card(slide, c_left, c_top, c_w, c_h, name, f"{desc}\ne.g. {ex}", border_color=col, fill=DIM_FILL, header_size=9)

    # JSON code block
    json_code = '''{
  "displayName": "Correlate host infrastructure entities by host",
  "enabled": true,
  "timeWindow": "15m",
  "groupingFields": ["dt.smartscape.host"],
  "matchingCondition": "matchesValue(dt.smartscape_source.type, {\\"CONTAINER\\", \\"DISK\\", ...})",
  "correlationNamespace": "default-dt.smartscape.host"
}'''
    code_top = c_top + c_h + Inches(0.15)
    _simple_code_box(slide, json_code, ML, code_top, SW - ML - MR, Inches(1.6), title="Example rule JSON")


def rebuild_slide16(slide):
    """pptx slide 17 — Exercise 1"""
    _remove_content_shapes(slide, keep_images=False)

    instr = _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.35), fill_color=TEAL_DIM)
    _add_textbox(slide, "3 scenario cards per team  ·  Everyone does all cards  ·  7 minutes  ·  Debrief together",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.06), SW - ML - MR - Inches(0.3), Inches(0.28),
                 font=F_HEAD, size=10, bold=True, color=WHITE)

    cards = [
        (TEAL,        "Card A · Audit",  "A host's memory, disk and process problems — check the built-in rule first."),
        (AMBER,       "Card B · Trap",   "Per-region error-rate events with region in both the name and entity — write the rule."),
        (PURPLE_LIGHT,"Card C · Build",  "A release across three services — deployment markers, warnings and errors under one namespace."),
    ]
    c_top = CONTENT_TOP + Inches(0.5)
    c_w = (SW - ML - MR - Inches(0.3)) / 3
    for i, (col, hdr, body) in enumerate(cards):
        _card(slide, ML + i * (c_w + Inches(0.15)), c_top, c_w, Inches(3.2), hdr, body, border_color=col, fill=DIM_FILL, header_size=11, body_size=10)

    # Timer
    _add_textbox(slide, "7 min", ML, CONTENT_TOP + Inches(3.9), SW - ML - MR, Inches(1.2),
                 font=F_HEAD, size=60, bold=True, color=TEAL, align=PP_ALIGN.CENTER)


def rebuild_slide17(slide):
    """pptx slide 18 — Card A lesson"""
    _update_placeholder(slide, 0, "Card A lesson — the audit reflex")
    _update_eyebrow(slide, "Check what ships before you build", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    half_w = Inches(5.8)
    gap = Inches(0.3)

    # Left panel
    bullets = [
        "The platform already ships rules for same-entity, vertical stack, Kubernetes, host-cloud grouping.",
        "Best practice: never edit a shipped rule directly.",
        "Disable it and create a new one with narrower scope — keeps a one-click revert path.",
    ]
    _section_label(slide, "The discipline", ML, CONTENT_TOP, half_w, color=TEAL)
    for i, b in enumerate(bullets):
        _add_textbox(slide, f"• {b}", ML, CONTENT_TOP + Inches(0.35) + i * Inches(0.75), half_w, Inches(0.7),
                     font=F_BODY, size=10, color=WHITE, wrap=True)

    # Right panel — code block
    json_code = '''{
  "displayName": "Correlate host infrastructure by host",
  "enabled": true,
  "timeWindow": "15m",
  "groupingFields": ["dt.smartscape.host"],
  "matchingCondition": "matchesValue(dt.smartscape_source.type,
    {\\"CONTAINER\\", \\"DISK\\", \\"NETWORK_INTERFACE\\",
     \\"PROCESS\\", \\"HOST\\", \\"OS_SERVICE\\"})",
  "correlationNamespace": "default-dt.smartscape.host"
}'''
    code_left = ML + half_w + gap
    _simple_code_box(slide, json_code, code_left, CONTENT_TOP, half_w, Inches(3.5), title="Built-in rule — do not edit directly")


def rebuild_slide18(slide):
    """pptx slide 19 — Card B lesson"""
    _update_placeholder(slide, 0, "Card B lesson — the rule that cannot exist")
    _update_eyebrow(slide, "You cannot fix improper event design with a correlation rule", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    half_w = Inches(5.6)
    gap = Inches(0.3)

    # Left — volatile identity
    lp = _add_rect(slide, ML, CONTENT_TOP, half_w, Inches(3.5), fill_color=DIM_FILL, line_color=AMBER, line_width_pt=2)
    _add_textbox(slide, "VOLATILE IDENTITY", ML + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.25),
                 font=F_HEAD, size=10, bold=True, color=AMBER)
    _add_textbox(slide, "event.name contains region", ML + Inches(0.15), CONTENT_TOP + Inches(0.4), half_w - Inches(0.3), Inches(0.3),
                 font=F_CODE, size=11, bold=True, color=ORANGE_KW)
    _add_textbox(slide, "source entity differs by region", ML + Inches(0.15), CONTENT_TOP + Inches(0.78), half_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=10, color=WHITE)
    _add_textbox(slide, "The transferable question: Which field in the identity tuple is carrying a value that changes?",
                 ML + Inches(0.15), CONTENT_TOP + Inches(1.2), half_w - Inches(0.3), Inches(0.8),
                 font=F_BODY, size=10, color=MUTED, wrap=True, italic=True)

    # Right — fix upstream
    rp_left = ML + half_w + gap
    rp = _add_rect(slide, rp_left, CONTENT_TOP, half_w, Inches(3.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "FIX UPSTREAM", rp_left + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.25),
                 font=F_HEAD, size=10, bold=True, color=TEAL)
    fixes = [
        "Fix at Stage 4: remove region from event.name.",
        "Bind to the real entity.",
        "Carry regions in event.description.",
        "Problem title inherits the first event's name — volatile values always mislead.",
        "Even with a rule, you'd get a problem titled with the region name.",
    ]
    for i, fix in enumerate(fixes):
        _add_textbox(slide, f"• {fix}", rp_left + Inches(0.15), CONTENT_TOP + Inches(0.42) + i * Inches(0.55), half_w - Inches(0.3), Inches(0.5),
                     font=F_BODY, size=9, color=WHITE, wrap=True)


def rebuild_slide19(slide):
    """pptx slide 20 — Card C lesson"""
    _update_placeholder(slide, 0, "Card C lesson — what the namespace solves")
    _update_eyebrow(slide, "The namespace pulls context into the problem", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Summary bar
    _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.4), fill_color=TEAL_DIM)
    _add_textbox(slide, "One shared namespace (release-impact-correlation) ties deployment markers + warnings + errors into one problem.",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.06), SW - ML - MR - Inches(0.3), Inches(0.32),
                 font=F_BODY, size=10, color=WHITE, wrap=True)

    # 3 code panels
    rules = [
        ("Rule 1 — deployment marker", 'timeWindow: "1h"\nmatchingCondition:\n  event.type == CUSTOM_INFO\ngroupingFields:\n  [release_version]'),
        ("Rule 2 — warning",           'timeWindow: "30m"\nmatchingCondition:\n  event.type == WARNING\ngroupingFields:\n  [release_version]'),
        ("Rule 3 — error/perf",        'timeWindow: "15m"\nmatchingCondition:\n  event.type in\n  [ERROR_EVENT,\n   PERFORMANCE_EVENT]\ngroupingFields:\n  [release_version]'),
    ]
    c_top = CONTENT_TOP + Inches(0.52)
    c_w = (SW - ML - MR - Inches(0.4)) / 3
    for i, (title, code) in enumerate(rules):
        c_left = ML + i * (c_w + Inches(0.2))
        _simple_code_box(slide, code, c_left, c_top, c_w, Inches(2.3), title=title)

    # Footer
    _footer_bar(slide, "Windows: Deployment 1h, warning 30m, error 15m. The error opens the problem; other windows look back from that start. Without the namespace, different group keys — the explanatory context is lost.")


def rebuild_slide20(slide):
    """pptx slide 21 — Before/After + question + DQL"""
    _remove_content_shapes(slide, keep_images=True)

    # Question box
    q_top = Inches(4.6)
    _add_rect(slide, ML, q_top, SW - ML - MR, Inches(0.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "Which field on the events was the grouping key, and how did you know it existed before writing the rule?",
                 ML + Inches(0.15), q_top + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.38),
                 font=F_BODY, size=10, bold=True, color=WHITE, wrap=True)

    # DQL
    dql = "fetch dt.davis.events\n| filter isNotNull(release_version)\n| summarize count(), by: {release_version, event.name}"
    _simple_code_box(slide, dql, ML, q_top + Inches(0.6), Inches(6.5), Inches(1.1), title="DQL — check the field exists")


def rebuild_slide22(slide):
    """pptx slide 23 (after deletion of 22 becomes 22) — When RCA is legitimately empty"""
    _update_placeholder(slide, 0, "When RCA is legitimately empty")
    _update_eyebrow(slide, "'No root cause' in a Problem could be expected behaviour", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Opening bar
    _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.45), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "Causal RCA exists only where one event explains another. No chain to build → the panel correctly shows 'No root cause.'",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.06), SW - ML - MR - Inches(0.3), Inches(0.35),
                 font=F_BODY, size=10, color=WHITE, wrap=True)

    # 5 tiles
    tiles = [
        ("Single-event problem",       "One event, one entity. It is both fault and cause. Nothing upstream to trace."),
        ("Same-entity merge",          "Multiple events, one entity. No second entity for a chain."),
        ("Pure infrastructure stack",  "Host, disk, process — no service layer, no topology walk needed."),
        ("Custom alert — review needed","Two checks: right event.category? Entity reference present?"),
        ("Non-standard entity types",  "Network, custom, extension entities — outside deterministic traversal."),
    ]
    tile_top = CONTENT_TOP + Inches(0.6)
    tile_w = (SW - ML - MR - Inches(0.6)) / 5
    for i, (name, desc) in enumerate(tiles):
        t_left = ML + i * (tile_w + Inches(0.15))
        _card(slide, t_left, tile_top, tile_w, Inches(2.8), name, desc, border_color=TEAL, fill=DIM_FILL, header_size=9)

    _footer_bar(slide, "Do not tune detection to force a root cause. Confirm the empty is expected before treating it as a defect.", fill=RGBColor(0x1A, 0x10, 0x00))


def rebuild_slide23(slide):
    """pptx slide 24 — Three dimensions of RCA quality"""
    _update_placeholder(slide, 0, "Three dimensions of RCA quality")
    _update_eyebrow(slide, "Populated RCA is not the same as correct", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    bars = [
        (TEAL,        1.0, "1 · Attached",    "Is the root cause field populated? A query answers this.",           "QUERY"),
        (PURPLE_LIGHT,0.75,"2 · Correct",      "The named entity is the right one? Needs service-owner knowledge.", "DOMAIN"),
        (MUTED,       0.5, "3 · Actionable",   "Given a correct root cause, could someone act? Depends on #2.",    "PEOPLE"),
    ]
    max_w = SW - ML - MR
    bar_top = CONTENT_TOP + Inches(0.2)
    bar_h = Inches(1.1)
    bar_gap = Inches(0.3)
    for i, (color, pct, title, desc, badge) in enumerate(bars):
        b_top = bar_top + i * (bar_h + bar_gap)
        b_w = int(max_w * pct)
        _add_rect(slide, ML, b_top, b_w, bar_h, fill_color=DIM_FILL, line_color=color, line_width_pt=2)
        # colour strip on left
        _add_rect(slide, ML, b_top, Inches(0.25), bar_h, fill_color=color)
        _add_textbox(slide, title, ML + Inches(0.35), b_top + Inches(0.08), Inches(3.0), Inches(0.35),
                     font=F_HEAD, size=14, bold=True, color=color)
        _add_textbox(slide, desc, ML + Inches(0.35), b_top + Inches(0.5), b_w - Inches(1.5), Inches(0.5),
                     font=F_BODY, size=10, color=WHITE, wrap=True)
        # badge
        _add_rect(slide, ML + b_w - Inches(1.1), b_top + Inches(0.3), Inches(1.0), Inches(0.35), fill_color=color)
        _add_textbox(slide, badge, ML + b_w - Inches(1.05), b_top + Inches(0.35), Inches(0.9), Inches(0.28),
                     font=F_HEAD, size=9, bold=True, color=NAVY)

    _footer_bar(slide, "A populated field only clears the first of three bars.  |  CoE framework — not a native Dynatrace concept.", fill=RGBColor(0x1A, 0x10, 0x00))


def rebuild_slide24_dql(slide):
    """pptx slide 25 — Master DQL. Screenshot stays on right, DQL on left."""
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
        and arraySize(dt.davis.impact_level) == 1, "expected_infra_stack",
    else: if(event.category == "CUSTOM_ALERT", "review_custom_alert",
    else: if(affected_count > 0 and not(has_classic),
        "expected_non_classic_entity",
    else: "should_have_had")))))
| summarize problems = count(),
    avg_events = round(avg(event_count), decimals: 1),
    by: {bucket}
| sort problems desc"""
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.2), Inches(4.8), title="Master classification query")


def rebuild_slide25_dql(slide):
    """pptx slide 26 — Case 1 DQL on left; screenshot on right."""
    _remove_content_shapes(slide, keep_images=True)
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
    and isNull(root_cause.smartscape_entity.id)
| fieldsAdd event_count = arraySize(dt.davis.event_ids)
| filter event_count <= 1"""
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.0), Inches(2.0), title="Case 1 — expected_single_event")


def rebuild_slide26_dql(slide):
    """pptx slide 27 — Case 2 DQL on left."""
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
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.0), Inches(2.2), title="Case 2 — expected_same_entity")


def rebuild_slide27_dql(slide):
    """pptx slide 28 — Case 3 DQL on left."""
    _remove_content_shapes(slide, keep_images=True)
    dql = """fetch dt.davis.problems
| filter not(dt.davis.is_duplicate)
| filter isNull(root_cause_entity_id)
| fieldsAdd affected_count = arraySize(smartscape.affected_entities)
| filter in("Infrastructure", dt.davis.impact_level)
    and arraySize(dt.davis.impact_level) == 1"""
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.0), Inches(2.2), title="Case 3 — expected_infra_stack")


def rebuild_slide28_dql(slide):
    """pptx slide 29 — Case 4 on left; screenshot stays on right."""
    _remove_content_shapes(slide, keep_images=True)
    dql = """| filter bucket == "should_have_had"
| filter event_count > 1
| filter arraySize(dt.davis.impact_level) > 1
| filter in("Services", dt.davis.impact_level)
    and in("Infrastructure", dt.davis.impact_level)"""
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.0), Inches(2.0), title="Case 4 — should_have_had key filters")


def rebuild_slide29(slide):
    """pptx slide 30 — Exercise 2"""
    _remove_content_shapes(slide, keep_images=False)

    # Instruction bar
    instr_rect = _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.5), fill_color=RGBColor(0x1A, 0x10, 0x00), line_color=AMBER, line_width_pt=2)
    _add_textbox(slide, "The tests are ordered deliberately. A problem matches one bucket only — the first one it hits. Do not skip rows.",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.1), SW - ML - MR - Inches(0.3), Inches(0.35),
                 font=F_BODY, size=10, bold=True, color=AMBER, wrap=True)

    # Two task panels
    task_top = CONTENT_TOP + Inches(0.65)
    task_w = (SW - ML - MR - Inches(0.3)) / 2
    _card(slide, ML, task_top, task_w, Inches(2.2), "Task 1 · Classify",
          "Assign every problem card scenario to the correct bucket. Work top to bottom — stop at the first match. 5 problem cards.",
          border_color=TEAL, fill=DIM_FILL, header_size=12, body_size=10)
    _card(slide, ML + task_w + Inches(0.3), task_top, task_w, Inches(2.2), "Task 2 · Investigate",
          "For the problem you classify as should_have_had (the genuine gap), write down two diagnostic steps.",
          border_color=PURPLE_LIGHT, fill=DIM_FILL, header_size=12, body_size=10)

    _add_textbox(slide, "6 min", ML, task_top + Inches(2.4), SW - ML - MR, Inches(1.2),
                 font=F_HEAD, size=60, bold=True, color=TEAL, align=PP_ALIGN.CENTER)


def rebuild_slide30(slide):
    """pptx slide 31 — Two tables"""
    _update_placeholder(slide, 0, "Classification reference + five records")
    _update_eyebrow(slide, "The five buckets and five problem records", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Top table — classification reference
    ref_rows = [
        ("1", "expected_single_event",       "event_count <= 1"),
        ("2", "expected_same_entity",         "affected_count == 1"),
        ("3", "expected_infra_stack",          'in("Infrastructure", impact_level) AND arraySize == 1'),
        ("4", "review_custom_alert",           'event.category == "CUSTOM_ALERT"'),
        ("5", "expected_non_classic_entity",  "affected_count > 0 AND not(has_classic)"),
        ("6", "should_have_had",               "Falls through all above"),
    ]
    tbl_w = SW - ML - MR
    tbl = slide.shapes.add_table(len(ref_rows) + 1, 3, ML, CONTENT_TOP, tbl_w, Inches(2.0)).table
    tbl.columns[0].width = Inches(0.5)
    tbl.columns[1].width = Inches(4.0)
    tbl.columns[2].width = tbl_w - Inches(4.5)
    for j, h in enumerate(["#", "Bucket", "Rule"]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=9, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, (num, bucket, rule) in enumerate(ref_rows):
        fill = RGBColor(0x1A, 0x10, 0x00) if bucket == "should_have_had" else DIM_FILL
        bdr = AMBER if bucket == "should_have_had" else CODE_BORDER
        for j, val in enumerate([num, bucket, rule]):
            c = tbl.cell(i + 1, j)
            col = AMBER if bucket == "should_have_had" else WHITE
            _cell_text(c, val, font=F_CODE if j == 1 else F_BODY, size=8, color=col)
            _set_cell_fill(c, fill)

    # Bottom table — problem records
    p_top = CONTENT_TOP + Inches(2.15)
    prows = [
        ("P-1", "1", "1", "[Services]",             "CUSTOM_ALERT",         "review_custom_alert"),
        ("P-2", "3", "2", "[Services, Infrastructure]","ERROR_EVENT",        "should_have_had"),
        ("P-3", "4", "1", "[Services]",              "ERROR_EVENT",          "expected_same_entity"),
        ("P-4", "2", "2", "[Infrastructure]",        "RESOURCE_CONTENTION",  "expected_infra_stack"),
        ("P-5", "1", "1", "[Infrastructure]",        "AVAILABILITY",         "expected_single_event"),
    ]
    ptbl = slide.shapes.add_table(len(prows) + 1, 6, ML, p_top, tbl_w, Inches(2.2)).table
    col_ws2 = [Inches(0.7), Inches(0.7), Inches(0.8), Inches(3.0), Inches(2.5), Inches(2.0)]
    for k, w in enumerate(col_ws2):
        ptbl.columns[k].width = w
    for j, h in enumerate(["#", "Events", "Entities", "Impact level", "Category", "Bucket"]):
        _cell_text(ptbl.cell(0, j), h, font=F_HEAD, size=9, bold=True, color=TEAL)
        _set_cell_fill(ptbl.cell(0, j), HDR_FILL)
    for i, row_vals in enumerate(prows):
        fill = RGBColor(0x1A, 0x10, 0x00) if row_vals[0] == "P-2" else DIM_FILL
        bdr = AMBER if row_vals[0] == "P-2" else CODE_BORDER
        for j, val in enumerate(row_vals):
            c = ptbl.cell(i + 1, j)
            col = AMBER if row_vals[0] == "P-2" else WHITE
            _cell_text(c, val, font=F_BODY, size=8, color=col)
            _set_cell_fill(c, fill)


def rebuild_slide31(slide):
    """pptx slide 32 — Answer reveal table"""
    _update_placeholder(slide, 0, "Answer reveal — which problem goes where")
    _update_eyebrow(slide, "The answers explained", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    answers = [
        ("P-1", "review_custom_alert",  "Custom alert — apply the two-check test."),
        ("P-2", "should_have_had",      "Multiple entities, service impact, no root cause — this is the genuine gap."),
        ("P-3", "expected_same_entity", "Four events, one service — merging is working."),
        ("P-4", "expected_infra_stack", "Infrastructure-only, arraySize == 1."),
        ("P-5", "expected_single_event","One event, one entity."),
    ]
    tbl_w = SW - ML - MR
    tbl = slide.shapes.add_table(len(answers) + 1, 3, ML, CONTENT_TOP, tbl_w, Inches(3.5)).table
    tbl.columns[0].width = Inches(0.8)
    tbl.columns[1].width = Inches(3.5)
    tbl.columns[2].width = tbl_w - Inches(4.3)
    for j, h in enumerate(["Problem", "Bucket", "Explanation"]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=10, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    for i, (prob, bucket, expl) in enumerate(answers):
        fill = PAIR_FILL if prob == "P-2" else DIM_FILL
        bdr = PAIR_BORDER if prob == "P-2" else CODE_BORDER
        for j, val in enumerate([prob, bucket, expl]):
            c = tbl.cell(i + 1, j)
            col = TEAL if prob == "P-2" else WHITE
            _cell_text(c, val, font=F_CODE if j == 1 else F_BODY, size=9, color=col)
            _set_cell_fill(c, fill)
            if prob == "P-2":
                _set_cell_border(c, PAIR_BORDER, 1.0)


def rebuild_slide32(slide):
    """pptx slide 33 — P-4 guard"""
    _update_placeholder(slide, 0, "P-4 — the guard")
    _update_eyebrow(slide, "What the array-size guard prevents", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    half_w = Inches(5.4)
    gap = Inches(0.3)

    # Left — PASSES
    _add_rect(slide, ML, CONTENT_TOP, half_w, Inches(2.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "PASSES", ML + Inches(0.1), CONTENT_TOP + Inches(0.08), Inches(2), Inches(0.25),
                 font=F_HEAD, size=10, bold=True, color=TEAL)
    lines_l = ['impact_level = ["Infrastructure"]', "arraySize = 1", "→ expected_infra_stack"]
    colors_l = [WHITE, WHITE, GOOD]
    for i, (l, col) in enumerate(zip(lines_l, colors_l)):
        _add_textbox(slide, l, ML + Inches(0.15), CONTENT_TOP + Inches(0.45) + i * Inches(0.5), half_w - Inches(0.3), Inches(0.4),
                     font=F_CODE, size=11, color=col)

    # Right — FAILS
    rp_left = ML + half_w + gap
    _add_rect(slide, rp_left, CONTENT_TOP, half_w, Inches(2.5), fill_color=DIM_FILL, line_color=AMBER, line_width_pt=2)
    _add_textbox(slide, "CORRECTLY FAILS", rp_left + Inches(0.1), CONTENT_TOP + Inches(0.08), Inches(3), Inches(0.25),
                 font=F_HEAD, size=10, bold=True, color=AMBER)
    lines_r = ['impact_level = ["Services", "Infrastructure"]', "arraySize = 2", "→ does NOT match expected_infra_stack"]
    colors_r = [WHITE, WHITE, ERR]
    for i, (l, col) in enumerate(zip(lines_r, colors_r)):
        _add_textbox(slide, l, rp_left + Inches(0.15), CONTENT_TOP + Inches(0.45) + i * Inches(0.5), half_w - Inches(0.3), Inches(0.4),
                     font=F_CODE, size=10, color=col)

    # Centre question + answer
    q_top = CONTENT_TOP + Inches(2.7)
    _add_rect(slide, ML, q_top, SW - ML - MR, Inches(0.45), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "What does arraySize(impact_level) == 1 actually prevent?",
                 ML + Inches(0.15), q_top + Inches(0.07), SW - ML - MR - Inches(0.3), Inches(0.32),
                 font=F_BODY, size=11, bold=True, color=TEAL)
    _add_textbox(slide,
        "Without it, [Services, Infrastructure] would also match the in('Infrastructure', ...) condition and get labelled expected_infra_stack — masking a genuine gap.",
        ML, q_top + Inches(0.55), SW - ML - MR, Inches(0.7),
        font=F_BODY, size=10, color=WHITE, wrap=True)


def rebuild_slide33(slide):
    """pptx slide 34 — P-1 two checks"""
    _update_placeholder(slide, 0, "P-1 — the two checks")
    _update_eyebrow(slide, "Custom alert: check the category first, then the reference", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    half_w = Inches(5.4)
    gap = Inches(0.3)
    panel_h = Inches(3.0)

    # Step 1
    _add_rect(slide, ML, CONTENT_TOP, half_w, panel_h, fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "Step 1 — Was CUSTOM_ALERT the right category?",
                 ML + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.35),
                 font=F_HEAD, size=11, bold=True, color=TEAL)
    _add_textbox(slide, "Yes: A computed business ratio has no built-in category — CUSTOM_ALERT is correct.",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.55), half_w - Inches(0.3), Inches(0.6),
                 font=F_BODY, size=10, color=GOOD, wrap=True)
    _add_textbox(slide, "No: A service's traffic dropping to zero is AVAILABILITY — using CUSTOM_ALERT is a governance miss.",
                 ML + Inches(0.15), CONTENT_TOP + Inches(1.25), half_w - Inches(0.3), Inches(0.6),
                 font=F_BODY, size=10, color=ERR, wrap=True)

    # Step 2
    rp_left = ML + half_w + gap
    _add_rect(slide, rp_left, CONTENT_TOP, half_w, panel_h, fill_color=DIM_FILL, line_color=PURPLE_LIGHT, line_width_pt=2)
    _add_textbox(slide, "Step 2 — Does the template carry an entity reference?",
                 rp_left + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.35),
                 font=F_HEAD, size=11, bold=True, color=PURPLE_LIGHT)
    _add_textbox(slide, "Empty is correct when nothing can be bound.",
                 rp_left + Inches(0.15), CONTENT_TOP + Inches(0.55), half_w - Inches(0.3), Inches(0.4),
                 font=F_BODY, size=10, color=GOOD, wrap=True)
    _add_textbox(slide, "But if a Smartscape entity can be mapped, the template must not miss it.",
                 rp_left + Inches(0.15), CONTENT_TOP + Inches(1.05), half_w - Inches(0.3), Inches(0.5),
                 font=F_BODY, size=10, color=AMBER, wrap=True)

    # Result box
    result_top = CONTENT_TOP + panel_h + Inches(0.25)
    _add_rect(slide, ML, result_top, SW - ML - MR, Inches(0.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "P-1 outcome: review_custom_alert — apply both checks before marking it resolved.",
                 ML + Inches(0.15), result_top + Inches(0.1), SW - ML - MR - Inches(0.3), Inches(0.35),
                 font=F_BODY, size=11, bold=True, color=WHITE)


def rebuild_slide34(slide):
    """pptx slide 35 — P-2 investigation: keep image, add DQL"""
    _remove_content_shapes(slide, keep_images=True)
    dql = """| filter bucket == "should_have_had"
| filter event_count > 1
| filter arraySize(dt.davis.impact_level) > 1
| filter in("Services", dt.davis.impact_level) and in("Infrastructure", dt.davis.impact_level)"""
    _simple_code_box(slide, dql, ML, CONTENT_TOP, Inches(6.5), Inches(1.8), title="Key filter — P-2 pattern")

    # Finding box
    find_top = CONTENT_TOP + Inches(1.95)
    _add_rect(slide, ML, find_top, Inches(6.5), Inches(0.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "Finding: a relationship exists — this is not a missing-topology gap. The link is there.",
                 ML + Inches(0.15), find_top + Inches(0.1), Inches(6.3), Inches(0.35),
                 font=F_BODY, size=10, bold=True, color=WHITE)

    # Note
    note_top = find_top + Inches(0.6)
    _add_textbox(slide,
        "The hypothesis about which edge types are treated as causal is not confirmed with the product team. Do not assert it as fact.",
        ML, note_top, Inches(6.5), Inches(0.5),
        font=F_BODY, size=9, color=AMBER, italic=True, wrap=True)


def insert_tcca_slide(prs, after_idx):
    """Insert new TCCA slide after the slide at after_idx (0-based)."""
    layout = _find_layout(prs, "Title+eyebrow_left")
    new_slide = prs.slides.add_slide(layout)

    # Move to correct position (it's added at the end)
    xml_slides = prs.slides._sldIdLst
    # The new slide is at the last position; move it to after_idx + 1
    last_el = xml_slides[-1]
    xml_slides.remove(last_el)
    xml_slides.insert(after_idx + 1, last_el)

    _update_placeholder(new_slide, 0, "P-2 — Answer is TCCA", font=F_HEAD, bold=True)
    _update_placeholder(new_slide, 1, "How TopologyCausalChainAnalyzer builds a causal chain", font=F_BODY, color=MUTED)

    # Definition bar
    _add_rect(new_slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.65), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(new_slide,
        "TCCA is what turns five separate alerts into one problem with a root cause and a Visual Resolution Path. It only draws a causal link when two things are both connected by topology AND there is a causal signal.",
        ML + Inches(0.15), CONTENT_TOP + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.55),
        font=F_BODY, size=10, color=WHITE, wrap=True)

    # Three condition panels
    conditions = [
        (TEAL,        "Topology link",   "A directed relationship exists between the affected entities in Smartscape."),
        (PURPLE_LIGHT,"Causal signal",   "One entity's anomaly temporally precedes the other's in a way consistent with causation."),
        (CYAN,        "Traversal",       "TCCA walks the topology graph to find the root cause entity."),
    ]
    c_top = CONTENT_TOP + Inches(0.85)
    c_w = (SW - ML - MR - Inches(0.4)) / 3
    for i, (col, hdr, body) in enumerate(conditions):
        c_left = ML + i * (c_w + Inches(0.2))
        _card(new_slide, c_left, c_top, c_w, Inches(1.8), hdr, body, border_color=col, fill=DIM_FILL, header_size=11, body_size=10)

    # Key note
    _add_rect(new_slide, ML, CONTENT_TOP + Inches(2.85), SW - ML - MR, Inches(0.55), fill_color=DIM_FILL, line_color=AMBER, line_width_pt=1.5)
    _add_textbox(new_slide,
        "A relationship exists but TCCA didn't draw the chain → the edge type may not be treated as causal. This is the interesting finding — not a missing-topology gap.",
        ML + Inches(0.15), CONTENT_TOP + Inches(2.95), SW - ML - MR - Inches(0.3), Inches(0.42),
        font=F_BODY, size=10, color=AMBER, wrap=True)

    # Speaker notes
    try:
        notes_slide = new_slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.clear()
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = ("TCCA is what turns five separate alerts into one problem with a root cause and a Visual Resolution Path. "
                  "It only draws a causal link when two things are both connected by topology AND there is a causal signal. "
                  "The key insight here: in P-2, the topology link exists (the service belongs_to the deployment), "
                  "but the causal signal may not satisfy TCCA's requirements for that edge type.")
    except Exception:
        pass

    return new_slide


def rebuild_slide35(slide):
    """pptx slide 36 — Honest conversation / Commit to the trend"""
    _update_placeholder(slide, 0, "Customer conversation and the honest position")
    _update_eyebrow(slide, "Commit to the trend, not the number", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    left_w = Inches(3.8)
    right_w = SW - ML - MR - left_w - Inches(0.3)
    panel_h = Inches(4.0)
    panel_top = CONTENT_TOP

    # Left — headline alone
    _add_rect(slide, ML, panel_top, left_w, panel_h, fill_color=DIM_FILL, line_color=CODE_BORDER, line_width_pt=1.5)
    _add_textbox(slide, "20%?", ML + Inches(0.15), panel_top + Inches(0.2), left_w - Inches(0.3), Inches(1.2),
                 font=F_HEAD, size=48, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    _add_textbox(slide, "RCA attachment", ML + Inches(0.15), panel_top + Inches(1.3), left_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=10, color=MUTED, align=PP_ALIGN.CENTER)
    _add_textbox(slide, '"Headline alone"', ML + Inches(0.15), panel_top + Inches(1.68), left_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=9, color=AMBER, italic=True, align=PP_ALIGN.CENTER)
    _add_textbox(slide, '"We are at 20% RCA attachment — the product is broken."',
                 ML + Inches(0.15), panel_top + Inches(2.1), left_w - Inches(0.3), Inches(0.8),
                 font=F_BODY, size=10, color=WHITE, italic=True, wrap=True)

    # Right — breakdown
    rp_left = ML + left_w + Inches(0.3)
    _add_rect(slide, rp_left, panel_top, right_w, panel_h, fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "20% broken down:", rp_left + Inches(0.15), panel_top + Inches(0.15), right_w - Inches(0.3), Inches(0.28),
                 font=F_HEAD, size=11, bold=True, color=TEAL)
    buckets_breakdown = [
        (TEAL,        "40%", "expected_single_event"),
        (CYAN,        "30%", "expected_same_entity"),
        (PURPLE_LIGHT,"15%", "expected_infra_stack"),
        (AMBER,       "10%", "review_custom_alert"),
        (ERR,         "5%",  "should_have_had  ← genuine work"),
    ]
    for i, (col, pct, label) in enumerate(buckets_breakdown):
        bar_top = panel_top + Inches(0.52) + i * Inches(0.64)
        _add_rect(slide, rp_left + Inches(0.15), bar_top, Inches(0.5), Inches(0.35), fill_color=col)
        _add_textbox(slide, f"{pct}  {label}", rp_left + Inches(0.75), bar_top + Inches(0.04), right_w - Inches(1.0), Inches(0.3),
                     font=F_BODY, size=10, color=col if label.startswith("should") else WHITE)

    _add_textbox(slide, "Trending down — the metric that stays true",
                 rp_left + Inches(0.15), panel_top + Inches(3.55), right_w - Inches(0.3), Inches(0.3),
                 font=F_BODY, size=10, bold=True, color=TEAL)

    # Three key points
    kp_top = panel_top + panel_h + Inches(0.2)
    points = [
        "First move: run the bucket query before agreeing. Headline alone cannot distinguish healthy from broken.",
        "A fixed percentage gets disproven the moment the environment changes — new service, topology shift, detector added.",
        "A trend stays true.",
    ]
    for i, pt in enumerate(points):
        _add_textbox(slide, f"• {pt}", ML, kp_top + i * Inches(0.35), SW - ML - MR, Inches(0.32),
                     font=F_BODY, size=9, color=WHITE, wrap=True)


def rebuild_slide36(slide):
    """pptx slide 37 — Agentic comparison"""
    _update_placeholder(slide, 0, "Agentic — two surfaces, shared reasoning, hard rule")
    _update_eyebrow(slide, "SRE Agent vs. Assist", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Rule bar
    _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.4), fill_color=RGBColor(0x1A, 0x10, 0x00), line_color=AMBER, line_width_pt=2)
    _add_textbox(slide,
        "The Smartscape root cause panel and Visual Resolution Path are written exclusively by the deterministic analysis. Agentic output never writes there.",
        ML + Inches(0.15), CONTENT_TOP + Inches(0.05), SW - ML - MR - Inches(0.3), Inches(0.32),
        font=F_BODY, size=10, bold=True, color=AMBER, wrap=True)

    # Comparison table
    tbl_top = CONTENT_TOP + Inches(0.55)
    tbl_w = SW - ML - MR
    tbl = slide.shapes.add_table(5, 3, ML, tbl_top, tbl_w, Inches(1.8)).table
    tbl.columns[0].width = Inches(2.5)
    tbl.columns[1].width = Inches(4.0)
    tbl.columns[2].width = tbl_w - Inches(6.5)
    for j, h in enumerate(["Feature", "Assist", "SRE Agent"]):
        _cell_text(tbl.cell(0, j), h, font=F_HEAD, size=10, bold=True, color=TEAL)
        _set_cell_fill(tbl.cell(0, j), HDR_FILL)
    rows = [
        ("Invocation",  "On-demand, conversational",   "Autonomous, workflow-triggered"),
        ("License",     "Included",                    "Paid"),
        ("Output",      "Explanation in chat",          "Comment on problem"),
        ("Best for",    '"Explain this to me"',          "Automatic first-response"),
    ]
    for i, row_vals in enumerate(rows):
        for j, val in enumerate(row_vals):
            c = tbl.cell(i + 1, j)
            _cell_text(c, val, font=F_BODY, size=9, color=WHITE)
            _set_cell_fill(c, DIM_FILL)

    # Schema
    schema_top = tbl_top + Inches(2.0)
    half_w = (SW - ML - MR - Inches(0.3)) / 2
    _add_rect(slide, ML, schema_top, half_w, Inches(1.5), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide, "Deterministic writes:", ML + Inches(0.1), schema_top + Inches(0.08), half_w - Inches(0.2), Inches(0.28),
                 font=F_HEAD, size=9, bold=True, color=TEAL)
    _add_textbox(slide, "Root cause panel + Visual Resolution Path",
                 ML + Inches(0.1), schema_top + Inches(0.42), half_w - Inches(0.2), Inches(0.5),
                 font=F_BODY, size=10, color=WHITE, wrap=True)

    rp_left = ML + half_w + Inches(0.3)
    _add_rect(slide, rp_left, schema_top, half_w, Inches(1.5), fill_color=DIM_FILL, line_color=PURPLE_LIGHT, line_width_pt=1.5)
    _add_textbox(slide, "Agentic writes:", rp_left + Inches(0.1), schema_top + Inches(0.08), half_w - Inches(0.2), Inches(0.28),
                 font=F_HEAD, size=9, bold=True, color=PURPLE_LIGHT)
    _add_textbox(slide, "Comment / annotation",
                 rp_left + Inches(0.1), schema_top + Inches(0.42), half_w - Inches(0.2), Inches(0.4),
                 font=F_BODY, size=10, color=WHITE)
    _add_textbox(slide, "HARD RULE: NO CROSSOVER", rp_left + Inches(0.1), schema_top + Inches(0.9), half_w - Inches(0.2), Inches(0.35),
                 font=F_HEAD, size=10, bold=True, color=AMBER)

    _footer_bar(slide, "What they share: Multi-step reasoning over logs, traces, metrics, events, deployments. Narrative with evidence. Assist is free. The workflow requires a license.")


def rebuild_slide37(slide):
    """pptx slide 38 — SRE Agent deployment"""
    _update_placeholder(slide, 0, "SRE Agent deployment")
    _update_eyebrow(slide, "Never run the template as-is", color=MUTED)
    _remove_content_shapes(slide, keep_images=True)

    # Warning bar
    _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(0.4), fill_color=RGBColor(0x1A, 0x10, 0x00), line_color=AMBER, line_width_pt=2)
    _add_textbox(slide, "The SRE Agent ships as a workflow template. Customize before running in production.",
                 ML + Inches(0.15), CONTENT_TOP + Inches(0.07), SW - ML - MR - Inches(0.3), Inches(0.28),
                 font=F_BODY, size=10, bold=True, color=AMBER, wrap=True)

    # Three concern panels
    concerns = [
        (ERR,   "Trigger scope",       "Unscoped = budget gone. Preview invocation count with 'Query past events' first."),
        (AMBER, "Prompt and tools",    "Tight prompt + few tools completes in time. Broad prompt hits timeout."),
        (TEAL,  "Duplicates excluded", "Mandatory. Without it, paying twice for the same incident."),
    ]
    c_top = CONTENT_TOP + Inches(0.55)
    c_w = (SW - ML - MR - Inches(0.4)) / 3
    for i, (col, hdr, body) in enumerate(concerns):
        c_left = ML + i * (c_w + Inches(0.2))
        _card(slide, c_left, c_top, c_w, Inches(1.8), hdr, body, border_color=col, fill=DIM_FILL, header_size=10, body_size=9)


def rebuild_slide38(slide):
    """pptx slide 39 — Assist on populated root cause"""
    _remove_content_shapes(slide, keep_images=True)
    _add_textbox(slide, "Assist interprets and explains what deterministic RCA already named.",
                 ML, CONTENT_TOP, SW - ML - MR, Inches(0.35),
                 font=F_BODY, size=10, color=TEAL, italic=True)
    _footer_bar(slide,
        "A problem where deterministic RCA already named the root cause entity — Assist adds the narrative layer on top.")


def rebuild_slide39(slide):
    """pptx slide 40 — SRE Agent on empty RCA"""
    _update_placeholder(slide, 0, "SRE Agent on an empty root cause")
    _update_eyebrow(slide, "Agentic narrative where deterministic correctly has nothing", color=MUTED)
    _remove_content_shapes(slide, keep_images=True)

    bullets_top = CONTENT_TOP
    panel_w = Inches(5.0)
    _add_rect(slide, ML, bullets_top, panel_w, Inches(2.0), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.5)
    _section_label(slide, "Agent output includes:", ML + Inches(0.1), bullets_top + Inches(0.08), panel_w - Inches(0.2), color=TEAL)
    items = [
        "Verdict — recurring structural problem, not a transient spike.",
        "Recurrence history — the agent queried and found the same entity failing days earlier.",
        "Recommendations — in priority order, including the governance call.",
    ]
    for i, item in enumerate(items):
        _add_textbox(slide, f"• {item}", ML + Inches(0.1), bullets_top + Inches(0.42) + i * Inches(0.48), panel_w - Inches(0.2), Inches(0.45),
                     font=F_BODY, size=10, color=WHITE, wrap=True)


def rebuild_slide40(slide):
    """pptx slide 41 — Agentic analysis / auditability"""
    _update_placeholder(slide, 0, "Agentic Analysis")
    _update_eyebrow(slide, "The auditability beat", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Big quote
    q_rect = _add_rect(slide, ML, CONTENT_TOP, SW - ML - MR, Inches(1.6), fill_color=NAVY, line_color=TEAL, line_width_pt=3)
    _add_textbox(slide,
        '"No application-level logs were returned — the pod is crash-looping before it can emit logs. No distributed traces — the workload does not expose them."',
        ML + Inches(0.3), CONTENT_TOP + Inches(0.25), SW - ML - MR - Inches(0.6), Inches(1.1),
        font=F_BODY, size=13, italic=True, color=WHITE, wrap=True)

    # Two rule boxes
    half_w = (SW - ML - MR - Inches(0.3)) / 2
    box_top = CONTENT_TOP + Inches(1.8)
    _card(slide, ML, box_top, half_w, Inches(1.5), "The rule",
          "An agent that says 'no logs, because crash-looping' is auditable. An agent that stays silent is not.",
          border_color=TEAL, fill=DIM_FILL, header_size=11, body_size=10)
    _card(slide, ML + half_w + Inches(0.3), box_top, half_w, Inches(1.5), "The action",
          "One prompt instruction to SRE Agent: if evidence is unavailable, state so explicitly and explain why.",
          border_color=PURPLE_LIGHT, fill=DIM_FILL, header_size=11, body_size=10)

    # Closing bar
    _add_rect(slide, ML, box_top + Inches(1.65), SW - ML - MR, Inches(0.55), fill_color=NAVY, line_color=TEAL, line_width_pt=1.5)
    _add_textbox(slide,
        "Deterministic tells you which entity. Agentic tells you the story — and tells you what it couldn't find. Assist is free. The workflow requires a license.",
        ML + Inches(0.15), box_top + Inches(1.72), SW - ML - MR - Inches(0.3), Inches(0.42),
        font=F_BODY, size=10, color=MUTED, wrap=True)


def rebuild_slide41(slide):
    """pptx slide 42 — Routing"""
    _update_placeholder(slide, 0, "Routing — the handoff")
    _update_eyebrow(slide, "Where a clean problem goes next", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Schema row
    schema_top = CONTENT_TOP + Inches(0.3)
    boxes = ["Stages 2-6", "Clean problem", "ITSM ticket", "On-call alert", "Chat notification"]
    colors = [TEAL_DIM, TEAL, DIM_FILL, DIM_FILL, DIM_FILL]
    bdr = [TEAL, TEAL, TEAL, AMBER, PURPLE_LIGHT]
    box_w = Inches(2.0)
    for i, (label, fill, border) in enumerate(zip(boxes, colors, bdr)):
        b_left = ML + i * (box_w + Inches(0.3))
        _add_rect(slide, b_left, schema_top, box_w, Inches(0.65), fill_color=fill, line_color=border, line_width_pt=1.5)
        _add_textbox(slide, label, b_left + Inches(0.1), schema_top + Inches(0.18), box_w - Inches(0.2), Inches(0.32),
                     font=F_HEAD, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i > 0 and i < len(boxes) - 1:
            _add_textbox(slide, "→", b_left - Inches(0.28), schema_top + Inches(0.18), Inches(0.25), Inches(0.32),
                         font=F_HEAD, size=14, color=TEAL)

    bullets = [
        "Pattern: problem trigger → filter on alert group + severity → open/close pairing.",
        "Clean, merged, deduplicated problems are what make downstream consumption work.",
        "Routing noise faster only scales the noise — Stage 7 sits after Stages 2–6.",
    ]
    b_top = schema_top + Inches(0.85)
    for i, b in enumerate(bullets):
        _add_textbox(slide, f"• {b}", ML, b_top + i * Inches(0.5), SW - ML - MR, Inches(0.45),
                     font=F_BODY, size=10, color=WHITE, wrap=True)


def rebuild_slide42(slide):
    """pptx slide 43 — Honest ITSM position + DQL"""
    _update_placeholder(slide, 0, "The honest ITSM position + DQL")
    _update_eyebrow(slide, "What downstream systems can consume today", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    half_w = Inches(5.5)
    gap = Inches(0.3)
    panel_h = Inches(2.2)

    # Left
    _add_rect(slide, ML, CONTENT_TOP, half_w, panel_h, fill_color=DIM_FILL, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide, "CONSUMABLE TODAY", ML + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.28),
                 font=F_HEAD, size=10, bold=True, color=TEAL)
    items_l = [
        "Deterministic root cause: queryable root cause field on problem record.",
        "Consumable directly — no parsing needed.",
        "Direct ITSM consumption.",
    ]
    for i, item in enumerate(items_l):
        _add_textbox(slide, f"• {item}", ML + Inches(0.15), CONTENT_TOP + Inches(0.45) + i * Inches(0.5), half_w - Inches(0.3), Inches(0.45),
                     font=F_BODY, size=9, color=WHITE, wrap=True)

    # Right
    rp_left = ML + half_w + gap
    _add_rect(slide, rp_left, CONTENT_TOP, half_w, panel_h, fill_color=DIM_FILL, line_color=AMBER, line_width_pt=2)
    _add_textbox(slide, "NEEDS A BRIDGE", rp_left + Inches(0.1), CONTENT_TOP + Inches(0.08), half_w - Inches(0.2), Inches(0.28),
                 font=F_HEAD, size=10, bold=True, color=AMBER)
    items_r = [
        "Agentic output: today a comment on problem (CUSTOM_ANNOTATION).",
        "Consuming it means parsing job via standard workflow.",
        "Native field targeted for December Rally release.",
    ]
    for i, item in enumerate(items_r):
        _add_textbox(slide, f"• {item}", rp_left + Inches(0.15), CONTENT_TOP + Inches(0.45) + i * Inches(0.5), half_w - Inches(0.3), Inches(0.45),
                     font=F_BODY, size=9, color=WHITE, wrap=True)

    # DQL below
    dql = '''fetch dt.davis.events.snapshots, from: "DATE", to: now()
| filter event.type == "CUSTOM_ANNOTATION"
| filter in(annotation.problem_ids, "YOUR_PROBLEM_ID")'''
    _simple_code_box(slide, dql, ML, CONTENT_TOP + panel_h + Inches(0.2), SW - ML - MR, Inches(1.4), title="Query agentic annotations on a problem")


def rebuild_slide43(slide):
    """pptx slide 44 — Governance cycle"""
    _update_placeholder(slide, 0, "Governance — Wayfinder Stage 9")
    _update_eyebrow(slide, "One-off tuning delivers temporary improvement only", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    cards = [
        (TEAL,        "Quarterly", "Detector justification", "Review all active detectors."),
        (PURPLE_LIGHT,"Monthly",   "Noise review",           "Check problem volume trend, new noisy sources."),
        (AMBER,       "On change", "Reassessment",           "New service, topology change, detector added."),
    ]
    c_top = CONTENT_TOP + Inches(0.3)
    c_w = (SW - ML - MR - Inches(0.4)) / 3
    for i, (col, freq, title, desc) in enumerate(cards):
        c_left = ML + i * (c_w + Inches(0.2))
        _add_rect(slide, c_left, c_top, c_w, Inches(2.8), fill_color=DIM_FILL, line_color=col, line_width_pt=2)
        _add_textbox(slide, freq, c_left + Inches(0.15), c_top + Inches(0.12), c_w - Inches(0.3), Inches(0.4),
                     font=F_HEAD, size=18, bold=True, color=col, align=PP_ALIGN.CENTER)
        _add_textbox(slide, title, c_left + Inches(0.15), c_top + Inches(0.62), c_w - Inches(0.3), Inches(0.35),
                     font=F_HEAD, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _add_textbox(slide, desc, c_left + Inches(0.15), c_top + Inches(1.05), c_w - Inches(0.3), Inches(0.6),
                     font=F_BODY, size=10, color=MUTED, wrap=True, align=PP_ALIGN.CENTER)

    _footer_bar(slide, "Configs decay as the environment changes. Wayfinder Stage 9 is mandatory, not optional.", fill=TEAL_DIM)


def rebuild_slide44(slide):
    """pptx slide 45 — Reinforcement Q&A"""
    _update_placeholder(slide, 0, "Reinforcement")
    _update_eyebrow(slide, "Quick recall", color=MUTED)
    _remove_content_shapes(slide, keep_images=False)

    # Pathway chip
    _add_rect(slide, SW - MR - Inches(2.2), CONTENT_TOP, Inches(2.2), Inches(0.35), fill_color=DIM_FILL, line_color=TEAL, line_width_pt=1.0)
    _add_textbox(slide, "Pathway diagram", SW - MR - Inches(2.15), CONTENT_TOP + Inches(0.05), Inches(2.1), Inches(0.28),
                 font=F_BODY, size=9, color=TEAL, align=PP_ALIGN.CENTER)

    qa_top = CONTENT_TOP + Inches(0.5)
    qa_h = Inches(2.2)

    # Q1
    _add_rect(slide, ML, qa_top, SW - ML - MR, Inches(0.55), fill_color=TEAL_DIM)
    _add_textbox(slide, "Q1 — A problem has four events, all on one service, and no deterministic root cause. Expected, or a gap?",
                 ML + Inches(0.15), qa_top + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.42),
                 font=F_BODY, size=11, bold=True, color=WHITE, wrap=True)
    a1_top = qa_top + Inches(0.65)
    _add_rect(slide, ML, a1_top, SW - ML - MR, Inches(0.5), fill_color=DIM_FILL, line_color=GOOD, line_width_pt=2)
    _add_textbox(slide, "A: Expected (same-entity merge). Multiple symptoms, one entity, merging working. No second entity for a chain.",
                 ML + Inches(0.15), a1_top + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.38),
                 font=F_BODY, size=10, color=WHITE, wrap=True)

    # Q2
    q2_top = qa_top + qa_h
    _add_rect(slide, ML, q2_top, SW - ML - MR, Inches(0.55), fill_color=TEAL_DIM)
    _add_textbox(slide, "Q2 — A detector fires one event per Kubernetes pod. Which stage owns the fix?",
                 ML + Inches(0.15), q2_top + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.42),
                 font=F_BODY, size=11, bold=True, color=WHITE, wrap=True)
    a2_top = q2_top + Inches(0.65)
    _add_rect(slide, ML, a2_top, SW - ML - MR, Inches(0.5), fill_color=DIM_FILL, line_color=GOOD, line_width_pt=2)
    _add_textbox(slide, "A: Stage 4 — event design. The pod name is a volatile value in the identity tuple.",
                 ML + Inches(0.15), a2_top + Inches(0.08), SW - ML - MR - Inches(0.3), Inches(0.38),
                 font=F_BODY, size=10, color=WHITE, wrap=True)


def rebuild_slide45(slide):
    """pptx slide 46 — Closing question"""
    _remove_content_shapes(slide, keep_images=False)

    # Single large centred panel
    panel_top = CONTENT_TOP + Inches(0.5)
    panel_h = Inches(4.0)
    _add_rect(slide, ML, panel_top, SW - ML - MR, panel_h, fill_color=NAVY, line_color=TEAL, line_width_pt=2)
    _add_textbox(slide,
        "The next time a customer says RCA doesn't work —\nwhat's your first move before you agree with them?",
        ML + Inches(0.5), panel_top + Inches(1.0), SW - ML - MR - Inches(1.0), Inches(2.0),
        font=F_HEAD, size=22, bold=True, color=TEAL, align=PP_ALIGN.CENTER, wrap=True)


# ============================================================
# Delete a slide by 0-based index
# ============================================================

def delete_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slide = prs.slides[idx]
    rId = xml_slides[idx].get(qn('r:id'))
    prs.part.drop_rel(rId)
    del xml_slides[idx]
    print(f"  Deleted slide at original idx={idx}")


# ============================================================
# Main
# ============================================================

def main():
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    prs = Presentation(INPUT)
    print(f"Loaded deck: {len(prs.slides)} slides")

    # Apply all per-slide improvements (use original indices before deletion)
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
        # skip 21 — will delete
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
            print(f"  Rebuilt slide idx={idx}")
        except Exception as e:
            print(f"  ERROR on slide idx={idx}: {e}")

    # Insert TCCA slide after idx=34 (before deletion adjusts indices)
    insert_tcca_slide(prs, after_idx=34)
    print("  Inserted TCCA slide after idx=34")

    # Delete slide at original idx=21 (Merged problem – Agentic RCA)
    # After insertion, idx=21 is unchanged (TCCA was inserted at 35)
    delete_slide(prs, 21)

    prs.save(OUTPUT)
    print(f"\nSaved → {OUTPUT}")
    print(f"Total slides after changes: {len(prs.slides)}")


if __name__ == "__main__":
    main()
