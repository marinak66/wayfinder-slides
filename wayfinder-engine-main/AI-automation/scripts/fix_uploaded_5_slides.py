#!/usr/bin/env python3
"""Fix the 5 slides in the uploaded session-8-v3.pptx:
 1. Query hygiene + tolerance bands  — add full-width teal line
 2. Which stage owned it?            — add full-width teal line
 3. Match the category to condition  — redesign as 3-col table (reference style)
 4. Fragmentation patterns           — add labelled headers above images
 5. Routing — the handoff            — redesign as non-linear diagram
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# ---------- Palette ----------
TEAL        = RGBColor(0x57, 0xE5, 0xE3)
TEAL_DIM    = RGBColor(0x2A, 0x7F, 0x7D)
MUTED       = RGBColor(0xA9, 0xB6, 0xC4)
CODE_BG     = RGBColor(0x07, 0x13, 0x2A)
CODE_BORDER = RGBColor(0x1A, 0x35, 0x55)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DIM_FILL    = RGBColor(0x0A, 0x1E, 0x35)
HDR_FILL    = RGBColor(0x08, 0x33, 0x66)
NAVY        = RGBColor(0x0B, 0x1E, 0x3A)
AMBER       = RGBColor(0xFF, 0xB3, 0x00)
DARK_RED    = RGBColor(0x2A, 0x05, 0x05)
ERR_RED     = RGBColor(0xFF, 0x45, 0x45)
GOOD        = RGBColor(0x4C, 0xAF, 0x50)
ROW_A       = RGBColor(0x07, 0x1C, 0x36)
ROW_B       = RGBColor(0x09, 0x22, 0x40)
ORANGE_KW   = RGBColor(0xFF, 0xA0, 0x47)

# ---------- Geometry ----------
SW = 12192000   # 13.333"
SH = 6858000    # 7.5"
LEFT_MARGIN  = Inches(0.62)
RIGHT_MARGIN = Inches(0.62)
CONTENT_TOP  = Inches(1.78)
CONTENT_W    = SW - LEFT_MARGIN - RIGHT_MARGIN

LINE_Y = Inches(1.64)   # y-position of the teal divider line

# ---------- Fonts ----------
F_HEAD  = "DT Flow Extrabold"
F_HEAVY = "DT Flow Heavy"
F_BODY  = "DT Flow"
F_CODE  = "Consolas"

KEEP_PH_IDX = {0, 1, 2, 10, 11, 12, 14, 15, 17}

INPUT   = "/root/.claude/uploads/4c9f84f5-5bfc-5394-95dc-07a5546c7be9/3abc6b31-session-8-v3.pptx"
OUTPUT  = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-slides-fixed.pptx"


# ===================================================================
# Helpers
# ===================================================================

def _remove_content_shapes(slide, keep_images=True):
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


def _teal_line(slide):
    """Add a full-width teal divider line at LINE_Y."""
    _rect(slide, LEFT_MARGIN, LINE_Y, CONTENT_W, Pt(2.5), fill=TEAL)


def _footer(slide, text, fill=TEAL_DIM, size=11):
    top = Inches(6.48)
    h = Inches(0.38)
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, h, fill=fill)
    _tb(slide, text, LEFT_MARGIN + Inches(0.18), top + Inches(0.04),
        CONTENT_W - Inches(0.36), h - Inches(0.08),
        font=F_BODY, size=size, color=WHITE, wrap=True)


def _line_shape(slide, x1, y1, x2, y2, color=TEAL, pt=2):
    """Draw a thin line as a rectangle."""
    if abs(x2 - x1) >= abs(y2 - y1):   # horizontal
        h = Pt(pt)
        cy = (y1 + y2) / 2
        _rect(slide, min(x1, x2), cy - h / 2, abs(x2 - x1), h, fill=color)
    else:                                 # vertical
        w = Pt(pt)
        cx = (x1 + x2) / 2
        _rect(slide, cx - w / 2, min(y1, y2), w, abs(y2 - y1), fill=color)


# ===================================================================
# Slide 1 — Query hygiene: just add the full-width teal line
# ===================================================================

def fix_slide1(slide):
    _teal_line(slide)
    # Add eyebrow label if missing (no idx=14 on this slide)
    has14 = any(ph.placeholder_format.idx == 14
                for ph in slide.placeholders if ph.placeholder_format)
    if not has14:
        _tb(slide, "Two rules for every measurement query",
            LEFT_MARGIN, Inches(0.44), Inches(10.0), Inches(0.28),
            font=F_BODY, size=10, color=MUTED)


# ===================================================================
# Slide 2 — Which stage owned it: just add the full-width teal line
# ===================================================================

def fix_slide2(slide):
    _teal_line(slide)


# ===================================================================
# Slide 3 — Match the category to the condition
#           Redesign as 3-column reference table (images 11/12)
# ===================================================================

def fix_slide3(slide):
    _remove_content_shapes(slide, keep_images=False)
    _teal_line(slide)

    C1 = Inches(5.1)    # CONDITION col width
    C2 = Inches(4.0)    # EVENT.TYPE col width
    C3 = CONTENT_W - C1 - C2   # EVENT.CATEGORY (remainder ~3.0")

    # ── Warning bar ──────────────────────────────────────────────
    warn_y = CONTENT_TOP
    warn_h = Inches(0.38)
    _rect(slide, LEFT_MARGIN, warn_y, CONTENT_W, warn_h,
          fill=DARK_RED, line_color=AMBER, line_pt=1)
    _tb(slide, "⚠  Do not default to CUSTOM_ALERT for error, slowdown or availability signals.",
        LEFT_MARGIN + Inches(0.15), warn_y + Inches(0.05),
        CONTENT_W - Inches(0.3), warn_h - Inches(0.1),
        font=F_BODY, size=11, bold=True, color=AMBER)

    # ── Column headers ────────────────────────────────────────────
    hdr_y = warn_y + warn_h + Inches(0.06)
    hdr_h = Inches(0.35)
    _rect(slide, LEFT_MARGIN, hdr_y, CONTENT_W, hdr_h, fill=HDR_FILL)
    for x, w, label in [
        (LEFT_MARGIN + Inches(0.55), C1 - Inches(0.55), "CONDITION"),
        (LEFT_MARGIN + C1, C2, "EVENT.TYPE"),
        (LEFT_MARGIN + C1 + C2, C3, "EVENT.CATEGORY"),
    ]:
        _tb(slide, label, x, hdr_y + Inches(0.06), w, hdr_h - Inches(0.12),
            font=F_HEAD, size=10, bold=True, color=MUTED)

    # ── Tier 1 rows: Opens a Problem ──────────────────────────────
    _tb(slide, "OPENS A PROBLEM",
        LEFT_MARGIN, hdr_y + hdr_h + Inches(0.04), Inches(4.0), Inches(0.22),
        font=F_HEAD, size=8, bold=True, color=TEAL)
    row_start = hdr_y + hdr_h + Inches(0.28)
    row_h = Inches(0.52)
    row_gap = Inches(0.04)

    tier1 = [
        ("⊘", "Outage / health-check failure",      "AVAILABILITY_EVENT",        "AVAILABILITY"),
        ("⚠", "Error-rate spike (5xx, exceptions)", "ERROR_EVENT",               "ERROR"),
        ("⏱", "Latency spike / throughput drop",    "SERVICE_SLOWDOWN",          "SLOWDOWN"),
        ("⬛", "Resource saturation (CPU, memory)",  "RESOURCE_CONTENTION_EVENT", "RESOURCE_CONTENTION"),
    ]
    row_fills = [ROW_A, ROW_B, ROW_A, ROW_B]

    for i, (icon, cond, etype, ecat) in enumerate(tier1):
        y = row_start + i * (row_h + row_gap)
        f = row_fills[i]
        _rect(slide, LEFT_MARGIN, y, CONTENT_W, row_h, fill=f, line_color=TEAL_DIM, line_pt=0.5)
        # icon circle bg
        _rect(slide, LEFT_MARGIN, y, Inches(0.46), row_h, fill=DIM_FILL)
        _tb(slide, icon, LEFT_MARGIN, y + Inches(0.1), Inches(0.46), row_h - Inches(0.2),
            font=F_BODY, size=16, color=TEAL, align=PP_ALIGN.CENTER)
        # Condition
        _tb(slide, cond, LEFT_MARGIN + Inches(0.52), y + Inches(0.1),
            C1 - Inches(0.58), row_h - Inches(0.2),
            font=F_BODY, size=12, color=WHITE)
        # Event type (teal monospace)
        _tb(slide, etype, LEFT_MARGIN + C1, y + Inches(0.1),
            C2 - Inches(0.1), row_h - Inches(0.2),
            font=F_CODE, size=11, bold=True, color=TEAL)
        # Event category
        _tb(slide, ecat, LEFT_MARGIN + C1 + C2, y + Inches(0.1),
            C3 - Inches(0.1), row_h - Inches(0.2),
            font=F_CODE, size=11, bold=True, color=TEAL)

    # CUSTOM_ALERT row (amber / last-resort)
    ca_y = row_start + 4 * (row_h + row_gap)
    _rect(slide, LEFT_MARGIN, ca_y, CONTENT_W, row_h,
          fill=RGBColor(0x18, 0x10, 0x00), line_color=AMBER, line_pt=1)
    _rect(slide, LEFT_MARGIN, ca_y, Inches(0.46), row_h, fill=DIM_FILL)
    _tb(slide, "⋯", LEFT_MARGIN, ca_y + Inches(0.1), Inches(0.46), row_h - Inches(0.2),
        font=F_BODY, size=16, color=AMBER, align=PP_ALIGN.CENTER)
    _tb(slide, "Fits none of the above (last resort)",
        LEFT_MARGIN + Inches(0.52), ca_y + Inches(0.1), C1 - Inches(0.58), row_h - Inches(0.2),
        font=F_BODY, size=12, color=MUTED)
    _tb(slide, "CUSTOM_ALERT", LEFT_MARGIN + C1, ca_y + Inches(0.1),
        C2 - Inches(0.1), row_h - Inches(0.2),
        font=F_CODE, size=11, bold=True, color=AMBER)
    _tb(slide, "CUSTOM", LEFT_MARGIN + C1 + C2, ca_y + Inches(0.1),
        C3 - Inches(0.1), row_h - Inches(0.2),
        font=F_CODE, size=11, bold=True, color=AMBER)

    # ── Separator: Does NOT open a Problem ───────────────────────
    sep_y = ca_y + row_h + Inches(0.08)
    sep_h = Inches(0.3)
    _rect(slide, LEFT_MARGIN, sep_y, CONTENT_W, sep_h, fill=RGBColor(0x14, 0x06, 0x06))
    _rect(slide, LEFT_MARGIN, sep_y + sep_h / 2 - Pt(0.75), CONTENT_W, Pt(1.5), fill=MUTED)
    _tb(slide, "DOES NOT OPEN A PROBLEM",
        LEFT_MARGIN + Inches(4.5), sep_y + Inches(0.04), Inches(4.0), sep_h - Inches(0.08),
        font=F_HEAD, size=9, bold=True, color=ERR_RED, align=PP_ALIGN.CENTER)

    # ── Tier 2 rows: No problem ───────────────────────────────────
    tier2 = [
        ("🔔", "Early warning / calibration period", "WARNING",     "WARNING"),
        ("ℹ",  "Deployment / config-change marker",  "CUSTOM_INFO", "INFO"),
    ]
    t2_start = sep_y + sep_h + Inches(0.04)
    t2_row_h = Inches(0.45)
    for i, (icon, cond, etype, ecat) in enumerate(tier2):
        y = t2_start + i * (t2_row_h + Inches(0.03))
        _rect(slide, LEFT_MARGIN, y, CONTENT_W, t2_row_h, fill=CODE_BG)
        _rect(slide, LEFT_MARGIN, y, Inches(0.46), t2_row_h, fill=DIM_FILL)
        _tb(slide, icon, LEFT_MARGIN, y + Inches(0.06), Inches(0.46), t2_row_h - Inches(0.12),
            font=F_BODY, size=14, color=MUTED, align=PP_ALIGN.CENTER)
        _tb(slide, cond, LEFT_MARGIN + Inches(0.52), y + Inches(0.08),
            C1 - Inches(0.58), t2_row_h - Inches(0.16), font=F_BODY, size=11, color=MUTED)
        _tb(slide, etype, LEFT_MARGIN + C1, y + Inches(0.08),
            C2 - Inches(0.1), t2_row_h - Inches(0.16), font=F_CODE, size=10, color=MUTED)
        _tb(slide, ecat, LEFT_MARGIN + C1 + C2, y + Inches(0.08),
            C3 - Inches(0.1), t2_row_h - Inches(0.16), font=F_CODE, size=10, color=MUTED)

    _footer(slide, "CUSTOM_ALERT is the last resort, not the default.")


# ===================================================================
# Slide 4 — Fragmentation patterns: add labels above images
# ===================================================================

def fix_slide4(slide):
    # Keep both images, remove only non-image shapes
    _remove_content_shapes(slide, keep_images=True)
    _teal_line(slide)

    # Add eyebrow
    _tb(slide, "Fragmentation cases shown",
        LEFT_MARGIN, Inches(0.44), Inches(10.0), Inches(0.28),
        font=F_BODY, size=10, color=MUTED)

    # Images are at:
    #   Left:  x=0.92"  top=2.70"  w=3.43"  h=4.30"
    #   Right: x=6.78"  top=2.38"  w=3.44"  h=4.29"
    # Add header strip ABOVE each image

    # ── Header above left image ───────────────────────────────────
    lbl_h = Inches(0.52)
    lbl1_y = Inches(2.70) - lbl_h - Inches(0.05)   # just above left image
    _rect(slide, Inches(0.92), lbl1_y, Inches(3.43), lbl_h,
          fill=HDR_FILL, line_color=TEAL, line_pt=1.5)
    _tb(slide, "Pattern 1",
        Inches(0.92) + Inches(0.12), lbl1_y + Inches(0.04),
        Inches(3.2), Inches(0.22),
        font=F_HEAD, size=10, bold=True, color=TEAL)
    _tb(slide, "Volume fragmentation",
        Inches(0.92) + Inches(0.12), lbl1_y + Inches(0.26),
        Inches(3.2), Inches(0.22),
        font=F_BODY, size=10, color=WHITE)

    # ── Header above right image ──────────────────────────────────
    lbl2_y = Inches(2.38) - lbl_h - Inches(0.05)   # just above right image
    _rect(slide, Inches(6.78), lbl2_y, Inches(3.44), lbl_h,
          fill=HDR_FILL, line_color=ORANGE_KW, line_pt=1.5)
    _tb(slide, "Pattern 2",
        Inches(6.78) + Inches(0.12), lbl2_y + Inches(0.04),
        Inches(3.2), Inches(0.22),
        font=F_HEAD, size=10, bold=True, color=ORANGE_KW)
    _tb(slide, "Over-differentiated event.name",
        Inches(6.78) + Inches(0.12), lbl2_y + Inches(0.26),
        Inches(3.2), Inches(0.22),
        font=F_BODY, size=10, color=WHITE)

    # ── Rule of thumb (centre column, between images) ─────────────
    # Gap between images: x=4.35" to x=6.78", width=2.43"
    mid_x = Inches(4.40)
    mid_w = Inches(2.3)
    mid_y = CONTENT_TOP + Inches(0.1)
    mid_h = Inches(4.5)
    _rect(slide, mid_x, mid_y, mid_w, mid_h,
          fill=DIM_FILL, line_color=TEAL_DIM, line_pt=1)
    _tb(slide, "RULE OF THUMB",
        mid_x + Inches(0.12), mid_y + Inches(0.12),
        mid_w - Inches(0.24), Inches(0.25),
        font=F_HEAD, size=9, bold=True, color=TEAL)
    _tb(slide,
        "If one anomaly manifests across many dimensions — "
        "cities, partitions, pods, regions — "
        "that dimension is data on one event, "
        "not a reason to emit one event per value.\n\n"
        "Keep high-cardinality context in\nevent.description.",
        mid_x + Inches(0.12), mid_y + Inches(0.42),
        mid_w - Inches(0.24), Inches(3.8),
        font=F_BODY, size=10, color=WHITE, wrap=True)

    _footer(slide,
            "Fragmentation = one incident split into multiple problems that should have been merged into one.")


# ===================================================================
# Slide 5 — Routing — the handoff: non-linear diagram
#           Matches reference image: down → right → down (L-shape)
# ===================================================================

def fix_slide5(slide):
    _remove_content_shapes(slide, keep_images=False)
    _teal_line(slide)

    # ── Left column: Stages 2-6 block ────────────────────────────
    LX = LEFT_MARGIN
    LW = Inches(4.6)
    BH = Inches(0.95)   # box height

    stages_y = CONTENT_TOP + Inches(0.05)

    # "Stages 2-6" group box
    _rect(slide, LX, stages_y, LW, BH,
          fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _rect(slide, LX, stages_y, Inches(0.5), BH, fill=TEAL)
    _tb(slide, "2–6", LX, stages_y + Inches(0.22), Inches(0.5), Inches(0.52),
        font=F_HEAD, size=18, bold=True, color=CODE_BG, align=PP_ALIGN.CENTER)
    _tb(slide, "Stages 2–6  ·  This session",
        LX + Inches(0.62), stages_y + Inches(0.08),
        LW - Inches(0.72), Inches(0.3),
        font=F_HEAD, size=13, bold=True, color=WHITE)
    _tb(slide, "Ensure readiness · Tune detection · Custom alerts · Validate stream · Measure RCA",
        LX + Inches(0.62), stages_y + Inches(0.42),
        LW - Inches(0.72), Inches(0.46),
        font=F_BODY, size=10, color=MUTED, wrap=True)

    # Arrow down (left column center)
    arr_cx = LX + LW / 2
    arr1_top = stages_y + BH
    arr1_bot = arr1_top + Inches(0.32)
    _line_shape(slide, arr_cx, arr1_top, arr_cx, arr1_bot, color=TEAL, pt=2)
    # Arrowhead (small downward triangle via thin rect cluster)
    _rect(slide, arr_cx - Inches(0.05), arr1_bot - Inches(0.02),
          Inches(0.1), Inches(0.1), fill=TEAL)

    # "Clean problem" output box
    clean_y = arr1_bot + Inches(0.04)
    _rect(slide, LX, clean_y, LW, BH,
          fill=DIM_FILL, line_color=GOOD, line_pt=2)
    _rect(slide, LX, clean_y, Inches(0.5), BH, fill=GOOD)
    _tb(slide, "✓", LX, clean_y + Inches(0.22), Inches(0.5), Inches(0.52),
        font=F_BODY, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _tb(slide, "Clean, merged, deduplicated problem",
        LX + Inches(0.62), clean_y + Inches(0.08),
        LW - Inches(0.72), Inches(0.3),
        font=F_HEAD, size=12, bold=True, color=WHITE)
    _tb(slide, "With impact context, root-cause field populated, duplicates suppressed",
        LX + Inches(0.62), clean_y + Inches(0.42),
        LW - Inches(0.72), Inches(0.46),
        font=F_BODY, size=10, color=MUTED, wrap=True)

    # ── Horizontal arrow: clean problem → Stage 7 ─────────────────
    arr2_y = clean_y + BH / 2
    arr2_left = LX + LW
    RX = arr2_left + Inches(0.42)   # right column x start
    _line_shape(slide, arr2_left, arr2_y, RX - Inches(0.04), arr2_y, color=TEAL, pt=2)
    _rect(slide, RX - Inches(0.1), arr2_y - Inches(0.05),
          Inches(0.1), Inches(0.1), fill=TEAL)

    # ── Right column: Stage 7 box ─────────────────────────────────
    RW = CONTENT_W - LW - Inches(0.42)
    s7_y = clean_y   # same level as clean problem box
    _rect(slide, RX, s7_y, RW, BH,
          fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _rect(slide, RX, s7_y, Inches(0.5), BH, fill=TEAL)
    _tb(slide, "7", RX, s7_y + Inches(0.22), Inches(0.5), Inches(0.52),
        font=F_HEAD, size=18, bold=True, color=CODE_BG, align=PP_ALIGN.CENTER)
    _tb(slide, "Stage 7  ·  Route via Workflows",
        RX + Inches(0.62), s7_y + Inches(0.08),
        RW - Inches(0.72), Inches(0.3),
        font=F_HEAD, size=13, bold=True, color=WHITE)
    _tb(slide, "Pattern: problem trigger → filter on alert group + severity → open/close pairing",
        RX + Inches(0.62), s7_y + Inches(0.42),
        RW - Inches(0.72), Inches(0.46),
        font=F_BODY, size=10, color=MUTED, wrap=True)

    # Arrow down from Stage 7 center
    arr3_cx = RX + RW / 2
    arr3_top = s7_y + BH
    arr3_bot = arr3_top + Inches(0.28)
    _line_shape(slide, arr3_cx, arr3_top, arr3_cx, arr3_bot, color=TEAL, pt=2)

    # ── Channel boxes ─────────────────────────────────────────────
    ch_y = arr3_bot + Inches(0.04)
    ch_h = Inches(0.75)
    ch_cnt = 3
    ch_w = (RW - Inches(0.2)) / ch_cnt

    channels = [
        (TEAL,       "🖥", "ServiceNow",   "ITSM ticket / incident"),
        (AMBER,      "📱", "PagerDuty",    "On-call alert"),
        (GOOD,       "💬", "Slack / Teams","Chat notification"),
    ]
    for i, (col, icon, name, sub) in enumerate(channels):
        cx = RX + i * (ch_w + Inches(0.1))
        _rect(slide, cx, ch_y, ch_w, ch_h,
              fill=CODE_BG, line_color=col, line_pt=2)
        _tb(slide, icon, cx + Inches(0.1), ch_y + Inches(0.1),
            Inches(0.35), ch_h - Inches(0.2),
            font=F_BODY, size=20, color=col, align=PP_ALIGN.CENTER)
        _tb(slide, name, cx + Inches(0.5), ch_y + Inches(0.1),
            ch_w - Inches(0.6), Inches(0.28),
            font=F_HEAD, size=11, bold=True, color=col)
        _tb(slide, sub, cx + Inches(0.5), ch_y + Inches(0.38),
            ch_w - Inches(0.6), Inches(0.28),
            font=F_BODY, size=10, color=MUTED)

    # ── Key point below ───────────────────────────────────────────
    note_y = ch_y + ch_h + Inches(0.22)
    _rect(slide, RX, note_y, RW, Inches(0.48),
          fill=DIM_FILL, line_color=TEAL_DIM, line_pt=1)
    _tb(slide,
        "Routing noise faster only scales the noise. "
        "Stage 7 sits after Stages 2–6, not before.",
        RX + Inches(0.15), note_y + Inches(0.08),
        RW - Inches(0.3), Inches(0.34),
        font=F_BODY, size=11, italic=True, color=MUTED, wrap=True)

    _footer(slide, "Clean, merged, deduplicated problems are what make downstream ITSM and on-call consumption reliable.")


# ===================================================================
# Main
# ===================================================================

def main():
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    prs = Presentation(INPUT)
    print(f"Loaded: {len(prs.slides)} slides")

    fixes = [
        (0, fix_slide1, "Query hygiene — teal line"),
        (1, fix_slide2, "Which stage owned it — teal line"),
        (2, fix_slide3, "Match category to condition — 3-col table"),
        (3, fix_slide4, "Fragmentation patterns — headers above images"),
        (4, fix_slide5, "Routing — non-linear diagram"),
    ]
    for idx, fn, label in fixes:
        try:
            fn(prs.slides[idx])
            print(f"  ✓ Slide {idx+1}: {label}")
        except Exception as e:
            import traceback
            print(f"  ✗ Slide {idx+1}: {e}")
            traceback.print_exc()

    prs.save(OUTPUT)
    print(f"\nSaved → {OUTPUT}")


if __name__ == "__main__":
    main()
