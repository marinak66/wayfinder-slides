#!/usr/bin/env python3
"""Fix slides 1-5 only in session-8-v3.pptx.
Outputs:
  session-8-v4.pptx        — full deck with fixed slides 1-5
  session-8-slides-1-5.pptx — preview of only those 5 slides
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
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
ORANGE_KW   = RGBColor(0xFF, 0xA0, 0x47)
NAVY        = RGBColor(0x0B, 0x1E, 0x3A)
AMBER       = RGBColor(0xFF, 0xB3, 0x00)
DARK_RED    = RGBColor(0x2A, 0x05, 0x05)
GOOD        = RGBColor(0x4C, 0xAF, 0x50)
ERR         = RGBColor(0xF4, 0x43, 0x36)

# ---------- Geometry ----------
SW = 12192000   # 13.333"
SH = 6858000    # 7.5"
LEFT_MARGIN  = Inches(0.62)
RIGHT_MARGIN = Inches(0.62)
CONTENT_TOP  = Inches(1.78)
CONTENT_W    = SW - LEFT_MARGIN - RIGHT_MARGIN

# ---------- Fonts ----------
F_HEAD  = "DT Flow Extrabold"
F_HEAVY = "DT Flow Heavy"
F_BODY  = "DT Flow"
F_CODE  = "Consolas"

KEEP_PH_IDX = {0, 1, 2, 10, 11, 12, 14, 15, 17}

INPUT   = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-v3.pptx"
OUTPUT  = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-v4.pptx"
PREVIEW = "/home/user/wayfinder-slides/wayfinder-engine-main/AI-automation/outputs/projects/session-8/session-8-slides-1-5.pptx"


# ===================================================================
# Helpers (copied / trimmed from v3)
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


def _set_title(slide, text):
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


def _footer(slide, text, fill=TEAL_DIM, text_color=WHITE, size=11):
    top = Inches(6.45)
    h = Inches(0.4)
    _rect(slide, LEFT_MARGIN, top, CONTENT_W, h, fill=fill)
    _tb(slide, text, LEFT_MARGIN + Inches(0.2), top + Inches(0.05),
        CONTENT_W - Inches(0.4), h - Inches(0.1),
        font=F_BODY, size=size, color=text_color, wrap=True)


def _line(slide, x1, y1, x2, y2, color=TEAL, pt=2):
    """Draw a thin rectangle as a line segment (horizontal or vertical)."""
    if abs(x2 - x1) >= abs(y2 - y1):  # horizontal
        h = Pt(pt)
        cy = (y1 + y2) / 2
        _rect(slide, min(x1, x2), cy - h / 2, abs(x2 - x1), h, fill=color)
    else:  # vertical
        w = Pt(pt)
        cx = (x1 + x2) / 2
        _rect(slide, cx - w / 2, min(y1, y2), w, abs(y2 - y1), fill=color)


# ===================================================================
# Slide 1  — Add teal divider line (slide is otherwise fine)
# ===================================================================

def fix_slide1(slide):
    """Add explicit teal divider line below title + eyebrow label."""
    # Add teal divider line at y=1.55" (just below title, above content)
    _rect(slide, LEFT_MARGIN, Inches(1.56), CONTENT_W, Pt(2.5), fill=TEAL)

    # Add small eyebrow label above the line if no idx=14 placeholder
    has_eyebrow = False
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.idx == 14:
                has_eyebrow = True
                break
        except Exception:
            pass
    if not has_eyebrow:
        _tb(slide, "AGENDA", LEFT_MARGIN, Inches(1.32), Inches(3.0), Inches(0.22),
            font=F_BODY, size=10, color=MUTED)


# ===================================================================
# Slide 2  — Course value / Objectives
# ===================================================================

def fix_slide2(slide):
    _set_title(slide, "Course value — what this session delivers")
    _set_eyebrow(slide, "What you will take away")
    _remove_content_shapes(slide, keep_images=False)

    objectives = [
        ("1", "Assess a noisy problem stream diagnostically — naming the pathway stage that owns each fix"),
        ("2", "Distinguish duplicates and fragments from genuine separate incidents"),
        ("3", "Classify expected-empty root cause from genuine RCA gaps — never misread healthy signal as product failure"),
        ("4", "Build an action plan using Wayfinder to improve signal quality, from detection tuning to routing"),
        ("5", "Prepare your customer's problem stream for incident routing and agentic RCA at scale"),
    ]

    box_h = Inches(0.72)
    gap = Inches(0.1)
    badge_w = Inches(0.56)

    for i, (num, text) in enumerate(objectives):
        y = CONTENT_TOP + i * (box_h + gap)
        _rect(slide, LEFT_MARGIN, y, CONTENT_W, box_h, fill=DIM_FILL, line_color=TEAL, line_pt=1.5)
        _rect(slide, LEFT_MARGIN, y, badge_w, box_h, fill=TEAL)
        _tb(slide, num, LEFT_MARGIN, y + Inches(0.18), badge_w, Inches(0.36),
            font=F_HEAD, size=20, bold=True, color=CODE_BG, align=PP_ALIGN.CENTER)
        _tb(slide, text, LEFT_MARGIN + badge_w + Inches(0.18), y + Inches(0.13),
            CONTENT_W - badge_w - Inches(0.28), box_h - Inches(0.26),
            font=F_BODY, size=13, color=WHITE, wrap=True)

    _footer(slide, "After this session you read the problem stream the way a doctor reads a chart — diagnostically.")


# ===================================================================
# Slide 3  — Cold open / The symptom (keep screenshots)
# ===================================================================

def fix_slide3(slide):
    _set_title(slide, "Cold open — the symptom")
    _set_eyebrow(slide, "Problem Stream")
    _remove_content_shapes(slide, keep_images=True)

    # Large framed quote
    quote = ('"A tenant generates hundreds of auto-closing problems a week '
             'and root cause analysis attaches nothing. '
             'The customer concludes the product is broken."')
    quote_h = Inches(1.25)
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, CONTENT_W, quote_h,
          fill=DIM_FILL, line_color=TEAL, line_pt=2.5)
    # Left accent bar
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, Inches(0.12), quote_h, fill=TEAL)
    _tb(slide, quote,
        LEFT_MARGIN + Inches(0.28), CONTENT_TOP + Inches(0.12),
        CONTENT_W - Inches(0.38), quote_h - Inches(0.24),
        font=F_HEAVY, size=14, color=WHITE, italic=True, wrap=True)

    # Three symptom indicator rows (left of screenshots)
    # Screenshots from inspection are at x~5.96" and x~8.81" — keep left col to 5.4"
    left_w = Inches(5.3)
    row_h  = Inches(0.68)
    row_gap = Inches(0.08)
    sym_top = CONTENT_TOP + quote_h + Inches(0.2)

    symptoms = [
        ("📊", "PROBLEM VOLUME",   "Hundreds of auto-closing problems every week"),
        ("🔍", "RCA ATTACHMENT",   "Root cause attaches nothing — 0 % coverage"),
        ("⚠️", "CUSTOMER VERDICT", "The platform is not working as expected"),
    ]
    fills = [RGBColor(0x07, 0x1E, 0x3A), RGBColor(0x08, 0x24, 0x44), RGBColor(0x06, 0x19, 0x32)]

    for i, (icon, label, value) in enumerate(symptoms):
        y = sym_top + i * (row_h + row_gap)
        _rect(slide, LEFT_MARGIN, y, left_w, row_h, fill=fills[i], line_color=TEAL_DIM, line_pt=1)
        _tb(slide, icon, LEFT_MARGIN + Inches(0.1), y + Inches(0.1),
            Inches(0.5), row_h - Inches(0.2),
            font=F_BODY, size=22, color=TEAL, align=PP_ALIGN.CENTER)
        _tb(slide, label, LEFT_MARGIN + Inches(0.72), y + Inches(0.08),
            left_w - Inches(0.82), Inches(0.24),
            font=F_HEAD, size=9, bold=True, color=MUTED)
        _tb(slide, value, LEFT_MARGIN + Inches(0.72), y + Inches(0.32),
            left_w - Inches(0.82), Inches(0.3),
            font=F_BODY, size=13, color=WHITE, wrap=True)

    _footer(slide, "This symptom has two completely different root causes — naming the right one is the skill this session builds.")


# ===================================================================
# Slide 4  — Two entry points + governing principle
# ===================================================================

def fix_slide4(slide):
    _set_title(slide, "Two entry points + governing principle")
    _set_eyebrow(slide, "Same symptom, two different causes")
    _remove_content_shapes(slide, keep_images=False)

    half = (CONTENT_W - Inches(0.3)) / 2

    # ---- Case 1: New tenant (left) ----
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, Inches(3.9),
          fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _rect(slide, LEFT_MARGIN, CONTENT_TOP, half, Inches(0.52), fill=HDR_FILL)
    _tb(slide, "Case 1 · New tenant",
        LEFT_MARGIN + Inches(0.18), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35),
        font=F_HEAD, size=14, bold=True, color=TEAL)

    case1 = [
        "Out-of-the-box detection defaults — deliberately sensitive to maximise coverage",
        "Correct for onboarding; wrong for sustained production",
        "Fix: calibrate sensitivity at Stage 3 as entities mature",
    ]
    for i, item in enumerate(case1):
        _tb(slide, f"•  {item}",
            LEFT_MARGIN + Inches(0.2), CONTENT_TOP + Inches(0.65) + i * Inches(0.82),
            half - Inches(0.35), Inches(0.75),
            font=F_BODY, size=13, color=WHITE, wrap=True)

    # ---- Case 2: Mature estate (right) ----
    rx = LEFT_MARGIN + half + Inches(0.3)
    _rect(slide, rx, CONTENT_TOP, half, Inches(3.9),
          fill=DIM_FILL, line_color=ORANGE_KW, line_pt=2)
    _rect(slide, rx, CONTENT_TOP, half, Inches(0.52), fill=RGBColor(0x1E, 0x12, 0x00))
    _tb(slide, "Case 2 · Mature estate",
        rx + Inches(0.18), CONTENT_TOP + Inches(0.1),
        half - Inches(0.3), Inches(0.35),
        font=F_HEAD, size=14, bold=True, color=ORANGE_KW)

    case2 = [
        "Critical entities never reviewed — still on defaults",
        "Custom detection duplicating native coverage",
        "Per-team configuration sprawl nobody governs",
        "Topology and enrichment gaps that never blocked dashboards",
    ]
    for i, item in enumerate(case2):
        _tb(slide, f"•  {item}",
            rx + Inches(0.2), CONTENT_TOP + Inches(0.65) + i * Inches(0.74),
            half - Inches(0.35), Inches(0.66),
            font=F_BODY, size=12, color=WHITE, wrap=True)

    # ---- Governing principle footer bar ----
    fy = CONTENT_TOP + Inches(4.05)
    _rect(slide, LEFT_MARGIN, fy, CONTENT_W, Inches(0.85),
          fill=RGBColor(0x06, 0x1E, 0x38), line_color=TEAL, line_pt=2)
    _tb(slide,
        "An existing configuration is not evidence of a deliberate, correct configuration.  "
        "Case 2 work is auditing and rationalising — not setting up.",
        LEFT_MARGIN + Inches(0.25), fy + Inches(0.1),
        CONTENT_W - Inches(0.5), Inches(0.65),
        font=F_BODY, size=13, bold=True, color=WHITE, wrap=True)


# ===================================================================
# Slide 5  — The Wayfinder Pathway (non-linear fork diagram)
# ===================================================================

def fix_slide5(slide):
    _set_title(slide, "The Wayfinder Pathway")
    _set_eyebrow(slide, "Nine stages to a clean problem stream")
    _remove_content_shapes(slide, keep_images=False)

    # ── Layout constants ──────────────────────────────────────────
    LEFT_COL_W = Inches(4.85)
    COL_GAP    = Inches(0.35)
    RIGHT_X    = LEFT_MARGIN + LEFT_COL_W + COL_GAP
    RIGHT_W    = CONTENT_W - LEFT_COL_W - COL_GAP
    SH         = Inches(0.58)   # stage box height
    SG         = Inches(0.09)   # gap between stage boxes

    # ── LEFT COLUMN: Stages 1-6 ───────────────────────────────────
    stages_left = [
        (1, "Confirm Platform Foundations",
         "IAM, enrichment at source, central tagging",
         True),   # True = prior session (greyed)
        (2, "Ensure Readiness",
         "Smartscape complete, tracing propagating, baselines settled",
         False),
        (3, "Fine-Tune Built-in Detection",
         "Calibrate native sensitivity per entity — highest-volume fix",
         False),
        (4, "Configure Custom Alerts",
         "Author DQL detectors only where built-in coverage doesn't exist",
         False),
        (5, "Validate the Problem Stream",
         "Confirm events merge, duplicates suppress, one anomaly → one problem",
         False),
        (6, "Measure Root Cause Attachment",
         "Baseline attachment; separate expected-empty from genuine gaps",
         False),
    ]

    badge_w = Inches(0.52)
    s6_center_y = None

    for i, (num, title, desc, prior) in enumerate(stages_left):
        y = CONTENT_TOP + i * (SH + SG)
        border = MUTED if prior else TEAL
        fill   = RGBColor(0x10, 0x10, 0x1C) if prior else DIM_FILL
        t_col  = MUTED if prior else WHITE
        badge_fill = MUTED if prior else TEAL
        badge_text = CODE_BG

        _rect(slide, LEFT_MARGIN, y, LEFT_COL_W, SH,
              fill=fill, line_color=border, line_pt=1.5 if not prior else 1)
        _rect(slide, LEFT_MARGIN, y, badge_w, SH, fill=badge_fill)
        _tb(slide, str(num),
            LEFT_MARGIN, y + Inches(0.14), badge_w, SH - Inches(0.28),
            font=F_HEAD, size=16, bold=True, color=badge_text, align=PP_ALIGN.CENTER)
        # Stage title
        _tb(slide, title,
            LEFT_MARGIN + badge_w + Inches(0.12), y + Inches(0.05),
            LEFT_COL_W - badge_w - Inches(0.18), Inches(0.26),
            font=F_HEAD, size=11, bold=True, color=t_col)
        # Description
        _tb(slide, desc,
            LEFT_MARGIN + badge_w + Inches(0.12), y + Inches(0.3),
            LEFT_COL_W - badge_w - Inches(0.18), Inches(0.24),
            font=F_BODY, size=9, color=MUTED)

        if prior:
            _tb(slide, "◄ Prior session",
                LEFT_MARGIN + badge_w + Inches(2.0), y + Inches(0.17),
                LEFT_COL_W - badge_w - Inches(2.1), Inches(0.22),
                font=F_BODY, size=9, italic=True, color=MUTED)

        if num == 6:
            s6_center_y = y + SH / 2

    # ── RIGHT COLUMN: fork into two parallel branches ─────────────
    # Branch A: Routing (Stage 7) — teal
    # Branch B: Operational (Stages 8-9) — amber

    # "Fork" label + spine
    fork_label_y = CONTENT_TOP
    _tb(slide, "PARALLEL BRANCHES FROM STAGE 6",
        RIGHT_X, fork_label_y, RIGHT_W, Inches(0.25),
        font=F_HEAD, size=8, bold=True, color=MUTED)

    # ── Branch A header ──────────────────────────────────────────
    brA_top = CONTENT_TOP + Inches(0.3)
    brA_label_h = Inches(0.28)
    _tb(slide, "ROUTING", RIGHT_X, brA_top, RIGHT_W, brA_label_h,
        font=F_HEAD, size=9, bold=True, color=TEAL)

    # Stage 7 box
    s7_y = brA_top + brA_label_h + Inches(0.04)
    s7_h = Inches(0.65)
    _rect(slide, RIGHT_X, s7_y, RIGHT_W, s7_h,
          fill=DIM_FILL, line_color=TEAL, line_pt=2)
    _rect(slide, RIGHT_X, s7_y, badge_w, s7_h, fill=TEAL)
    _tb(slide, "7", RIGHT_X, s7_y + Inches(0.16), badge_w, s7_h - Inches(0.3),
        font=F_HEAD, size=16, bold=True, color=CODE_BG, align=PP_ALIGN.CENTER)
    _tb(slide, "Route Confirmed Problems",
        RIGHT_X + badge_w + Inches(0.12), s7_y + Inches(0.06),
        RIGHT_W - badge_w - Inches(0.18), Inches(0.28),
        font=F_HEAD, size=11, bold=True, color=WHITE)
    _tb(slide, "Route clean problems via workflow notifications to the right team",
        RIGHT_X + badge_w + Inches(0.12), s7_y + Inches(0.34),
        RIGHT_W - badge_w - Inches(0.18), Inches(0.25),
        font=F_BODY, size=9, color=MUTED, wrap=True)

    # Channel boxes below Stage 7
    ch_y   = s7_y + s7_h + Inches(0.12)
    ch_h   = Inches(0.58)
    ch_cnt = 3
    ch_w   = (RIGHT_W - Inches(0.16)) / ch_cnt

    channels = [
        ("🖥", "ServiceNow", "incident"),
        ("📱", "PagerDuty",  "on-call"),
        ("💬", "Slack / Teams", "notification"),
    ]
    for i, (icon, name, sub) in enumerate(channels):
        cx = RIGHT_X + i * (ch_w + Inches(0.08))
        _rect(slide, cx, ch_y, ch_w, ch_h,
              fill=CODE_BG, line_color=TEAL_DIM, line_pt=1)
        _tb(slide, icon, cx + Inches(0.08), ch_y + Inches(0.07),
            Inches(0.32), ch_h - Inches(0.14),
            font=F_BODY, size=18, color=TEAL, align=PP_ALIGN.CENTER)
        _tb(slide, name, cx + Inches(0.44), ch_y + Inches(0.06),
            ch_w - Inches(0.5), Inches(0.26),
            font=F_HEAD, size=10, bold=True, color=TEAL)
        _tb(slide, sub, cx + Inches(0.44), ch_y + Inches(0.3),
            ch_w - Inches(0.5), Inches(0.2),
            font=F_BODY, size=9, color=MUTED)

    # ── Branch B: Operational (Stages 8-9) ──────────────────────
    brB_top = ch_y + ch_h + Inches(0.2)
    _tb(slide, "OPERATIONAL", RIGHT_X, brB_top, RIGHT_W, Inches(0.25),
        font=F_HEAD, size=9, bold=True, color=AMBER)

    s89 = [
        (8, "Measure Signal Quality (KPIs)",
         "Read the Stage 8 dashboard as a trend — evidence that tuning worked"),
        (9, "Establish Governance",
         "Quarterly / monthly / on-change review — tuning must not decay"),
    ]
    s89_w = (RIGHT_W - Inches(0.12)) / 2
    for i, (num, title, desc) in enumerate(s89):
        bx = RIGHT_X + i * (s89_w + Inches(0.12))
        by = brB_top + Inches(0.28)
        _rect(slide, bx, by, s89_w, Inches(0.72),
              fill=DIM_FILL, line_color=AMBER, line_pt=1.5)
        _rect(slide, bx, by, badge_w, Inches(0.72), fill=AMBER)
        _tb(slide, str(num), bx, by + Inches(0.2), badge_w, Inches(0.32),
            font=F_HEAD, size=14, bold=True, color=CODE_BG, align=PP_ALIGN.CENTER)
        _tb(slide, title, bx + badge_w + Inches(0.1), by + Inches(0.05),
            s89_w - badge_w - Inches(0.15), Inches(0.28),
            font=F_HEAD, size=9, bold=True, color=WHITE)
        _tb(slide, desc, bx + badge_w + Inches(0.1), by + Inches(0.34),
            s89_w - badge_w - Inches(0.15), Inches(0.32),
            font=F_BODY, size=8, color=MUTED, wrap=True)

    # ── Spine connector: from Stage 6 right edge → splits into 7 and 8-9 ──
    # Horizontal stub from Stage 6 right edge to spine
    spine_x = RIGHT_X - Inches(0.18)
    _line(slide, LEFT_MARGIN + LEFT_COL_W, s6_center_y,
          spine_x, s6_center_y, color=TEAL, pt=2)

    # Spine: vertical bar from Stage 7 center to Stage 8-9 center
    s7_center_y = s7_y + s7_h / 2
    s89_by = brB_top + Inches(0.28) + Inches(0.36)   # center of 8-9 boxes
    _line(slide, spine_x, min(s7_center_y, s6_center_y),
          spine_x, max(s89_by, s6_center_y), color=TEAL, pt=2)

    # Stub → Stage 7
    _line(slide, spine_x, s7_center_y,
          RIGHT_X, s7_center_y, color=TEAL, pt=2)

    # Stub → Stage 8-9
    _line(slide, spine_x, s89_by,
          RIGHT_X, s89_by, color=AMBER, pt=2)

    # Vertical stub from Stage 7 down to channels
    _line(slide, RIGHT_X + RIGHT_W / 2, s7_y + s7_h,
          RIGHT_X + RIGHT_W / 2, ch_y, color=TEAL, pt=2)

    _footer(slide, "Stages 7 and 8-9 are parallel consumers of a clean problem stream — not sequential. The fork at Stage 6 is deliberate.")


# ===================================================================
# Slide deletion helper (for preview)
# ===================================================================

def delete_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    rId = xml_slides[idx].get(qn('r:id'))
    prs.part.drop_rel(rId)
    del xml_slides[idx]


# ===================================================================
# Main
# ===================================================================

def main():
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

    # ── Apply fixes to the full deck ─────────────────────────────
    prs = Presentation(INPUT)
    print(f"Loaded: {len(prs.slides)} slides from {INPUT}")

    fixes = [
        (0, fix_slide1),
        (1, fix_slide2),
        (2, fix_slide3),
        (3, fix_slide4),
        (4, fix_slide5),
    ]
    for idx, fn in fixes:
        try:
            fn(prs.slides[idx])
            print(f"  Fixed slide idx={idx} (pptx {idx+1})")
        except Exception as e:
            import traceback
            print(f"  ERROR slide idx={idx}: {e}")
            traceback.print_exc()

    prs.save(OUTPUT)
    print(f"\nSaved full deck → {OUTPUT}")
    print(f"Total slides: {len(prs.slides)}")

    # ── 5-slide preview: delete slides 5..end ───────────────────
    prs2 = Presentation(OUTPUT)
    while len(prs2.slides) > 5:
        delete_slide(prs2, len(prs2.slides) - 1)
    prs2.save(PREVIEW)
    print(f"Saved 5-slide preview → {PREVIEW}")


if __name__ == "__main__":
    main()
