# Wayfinder Slides

AI-assisted slide deck generation for **Dynatrace D1 Event 2026** technical sessions.

## What this repo contains

```
wayfinder-engine-main/
└── AI-automation/
    ├── inputs/          # Source PPTX drafts + build instruction DOCX files
    ├── outputs/         # Generated PPTX files (gitignored)
    └── scripts/         # Python scripts that rebuild and improve slides
```

## Scripts

| Script | Purpose |
|--------|---------|
| `improve_slides.py` | Initial prototype — rebuilds select slides from Session 8 draft |
| `improve_slides_v2.py` | Full rebuild of Session 8 (46 slides); fixes fonts, eyebrow placeholder, DQL panel layout |
| `improve_slides_v3.py` | Comprehensive overhaul: full-width teal divider line on all slides, correct font hierarchy, no DQL/image overlap on slides 24-28, reference-style horizontal band layouts |
| `fix_slides_1_5.py` | Targeted fixes for pptx slides 1-5 from the full session-8 deck |
| `fix_uploaded_5_slides.py` | Fixes for a standalone 5-slide review file (query hygiene, reveal table, event category table, fragmentation patterns, routing diagram) |

## Session 8 — From Noise to Signal

**Problem Quality & Root Cause Analysis** · D1 Event 2026 · Technical Deep Dive Track · 75 min

Covers the nine-stage Wayfinder pathway: Confirm Platform Foundations → Ensure Readiness → Fine-Tune Detection → Configure Custom Alerts → Validate Problem Stream → Measure RCA Attachment → Route Problems → Measure KPIs → Establish Governance.

## Design conventions

- **Dark background** with Dynatrace brand palette (teal `#57E5E3`, navy `#0B1E3A`)
- **Fonts**: DT Flow (body/eyebrow), DT Flow Extrabold (section headers), DT Flow Heavy (large stats), Consolas (code)
- **Teal divider line** (`ph idx=15`) preserved on every slide below the title
- **Eyebrow placeholder** (`ph idx=14`) carries the slide sub-header in DT Flow body font
- Title placeholder (`ph idx=0`) — no explicit font override; theme applies the correct DT Flow weight

## Requirements

```bash
pip install python-pptx python-docx
```

## Usage

```bash
cd wayfinder-engine-main/AI-automation
python3 scripts/improve_slides_v3.py
# Output → outputs/projects/session-8/session-8-v3.pptx
```
