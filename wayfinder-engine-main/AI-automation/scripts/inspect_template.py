"""Inspect the dt-template.pptx and the reference deck to see available
slide layouts, master theme fonts, and default colors."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "dt-template" / "reduced-template.pptx"
REFERENCE = ROOT / "inputs" / "Reference slide-deck.pptx"


def dump_pres(path: Path, tag: str) -> None:
    print(f"\n========== {tag}: {path.name} ==========")
    if not path.exists():
        print(f"  MISSING: {path}")
        return
    prs = Presentation(str(path))
    print(f"  slide size: {Emu(prs.slide_width).inches:.2f}\" x {Emu(prs.slide_height).inches:.2f}\"")
    print(f"  slide masters: {len(prs.slide_masters)}")
    for mi, master in enumerate(prs.slide_masters):
        print(f"  [Master {mi}] name={master.name!r}  layouts={len(master.slide_layouts)}")
        for li, layout in enumerate(master.slide_layouts):
            phs = [(ph.placeholder_format.idx, ph.placeholder_format.type, ph.name) for ph in layout.placeholders]
            print(f"     - layout[{li}] name={layout.name!r}  placeholders={phs}")

    print(f"  existing slides: {len(prs.slides)}")
    for si, slide in enumerate(prs.slides):
        print(f"  [slide {si}] layout={slide.slide_layout.name!r}")
        for shp in slide.shapes:
            tf = shp.text_frame.text[:80].replace("\n", " | ") if shp.has_text_frame else ""
            print(f"       shape name={shp.name!r} type={shp.shape_type} text={tf!r}")


if __name__ == "__main__":
    dump_pres(TEMPLATE, "TEMPLATE")
    dump_pres(REFERENCE, "REFERENCE")
