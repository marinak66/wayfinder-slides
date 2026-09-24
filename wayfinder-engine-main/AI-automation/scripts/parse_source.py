"""Parse the Dynatrace-ServiceNow source .docx into an ordered list of Slide
objects.

Structure in the source doc:
- Heading 1                                 -> deck-level title (ignored per slide)
- Heading 2 "N - <title> - Slide Header"    -> starts a new slide
- Normal (Web) starting with "Key message:" -> slide subtitle (key message)
- Heading 3 "On Slide:" (or with suffix)    -> marks start of body
- Normal (Web) after "On Slide"             -> body paragraphs (verbatim bullets)
- Tables interleaved                        -> belong to the enclosing slide,
                                               in document order

The parser walks the docx body in real document order so tables land inside
the correct slide.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Optional

from docx import Document
from docx.document import Document as DocxDoc
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

SLIDE_HEADER_SUFFIX_RE = re.compile(r"\s*-\s*Slide Header\s*:?\s*$", re.IGNORECASE)
SLIDE_HEADER_ALT_RE = re.compile(r"\s*Slide Header\s*:?\s*$", re.IGNORECASE)
LEADING_NUM_RE = re.compile(r"^\s*\d+\s*[-–.:]\s*")
KEY_MESSAGE_PREFIX_RE = re.compile(r"^\s*Key message\s*:\s*", re.IGNORECASE)
ON_SLIDE_PREFIX_RE = re.compile(r"^\s*On Slide\s*:?\s*", re.IGNORECASE)


@dataclass
class TableBlock:
    """Body element: a table."""

    rows: list[list[str]]  # list of rows, each row is a list of cell strings

    @property
    def header(self) -> list[str]:
        return self.rows[0] if self.rows else []

    @property
    def data_rows(self) -> list[list[str]]:
        return self.rows[1:] if len(self.rows) > 1 else []


@dataclass
class TextBlock:
    """Body element: a paragraph (candidate for a bullet)."""

    text: str


@dataclass
class Slide:
    number: str  # e.g. "0", "1", "12a"
    title: str  # cleaned title (leading number + trailing "Slide Header" stripped)
    key_message: str  # cleaned key message (prefix stripped)
    body: list[TextBlock | TableBlock] = field(default_factory=list)
    raw_title: str = ""  # original heading text, for debugging


def _iter_block_items(doc: DocxDoc) -> Iterator[Paragraph | Table]:
    """Yield paragraphs and tables in the order they appear in the document body."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _extract_table(tbl: Table) -> TableBlock:
    rows: list[list[str]] = []
    for row in tbl.rows:
        cells: list[str] = []
        for cell in row.cells:
            text = cell.text.strip()
            cells.append(text)
        rows.append(cells)
    return TableBlock(rows=rows)


def _clean_title(raw: str) -> tuple[str, str]:
    """Return (number, cleaned_title) from a raw slide header string.

    Raw format (per doc):  "0 - Agenda & Purpose - Slide Header"
                           "12 - Appendix: FAQ Cheat Sheet (1min) - Slide Header"
                           "12 - QA Live (5-10 min) Slide Header"   <- variant without dash
    """
    text = raw.strip()
    # Strip trailing " - Slide Header" or " Slide Header"
    text = SLIDE_HEADER_SUFFIX_RE.sub("", text)
    text = SLIDE_HEADER_ALT_RE.sub("", text)
    # Extract leading number
    m = re.match(r"^\s*(\d+)\s*[-–.:]\s*(.*)$", text)
    if m:
        return m.group(1), m.group(2).strip()
    return "", text.strip()


def parse_source(path: Path) -> list[Slide]:
    doc = Document(str(path))
    slides: list[Slide] = []
    current: Optional[Slide] = None
    in_body = False  # True once we've hit the "On Slide:" heading for current slide

    for block in _iter_block_items(doc):
        if isinstance(block, Table):
            if current is not None:
                current.body.append(_extract_table(block))
            continue

        # It's a Paragraph
        para = block
        text = para.text.strip()
        style = para.style.name if para.style else ""

        if not text:
            continue

        # New slide starts at Heading 2
        if style == "Heading 2":
            number, cleaned = _clean_title(text)
            current = Slide(number=number, title=cleaned, key_message="", raw_title=text)
            slides.append(current)
            in_body = False
            continue

        # "On Slide:" marker (Heading 3) — sometimes has trailing note ("leave it empty, ...")
        if style == "Heading 3":
            if current is None:
                continue
            in_body = True
            # If the "On Slide:" line contains a note (e.g., "leave it empty, populated by ..."),
            # capture that note as the first body item.
            remainder = ON_SLIDE_PREFIX_RE.sub("", text).strip()
            if remainder:
                current.body.append(TextBlock(text=remainder))
            continue

        # Regular paragraph
        if current is None:
            # Pre-slide text (e.g. Heading 1 or intro paragraph) — ignore for slide content
            continue

        if not in_body:
            # Between the slide header and "On Slide:" — expect the "Key message:" line
            if KEY_MESSAGE_PREFIX_RE.match(text):
                current.key_message = KEY_MESSAGE_PREFIX_RE.sub("", text).strip()
            else:
                # Fallback: treat as key message if not set yet, else as body
                if not current.key_message:
                    current.key_message = text
                else:
                    current.body.append(TextBlock(text=text))
            continue

        # In body: append as a text block. Preserve line breaks inside a single paragraph.
        current.body.append(TextBlock(text=text))

    return slides


def _debug_print(slides: list[Slide]) -> None:
    for s in slides:
        print(f"\n===== Slide {s.number}: {s.title} =====")
        print(f"  key_message: {s.key_message}")
        print(f"  body ({len(s.body)} blocks):")
        for i, blk in enumerate(s.body):
            if isinstance(blk, TextBlock):
                preview = blk.text if len(blk.text) <= 140 else blk.text[:140] + "..."
                print(f"    [{i}] TEXT: {preview}")
            else:
                print(f"    [{i}] TABLE: {len(blk.rows)} rows x {len(blk.rows[0]) if blk.rows else 0} cols")
                for r in blk.rows[:2]:
                    print(f"          row: {[c[:40] for c in r]}")


if __name__ == "__main__":
    here = Path(__file__).resolve().parents[1]
    src = here / "inputs" / "For+Slide-Generator+skill.docx"
    slides = parse_source(src)
    print(f"Parsed {len(slides)} slides")
    _debug_print(slides)
