"""
Parse structured voice scripts from .docx files (Format A: flat list, Format B: animation groups).
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterator, List, Optional

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

MAX_ENTRIES = 1000
AUDIO_EXTENSION = "wav"


def _is_animation_header(line: str) -> bool:
    return bool(re.match(r"^Animation\s+Voice\s+", line.strip(), re.IGNORECASE))


def _is_voice_label(line: str) -> bool:
    s = line.strip()
    if _is_animation_header(s):
        return False
    return bool(re.match(r"^Voice\s+", s, re.IGNORECASE))


@dataclass
class _ParseState:
    current_group: Optional[str] = None
    current_voice: Optional[str] = None
    lines: List[str] = None
    saw_animation_header: bool = False

    def __post_init__(self):
        if self.lines is None:
            self.lines = []


def _iter_lines_from_doc(doc: DocumentObject) -> Iterator[str]:
    """Walk body in document order: paragraphs and table cells."""
    body = doc.element.body
    for child in body:
        if child.tag == qn("w:p"):
            p = Paragraph(child, doc)
            for raw in p.text.splitlines():
                yield raw
        elif child.tag == qn("w:tbl"):
            tbl = Table(child, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for raw in p.text.splitlines():
                            yield raw


def parse_voice_docx(file_bytes: bytes, audio_extension: str = AUDIO_EXTENSION) -> dict:
    """
    Returns:
      format: "A" | "B"
      audio_extension: str
      entries: list of { order, group, voice_label, text, relative_path }
    """
    doc = Document(io.BytesIO(file_bytes))
    state = _ParseState()
    entries: List[dict] = []

    def flush() -> None:
        if not state.current_voice:
            state.lines.clear()
            return
        text = "\n".join(state.lines).strip()
        group = state.current_group
        label = state.current_voice
        if group:
            rel = f"{group}/{label}.{audio_extension}"
        else:
            rel = f"{label}.{audio_extension}"
        entries.append(
            {
                "order": len(entries),
                "group": group,
                "voice_label": label,
                "text": text,
                "relative_path": rel,
            }
        )
        state.lines.clear()

    for raw_line in _iter_lines_from_doc(doc):
        s = raw_line.strip()
        if not s:
            continue
        if _is_animation_header(s):
            flush()
            state.current_group = s.strip()
            state.current_voice = None
            state.saw_animation_header = True
            state.lines.clear()
            continue
        if _is_voice_label(s):
            flush()
            state.current_voice = s.strip()
            state.lines.clear()
            continue
        if state.current_voice:
            state.lines.append(s)

    flush()

    if not entries:
        raise ValueError(
            "No voice entries found. Use lines like 'Voice 1' as labels, "
            "optionally under 'Animation Voice …' group headers."
        )
    if len(entries) > MAX_ENTRIES:
        raise ValueError(f"Too many entries ({len(entries)}); maximum is {MAX_ENTRIES}.")

    fmt = "B" if state.saw_animation_header else "A"
    return {
        "format": fmt,
        "audio_extension": audio_extension,
        "entry_count": len(entries),
        "entries": entries,
    }
