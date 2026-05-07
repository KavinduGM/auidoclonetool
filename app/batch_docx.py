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

# Leading "1." / "2)" / "a." / bullets — common in Word; strip only for *matching*, not from stored labels.
_ENUM_PREFIX = re.compile(
    r"^\s*(?:(?:\d+|[a-zA-Z])[\.\)]\s+|[•\-\*·▪▸]\s+)+",
    re.UNICODE,
)


def _normalized_display(raw: str) -> str:
    """Normalize odd Word spaces so labels match the real text (minus invisible junk)."""
    if not raw:
        return ""
    s = (
        raw.replace("\u00a0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\u2009", " ")
        .replace("\ufeff", "")
    )
    return s.strip()


def _match_core(s: str) -> str:
    """Text used to detect Animation / Voice headers after list/bullet prefixes."""
    t = _normalized_display(s)
    if not t:
        return ""
    m = _ENUM_PREFIX.match(t)
    if m:
        t = t[m.end() :].strip()
    return t


def _is_animation_header_line(raw: str) -> bool:
    core = _match_core(raw)
    if not core:
        return False
    # Match any "Animation …" header: "Animation 1", "Animation Voice 1", "Animation - Intro", etc.
    return bool(re.match(r"^Animation\b", core, re.IGNORECASE))


def _is_voice_label_line(raw: str) -> bool:
    core = _match_core(raw)
    if not core:
        return False
    if re.match(r"^Animation\b", core, re.IGNORECASE):
        return False
    if _voice_name_value(raw) is not None:
        return False
    # "Voice 1", "Voice 12", optional tight "Voice1"
    return bool(
        re.match(r"^Voice\s+\S", core, re.IGNORECASE)
        or re.match(r"^Voice\d+\b", core, re.IGNORECASE)
    )


_VOICE_NAME_PREFIX_RE = re.compile(r"^Voice\s+Name\s*:\s*(.*)$", re.IGNORECASE)


def _voice_name_value(raw: str) -> Optional[str]:
    """If the line is 'Voice Name: <value>', return the trimmed value (preserving inner colons). Else None."""
    core = _match_core(raw)
    if not core:
        return None
    m = _VOICE_NAME_PREFIX_RE.match(core)
    if not m:
        return None
    val = m.group(1).strip()
    return val or None


@dataclass
class _ParseState:
    current_group: Optional[str] = None
    current_voice: Optional[str] = None
    current_voice_name: Optional[str] = None
    pending_voice_name: Optional[str] = None
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
      entries: list of { order, group, voice_label, text, relative_path, path_parts }
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
        # Format A only: a "Voice Name: …" line preceding this block overrides the file stem.
        # Format B keeps "Voice N" as the filename so the existing layout does not change.
        custom = state.current_voice_name if not group else None
        file_stem = custom if custom else label
        if group:
            rel = f"{group}/{file_stem}.{audio_extension}"
            # Explicit segments so the web client always creates one subfolder per group (Format B).
            path_parts = [group, f"{file_stem}.{audio_extension}"]
        else:
            rel = f"{file_stem}.{audio_extension}"
            path_parts = [f"{file_stem}.{audio_extension}"]
        entries.append(
            {
                "order": len(entries),
                "group": group,
                "voice_label": label,
                "voice_name": custom,
                "file_stem": file_stem,
                "text": text,
                "relative_path": rel,
                "path_parts": path_parts,
            }
        )
        state.current_voice_name = None
        state.lines.clear()

    for raw_line in _iter_lines_from_doc(doc):
        display = _normalized_display(raw_line)
        if not display:
            continue
        if _is_animation_header_line(raw_line):
            flush()
            state.current_group = display
            state.current_voice = None
            state.current_voice_name = None
            state.pending_voice_name = None
            state.saw_animation_header = True
            state.lines.clear()
            continue
        voice_name = _voice_name_value(raw_line)
        if voice_name is not None:
            # Stash for the *next* "Voice N" block. Do not flush — text still belongs to the prior voice.
            state.pending_voice_name = voice_name
            continue
        if _is_voice_label_line(raw_line):
            flush()
            state.current_voice = display
            # Lock in the name that was queued *before* this block; clear pending so it
            # cannot leak into the next voice if no new "Voice Name:" line shows up.
            state.current_voice_name = state.pending_voice_name
            state.pending_voice_name = None
            state.lines.clear()
            continue
        if state.current_voice:
            state.lines.append(display)

    flush()

    if not entries:
        raise ValueError(
            "No voice entries found. Use lines like 'Voice 1' as labels, "
            "optionally under 'Animation …' group headers."
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
