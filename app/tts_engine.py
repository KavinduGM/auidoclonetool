"""
XTTS v2 voice cloning + generation engine.
Loads the model once at startup, reuses for all requests.
"""
import os
import io
import re
import uuid
import torch
import numpy as np
import soundfile as sf
from typing import List
from TTS.api import TTS

# Accept Coqui license automatically (non-interactive server use)
os.environ.setdefault("COQUI_TOS_AGREED", "1")


# Marks "12. " list-style periods so sentence splitting does not break after the dot.
_LIST_PERIOD_MARK = "\uE000"


class VoiceEngine:
    def __init__(self):
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        print(f"[tts_engine] Loading XTTS v2 on device: {self.device}")
        self.tts = TTS("tts_models/multilingual/multi-dataset/xtts_v2").to(self.device)
        self.language = "en"
        print("[tts_engine] Model loaded.")

    @staticmethod
    def _protect_list_periods(text: str) -> str:
        """
        Hide periods after digits when followed by whitespace so they are not treated
        as sentence ends (e.g. '2. __' must stay one phrase, not ['2.', '__']).
        Decimals like '3.14' are unchanged (no whitespace after the first dot).
        """
        return re.sub(
            r"(\d+)\.(\s)",
            lambda m: f"{m.group(1)}{_LIST_PERIOD_MARK}{m.group(2)}",
            text,
        )

    @staticmethod
    def _restore_list_periods(text: str) -> str:
        return text.replace(_LIST_PERIOD_MARK, ".")

    # ---------- text chunking ----------
    @staticmethod
    def _split_text(text: str, max_chars: int = 240) -> List[str]:
        """
        Split text into chunks that respect sentence boundaries.
        XTTS handles ~250 chars per inference cleanly; longer produces artifacts.
        """
        text = text.strip().replace("\r", "")
        if not text:
            return []

        text = VoiceEngine._protect_list_periods(text)
        # split into sentences (list markers like "2. " no longer end with '.' for this step)
        sentences = re.split(r'(?<=[.!?])\s+', text)
        chunks: List[str] = []
        buf = ""
        for s in sentences:
            s = VoiceEngine._restore_list_periods(s).strip()
            if not s:
                continue
            # if sentence itself too long, hard-split on commas / spaces
            if len(s) > max_chars:
                parts = re.split(r'(?<=[,;:])\s+', s)
                for p in parts:
                    if len(p) > max_chars:
                        # Flush buf before word-splitting: otherwise chunks from this
                        # part get appended while buf still holds an earlier clause.
                        if buf:
                            chunks.append(buf.strip())
                            buf = ""
                        # last resort: break on spaces
                        words = p.split()
                        cur = ""
                        for w in words:
                            if len(cur) + len(w) + 1 > max_chars:
                                if cur:
                                    chunks.append(cur.strip())
                                cur = w
                            else:
                                cur = (cur + " " + w).strip()
                        if cur:
                            if buf and len(buf) + len(cur) + 1 <= max_chars:
                                buf = (buf + " " + cur).strip()
                            else:
                                if buf:
                                    chunks.append(buf.strip())
                                buf = cur
                    else:
                        if buf and len(buf) + len(p) + 1 <= max_chars:
                            buf = (buf + " " + p).strip()
                        else:
                            if buf:
                                chunks.append(buf.strip())
                            buf = p
                continue

            if buf and len(buf) + len(s) + 1 <= max_chars:
                buf = (buf + " " + s).strip()
            else:
                if buf:
                    chunks.append(buf.strip())
                buf = s
        if buf:
            chunks.append(buf.strip())
        return [VoiceEngine._restore_list_periods(c) for c in chunks]

    # ---------- generation ----------
    def generate(
        self,
        text: str,
        speaker_wav: str,
        speed: float = 1.0,
    ) -> tuple[np.ndarray, int]:
        """
        Generate speech audio from text using the given reference wav.
        Returns (float32 waveform mono, sample_rate).
        """
        speed = float(max(0.5, min(2.0, speed)))
        chunks = self._split_text(text)
        if not chunks:
            raise ValueError("Empty text")

        sample_rate = 24000  # XTTS v2 native
        audio_parts: List[np.ndarray] = []
        silence = np.zeros(int(0.25 * sample_rate), dtype=np.float32)  # 250ms pause

        for i, chunk in enumerate(chunks):
            print(f"[tts_engine] chunk {i+1}/{len(chunks)} ({len(chunk)} chars)")
            wav = self.tts.tts(
                text=chunk,
                speaker_wav=speaker_wav,
                language=self.language,
                speed=speed,
            )
            wav = np.array(wav, dtype=np.float32)
            audio_parts.append(wav)
            if i < len(chunks) - 1:
                audio_parts.append(silence)

        full = np.concatenate(audio_parts) if audio_parts else np.zeros(0, dtype=np.float32)
        # normalize lightly to prevent clipping
        peak = np.max(np.abs(full)) if full.size else 1.0
        if peak > 0.99:
            full = full * (0.99 / peak)
        return full, sample_rate

    @staticmethod
    def to_wav_bytes(wav: np.ndarray, sample_rate: int) -> bytes:
        buf = io.BytesIO()
        sf.write(buf, wav, sample_rate, format="WAV", subtype="PCM_16")
        return buf.getvalue()
