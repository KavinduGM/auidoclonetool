"""
FastAPI server: web UI + JSON API for voice cloning and generation.
"""
import io
import os
import re
import tempfile
import zipfile
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Depends
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from .tts_engine import VoiceEngine
from . import voice_store
from . import auth
from . import batch_docx

app = FastAPI(title="Voice Clone Tool", version="1.1.0")

# Split pasted “many questions” text into separate TTS jobs (own line, exact match).
QUESTION_SPLIT_MARKER = "<<<QUESTION>>>"
MAX_BATCH_QUESTIONS = 500

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

engine: Optional[VoiceEngine] = None


@app.on_event("startup")
def _startup():
    global engine
    engine = VoiceEngine()


# ---------- static web UI ----------
WEB_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "web")


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(WEB_DIR, "index.html"), "r") as f:
        return f.read()


@app.get("/style.css")
def _css():
    return Response(
        content=open(os.path.join(WEB_DIR, "style.css")).read(),
        media_type="text/css",
    )


@app.get("/app.js")
def _js():
    return Response(
        content=open(os.path.join(WEB_DIR, "app.js")).read(),
        media_type="application/javascript",
    )


@app.get("/batch_generate.js")
def _batch_js():
    return Response(
        content=open(os.path.join(WEB_DIR, "batch_generate.js")).read(),
        media_type="application/javascript",
    )


app.mount("/static", StaticFiles(directory=WEB_DIR), name="static")


# ---------- auth status (public) ----------
@app.get("/api/auth/status")
def auth_status():
    """Tells the UI whether auth is required yet (bootstrap vs normal mode)."""
    return {"auth_required": auth.any_keys_exist()}


# ---------- API keys management (protected) ----------
@app.get("/api/keys", dependencies=[Depends(auth.require_auth)])
def api_list_keys():
    return {"keys": auth.list_keys_safe()}


@app.post("/api/keys", dependencies=[Depends(auth.require_auth)])
def api_create_key(name: str = Form(...)):
    try:
        entry = auth.create_key(name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    # return the full key once on creation
    return {
        "id": entry["id"],
        "name": entry["name"],
        "key": entry["key"],
        "created_at": entry["created_at"],
    }


@app.delete("/api/keys/{key_id}", dependencies=[Depends(auth.require_auth)])
def api_delete_key(key_id: str):
    if not auth.delete_key(key_id):
        raise HTTPException(404, "key not found")
    return {"deleted": key_id}


# ---------- Voices (protected) ----------
@app.get("/api/voices", dependencies=[Depends(auth.require_auth)])
def api_list_voices():
    return {"voices": voice_store.list_voices()}


@app.post("/api/voices", dependencies=[Depends(auth.require_auth)])
async def api_create_voice(
    name: str = Form(...),
    description: str = Form(""),
    file: UploadFile = File(...),
):
    if not name.strip():
        raise HTTPException(400, "name required")

    suffix = os.path.splitext(file.filename or "upload.wav")[1] or ".wav"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        meta = voice_store.create_voice(name.strip(), description.strip(), tmp_path)
    except Exception as e:
        raise HTTPException(500, f"Failed to process audio: {e}")
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
    return meta


@app.delete("/api/voices/{voice_id}", dependencies=[Depends(auth.require_auth)])
def api_delete_voice(voice_id: str):
    if not voice_store.delete_voice(voice_id):
        raise HTTPException(404, "voice not found")
    return {"deleted": voice_id}


# ---------- Generate (protected) ----------
@app.post("/api/generate", dependencies=[Depends(auth.require_auth)])
async def api_generate(
    voice_id: str = Form(...),
    text: str = Form(...),
    speed: float = Form(1.0),
):
    if engine is None:
        raise HTTPException(503, "engine not ready")
    ref = voice_store.get_reference_wav(voice_id)
    if not ref:
        raise HTTPException(404, "voice_id not found")
    if not text.strip():
        raise HTTPException(400, "text required")
    try:
        wav, sr = engine.generate(text=text, speaker_wav=ref, speed=speed)
    except Exception as e:
        raise HTTPException(500, f"generation failed: {e}")
    data = VoiceEngine.to_wav_bytes(wav, sr)
    return Response(
        content=data,
        media_type="audio/wav",
        headers={"Content-Disposition": f'attachment; filename="{voice_id}.wav"'},
    )


def _split_questions_blob(raw: str) -> List[str]:
    """Split user text on QUESTION_SPLIT_MARKER lines; trim and drop empties."""
    parts = re.split(
        rf"(?:^\s*{re.escape(QUESTION_SPLIT_MARKER)}\s*$)",
        raw,
        flags=re.MULTILINE,
    )
    out = [p.strip() for p in parts if p.strip()]
    return out


@app.post("/api/generate-batch", dependencies=[Depends(auth.require_auth)])
async def api_generate_batch(
    voice_id: str = Form(...),
    text: str = Form(...),
    speed: float = Form(1.0),
):
    """
    One WAV per question, returned inside a single ZIP.
    Separate questions in `text` with a line containing only: <<<QUESTION>>>
    """
    if engine is None:
        raise HTTPException(503, "engine not ready")
    ref = voice_store.get_reference_wav(voice_id)
    if not ref:
        raise HTTPException(404, "voice_id not found")
    raw = (text or "").strip()
    if not raw:
        raise HTTPException(400, "text required")

    questions = _split_questions_blob(raw)
    if not questions:
        raise HTTPException(400, "no question text after splitting")
    if len(questions) > MAX_BATCH_QUESTIONS:
        raise HTTPException(
            400,
            f"Too many questions ({len(questions)}). Maximum is {MAX_BATCH_QUESTIONS}.",
        )

    buf = io.BytesIO()
    try:
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, q in enumerate(questions):
                try:
                    wav, sr = engine.generate(text=q, speaker_wav=ref, speed=speed)
                except Exception as e:
                    raise HTTPException(
                        500,
                        f"generation failed on question {i + 1} of {len(questions)}: {e}",
                    ) from e
                name = f"question_{i + 1:03d}.wav"
                zf.writestr(name, VoiceEngine.to_wav_bytes(wav, sr))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"batch zip failed: {e}") from e

    data = buf.getvalue()
    return Response(
        content=data,
        media_type="application/zip",
        headers={
            "Content-Disposition": 'attachment; filename="questions.zip"',
        },
    )


@app.post("/api/batch/parse-docx", dependencies=[Depends(auth.require_auth)])
async def api_batch_parse_docx(file: UploadFile = File(...)):
    """
    Parse a .docx voice script (Format A: Voice N list; Format B: Animation Voice groups).
    Returns JSON entries for client-side batch generation against /api/generate.
    """
    fname = (file.filename or "").lower()
    if not fname.endswith(".docx"):
        raise HTTPException(400, "Upload a .docx file.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "Empty file.")
    try:
        return batch_docx.parse_voice_docx(data)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"Could not read document: {e}")


# ---------- Health (public) ----------
@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "device": engine.device if engine else "loading",
        "voices": len(voice_store.list_voices()),
        "auth_required": auth.any_keys_exist(),
    }
