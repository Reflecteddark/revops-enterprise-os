import os
import tempfile
import urllib.request
import logging
from typing import Optional
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from faster_whisper import WhisperModel

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s: %(message)s")
logger = logging.getLogger("transcriber")

app = FastAPI(title="Voice Transcription Microservice", version="1.0.0")

# Initialize Whisper model lazily or on startup
# Using "base" or "small" model for high accuracy and fast transcription in Russian
MODEL_SIZE = os.environ.get("WHISPER_MODEL", "small")
DEVICE = os.environ.get("WHISPER_DEVICE", "auto")

_model: Optional[WhisperModel] = None

def get_model() -> WhisperModel:
    global _model
    if _model is None:
        logger.info(f"Loading Faster-Whisper model '{MODEL_SIZE}' on device '{DEVICE}'...")
        try:
            _model = WhisperModel(MODEL_SIZE, device="cuda", compute_type="float16")
            logger.info("Successfully loaded Faster-Whisper on CUDA GPU.")
        except Exception as e:
            logger.warning(f"CUDA initialization failed ({e}), falling back to CPU...")
            _model = WhisperModel(MODEL_SIZE, device="cpu", compute_type="int8")
            logger.info("Successfully loaded Faster-Whisper on CPU.")
    return _model

class TranscribeRequest(BaseModel):
    url: Optional[str] = None
    audio_url: Optional[str] = None

@app.get("/")
@app.get("/health")
def health():
    return {"status": "ok", "service": "Voice Transcriber", "port": 8000}

@app.post("/transcribe-by-url")
def transcribe_by_url(req: TranscribeRequest):
    target_url = req.url or req.audio_url
    if not target_url:
        raise HTTPException(status_code=400, detail="Missing 'url' or 'audio_url' field in request body")
    
    logger.info(f"Received transcription request for URL: {target_url[:80]}...")
    
    # Download audio to temporary file
    suffix = ".mp3"
    if ".wav" in target_url.lower():
        suffix = ".wav"
    elif ".ogg" in target_url.lower():
        suffix = ".ogg"
        
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp_file:
        tmp_path = tmp_file.name

    try:
        req_headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        request = urllib.request.Request(target_url, headers=req_headers)
        with urllib.request.urlopen(request, timeout=60) as response, open(tmp_path, "wb") as out_file:
            out_file.write(response.read())
        
        file_size_mb = os.path.getsize(tmp_path) / 1e6
        logger.info(f"Downloaded audio file: {file_size_mb:.2f} MB. Starting transcription...")

        model = get_model()
        segments, info = model.transcribe(tmp_path, beam_size=5, language="ru")
        
        text_segments = []
        for segment in segments:
            text_segments.append(segment.text.strip())
            
        full_text = " ".join(text_segments).strip()
        logger.info(f"Transcription complete: {len(full_text.split())} words, language: {info.language}")
        
        return {
            "text": full_text,
            "language": info.language,
            "language_probability": round(info.language_probability, 3),
            "duration": round(info.duration, 2)
        }

    except Exception as e:
        logger.error(f"Error during transcription: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
