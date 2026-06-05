import os
import io
import json
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional

from formatter import apply_formatting
from nb_converter import convert_notebook
from ai_hints import get_suggestion

app = FastAPI(title="WordForge API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Health check ──────────────────────────────────────────────────────────────

@app.get("/api/healthz")
async def healthz():
    return {"status": "ok"}


# ── Format document ──────────────────────────────────────────────────────────

@app.post("/api/format-doc")
async def format_doc(
    settings: str = Form(...),
    file: Optional[UploadFile] = File(None),
):
    try:
        cfg = json.loads(settings)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid settings JSON")

    try:
        file_bytes = await file.read() if file else None
        result = apply_formatting(file_bytes, cfg)
        return StreamingResponse(
            io.BytesIO(result),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=wordforge_formatted.docx"},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ── Convert notebook ──────────────────────────────────────────────────────────

@app.post("/api/convert-notebook")
async def convert_nb(
    settings: str = Form("{}"),
    file: UploadFile = File(...),
):
    try:
        cfg = json.loads(settings)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid settings JSON")

    try:
        file_bytes = await file.read()
        result = convert_notebook(file_bytes, cfg)
        filename = Path(file.filename or "notebook").stem + ".docx"
        return StreamingResponse(
            io.BytesIO(result),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ── AI suggestions ────────────────────────────────────────────────────────────

class AiSuggestInput(BaseModel):
    trigger: str
    context: Optional[str] = None


@app.post("/api/ai-suggest")
async def ai_suggest(body: AiSuggestInput):
    try:
        suggestion = get_suggestion(body.trigger, body.context or "")
        return {"suggestion": suggestion}
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port)
