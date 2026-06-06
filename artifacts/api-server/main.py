import os
import io
import json
import zipfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional, List

from formatter import apply_formatting
from nb_converter import convert_notebook
from ai_hints import get_suggestion
from cover_page import add_cover_page
from sop_formatter import generate_sop_document
from docx import Document

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
        result_bytes = apply_formatting(file_bytes, cfg)

        # Optionally prepend a cover page
        cover_cfg = cfg.get("coverPage")
        if cover_cfg and cover_cfg.get("enabled"):
            doc = Document(io.BytesIO(result_bytes))
            add_cover_page(doc, cover_cfg)
            buf = io.BytesIO()
            doc.save(buf)
            result_bytes = buf.getvalue()

        filename = cfg.get("filename", "submission.docx")
        if not filename.endswith(".docx"):
            filename += ".docx"

        return StreamingResponse(
            io.BytesIO(result_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ── Convert notebook(s) ───────────────────────────────────────────────────────

@app.post("/api/convert-notebook")
async def convert_nb(
    settings: str = Form("{}"),
    files: List[UploadFile] = File(...),
):
    try:
        cfg = json.loads(settings)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid settings JSON")

    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    if len(files) > 5:
        raise HTTPException(status_code=400, detail="Maximum 5 notebooks per batch")

    try:
        if len(files) == 1:
            # Single file — return docx directly
            f = files[0]
            file_bytes = await f.read()
            result = convert_notebook(file_bytes, cfg)
            stem = Path(f.filename or "notebook").stem
            filename = cfg.get("filename") or f"{stem}.docx"
            if not filename.endswith(".docx"):
                filename += ".docx"
            return StreamingResponse(
                io.BytesIO(result),
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": f'attachment; filename="{filename}"'},
            )
        else:
            # Batch — return zip
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in files:
                    file_bytes = await f.read()
                    result = convert_notebook(file_bytes, cfg)
                    stem = Path(f.filename or "notebook").stem
                    zf.writestr(f"{stem}.docx", result)
            zip_buf.seek(0)
            return StreamingResponse(
                zip_buf,
                media_type="application/zip",
                headers={"Content-Disposition": 'attachment; filename="notebooks.zip"'},
            )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ── Format SOP (Gujarat Govt.) ────────────────────────────────────────────────

@app.post("/api/format-sop")
async def format_sop(
    chapter_number: str = Form("1"),
    chapter_title: str = Form(""),
    content: str = Form(""),
    reference_docx: Optional[UploadFile] = File(None),
):
    try:
        ref_bytes = await reference_docx.read() if reference_docx else None
        docx_bytes, report = generate_sop_document(
            content=content,
            chapter_number=chapter_number,
            chapter_title=chapter_title,
            reference_docx_bytes=ref_bytes,
        )
        import json as _json
        filename = f"SOP_Chapter_{chapter_number}.docx"
        return StreamingResponse(
            io.BytesIO(docx_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "X-SOP-Validation": _json.dumps(report),
                "Access-Control-Expose-Headers": "X-SOP-Validation",
            },
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
