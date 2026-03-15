"""
webapp/app.py - FastAPI Web Server
Port: 5050

Endpoints:
  GET  /                     - Main UI
  POST /api/convert          - Upload PDF(s), returns job_id
  GET  /api/status/{job_id}  - Poll conversion status
  GET  /api/files            - List all converted files
  GET  /api/download/{name}  - Download an Excel file
  DELETE /api/files/{name}   - Delete a converted file
  GET  /api/health           - Health check
"""

from __future__ import annotations

import asyncio
import logging
import os
import sys
import time
import traceback
import uuid
from pathlib import Path
from typing import Dict, List, Optional

# ── make invoice_parser importable ────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import aiofiles

from invoice_parser import convert, Config
from invoice_parser.config import DEFAULT_CONFIG

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("webapp")

# ── Directories ───────────────────────────────────────────────────────────────
UPLOAD_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── In-memory job store ───────────────────────────────────────────────────────
# job_id → { status, filename, output, error, progress, started_at, ended_at }
JOBS: Dict[str, dict] = {}

# ── FastAPI app ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="PDF to Excel Converter",
    description="Layout-preserving PDF → Excel conversion engine",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve static assets (CSS / JS)
STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


# ── Helper ────────────────────────────────────────────────────────────────────
def _file_info(path: Path) -> dict:
    stat = path.stat()
    return {
        "name":     path.name,
        "size_kb":  round(stat.st_size / 1024, 1),
        "modified": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat.st_mtime)),
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def index():
    """Serve the main HTML interface."""
    html_path = Path(__file__).parent / "templates" / "index.html"
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))


@app.get("/api/health")
async def health():
    return {
        "status": "ok",
        "version": "1.0.0",
        "jobs_active": sum(1 for j in JOBS.values() if j["status"] == "running"),
    }


@app.post("/api/convert")
async def api_convert(files: List[UploadFile] = File(...)):
    """
    Accept one or more PDF uploads.
    Returns a list of job IDs for polling.
    """
    if not files:
        raise HTTPException(400, "No files uploaded")

    job_ids = []
    for upload in files:
        if not upload.filename.lower().endswith(".pdf"):
            raise HTTPException(400, f"'{upload.filename}' is not a PDF")

        # Save upload to disk
        safe_name = Path(upload.filename).name
        pdf_path = UPLOAD_DIR / safe_name
        async with aiofiles.open(str(pdf_path), "wb") as f:
            content = await upload.read()
            await f.write(content)

        # Create job
        job_id = str(uuid.uuid4())[:8]
        JOBS[job_id] = {
            "status":     "queued",
            "filename":   safe_name,
            "pdf_path":   str(pdf_path),
            "output":     None,
            "error":      None,
            "progress":   0,
            "started_at": None,
            "ended_at":   None,
        }
        job_ids.append(job_id)

        # Kick off background conversion
        asyncio.create_task(_run_conversion(job_id))
        logger.info("Job %s queued for %s", job_id, safe_name)

    return {"job_ids": job_ids}


@app.get("/api/status/{job_id}")
async def api_status(job_id: str):
    """Poll the status of a conversion job."""
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, f"Job '{job_id}' not found")
    return job


@app.get("/api/jobs")
async def api_all_jobs():
    """Return all jobs (most recent first)."""
    sorted_jobs = sorted(
        [{"job_id": k, **v} for k, v in JOBS.items()],
        key=lambda j: j.get("started_at") or "",
        reverse=True,
    )
    return {"jobs": sorted_jobs}


@app.get("/api/files")
async def api_files():
    """List all Excel files in the output directory."""
    files = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return {"files": [_file_info(f) for f in files]}


@app.get("/api/download/{filename}")
async def api_download(filename: str):
    """Download an Excel output file."""
    path = OUTPUT_DIR / filename
    if not path.exists() or not path.name.endswith(".xlsx"):
        raise HTTPException(404, f"File '{filename}' not found")
    return FileResponse(
        path=str(path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )


@app.delete("/api/files/{filename}")
async def api_delete(filename: str):
    """Delete an output Excel file."""
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(404, f"File '{filename}' not found")
    path.unlink()
    # Also remove source PDF if it exists
    pdf = UPLOAD_DIR / filename.replace(".xlsx", ".pdf")
    if pdf.exists():
        pdf.unlink()
    return {"deleted": filename}


# ── Background conversion task ────────────────────────────────────────────────

async def _run_conversion(job_id: str) -> None:
    """Run the PDF→Excel pipeline in a thread pool (non-blocking)."""
    job = JOBS[job_id]
    job["status"]     = "running"
    job["started_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    job["progress"]   = 10

    pdf_path  = Path(job["pdf_path"])
    xlsx_name = pdf_path.stem + ".xlsx"
    xlsx_path = OUTPUT_DIR / xlsx_name

    try:
        loop = asyncio.get_event_loop()
        # Run blocking pipeline in thread pool so we don't block the event loop
        await loop.run_in_executor(
            None,
            _sync_convert,
            str(pdf_path), str(xlsx_path), job_id,
        )
        job["status"]   = "done"
        job["output"]   = xlsx_name
        job["progress"] = 100
        logger.info("Job %s completed: %s", job_id, xlsx_name)
    except Exception as exc:
        job["status"] = "error"
        job["error"]  = str(exc)
        job["progress"] = 0
        logger.error("Job %s failed: %s", job_id, exc)
        logger.debug(traceback.format_exc())
    finally:
        job["ended_at"] = time.strftime("%Y-%m-%d %H:%M:%S")


def _sync_convert(pdf_path: str, xlsx_path: str, job_id: str) -> None:
    """Synchronous wrapper called from thread pool."""
    JOBS[job_id]["progress"] = 30
    convert(pdf_path, xlsx_path, config=DEFAULT_CONFIG)
    JOBS[job_id]["progress"] = 90


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app:app",
        host="127.0.0.1",
        port=5050,
        reload=False,
        log_level="info",
    )
