"""
API questionnaire de satisfaction : PDFs → OCR → PPTX + CSV + JSON.
Gère : N PDFs de 1 page, ou 1 PDF multi-pages, ou un mix.

Deux entrées :
- POST /process     : multipart-form-data (legacy, fichiers binaires)
- POST /process_json: application/json (files: [{name, base64}], title)
"""

import base64
import binascii
import csv
import io
import os
import uuid
from typing import Optional

from fastapi import (
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
    status,
)
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from slowapi import Limiter
from slowapi.errors import RateLimitExceeded
from slowapi.util import get_remote_address
from starlette.middleware.base import BaseHTTPMiddleware

from logging_config import configure_logging, get_logger, request_id_ctx
from ocr import extract_from_pdf
from pptx_generator import generate_pptx
from security import require_api_key

ALLOWED_CONTENT_TYPES = {"application/pdf"}
ALLOWED_EXTENSIONS = {".pdf"}

MAX_FILE_SIZE_MB = int(os.environ.get("MAX_FILE_SIZE_MB", "20"))
MAX_FILES_PER_REQUEST = int(os.environ.get("MAX_FILES_PER_REQUEST", "100"))
RATE_LIMIT_PER_MINUTE = int(os.environ.get("RATE_LIMIT_PER_MINUTE", "10"))

configure_logging()
logger = get_logger()

limiter = Limiter(key_func=get_remote_address, default_limits=[])

app = FastAPI(title="Nefia Questionnaire Processor", version="2.2.0")
app.state.limiter = limiter


class RequestIdMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        rid = request.headers.get("X-Request-ID") or uuid.uuid4().hex[:16]
        token = request_id_ctx.set(rid)
        request.state.request_id = rid
        try:
            response = await call_next(request)
        finally:
            request_id_ctx.reset(token)
        response.headers["X-Request-ID"] = rid
        return response


app.add_middleware(RequestIdMiddleware)


@app.exception_handler(RateLimitExceeded)
async def rate_limit_handler(request: Request, exc: RateLimitExceeded):
    rid = getattr(request.state, "request_id", "-")
    logger.warning("rate_limit_exceeded", path=request.url.path)
    return JSONResponse(
        status_code=status.HTTP_429_TOO_MANY_REQUESTS,
        content={"error": "Rate limit exceeded", "request_id": rid},
        headers={"X-Request-ID": rid},
    )


@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    rid = getattr(request.state, "request_id", "-")
    return JSONResponse(
        status_code=exc.status_code,
        content={"error": exc.detail, "request_id": rid},
        headers={"X-Request-ID": rid},
    )


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    rid = getattr(request.state, "request_id", "-")
    logger.error(
        "unhandled_exception",
        exc_info=exc,
        path=request.url.path,
        method=request.method,
    )
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={
            "error": "Internal server error",
            "request_id": rid,
        },
        headers={"X-Request-ID": rid},
    )


@app.get("/health")
async def health():
    """Liveness/readiness probe. Public (no auth)."""
    deps = {
        "mistral_api_key": bool(os.environ.get("MISTRAL_API_KEY")),
        "api_key_configured": bool(os.environ.get("API_KEY")),
    }
    healthy = all(deps.values())
    return JSONResponse(
        status_code=200 if healthy else 503,
        content={"status": "ok" if healthy else "degraded", "deps": deps},
    )


def _validate_extension(filename: str) -> None:
    if not filename:
        return
    ext = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported extension for {filename!r}",
        )


def _validate_upload_file(file: UploadFile) -> None:
    if file.content_type not in ALLOWED_CONTENT_TYPES:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported content-type for {file.filename!r}: {file.content_type}",
        )
    _validate_extension(file.filename or "")


def _process_files(
    pdfs: list[tuple[str, bytes]],
    title: str,
    request_id: str,
) -> dict:
    """Pipeline OCR → PPTX + CSV. Reçoit une liste (filename, bytes)."""
    if len(pdfs) > MAX_FILES_PER_REQUEST:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Too many files: {len(pdfs)} > {MAX_FILES_PER_REQUEST}",
        )

    max_bytes = MAX_FILE_SIZE_MB * 1024 * 1024
    all_questionnaires: list[dict] = []
    errors: list[dict] = []

    logger.info("process_started", file_count=len(pdfs), title=title)

    for filename, pdf_bytes in pdfs:
        try:
            _validate_extension(filename)
            if len(pdf_bytes) > max_bytes:
                raise HTTPException(
                    status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                    detail=f"File {filename!r} too large ({len(pdf_bytes)} bytes > {max_bytes})",
                )
            questionnaires = extract_from_pdf(pdf_bytes)
            all_questionnaires.extend(questionnaires)
            logger.info(
                "file_processed",
                filename=filename,
                size_bytes=len(pdf_bytes),
                questionnaires_extracted=len(questionnaires),
            )
        except HTTPException:
            raise
        except Exception as e:
            logger.exception("file_processing_failed", filename=filename)
            errors.append({"file": filename, "error": str(e)})

    if not all_questionnaires:
        logger.warning("no_questionnaires_extracted", errors=errors)
        return {
            "_status_code": 400,
            "error": "Aucun questionnaire extrait",
            "details": errors,
            "request_id": request_id,
        }

    pptx_bytes = generate_pptx(all_questionnaires, title)

    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Participant", "Question", "Reponse", "Commentaire"])
    for q in all_questionnaires:
        participant = q.get("metadata", {}).get("participant", "?")
        for item in q.get("items", []):
            writer.writerow([
                participant,
                item.get("label", ""),
                item.get("response", ""),
                item.get("comment", "") or "",
            ])
    csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

    logger.info(
        "process_completed",
        questionnaires_processed=len(all_questionnaires),
        errors=len(errors),
    )

    return {
        "success": True,
        "questionnaires_processed": len(all_questionnaires),
        "errors": errors,
        "pptx_base64": base64.b64encode(pptx_bytes).decode(),
        "csv_base64": base64.b64encode(csv_bytes).decode(),
        "data": all_questionnaires,
        "request_id": request_id,
    }


@app.post("/process", dependencies=[Depends(require_api_key)])
@limiter.limit(f"{RATE_LIMIT_PER_MINUTE}/minute")
async def process_questionnaires(
    request: Request,
    files: list[UploadFile] = File(..., description="PDF questionnaires scannes"),
    title: Optional[str] = Form("Resultats de satisfaction"),
):
    """
    Multipart-form-data : N PDFs en binaires + title.
    """
    pdfs: list[tuple[str, bytes]] = []
    for f in files:
        _validate_upload_file(f)
        pdfs.append((f.filename or "questionnaire.pdf", await f.read()))

    result = _process_files(pdfs, title or "Resultats de satisfaction", request.state.request_id)
    if result.get("_status_code"):
        sc = result.pop("_status_code")
        return JSONResponse(status_code=sc, content=result)
    return result


class _JsonFile(BaseModel):
    name: str = Field(..., description="Nom du fichier (avec extension .pdf)")
    base64: str = Field(..., description="Contenu binaire encodé base64")


class _JsonProcessRequest(BaseModel):
    title: str = Field(default="Resultats de satisfaction")
    files: list[_JsonFile] = Field(..., min_length=1)


@app.post("/process_json", dependencies=[Depends(require_api_key)])
@limiter.limit(f"{RATE_LIMIT_PER_MINUTE}/minute")
async def process_questionnaires_json(
    request: Request,
    payload: _JsonProcessRequest,
):
    """
    Application/json : { title, files: [{name, base64}] }.
    Mêmes contraintes (taille, extension, rate limit) que /process.
    """
    pdfs: list[tuple[str, bytes]] = []
    for jf in payload.files:
        try:
            pdf_bytes = base64.b64decode(jf.base64, validate=True)
        except (binascii.Error, ValueError) as e:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid base64 for file {jf.name!r}: {e}",
            )
        pdfs.append((jf.name, pdf_bytes))

    result = _process_files(pdfs, payload.title, request.state.request_id)
    if result.get("_status_code"):
        sc = result.pop("_status_code")
        return JSONResponse(status_code=sc, content=result)
    return result
