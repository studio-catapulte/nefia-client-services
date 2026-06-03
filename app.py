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
import tempfile
import uuid
from pathlib import Path
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

from extract_pic import (
    Q_WORDINGS,
    aggregate_participants,
    extract_questionnaires_pic,
)
from inject_slides import inject_pic_analysis
from logging_config import configure_logging, get_logger, request_id_ctx
from ocr import extract_from_pdf
from pptx_generator import generate_pptx
from security import require_api_key

ALLOWED_CONTENT_TYPES = {"application/pdf"}
ALLOWED_EXTENSIONS = {".pdf"}
ALLOWED_PPTX_CONTENT_TYPES = {
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/octet-stream",  # certains uploaders (Tally) envoient octet-stream
}
ALLOWED_PPTX_EXTENSIONS = {".pptx"}

TEMPLATES_DIR = Path(__file__).parent / "templates"
ANALYSIS_TEMPLATE = TEMPLATES_DIR / "analysis-block.pptx"
CLOSING_TEMPLATE = TEMPLATES_DIR / "commercial-closing.pptx"

MAX_FILE_SIZE_MB = int(os.environ.get("MAX_FILE_SIZE_MB", "20"))
MAX_FILES_PER_REQUEST = int(os.environ.get("MAX_FILES_PER_REQUEST", "100"))
RATE_LIMIT_PER_MINUTE = int(os.environ.get("RATE_LIMIT_PER_MINUTE", "10"))

configure_logging()
logger = get_logger()

limiter = Limiter(key_func=get_remote_address, default_limits=[])

app = FastAPI(title="Nefia Questionnaire Processor", version="2.4.9")
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


# =====================================================================
# v3 — endpoint /inject_slides (Pic-Formation injection-driven)
# =====================================================================
#
# Le formateur upload son propre PPTX déjà rempli + les scans PDF des
# questionnaires. On extrait l'OCR (Pic-specific : 10 questions Likert
# + remarques + souhaits), on injecte 15 slides d'analyse avant la slide
# « Analyse des résultats » + 6 slides closing commercial.
#
# Réponse : pptx_base64 + csv_base64 + raw_data (consommé par n8n pour
# pousser dans la table NocoDB questionnaire_responses).


def _validate_pptx_upload(file: UploadFile) -> None:
    if file.content_type not in ALLOWED_PPTX_CONTENT_TYPES:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported content-type for {file.filename!r}: {file.content_type}",
        )
    fname = file.filename or ""
    ext = "." + fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
    if ext not in ALLOWED_PPTX_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=f"Unsupported extension for {fname!r}",
        )


def _build_csv_bytes(participants: list[dict]) -> bytes:
    """CSV : 1 ligne par (participant × question) — réutilisable analytics."""
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["Participant", "Question_ID", "Question", "Reponse", "Commentaire"])
    for p in participants:
        prenom = p.get("prenom", "")
        reponses = p.get("reponses", {}) or {}
        commentaires = p.get("commentaires_par_question", {}) or {}
        for q_idx in range(1, 11):
            q_key = f"Q{q_idx}"
            writer.writerow([
                prenom,
                q_idx,
                Q_WORDINGS[q_idx - 1],
                reponses.get(q_key, "Non renseigné"),
                commentaires.get(q_key, "") or "",
            ])
    return buf.getvalue().encode("utf-8-sig")


def _build_raw_data(aggregates: dict) -> dict:
    """Blob JSON complet stocké dans `executions.raw_data_json` (1 row/exécution).

    Contient counts par question/modalité, commentaires détaillés par question,
    remarques + souhaits libres avec prénoms. Tout ce qu'il faut pour reconstruire
    l'analytique a posteriori (drop du long-format intermédiaire).
    """
    questions_out = []
    for i, q in enumerate(aggregates.get("questions", [])[:10], start=1):
        questions_out.append({
            "question_id": i,
            "question_label": Q_WORDINGS[i - 1],
            "counts": q.get("counts", {}),
            "commentaires": q.get("commentaires", []),  # [{prenom, commentaire}]
        })
    return {
        "nb_participants": aggregates.get("nb_participants", 0),
        "questions": questions_out,
        "remarques_libres": aggregates.get("remarques_libres", []),
        "souhaits_libres": aggregates.get("souhaits_libres", []),
    }


def _run_inject_pipeline(
    pptx_bytes: bytes,
    pdf_files: list[tuple[str, bytes]],
    session_id: str,
    client: str,
    titre: str,
    request_id: str,
) -> dict:
    """Pipeline OCR → aggregate → inject. Partagé entre /inject_slides et /inject_slides_json."""
    if not ANALYSIS_TEMPLATE.exists() or not CLOSING_TEMPLATE.exists():
        logger.error(
            "templates_missing",
            analysis_exists=ANALYSIS_TEMPLATE.exists(),
            closing_exists=CLOSING_TEMPLATE.exists(),
        )
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Server misconfiguration: templates missing",
        )

    logger.info(
        "inject_slides_started",
        session_id=session_id,
        client=client,
        titre=titre,
        n_pdfs=len(pdf_files),
        pptx_size=len(pptx_bytes),
    )

    with tempfile.TemporaryDirectory(prefix="inject_") as tmpdir:
        tmpdir_path = Path(tmpdir)
        pptx_input_path = tmpdir_path / "formateur.pptx"
        pptx_input_path.write_bytes(pptx_bytes)

        all_participants: list[dict] = []
        ocr_errors: list[dict] = []
        for fname, pdf_bytes in pdf_files:
            pdf_path = tmpdir_path / fname
            try:
                pdf_path.write_bytes(pdf_bytes)
                extraction = extract_questionnaires_pic(str(pdf_path))
                participants = extraction.get("participants", [])
                all_participants.extend(participants)
                logger.info(
                    "ocr_done",
                    filename=fname,
                    participants_extracted=len(participants),
                )
            except Exception as e:
                logger.exception("ocr_failed", filename=fname)
                ocr_errors.append({"file": fname, "error": str(e)})

        if not all_participants:
            logger.warning("no_questionnaires_extracted", errors=ocr_errors)
            return {
                "_status_code": status.HTTP_400_BAD_REQUEST,
                "error": "Aucun questionnaire extrait",
                "details": ocr_errors,
                "request_id": request_id,
            }

        aggregates = aggregate_participants({"participants": all_participants})

        out_path = tmpdir_path / "out.pptx"
        try:
            inject_pic_analysis(
                formateur_pptx_path=pptx_input_path,
                aggregates=aggregates,
                output_path=out_path,
                analysis_template_path=ANALYSIS_TEMPLATE,
                closing_template_path=CLOSING_TEMPLATE,
            )
        except Exception as e:
            logger.exception("inject_failed")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Injection failed: {type(e).__name__}: {e}",
            )

        out_pptx_bytes = out_path.read_bytes()

    csv_bytes = _build_csv_bytes(all_participants)

    logger.info(
        "inject_slides_completed",
        session_id=session_id,
        nb_participants=aggregates["nb_participants"],
        ocr_errors=len(ocr_errors),
        output_size=len(out_pptx_bytes),
    )

    return {
        "success": True,
        "session_id": session_id,
        "client": client,
        "titre": titre,
        "nb_participants": aggregates["nb_participants"],
        "errors": ocr_errors,
        "pptx_base64": base64.b64encode(out_pptx_bytes).decode(),
        "csv_base64": base64.b64encode(csv_bytes).decode(),
        "raw_data": _build_raw_data(aggregates),
        "request_id": request_id,
    }


@app.post("/inject_slides", dependencies=[Depends(require_api_key)])
@limiter.limit(f"{RATE_LIMIT_PER_MINUTE}/minute")
async def inject_slides_endpoint(
    request: Request,
    pptx_input: UploadFile = File(..., description="PPTX formateur (déjà rempli)"),
    scans_pdf: list[UploadFile] = File(..., description="Scans PDF des questionnaires"),
    session_id: str = Form(..., description="Identifiant de session (slug humain)"),
    client: str = Form(..., description="Nom du client / structure"),
    titre: str = Form(..., description="Titre de la formation"),
):
    """v3 Pic-Formation (multipart) : injection bloc analyse + closing commercial dans le PPTX formateur."""
    _validate_pptx_upload(pptx_input)
    if len(scans_pdf) > MAX_FILES_PER_REQUEST:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Too many files: {len(scans_pdf)} > {MAX_FILES_PER_REQUEST}",
        )
    for f in scans_pdf:
        _validate_upload_file(f)

    max_bytes = MAX_FILE_SIZE_MB * 1024 * 1024
    pptx_bytes = await pptx_input.read()
    if len(pptx_bytes) > max_bytes:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"PPTX input too large ({len(pptx_bytes)} > {max_bytes})",
        )

    pdf_files: list[tuple[str, bytes]] = []
    for f in scans_pdf:
        b = await f.read()
        if len(b) > max_bytes:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail=f"File {f.filename!r} too large ({len(b)} > {max_bytes})",
            )
        pdf_files.append((f.filename or "questionnaire.pdf", b))

    result = _run_inject_pipeline(
        pptx_bytes, pdf_files, session_id, client, titre, request.state.request_id
    )
    if result.get("_status_code"):
        sc = result.pop("_status_code")
        return JSONResponse(status_code=sc, content=result)
    return result


class _InjectJsonRequest(BaseModel):
    session_id: str = Field(..., description="Identifiant de session (slug humain)")
    client: str = Field(..., description="Nom du client / structure")
    titre: str = Field(..., description="Titre de la formation")
    pptx_input_base64: str = Field(..., description="PPTX formateur (base64)")
    scans_pdf: list[_JsonFile] = Field(..., min_length=1)


@app.post("/inject_slides_json", dependencies=[Depends(require_api_key)])
@limiter.limit(f"{RATE_LIMIT_PER_MINUTE}/minute")
async def inject_slides_json_endpoint(
    request: Request,
    payload: _InjectJsonRequest,
):
    """v3 Pic-Formation (JSON+base64) : variante consommée par n8n.

    Évite le multipart depuis n8n (fragile avec arrays de binaries) et le
    sandbox du Code node 2.19+ qui interdit require('form-data') et
    httpRequestWithAuthentication.
    """
    max_bytes = MAX_FILE_SIZE_MB * 1024 * 1024

    try:
        pptx_bytes = base64.b64decode(payload.pptx_input_base64, validate=True)
    except (binascii.Error, ValueError) as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Invalid base64 for pptx_input: {e}",
        )
    if len(pptx_bytes) > max_bytes:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"PPTX input too large ({len(pptx_bytes)} > {max_bytes})",
        )

    if len(payload.scans_pdf) > MAX_FILES_PER_REQUEST:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Too many files: {len(payload.scans_pdf)} > {MAX_FILES_PER_REQUEST}",
        )

    pdf_files: list[tuple[str, bytes]] = []
    for jf in payload.scans_pdf:
        try:
            pdf_bytes = base64.b64decode(jf.base64, validate=True)
        except (binascii.Error, ValueError) as e:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid base64 for file {jf.name!r}: {e}",
            )
        _validate_extension(jf.name)
        if len(pdf_bytes) > max_bytes:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail=f"File {jf.name!r} too large ({len(pdf_bytes)} > {max_bytes})",
            )
        pdf_files.append((jf.name, pdf_bytes))

    result = _run_inject_pipeline(
        pptx_bytes,
        pdf_files,
        payload.session_id,
        payload.client,
        payload.titre,
        request.state.request_id,
    )
    if result.get("_status_code"):
        sc = result.pop("_status_code")
        return JSONResponse(status_code=sc, content=result)
    return result
