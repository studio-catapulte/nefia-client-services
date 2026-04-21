"""
API questionnaire de satisfaction : PDFs → OCR → PPTX + CSV + JSON.
Gère : N PDFs de 1 page, ou 1 PDF multi-pages, ou un mix.
"""

import base64
import csv
import io
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse

from ocr import extract_from_pdf
from pptx_generator import generate_pptx

app = FastAPI(title="Nefia Questionnaire Processor", version="2.0.0")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/process")
async def process_questionnaires(
    files: list[UploadFile] = File(..., description="PDF questionnaires scannes"),
    title: Optional[str] = Form("Resultats de satisfaction"),
):
    """
    Recoit N PDFs (chacun pouvant contenir plusieurs pages/questionnaires).
    OCR chaque page, agrege, retourne PPTX + CSV + JSON en base64.
    """
    all_questionnaires = []
    errors = []

    for file in files:
        try:
            pdf_bytes = await file.read()
            questionnaires = extract_from_pdf(pdf_bytes)
            all_questionnaires.extend(questionnaires)
        except Exception as e:
            errors.append({"file": file.filename, "error": str(e)})

    if not all_questionnaires:
        return JSONResponse(
            status_code=400,
            content={"error": "Aucun questionnaire extrait", "details": errors},
        )

    # Generer PPTX
    pptx_bytes = generate_pptx(all_questionnaires, title)

    # Generer CSV
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

    return {
        "success": True,
        "questionnaires_processed": len(all_questionnaires),
        "errors": errors,
        "pptx_base64": base64.b64encode(pptx_bytes).decode(),
        "csv_base64": base64.b64encode(csv_bytes).decode(),
        "data": all_questionnaires,
    }
