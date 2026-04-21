"""
API questionnaire de satisfaction : PDFs → OCR → PPTX + CSV + JSON.
"""

import base64
import csv
import io
import json
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse

from ocr import extract_from_pdf
from pptx_generator import generate_pptx

app = FastAPI(title="Nefia Questionnaire Processor", version="1.0.0")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/process")
async def process_questionnaires(
    files: list[UploadFile] = File(..., description="PDF questionnaires scannés"),
    title: Optional[str] = Form("Résultats de satisfaction"),
):
    """
    Reçoit N PDFs, les OCR, agrège les résultats, retourne PPTX + CSV + JSON en base64.
    """
    # Extraire chaque questionnaire
    all_data = []
    errors = []

    for i, file in enumerate(files):
        try:
            pdf_bytes = await file.read()
            data = extract_from_pdf(pdf_bytes)
            all_data.append(data)
        except Exception as e:
            errors.append({"file": file.filename, "error": str(e)})

    if not all_data:
        return JSONResponse(
            status_code=400,
            content={"error": "Aucun questionnaire extrait", "details": errors},
        )

    # Générer PPTX
    pptx_bytes = generate_pptx(all_data, title)

    # Générer CSV
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Questionnaire", "Catégorie", "Item", "Réponse"])
    for idx, data in enumerate(all_data, 1):
        for cat in data["categories"]:
            for item in cat["items"]:
                writer.writerow([f"Q{idx}", cat["name"], item["label"], item["response"]])
    csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")  # BOM pour Excel

    return {
        "success": True,
        "questionnaires_processed": len(all_data),
        "errors": errors,
        "pptx_base64": base64.b64encode(pptx_bytes).decode(),
        "csv_base64": base64.b64encode(csv_bytes).decode(),
        "data": all_data,
    }
