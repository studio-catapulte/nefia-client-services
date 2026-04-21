"""
OCR extraction via Mistral OCR + chat structuration.
Gère : 1 PDF multi-pages (1 page = 1 questionnaire) ou N PDFs de 1 page.
"""

import base64
import json
import os

from mistralai.client import Mistral


STRUCTURATION_PROMPT = """Voici le contenu OCR d'UN questionnaire de satisfaction rempli par UN participant.
Les cases cochées sont indiquées par ☑ ou ☒. Les cases vides sont vides ou ☐.

Extrais toutes les données et retourne UNIQUEMENT un JSON valide (sans markdown) dans ce format :
{
  "metadata": {
    "participant": "Nom Prénom du participant",
    "formation": "Nom de la formation",
    "lieu": "Lieu / établissement",
    "formateur": "Nom du formateur",
    "date": "Date du bilan"
  },
  "response_levels": ["Très satisfait", "Satisfait", "Déçu", "Très déçu"],
  "items": [
    {
      "number": 1,
      "label": "Libellé de la question",
      "response": "Très satisfait",
      "comment": "Commentaire du participant ou null"
    }
  ],
  "remarks": "Remarques et suggestions libres en fin de questionnaire ou null",
  "wishes": "Souhaits pour d'autres formations ou null"
}

IMPORTANT :
- "response_levels" doit contenir les niveaux exacts tels qu'écrits dans l'en-tête du tableau (du plus positif au plus négatif)
- "response" doit être exactement l'un de ces niveaux
- Si aucune case n'est cochée, "response" = "Non renseigné"
- "comment" = null si pas de commentaire pour cet item

Contenu OCR :
"""


def get_client():
    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        raise RuntimeError("MISTRAL_API_KEY non définie")
    return Mistral(api_key=api_key)


def _ocr_pdf(client, pdf_bytes: bytes) -> list[str]:
    """OCR un PDF et retourne le markdown de chaque page."""
    data_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    ocr_result = client.ocr.process(
        model="mistral-ocr-latest",
        document={
            "type": "document_url",
            "document_url": f"data:application/pdf;base64,{data_b64}",
        },
    )

    return [page.markdown for page in ocr_result.pages]


def _structure_page(client, markdown: str) -> dict:
    """Structuration d'une page de questionnaire via chat."""
    response = client.chat.complete(
        model="mistral-small-latest",
        messages=[
            {"role": "user", "content": STRUCTURATION_PROMPT + markdown},
        ],
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()

    return json.loads(raw)


def extract_from_pdf(pdf_bytes: bytes) -> list[dict]:
    """
    Extrait les données de tous les questionnaires d'un PDF.
    Retourne une liste de questionnaires (1 par page).
    """
    client = get_client()
    pages_md = _ocr_pdf(client, pdf_bytes)

    questionnaires = []
    for md in pages_md:
        if len(md.strip()) < 50:
            continue
        data = _structure_page(client, md)
        questionnaires.append(data)

    return questionnaires
