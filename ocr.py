"""
OCR extraction via Mistral OCR + chat structuration.
"""

import base64
import json
import os

from mistralai.client import Mistral


STRUCTURATION_PROMPT = """Voici le contenu OCR d'un questionnaire de satisfaction. Les cases cochées sont indiquées par des symboles comme ☑, ☒, ✗, X, x, ou [x].

Extrais toutes les réponses et retourne UNIQUEMENT un JSON valide (sans markdown) dans ce format :
{
  "title": "Questionnaire de satisfaction",
  "categories": [
    {
      "name": "Nom de la catégorie",
      "items": [
        {"label": "Libellé de l'item", "response": "Tout à fait satisfait"}
      ]
    }
  ]
}

Les réponses possibles sont exactement : "Tout à fait satisfait", "Satisfait", "Peu satisfait", "Pas du tout satisfait", "Non renseigné".

Contenu OCR :
"""


def get_client():
    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        raise RuntimeError("MISTRAL_API_KEY non définie")
    return Mistral(api_key=api_key)


def extract_from_pdf(pdf_bytes: bytes) -> dict:
    """Extrait les données d'un questionnaire PDF via Mistral OCR + chat."""
    client = get_client()
    data_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    # Étape 1 : OCR dédié
    ocr_result = client.ocr.process(
        model="mistral-ocr-latest",
        document={
            "type": "document_url",
            "document_url": f"data:application/pdf;base64,{data_b64}",
        },
    )

    markdown = "\n\n".join(page.markdown for page in ocr_result.pages)

    # Étape 2 : Structuration via chat
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
