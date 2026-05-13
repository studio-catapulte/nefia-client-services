"""
OCR + structuration des questionnaires Pic Formation via Mistral.

Format Pic Formation (10 questions standardisées) :
    - 10 questions à 4 modalités : Très satisfait / Satisfait / Déçu / Très déçu
    - Commentaire libre par question (zone "Mes commentaires")
    - 2 zones libres en fin : "Vos remarques et suggestions" + "Vos souhaits pour d'autres formations"
    - 1 PDF peut contenir 1 ou N questionnaires (1 par participant)

Output : list de dicts participants. Voir docstring de extract_questionnaires_pic.
"""

import base64
import json
import os
import sys
from pathlib import Path

from mistralai.client import Mistral


# Wording exact des 10 questions Pic Formation (figé)
Q_WORDINGS = [
    "Cette formation m'a apporté de nouvelles compétences et connaissances",
    "Ces compétences vont m'aider dans mon travail",
    "L'ambiance du groupe était bonne",
    "Les conditions d'accueil étaient adaptées (salle, restauration…)",
    "La durée et le rythme étaient bien adaptés",
    "Le formateur a bien expliqué (objectifs, contenus…)",
    "Les supports étaient adaptés (photos, vidéos…)",
    "Le formateur a pris en compte mes difficultés",
    "Les jeux & exercices pratiques m'ont aidé à bien comprendre",
    "En conclusion, vous êtes :",
]

# Modalités exactes (telles qu'imprimées sur le questionnaire PIC)
RESPONSES_VALID = ["Très satisfait", "Satisfait", "Déçu", "Très déçu"]

STRUCTURATION_PROMPT = f"""Tu reçois le contenu OCR d'un ou plusieurs questionnaires de satisfaction Pic Formation. Chaque questionnaire correspond à UN participant (le prénom est en haut). Les questionnaires sont remplis à la main, les cases cochées sont indiquées par des symboles comme ☒, ✗, X, x, [x], ou un noircissement de la case.

Format du questionnaire Pic Formation :
- 10 questions standardisées (toujours dans cet ordre, parfois numérotées 1-10 sur le scan)
- Pour chaque question, 4 modalités possibles : "Très satisfait", "Satisfait", "Déçu", "Très déçu"
- Pour chaque question, une zone "Mes commentaires" optionnelle (souvent vide ou avec quelques mots manuscrits)
- En fin de questionnaire :
  - "Vos remarques et suggestions" (texte libre, optionnel)
  - "Vos souhaits pour d'autres formations" (texte libre, optionnel)

Les 10 questions exactes sont :
1. Cette formation m'a apporté de nouvelles compétences et connaissances
2. Ces compétences vont m'aider dans mon travail
3. L'ambiance du groupe était bonne
4. Les conditions d'accueil étaient adaptées (salle, restauration…)
5. La durée et le rythme étaient bien adaptés
6. Le formateur a bien expliqué (objectifs, contenus…)
7. Les supports étaient adaptés (photos, vidéos…)
8. Le formateur a pris en compte mes difficultés
9. Les jeux & exercices pratiques m'ont aidé à bien comprendre
10. En conclusion, vous êtes :

Extrais TOUS les questionnaires présents dans le contenu OCR, dans l'ordre où ils apparaissent. Retourne UNIQUEMENT un JSON valide (sans markdown), au format suivant :

{{
  "participants": [
    {{
      "prenom": "Prénom du participant tel qu'écrit en haut du questionnaire",
      "reponses": {{
        "Q1": "Très satisfait",
        "Q2": "Satisfait",
        "Q3": "Très satisfait",
        "Q4": "Satisfait",
        "Q5": "Très satisfait",
        "Q6": "Très satisfait",
        "Q7": "Satisfait",
        "Q8": "Très satisfait",
        "Q9": "Très satisfait",
        "Q10": "Très satisfait"
      }},
      "commentaires_par_question": {{
        "Q1": "Texte du commentaire si présent, sinon omettre la clé",
        "Q5": "Trop court"
      }},
      "remarque_libre": "Contenu de la zone 'Vos remarques et suggestions' si rempli, sinon chaîne vide",
      "souhait_autre_formation": "Contenu de la zone 'Vos souhaits pour d'autres formations' si rempli, sinon chaîne vide"
    }}
  ]
}}

Règles strictes :
- Les valeurs de "reponses" doivent être EXACTEMENT l'une de : {RESPONSES_VALID!r} ou la chaîne "Non renseigné" si la case n'est pas cochée ou ambiguë.
- Les clés "Q1" à "Q10" doivent toutes être présentes dans "reponses".
- N'invente pas de commentaires : si la zone est vide ou illisible, omet la clé pour cette question (ne mets pas "" non plus).
- Préserve l'orthographe et la ponctuation manuscrites dans les commentaires (même les fautes — c'est le mot du participant).
- Pour le prénom : si illisible ou absent, mets "Anonyme N" où N est l'index (1, 2, 3…) du questionnaire.

Contenu OCR :
"""


def read_file_as_base64(file_path: str) -> tuple[str, str]:
    path = Path(file_path)
    suffix = path.suffix.lower()
    media_type_map = {
        ".pdf": "application/pdf",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
    }
    media_type = media_type_map.get(suffix)
    if not media_type:
        raise ValueError(f"Format non supporté : {suffix}")
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, media_type


def extract_questionnaires_pic(file_path: str) -> dict:
    """Pipeline OCR Mistral → JSON structuré au format Pic Formation.

    Output :
        {
            "participants": [
                {
                    "prenom": str,
                    "reponses": {"Q1": "Très satisfait", ..., "Q10": "..."},
                    "commentaires_par_question": {"Q1": "...", "Q5": "..."},  # clés présentes seulement si non vide
                    "remarque_libre": str,
                    "souhait_autre_formation": str,
                },
                ...
            ]
        }
    """
    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        raise RuntimeError("MISTRAL_API_KEY non définie")

    client = Mistral(api_key=api_key)
    data, media_type = read_file_as_base64(file_path)

    print("  OCR en cours...")
    if media_type == "application/pdf":
        document = {
            "type": "document_url",
            "document_url": f"data:{media_type};base64,{data}",
        }
    else:
        document = {
            "type": "image_url",
            "image_url": f"data:{media_type};base64,{data}",
        }

    ocr_result = client.ocr.process(
        model="mistral-ocr-latest",
        document=document,
    )
    markdown = "\n\n".join(page.markdown for page in ocr_result.pages)
    print(f"  OCR terminé ({len(ocr_result.pages)} page(s), {len(markdown)} chars)")

    print("  Structuration JSON (Mistral chat)...")
    response = client.chat.complete(
        model="mistral-small-latest",
        messages=[
            {"role": "user", "content": STRUCTURATION_PROMPT + markdown},
        ],
        response_format={"type": "json_object"},
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()

    parsed = json.loads(raw)

    # Sanity check
    if "participants" not in parsed:
        raise ValueError(f"Pas de clé 'participants' dans la réponse Mistral : {raw[:300]}")

    return parsed


def aggregate_participants(extraction: dict) -> dict:
    """Agrège la sortie de extract_questionnaires_pic en `data` pour pptx_generator.

    Convertit la liste de participants en :
        - nb_participants : int
        - questions : list[10] de {counts: dict, commentaires: list[{prenom, commentaire}]}
        - remarques_libres : list[{prenom, commentaire}]
        - souhaits_libres : list[{prenom, commentaire}]
    """
    participants = extraction.get("participants", [])
    nb = len(participants)

    questions = []
    for q_idx in range(1, 11):
        q_key = f"Q{q_idx}"
        counts = {r: 0 for r in RESPONSES_VALID}
        commentaires = []
        for p in participants:
            resp = ((p.get("reponses") or {}).get(q_key) or "").strip()
            if resp in counts:
                counts[resp] += 1
            comm = ((p.get("commentaires_par_question") or {}).get(q_key) or "").strip()
            if comm:
                commentaires.append({"prenom": p.get("prenom", ""), "commentaire": comm})
        questions.append({"counts": counts, "commentaires": commentaires})

    remarques_libres = [
        {"prenom": p.get("prenom", ""), "commentaire": (p.get("remarque_libre") or "").strip()}
        for p in participants
        if (p.get("remarque_libre") or "").strip()
    ]
    souhaits_libres = [
        {"prenom": p.get("prenom", ""), "commentaire": (p.get("souhait_autre_formation") or "").strip()}
        for p in participants
        if (p.get("souhait_autre_formation") or "").strip()
    ]

    return {
        "nb_participants": nb,
        "questions": questions,
        "remarques_libres": remarques_libres,
        "souhaits_libres": souhaits_libres,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract.py <fichier.pdf|image>")
        sys.exit(1)

    extraction = extract_questionnaires_pic(sys.argv[1])
    aggregated = aggregate_participants(extraction)

    print("\n--- Extraction brute ---")
    print(json.dumps(extraction, indent=2, ensure_ascii=False))

    print("\n--- Agrégation ---")
    print(json.dumps(aggregated, indent=2, ensure_ascii=False))
