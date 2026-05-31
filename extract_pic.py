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

STRUCTURATION_PROMPT = f"""Tu reçois le contenu OCR d'un ou plusieurs questionnaires de satisfaction "BILAN A CHAUD" Pic Formation. Chaque questionnaire correspond à UN participant (un seul scan = un seul participant). Les questionnaires sont remplis à la main, les cases cochées sont indiquées par des symboles comme ☒, ✗, X, x, [x], ou un noircissement de la case.

## Géométrie EXACTE du formulaire (à connaître pour bien router le texte)

Le formulaire papier a UNE seule page par participant, structurée ainsi (du haut vers le bas) :

```
┌─────────────────────────────────────────────────────────────────────────┐
│  Bandeau : "BILAN A CHAUD"  Formation: …  ESAT/Lieu: …  Formateur: …    │
│                                                                          │
│  NOM / Prénom : __________________________                               │
│                                                                          │
│  ┌──────────────────────────────────────┬───┬───┬───┬───┬─────────────┐ │
│  │ #  Question                          │TS │ S │ D │TD │Mes commentaires│
│  ├──────────────────────────────────────┼───┼───┼───┼───┼─────────────┤ │
│  │ 1. Cette formation m'a apporté…      │ X │   │   │   │  (souvent vide)│
│  │ 2. Ces compétences vont m'aider…     │ X │   │   │   │              │ │
│  │ … (10 lignes au total)               │   │   │   │   │              │ │
│  │ 10. En conclusion, vous êtes :       │ ☺ │ ☺ │ ☹ │ ☹ │              │ │
│  └──────────────────────────────────────┴───┴───┴───┴───┴─────────────┘ │
│                                                                          │
│  ┌─────────────────────────────────────────────────────────────────────┐ │
│  │ Commentaires : (ce que vous avez bien aimé / ou pas aimé…)          │ │
│  │ ____________________________________________________________________│ │
│  └─────────────────────────────────────────────────────────────────────┘ │
│                                                                          │
│  Vos souhaits pour d'autres formations : ___________________________     │
└─────────────────────────────────────────────────────────────────────────┘
```

Il y a donc TROIS zones de texte libre, et il est CRITIQUE de ne pas les confondre :

1. **Colonne "Mes commentaires"** (à DROITE de la grille des 10 questions, dans le tableau)
   - Chaque ligne de cette colonne est alignée avec UNE question précise (Q1 sur la 1ʳᵉ ligne, …, Q10 sur la 10ᵉ)
   - Contenu typique : très court, souvent 1 à 10 mots ("Trop court", "Temps", "rien", "vidéo concerne plus")
   - Le plus souvent VIDE pour la plupart des questions, voire pour TOUT le tableau
   - → va dans `commentaires_par_question["QN"]`

2. **Boîte "Commentaires : (ce que vous avez bien aimé / ou pas aimé…)"** (SOUS le tableau, sur 1-3 lignes manuscrites)
   - C'est UN SEUL bloc de texte général sur l'ensemble de la formation
   - Contenu typique : phrase(s) complète(s), souvent plus long que les commentaires colonne
   - → va dans `remarque_libre`

3. **Ligne "Vos souhaits pour d'autres formations :"** (TOUT EN BAS, sur 1 ligne)
   - Le texte qui suit le `:` jusqu'à la fin de la page
   - → va dans `souhait_autre_formation`

## Règles ANTI-CONFUSION (les bugs déjà observés)

- Le bandeau "Mes commentaires" (colonne) et le bandeau "Commentaires :" (boîte du bas) se ressemblent. **Ne JAMAIS** placer dans `commentaires_par_question` un texte qui se trouve sous la grille (= sous la ligne Q10). Tout ce qui est sous la grille appartient à `remarque_libre` ou `souhait_autre_formation`, JAMAIS à un commentaire par question.
- Si la colonne "Mes commentaires" est entièrement vide, `commentaires_par_question` doit être un objet **vide** ({{}}) — NE PAS aller chercher du texte ailleurs sur la page pour remplir Q1.
- Le texte de la boîte "Commentaires" du bas ne doit JAMAIS être splitté en deux entre `remarque_libre` et un commentaire par question. C'est un bloc indivisible.
- Le texte qui suit "Vos souhaits pour d'autres formations :" doit aller UNIQUEMENT dans `souhait_autre_formation`, jamais dans `remarque_libre`.
- En cas de doute sur le routage : par défaut → `remarque_libre`. Mieux vaut un remarque_libre trop riche qu'un commentaire fantôme sur Q1.

## Les 10 questions (wording exact, dans cet ordre)

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

## Format de sortie (JSON strict, pas de markdown)

{{
  "participants": [
    {{
      "prenom": "Romain",
      "reponses": {{
        "Q1": "Très satisfait", "Q2": "Satisfait", "Q3": "Très satisfait",
        "Q4": "Satisfait", "Q5": "Très satisfait", "Q6": "Très satisfait",
        "Q7": "Satisfait", "Q8": "Très satisfait", "Q9": "Très satisfait",
        "Q10": "Très satisfait"
      }},
      "commentaires_par_question": {{
        "Q5": "Trop court"
      }},
      "remarque_libre": "J'ai aimé la pratique sur le site de chantier.",
      "souhait_autre_formation": "Gestion du stress"
    }}
  ]
}}

## Règles de format

### `prenom`
- Le champ "NOM / Prénom :" contient typiquement DEUX mots dans cet ordre : NOM puis Prénom. Exemple : "LECLERC ROMAIN" → NOM="LECLERC", Prénom="ROMAIN".
- Extrais SEULEMENT le prénom (le 2ᵉ mot). Pas le nom de famille.
- Normalise en Title Case (1ʳᵉ lettre majuscule, reste minuscules) : "ROMAIN" → "Romain", "MATHIOS" → "Mathios".
- Si un seul mot est lisible, utilise-le comme prénom.
- Si rien n'est lisible, mets "Anonyme N" où N est l'index (1, 2, 3…) du questionnaire dans le PDF.
- N'invente JAMAIS un prénom — préfère "Anonyme N" si tu n'es pas sûr.

### `reponses`
- Valeurs autorisées (exactes) : {RESPONSES_VALID!r} ou "Non renseigné" si la case n'est pas cochée ou ambiguë.
- Les 10 clés "Q1" à "Q10" doivent TOUTES être présentes.
- Pour Q10 ("En conclusion"), l'émoticône entourée détermine la réponse : 😊 vert = "Très satisfait", 🙂 jaune = "Satisfait", ☹️ orange = "Déçu", 😠 rouge = "Très déçu".

### `commentaires_par_question`
- Clé `QN` présente UNIQUEMENT si la cellule "Mes commentaires" de la ligne N de la grille contient du texte manuscrit lisible.
- Si la colonne entière est vide → objet vide `{{}}`. NE PAS y mettre du texte qui vient d'ailleurs.
- Préserve l'orthographe et la ponctuation manuscrites (même les fautes — c'est la parole du participant).
- N'invente JAMAIS de commentaire. Si la cellule est vide ou illisible, omet la clé (pas de chaîne vide).

### `remarque_libre`
- Contenu de la boîte "Commentaires : (ce que vous avez bien aimé / ou pas aimé…)" sous la grille.
- Chaîne vide "" si la boîte est vide.
- Préserve les fautes manuscrites.

### `souhait_autre_formation`
- Contenu de la ligne "Vos souhaits pour d'autres formations :" tout en bas.
- Chaîne vide "" si rien après les deux points.

Extrais TOUS les questionnaires présents dans le contenu OCR, dans l'ordre où ils apparaissent.

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
