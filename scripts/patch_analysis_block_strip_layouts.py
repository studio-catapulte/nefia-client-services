"""
One-shot script : strip les <c:layout><c:manualLayout> de tous les charts
dans templates/analysis-block.pptx (plotArea + dLbls).

POURQUOI :
    Feedback Inès 14/05 — sur Q4 ("Les conditions d'accueil..."), micro
    décalage du camembert + label % décalé sur un côté plutôt que centré.

    Cause : le template a été créé à partir du PDF EXEMPLE original, et les
    positions des dLbls + plotArea ont été *manuellement ajustées* dans
    PowerPoint pour les valeurs spécifiques de CET exemple. Ces overrides
    `manualLayout` sont figés dans le XML.

    Quand on réutilise le template pour n'importe quel jeu de données :
    - le plotArea garde la position calibrée pour l'exemple → micro shift
    - les dLbl gardent leur position absolue → label peut atterrir n'importe où

SOLUTION :
    Supprimer les `<c:layout><c:manualLayout>` :
    - sur le plotArea : PowerPoint auto-centre le pie dans le chart frame
    - sur chaque dLbl : PowerPoint auto-positionne le label à côté/dans la slice

    Le rendu redevient cohérent quelles que soient les valeurs en entrée.

USAGE :
    python scripts/patch_analysis_block_strip_layouts.py

    Idempotent : si aucun manualLayout n'est trouvé, skip silencieux.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "analysis-block.pptx"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def strip_layouts_from_chart_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    """Strip <c:layout> from plotArea + every <c:dLbl> in a chart XML.

    Returns (new_bytes, n_stripped).
    """
    tree = etree.fromstring(xml_bytes)
    n = 0

    # plotArea layout
    for layout in tree.iter(f"{{{NS_C}}}plotArea"):
        for child in list(layout):
            if child.tag == f"{{{NS_C}}}layout":
                layout.remove(child)
                n += 1

    # dLbl layouts (each individual data label)
    for dlbl in tree.iter(f"{{{NS_C}}}dLbl"):
        for child in list(dlbl):
            if child.tag == f"{{{NS_C}}}layout":
                dlbl.remove(child)
                n += 1

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True), n


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Template introuvable : {TEMPLATE}")

    # Backup
    backup = TEMPLATE.with_suffix(".pptx.bak")
    shutil.copy(TEMPLATE, backup)

    # Read all entries
    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        entries = [(zi, zin.read(zi.filename)) for zi in zin.infolist()]

    total_stripped = 0
    modified = []

    new_entries = []
    for zi, content in entries:
        if zi.filename.startswith("ppt/charts/") and zi.filename.endswith(".xml") and "chartXml" not in zi.filename and "_rels" not in zi.filename:
            new_content, n = strip_layouts_from_chart_xml(content)
            if n > 0:
                modified.append((zi.filename, n))
                total_stripped += n
                new_entries.append((zi, new_content))
                continue
        new_entries.append((zi, content))

    if total_stripped == 0:
        print("Rien à faire — aucun manualLayout trouvé.")
        backup.unlink()
        return

    # Write back
    with zipfile.ZipFile(TEMPLATE, "w", zipfile.ZIP_DEFLATED) as zout:
        for zi, content in new_entries:
            zout.writestr(zi, content)

    for fname, n in modified:
        print(f"  [OK] {fname} : {n} <c:layout> stripped")
    print(f"\n{total_stripped} <c:layout> stripped au total sur {len(modified)} charts.")
    print(f"Backup : {backup}")


if __name__ == "__main__":
    main()
