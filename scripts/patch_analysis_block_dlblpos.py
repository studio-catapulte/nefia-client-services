"""
One-shot script : force <c:dLblPos val="ctr"/> sur chaque série de
templates/analysis-block.pptx (Q1-Q10).

POURQUOI :
    Feedback Inès 14/05 (2e itération) : après strip-layouts précédent,
    sur 100% Très satisfait, le label N (xx %) n'est plus centré sur
    la slice. Cause = sans manualLayout figé, PowerPoint utilise le
    dLblPos du chart pour positionner. Si dLblPos est absent ou
    "bestFit", PowerPoint décide selon une heuristique imprévisible.

    Forcer "ctr" (center) garantit que le label apparaît au CENTRE
    de la slice, quelle que soit la distribution :
    - 100% → label au centre du pie
    - 50/50 → label au centre de chaque demi-cercle
    - 17/83 → label au centre géométrique de chaque part

USAGE :
    python scripts/patch_analysis_block_dlblpos.py

    Idempotent : si dLblPos existe déjà avec val="ctr", skip.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "analysis-block.pptx"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def force_dlblpos_ctr(xml_bytes: bytes) -> tuple[bytes, int]:
    """Set <c:dLblPos val="ctr"/> on each <c:ser>/<c:dLbls> in chart XML.

    Insertion canonique : juste avant les show* flags.
    Si dLblPos existe déjà, on update val à "ctr".
    Returns (new_bytes, n_modified).
    """
    tree = etree.fromstring(xml_bytes)
    n = 0

    for dlbls in tree.iter(f"{{{NS_C}}}dLbls"):
        # On ne touche que les dLbls au niveau série (parent = c:ser),
        # pas ceux à l'intérieur d'un dLbl per-point (rare mais possible)
        parent = dlbls.getparent()
        if parent is None or parent.tag != f"{{{NS_C}}}ser":
            continue

        existing = dlbls.find(f"{{{NS_C}}}dLblPos")
        if existing is not None:
            if existing.get("val") == "ctr":
                continue
            existing.set("val", "ctr")
            n += 1
            continue

        # Créer le dLblPos et l'insérer canoniquement avant les show*
        new_el = etree.Element(f"{{{NS_C}}}dLblPos")
        new_el.set("val", "ctr")
        # Position canonique dans dLbls : ... → c:dLblPos → c:showLegendKey → c:showVal → c:showCatName → ...
        insert_before = None
        for tag_local in ("showLegendKey", "showVal", "showCatName", "showSerName",
                          "showPercent", "showBubbleSize", "separator", "extLst"):
            el = dlbls.find(f"{{{NS_C}}}{tag_local}")
            if el is not None:
                insert_before = el
                break
        if insert_before is not None:
            insert_before.addprevious(new_el)
        else:
            dlbls.append(new_el)
        n += 1

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True), n


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Template introuvable : {TEMPLATE}")

    backup = TEMPLATE.with_suffix(".pptx.bak")
    shutil.copy(TEMPLATE, backup)

    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        entries = [(zi, zin.read(zi.filename)) for zi in zin.infolist()]

    total = 0
    modified = []
    new_entries = []
    for zi, content in entries:
        if (
            zi.filename.startswith("ppt/charts/")
            and zi.filename.endswith(".xml")
            and "_rels" not in zi.filename
        ):
            new_content, n = force_dlblpos_ctr(content)
            if n > 0:
                modified.append((zi.filename, n))
                total += n
                new_entries.append((zi, new_content))
                continue
        new_entries.append((zi, content))

    if total == 0:
        print("Rien à faire — dLblPos=ctr déjà présent partout.")
        backup.unlink()
        return

    with zipfile.ZipFile(TEMPLATE, "w", zipfile.ZIP_DEFLATED) as zout:
        for zi, content in new_entries:
            zout.writestr(zi, content)

    for fname, n in modified:
        print(f"  [OK] {fname} : {n} <c:dLblPos val='ctr'/> added/updated")
    print(f"\n{total} dLblPos modifiés sur {len(modified)} charts.")
    print(f"Backup : {backup}")


if __name__ == "__main__":
    main()
