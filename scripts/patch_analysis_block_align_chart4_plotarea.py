"""
One-shot script : aligne le `<c:plotArea>/<c:layout>/<c:manualLayout>` de chart4
sur le pattern uniforme des 9 autres charts dans templates/analysis-block.pptx.

POURQUOI :
    Audit 14/05 (suite) — sur 10 charts du template, 9 ont un plotArea
    manualLayout uniforme (x=0.05625, y=0.08828, w=0.84063, h=0.82344). Seul
    chart4 (Q4 "conditions d'accueil") a un layout idiosyncratique
    (x=0.05313, y=0.09460) hérité du PDF EXEMPLE original.

    Conséquence : le DISQUE 3D entier de chart4 est positionné ~0.7 mm à
    gauche et ~0.8 mm plus bas que les autres charts. Le label, qui suit
    le centroid du disque, hérite de ce micro-décalage. Visible quand on
    fait défiler les slides Q1-Q10 — feedback Guillaume 14/05.

    Le patch précédent (`patch_analysis_block_align_chart4_dlbl.py`) ne
    s'occupait que du dLbl idx=0 (offset label depuis centroid). Ici on
    aligne le centroid lui-même via le plotArea.

SOLUTION :
    Patcher chart4/plotArea/layout/manualLayout → (x=0.05625, y=0.08828),
    copie exacte des valeurs présentes sur chart1/2/3/5..10. w et h
    sont déjà identiques sur les 10 charts → on ne touche pas.

USAGE :
    python scripts/patch_analysis_block_align_chart4_plotarea.py

    Idempotent : si chart4 plotArea est déjà aligné, skip silencieux.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "analysis-block.pptx"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# Valeurs cibles : copie exacte du plotArea manualLayout présent sur chart1/2/3/5..10
TARGET_X = "5.6250000000000001E-2"
TARGET_Y = "8.8281121537824719E-2"


def align_chart4_plotarea(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Force le manualLayout du plotArea sur (TARGET_X, TARGET_Y).

    Returns (new_bytes, changed).
    """
    tree = etree.fromstring(xml_bytes)
    changed = False

    plot_area = tree.find(f".//{{{NS_C}}}plotArea")
    if plot_area is None:
        return xml_bytes, False
    layout = plot_area.find(f"{{{NS_C}}}layout")
    if layout is None:
        return xml_bytes, False
    ml = layout.find(f"{{{NS_C}}}manualLayout")
    if ml is None:
        return xml_bytes, False

    x_el = ml.find(f"{{{NS_C}}}x")
    y_el = ml.find(f"{{{NS_C}}}y")
    if x_el is None or y_el is None:
        return xml_bytes, False

    if x_el.get("val") != TARGET_X:
        x_el.set("val", TARGET_X)
        changed = True
    if y_el.get("val") != TARGET_Y:
        y_el.set("val", TARGET_Y)
        changed = True

    if not changed:
        return xml_bytes, False
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True), True


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Template introuvable : {TEMPLATE}")

    backup = TEMPLATE.with_suffix(".pptx.bak")
    shutil.copy(TEMPLATE, backup)

    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        entries = [(zi, zin.read(zi.filename)) for zi in zin.infolist()]

    new_entries = []
    patched = False
    for zi, content in entries:
        if zi.filename == "ppt/charts/chart4.xml":
            new_content, changed = align_chart4_plotarea(content)
            if changed:
                patched = True
                new_entries.append((zi, new_content))
                continue
        new_entries.append((zi, content))

    if not patched:
        print("Rien à faire — chart4 plotArea est déjà aligné.")
        backup.unlink()
        return

    with zipfile.ZipFile(TEMPLATE, "w", zipfile.ZIP_DEFLATED) as zout:
        for zi, content in new_entries:
            zout.writestr(zi, content)

    print(f"[OK] chart4 plotArea aligné sur (x={TARGET_X}, y={TARGET_Y})")
    print(f"Backup : {backup}")


if __name__ == "__main__":
    main()
