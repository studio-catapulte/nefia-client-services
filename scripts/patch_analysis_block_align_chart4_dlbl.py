"""
One-shot script : aligne le `<c:manualLayout>` de chart4 (dLbl idx=0) sur le
pattern uniforme des 9 autres charts dans templates/analysis-block.pptx.

POURQUOI :
    Audit 14/05 — sur 10 charts du template, 9 ont un manualLayout idx=0
    uniforme (x=-0.0008, y=-0.4453) calibré pour positionner le label "N (P %)"
    au centre visuel du pie 3D. Seul chart4 (Q4 "conditions d'accueil") a un
    layout idiosyncratique (x=-0.227, y=-0.3506) hérité du PDF EXEMPLE original
    où la distribution Q4 avait plusieurs catégories non-nulles avec un cas où
    le label devait être à gauche.

    Conséquence : pour toute formation où Q4 a une distribution dominée par
    "Très satisfait" (cas majoritaire), le label "N (P %)" apparaît shifté
    de 23 % à gauche du centre visuel. Feedback Inès 14/05 = "label % décalé
    sur un côté plutôt que centré".

SOLUTION :
    Patcher uniquement chart4 → idx=0 manualLayout = (x=-7.8125e-4,
    y=-0.4453124726), copie exacte des valeurs présentes sur chart1, chart2,
    chart3, chart5..10.

    Approche conservative : on N'INTERVIENT PAS sur dLbl idx=1/2/3 (qui restent
    delete=1 dans le template). Le générateur `_set_dlbl_text` les un-delete
    quand val>0 et leur position fallback est gérée séparément si besoin
    (TODO ultérieur).

    Évite l'approche "strip global" tentée le 14/05 nuit (commit ec7e15d
    reverted) qui cassait le rendu sur toutes les configurations sans résoudre
    le pattern multi-cat sous-jacent.

USAGE :
    python scripts/patch_analysis_block_align_chart4_dlbl.py

    Idempotent : si chart4 est déjà aligné, skip silencieux.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "analysis-block.pptx"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# Valeurs cibles : copie exacte du manualLayout idx=0 présent sur chart1/2/3/5..10
TARGET_X = "-7.8125000000000004E-4"
TARGET_Y = "-0.44531247260627022"


def align_chart4_idx0_layout(xml_bytes: bytes) -> tuple[bytes, bool]:
    """Force le manualLayout du dLbl idx=0 sur (TARGET_X, TARGET_Y).

    Returns (new_bytes, changed).
    """
    tree = etree.fromstring(xml_bytes)
    changed = False

    for dlbl in tree.iter(f"{{{NS_C}}}dLbl"):
        idx_el = dlbl.find(f"{{{NS_C}}}idx")
        if idx_el is None or idx_el.get("val") != "0":
            continue

        layout = dlbl.find(f"{{{NS_C}}}layout")
        if layout is None:
            continue
        ml = layout.find(f"{{{NS_C}}}manualLayout")
        if ml is None:
            continue

        x_el = ml.find(f"{{{NS_C}}}x")
        y_el = ml.find(f"{{{NS_C}}}y")
        if x_el is None or y_el is None:
            continue

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
            new_content, changed = align_chart4_idx0_layout(content)
            if changed:
                patched = True
                new_entries.append((zi, new_content))
                continue
        new_entries.append((zi, content))

    if not patched:
        print("Rien à faire — chart4 idx=0 layout est déjà aligné.")
        backup.unlink()
        return

    with zipfile.ZipFile(TEMPLATE, "w", zipfile.ZIP_DEFLATED) as zout:
        for zi, content in new_entries:
            zout.writestr(zi, content)

    print(f"[OK] chart4 idx=0 layout aligné sur (x={TARGET_X}, y={TARGET_Y})")
    print(f"Backup : {backup}")


if __name__ == "__main__":
    main()
