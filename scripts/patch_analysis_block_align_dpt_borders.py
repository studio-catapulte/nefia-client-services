"""
One-shot script : aligne la couleur de bordure (`<a:ln>`) et de contour 3D
(`<a:sp3d>/<a:contourClr>`) de chaque `<c:dPt>` sur sa couleur de remplissage
(`<a:solidFill>`) dans templates/analysis-block.pptx.

POURQUOI :
    Audit 14/05 (suite v2.4.2) — dans les 10 charts du template, le `dPt idx=0`
    (slice "Très satisfait" verte) a sa bordure `<a:ln>` et son contour 3D
    `<a:sp3d>/<a:contourClr>` alignés sur sa couleur de remplissage (srgb
    00B050 vert). Les autres dPt (idx=1 jaune, idx=2 orange, idx=3 rouge,
    plus une 2e série fantôme en accent1/2/3/4) ont leurs ln + contourClr en
    `schemeClr lt1` (blanc du thème) — héritage du PDF EXEMPLE original Pic.

    Tant qu'une slice non-idx=0 est à 0 (supprimée par _update_pie3d_chart),
    invisible. Dès qu'une slice idx=1/2/3 apparaît (cas multi-cat), le liseré
    blanc devient visible autour de la slice ET un micro cercle blanc apparaît
    aux sommets 3D où plusieurs arêtes convergent (effet du `sp3d contourW=2pt`
    en blanc qui ressort sur fond coloré).

    Feedback Guillaume 14/05 (multi-cat Q1 jaune 1/17%) : "micro carré blanc à
    côté du label, pas dérangeant mais d'où il vient ?".

SOLUTION :
    Pour chaque `<c:dPt>` du template :
    - Lire la couleur de `<a:solidFill>` (srgbClr ou schemeClr)
    - Recopier cette couleur dans `<a:ln>/<a:solidFill>` (bordure)
    - Recopier dans `<a:sp3d>/<a:contourClr>` (contour 3D)

    Résultat : bordures et contours invisibles parce que même couleur que le
    fill. Pattern uniforme avec ce que Pic a fait nativement sur idx=0.

USAGE :
    python scripts/patch_analysis_block_align_dpt_borders.py

    Idempotent : si déjà aligné, skip silencieux.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "analysis-block.pptx"

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _color_kind_val(spPr):
    """Return (kind, val) où kind ∈ {'srgbClr', 'schemeClr'} ou None."""
    sf = spPr.find(f"{{{NS_A}}}solidFill")
    if sf is None:
        return None
    srgb = sf.find(f"{{{NS_A}}}srgbClr")
    if srgb is not None:
        return ("srgbClr", srgb.get("val"))
    scheme = sf.find(f"{{{NS_A}}}schemeClr")
    if scheme is not None:
        return ("schemeClr", scheme.get("val"))
    return None


def _same_color(parent, kind, val):
    """Test si parent contient déjà <a:{kind} val="{val}"/>."""
    el = parent.find(f"{{{NS_A}}}{kind}")
    return el is not None and el.get("val") == val and len(list(parent)) == 1


def align_dpt_borders(xml_bytes: bytes) -> tuple[bytes, int]:
    """Aligne ln/solidFill + sp3d/contourClr de chaque dPt sur son fill.

    Returns (new_bytes, n_changed).
    """
    tree = etree.fromstring(xml_bytes)
    changed = 0

    for dpt in tree.iter(f"{{{NS_C}}}dPt"):
        spPr = dpt.find(f"{{{NS_C}}}spPr")
        if spPr is None:
            continue
        color = _color_kind_val(spPr)
        if color is None:
            continue
        kind, val = color

        # Aligner <a:ln>/<a:solidFill>
        ln = spPr.find(f"{{{NS_A}}}ln")
        if ln is not None:
            ln_sf = ln.find(f"{{{NS_A}}}solidFill")
            if ln_sf is None or not _same_color(ln_sf, kind, val):
                for sf in ln.findall(f"{{{NS_A}}}solidFill"):
                    ln.remove(sf)
                # Insérer en première position pour respecter l'ordre canonique
                new_sf = etree.Element(f"{{{NS_A}}}solidFill")
                new_color = etree.SubElement(new_sf, f"{{{NS_A}}}{kind}")
                new_color.set("val", val)
                ln.insert(0, new_sf)
                changed += 1

        # Aligner <a:sp3d>/<a:contourClr>
        sp3d = spPr.find(f"{{{NS_A}}}sp3d")
        if sp3d is not None:
            cc = sp3d.find(f"{{{NS_A}}}contourClr")
            if cc is not None and not _same_color(cc, kind, val):
                for child in list(cc):
                    cc.remove(child)
                new_color = etree.SubElement(cc, f"{{{NS_A}}}{kind}")
                new_color.set("val", val)
                changed += 1

    if not changed:
        return xml_bytes, 0
    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True), changed


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Template introuvable : {TEMPLATE}")

    backup = TEMPLATE.with_suffix(".pptx.bak")
    shutil.copy(TEMPLATE, backup)

    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        entries = [(zi, zin.read(zi.filename)) for zi in zin.infolist()]

    new_entries = []
    total = 0
    patched_files = []
    for zi, content in entries:
        if zi.filename.startswith("ppt/charts/chart") and zi.filename.endswith(".xml"):
            new_content, n = align_dpt_borders(content)
            if n > 0:
                total += n
                patched_files.append(f"{zi.filename} ({n} refs)")
                new_entries.append((zi, new_content))
                continue
        new_entries.append((zi, content))

    if total == 0:
        print("Rien à faire — bordures dPt déjà alignées sur fills.")
        backup.unlink()
        return

    with zipfile.ZipFile(TEMPLATE, "w", zipfile.ZIP_DEFLATED) as zout:
        for zi, content in new_entries:
            zout.writestr(zi, content)

    print(f"[OK] {total} bordures/contours alignés sur fills :")
    for f in patched_files:
        print(f"  - {f}")
    print(f"Backup : {backup}")


if __name__ == "__main__":
    main()
