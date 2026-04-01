"""
services/pptx_extract.py — Raw PPTX Signal Extraction
======================================================
Layer 0 of the Semantic Ingest pipeline.

CONTRACT:
  Input:  PPTX file path (str or Path)
  Output: dict with raw signals per slide

RULES:
  - NO semantic interpretation
  - Only extract raw signals
  - Must not lose content
  - No UI code
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

# ── XML namespaces ────────────────────────────────────────────────────────────
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

_NS = {
    "a": NS_A,
    "p": NS_P,
    "r": NS_R,
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _all_text(elem: etree._Element) -> str:
    """Extract all text content from an XML element tree."""
    parts = []
    for t in elem.iter("{%s}t" % NS_A):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


def _shape_type(sp: etree._Element) -> str:
    """Determine shape type from XML."""
    nvsp = sp.find(".//{%s}nvSpPr" % NS_P) or sp.find(".//{%s}nvSpPr" % NS_A)
    if nvsp is not None:
        ph = nvsp.find(".//{%s}ph" % NS_P)
        if ph is not None:
            return ph.get("type", "body")
    pic = sp.find(".//{%s}pic" % NS_P)
    if pic is not None:
        return "picture"
    chart = sp.find(".//{%s}chart" % "http://schemas.openxmlformats.org/drawingml/2006/chart")
    if chart is not None:
        return "chart"
    return "freeform"


def _extract_colors(sp: etree._Element) -> list[str]:
    """Extract all color hex values from a shape element."""
    colors = []
    for solid in sp.iter("{%s}solidFill" % NS_A):
        srgb = solid.find("{%s}srgbClr" % NS_A)
        if srgb is not None:
            val = srgb.get("val", "")
            if val:
                colors.append(val.upper())
        scheme = solid.find("{%s}schemeClr" % NS_A)
        if scheme is not None:
            colors.append("scheme:" + scheme.get("val", ""))
    return colors


def _extract_table(tbl: etree._Element) -> list[list[str]]:
    """Extract table contents as list of rows."""
    rows = []
    for tr in tbl.findall("{%s}tr" % NS_A):
        row = []
        for tc in tr.findall("{%s}tc" % NS_A):
            row.append(_all_text(tc))
        rows.append(row)
    return rows


def _layout_signature(sp_tree: etree._Element) -> str:
    """
    Produce a rough layout signature: count and rough positions of shapes.
    Format: "<shape_count>sh|<has_table>|<has_image>|<has_chart>"
    """
    shapes = sp_tree.findall("{%s}sp" % NS_P)
    n_shapes = len(shapes)
    has_table = sp_tree.find(".//{%s}graphicFrame" % NS_P) is not None
    has_image = sp_tree.find(".//{%s}pic" % NS_P) is not None
    has_chart = sp_tree.find(
        ".//{%s}chart" % "http://schemas.openxmlformats.org/drawingml/2006/chart"
    ) is not None
    return f"{n_shapes}sh|{'T' if has_table else 'f'}|{'I' if has_image else 'f'}|{'C' if has_chart else 'f'}"


# ── Main extraction ───────────────────────────────────────────────────────────

def extract_raw_deck(pptx_path: str | Path) -> dict[str, Any]:
    """
    Extract all raw signals from a PPTX file.

    Returns:
        {
          "source": str,           # filename
          "slide_count": int,
          "slides": [
            {
              "slide_id": int,      # 1-based
              "raw_text": [str],    # all text runs from all shapes
              "shapes": [           # per-shape metadata
                {
                  "shape_name": str,
                  "shape_type": str,
                  "text": str,
                  "placeholder_type": str | None,
                }
              ],
              "tables": [           # extracted table data
                [[str]]             # rows of cells
              ],
              "colors": [str],      # all color hex codes seen
              "layout_sig": str     # rough layout signature
            }
          ]
        }
    """
    pptx_path = Path(pptx_path)
    slides_out = []

    with zipfile.ZipFile(pptx_path, "r") as z:
        # Find slide entries in order
        slide_keys = sorted(
            [k for k in z.namelist() if k.startswith("ppt/slides/slide") and k.endswith(".xml")],
            key=lambda k: int("".join(filter(str.isdigit, Path(k).stem)) or "0"),
        )

        for idx, slide_key in enumerate(slide_keys, start=1):
            raw_xml = z.read(slide_key)
            root = etree.fromstring(raw_xml)

            cSld = root.find("{%s}cSld" % NS_P)
            sp_tree = cSld.find("{%s}spTree" % NS_P) if cSld is not None else None

            raw_text: list[str] = []
            shapes_out: list[dict] = []
            tables_out: list[list[list[str]]] = []
            colors_out: list[str] = []

            if sp_tree is not None:
                # Extract shapes (sp elements)
                for sp in sp_tree.findall("{%s}sp" % NS_P):
                    name_elem = sp.find(".//{%s}cNvPr" % NS_P)
                    shape_name = name_elem.get("name", "") if name_elem is not None else ""

                    txt = _all_text(sp)
                    if txt:
                        raw_text.append(txt)

                    ph_type = None
                    nvsp = sp.find("{%s}nvSpPr" % NS_P)
                    if nvsp is not None:
                        ph = nvsp.find(".//{%s}ph" % NS_P)
                        if ph is not None:
                            ph_type = ph.get("type", "body")

                    s_type = "placeholder" if ph_type else "freeform"

                    # Check for picture inside sp
                    if sp.find(".//{%s}blipFill" % NS_A) is not None:
                        s_type = "picture_shape"

                    shapes_out.append({
                        "shape_name": shape_name,
                        "shape_type": s_type,
                        "placeholder_type": ph_type,
                        "text": txt,
                    })
                    colors_out.extend(_extract_colors(sp))

                # Extract graphic frames (tables, charts)
                for gf in sp_tree.findall("{%s}graphicFrame" % NS_P):
                    name_elem = gf.find(".//{%s}cNvPr" % NS_P)
                    shape_name = name_elem.get("name", "") if name_elem is not None else ""

                    # Table
                    tbl = gf.find(".//{%s}tbl" % NS_A)
                    if tbl is not None:
                        table_data = _extract_table(tbl)
                        tables_out.append(table_data)
                        # Also add table text to raw_text
                        for row in table_data:
                            for cell in row:
                                if cell:
                                    raw_text.append(cell)
                        shapes_out.append({
                            "shape_name": shape_name,
                            "shape_type": "table",
                            "placeholder_type": None,
                            "text": " | ".join(
                                " ".join(r) for r in table_data
                            ),
                        })

                    # Chart
                    chart_ref = gf.find(
                        ".//{%s}chart" % "http://schemas.openxmlformats.org/drawingml/2006/chart"
                    )
                    if chart_ref is not None:
                        shapes_out.append({
                            "shape_name": shape_name,
                            "shape_type": "chart",
                            "placeholder_type": None,
                            "text": "",
                        })

                # Picture shapes (pic elements at spTree level)
                for pic in sp_tree.findall("{%s}pic" % NS_P):
                    name_elem = pic.find(".//{%s}cNvPr" % NS_P)
                    shape_name = name_elem.get("name", "") if name_elem is not None else ""
                    shapes_out.append({
                        "shape_name": shape_name,
                        "shape_type": "picture",
                        "placeholder_type": None,
                        "text": "",
                    })

                layout_sig = _layout_signature(sp_tree)
            else:
                layout_sig = "0sh|f|f|f"

            slides_out.append({
                "slide_id": idx,
                "raw_text": raw_text,
                "shapes": shapes_out,
                "tables": tables_out,
                "colors": list(dict.fromkeys(colors_out)),  # deduplicated, order preserved
                "layout_sig": layout_sig,
            })

    return {
        "source": pptx_path.name,
        "slide_count": len(slides_out),
        "slides": slides_out,
    }
