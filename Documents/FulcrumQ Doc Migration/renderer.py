"""
renderer.py  —  Layer 3: Stateless PPTX Renderer
==================================================
FulcrumQ deck conversion pipeline.

Receives:
  - source PPTX path
  - schema_decisions.json (list[SchemaDecision])
  - FulcrumQ master PPTX path (provides layouts)

Produces:
  - output PPTX with FulcrumQ layouts applied
  - render_log.json (per-slide content mapping + warnings)

CONTRACT:
  Input:  source PPTX path
          list[SchemaDecision]  (from schema_engine.py)
          master PPTX path
  Output: output PPTX + list[RenderResult]

LAYER RULES — this module MUST NOT:
  - Make any layout-selection decisions (those come from schema decisions)
  - Re-classify slides
  - Write schema_decisions.json or classifier_output.json
"""

from __future__ import annotations

import json
import re
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Optional

from lxml import etree

NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS  = "http://schemas.openxmlformats.org/package/2006/content-types"

SLIDE_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
SLIDE_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
THEME_REL        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
SLIDE_LAYOUT_CT  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
SLIDE_MASTER_CT  = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
THEME_CT         = "application/vnd.openxmlformats-officedocument.theme+xml"


# ── PPTX file map helpers ──────────────────────────────────────────────────────

def _load_pptx(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(str(path)) as z:
        return {n: z.read(n) for n in z.namelist()}


def _save_pptx(file_map: dict[str, bytes], out_path: Path) -> None:
    with zipfile.ZipFile(str(out_path), "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in file_map.items():
            z.writestr(name, data)


def _xml(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def _ser(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ── Relationship parsing ───────────────────────────────────────────────────────

def _parse_rels(data: bytes) -> list[dict]:
    root = _xml(data)
    return [
        {"id": r.get("Id"), "type": r.get("Type"), "target": r.get("Target")}
        for r in root
    ]


def _rels_path(content_path: str) -> str:
    parts = content_path.rsplit("/", 1)
    if len(parts) == 2:
        return f"{parts[0]}/_rels/{parts[1]}.rels"
    return f"_rels/{content_path}.rels"


# ── Master layout extraction ───────────────────────────────────────────────────

def _extract_master_layouts(master_fm: dict[str, bytes]) -> dict[str, bytes]:
    """
    Returns {layout_name: layout_xml_bytes} from the master PPTX.
    Layout name is derived from p:cSld name attribute or filename fallback.
    """
    layout_map: dict[str, bytes] = {}

    content_types = _xml(master_fm["[Content_Types].xml"])
    for override in content_types.iter(f"{{{CT_NS}}}Override"):
        ct = override.get("ContentType", "")
        part = override.get("PartName", "").lstrip("/")
        if ct == SLIDE_LAYOUT_CT and part in master_fm:
            layout_root = _xml(master_fm[part])
            cSld = layout_root.find(f"{{{NS_P}}}cSld")
            name = (cSld.get("name") if cSld is not None else None) or Path(part).stem
            layout_map[name] = master_fm[part]

    return layout_map


def _find_layout_path_by_name(
    master_fm: dict[str, bytes],
    layout_name: str,
) -> Optional[str]:
    """Return the internal PPTX path for the named layout, or None."""
    content_types = _xml(master_fm["[Content_Types].xml"])
    for override in content_types.iter(f"{{{CT_NS}}}Override"):
        ct  = override.get("ContentType", "")
        part = override.get("PartName", "").lstrip("/")
        if ct != SLIDE_LAYOUT_CT or part not in master_fm:
            continue
        layout_root = _xml(master_fm[part])
        cSld = layout_root.find(f"{{{NS_P}}}cSld")
        name = (cSld.get("name") if cSld is not None else None) or Path(part).stem
        if name == layout_name:
            return part
    return None


# ── Content extraction from source slide ─────────────────────────────────────

def _extract_placeholder_content(slide_root: etree._Element) -> dict[str, list]:
    """
    Extract text content from standard placeholders.
    Returns: {ph_type: [paragraph_text, ...]}
    e.g. {"title": ["My Title"], "body": ["Bullet 1", "Bullet 2"]}
    """
    content: dict[str, list] = {}
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None:
            continue
        ph_type = ph.get("type", "body")
        ph_idx  = ph.get("idx", "")

        key = ph_type
        if ph_type == "body" and ph_idx == "12":
            key = "subtitle"

        paragraphs = []
        txBody = sp.find(f".//{{{NS_P}}}txBody")
        if txBody is None:
            txBody = sp.find(f".//{{{NS_A}}}txBody")
        if txBody is not None:
            for para in txBody.findall(f"{{{NS_A}}}p"):
                text = "".join(t.text or "" for t in para.iter(f"{{{NS_A}}}t")).strip()
                if text:
                    paragraphs.append(text)

        if paragraphs:
            content.setdefault(key, []).extend(paragraphs)

    return content


def _has_table(slide_root: etree._Element) -> bool:
    return bool(list(slide_root.iter(f"{{{NS_A}}}tbl")))


def _has_image(slide_root: etree._Element) -> bool:
    return bool(list(slide_root.iter(f"{{{NS_P}}}pic")))


# ── Layout remapping ───────────────────────────────────────────────────────────

def _remap_slide_layout(
    slide_path: str,
    layout_name: str,
    source_fm: dict[str, bytes],
    master_fm: dict[str, bytes],
    out_fm: dict[str, bytes],
    master_layout_offset: int,
) -> tuple[dict[str, list], list[str], list[str]]:
    """
    Rewrite the slide's relationship to point at the named FQ layout.
    Copies FQ layout + its rels into out_fm if not already present.

    Returns (content_mapped, missing_placeholders, warnings).
    """
    warnings: list[str] = []
    missing_placeholders: list[str] = []

    slide_root    = _xml(source_fm[slide_path])
    content_mapped = _extract_placeholder_content(slide_root)

    # Locate target layout in master
    layout_src_path = _find_layout_path_by_name(master_fm, layout_name)
    if layout_src_path is None:
        warnings.append(f"layout_not_found:{layout_name}")
        layout_name = "Content_Light"
        layout_src_path = _find_layout_path_by_name(master_fm, layout_name)
        if layout_src_path is None:
            warnings.append("fallback_layout_also_missing")
            return content_mapped, missing_placeholders, warnings

    # Determine destination path for this layout in out_fm
    layout_num  = master_layout_offset
    layout_dst  = f"ppt/slideLayouts/slideLayout{layout_num}.xml"

    if layout_dst not in out_fm:
        out_fm[layout_dst] = master_fm[layout_src_path]
        # Copy layout rels
        layout_rels_src = _rels_path(layout_src_path)
        layout_rels_dst = _rels_path(layout_dst)
        if layout_rels_src in master_fm:
            out_fm[layout_rels_dst] = master_fm[layout_rels_src]

    # Determine destination path for the slide in out_fm
    slide_dst = slide_path  # slide paths preserved 1:1

    # Rewrite slide rels to point at new layout
    slide_rels_src = _rels_path(slide_path)
    if slide_rels_src in source_fm:
        rels_root = _xml(source_fm[slide_rels_src])
        for rel in rels_root:
            if rel.get("Type") == SLIDE_LAYOUT_REL:
                # Compute relative path from ppt/slides/ to ppt/slideLayouts/
                rel.set("Target", f"../slideLayouts/slideLayout{layout_num}.xml")
        out_fm[_rels_path(slide_dst)] = _ser(rels_root)

    # Check which layout placeholders are satisfied.
    # Match by type only (ignore idx) — FQ layouts use non-zero idx for body slots.
    layout_root = _xml(out_fm[layout_dst])
    layout_ph_types_present = {
        ph.get("type", "body")
        for ph in layout_root.iter(f"{{{NS_P}}}ph")
    }
    for content_key in content_mapped:
        ph_type = "body" if content_key == "subtitle" else content_key
        if ph_type not in layout_ph_types_present:
            missing_placeholders.append(content_key)

    # Write (possibly modified) slide to out_fm
    out_fm[slide_dst] = source_fm[slide_path]

    return content_mapped, missing_placeholders, warnings


# ── Font remapping ────────────────────────────────────────────────────────────

# Source fonts to remap → FQ equivalents.
# Loaded from layout_rules.json if present; these are the compile-time defaults.
_DEFAULT_FONT_MAP: dict[str, str] = {
    # Heading weight → Aileron
    "Calibri": "Aileron",
    "Calibri Light": "Aileron",
    "Arial": "Aileron",
    "Helvetica": "Aileron",
    "Helvetica Neue": "Aileron",
    "Helvetica Neue Light": "Aileron",
    "Helvetica Neue Medium": "Aileron",
    "Helvetica Neue Bold": "Aileron",
    "Gill Sans": "Aileron",
    "Gill Sans MT": "Aileron",
    "Futura": "Aileron",
    "Lato": "Switzer",
    "Lato Light": "Switzer",
    "Lato Bold": "Switzer",
    "Open Sans": "Switzer",
    "Roboto": "Switzer",
    "Source Sans Pro": "Switzer",
    "Source Sans 3": "Switzer",
    "Proxima Nova": "Switzer",
    "Myriad Pro": "Switzer",
    "Georgia": "Aileron",
    "Garamond": "Aileron",
    "Times New Roman": "Aileron",
    "Cambria": "Aileron",
}

# Placeholder to inherit from master (clear explicit font so master theme wins).
_CLEAR_FONT_PH_TYPES = {"title", "ctrTitle", "body", "subTitle"}


def _remap_fonts(slide_root: etree._Element, font_map: dict[str, str]) -> int:
    """
    Replace non-FQ latin typefaces with their FQ equivalents.
    For placeholder shapes: clear the explicit latin element so the master theme governs.
    For free shapes: apply the mapped font.
    Returns count of replacements made.
    """
    count = 0
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        is_placeholder = ph is not None
        ph_type = ph.get("type", "body") if ph is not None else ""
        clear_mode = is_placeholder and ph_type in _CLEAR_FONT_PH_TYPES

        for rPr in sp.iter(f"{{{NS_A}}}rPr"):
            latin = rPr.find(f"{{{NS_A}}}latin")
            if latin is None:
                continue
            face = latin.get("typeface", "")
            if not face or face.startswith("+"):
                continue  # theme font reference, leave alone
            if clear_mode:
                rPr.remove(latin)
                count += 1
            elif face in font_map:
                latin.set("typeface", font_map[face])
                count += 1

        # Also clear defRPr and endParaRPr in placeholder txBodies
        if clear_mode:
            for txBody in sp.iter(f"{{{NS_P}}}txBody"):
                for tag in (f"{{{NS_A}}}defRPr", f"{{{NS_A}}}endParaRPr"):
                    for el in txBody.iter(tag):
                        latin = el.find(f"{{{NS_A}}}latin")
                        if latin is not None:
                            el.remove(latin)
                            count += 1

    return count


# Placeholder types that should render as clean labels (no bullets).
_LABEL_PH_TYPES = {"title", "ctrTitle", "subTitle"}


def _strip_bullets(slide_root: etree._Element) -> None:
    """
    Remove bullet markers from label-type placeholder shapes (title, ctrTitle, subTitle).
    These slots should never render as bulleted lists. We remove buChar/buFont/buAutoNum
    from each paragraph's pPr and insert buNone so PowerPoint suppresses bullets.
    """
    _BU_TAGS = {
        f"{{{NS_A}}}buChar",
        f"{{{NS_A}}}buFont",
        f"{{{NS_A}}}buAutoNum",
        f"{{{NS_A}}}buClr",
        f"{{{NS_A}}}buSzPct",
        f"{{{NS_A}}}buSzPts",
    }
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None or ph.get("type", "body") not in _LABEL_PH_TYPES:
            continue
        for txBody in sp.iter(f"{{{NS_P}}}txBody"):
            for para in txBody.findall(f"{{{NS_A}}}p"):
                pPr = para.find(f"{{{NS_A}}}pPr")
                if pPr is None:
                    pPr = etree.SubElement(para, f"{{{NS_A}}}pPr")
                    para.insert(0, pPr)
                # Remove any explicit bullet elements
                for child in list(pPr):
                    if child.tag in _BU_TAGS:
                        pPr.remove(child)
                # Ensure buNone is present to suppress inherited bullets
                if pPr.find(f"{{{NS_A}}}buNone") is None:
                    etree.SubElement(pPr, f"{{{NS_A}}}buNone")


# Layout families that should never show footer/page-number shapes
_NO_FOOTER_LAYOUTS = {
    "Divider_Dark", "Divider_Crystal",
    "Cover_Dark", "Cover_Light", "Cover_Photo",
    "Ending_PivotPurple", "Ending_Dark",
    "Blank_Dark", "Blank_Light",
}
_FOOTER_PH_TYPES = {"sldNum", "dt", "ftr"}


def _strip_footer_shapes(slide_root: etree._Element, layout_name: str) -> int:
    """
    For layouts that intentionally have no footer/page-number, remove any
    sldNum / dt / ftr placeholder shapes from the slide XML so they don't
    render over the clean layout background.
    Returns count of shapes removed.
    """
    if layout_name not in _NO_FOOTER_LAYOUTS:
        return 0

    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return 0

    removed = 0
    for sp in list(spTree):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "") in _FOOTER_PH_TYPES:
            spTree.remove(sp)
            removed += 1
    return removed


# ── Title lift pass ───────────────────────────────────────────────────────────

def _sp_area_emu(sp: etree._Element) -> int:
    """Return cx*cy EMU area of a shape, 0 if not found."""
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return 0
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return 0
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if ext is None:
        return 0
    try:
        return int(ext.get("cx", 0)) * int(ext.get("cy", 0))
    except (ValueError, TypeError):
        return 0


def _sp_top_emu(sp: etree._Element) -> int:
    """Return y-offset EMU of a shape (for topmost heuristic)."""
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return 999_999_999
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return 999_999_999
    off = xfrm.find(f"{{{NS_A}}}off")
    if off is None:
        return 999_999_999
    try:
        return int(off.get("y", 999_999_999))
    except (ValueError, TypeError):
        return 999_999_999


def _lift_title_placeholder(slide_root: etree._Element) -> bool:
    """
    Ensure the slide has a non-empty title placeholder.

    Case A — no title ph at all: promote the best candidate shape to type="title".
    Case B — title ph exists but is empty: copy text from the best free-shape
             candidate into the title ph's txBody, then remove the source shape.

    Candidate search order:
      1. Free text shapes (no ph at all) in the top 25% of the slide.
      2. Body-ph shapes (likely mis-typed title) — no zone filter since they
         rarely have explicit xfrm geometry.
      3. Any free text shape anywhere on the slide.

    Returns True if a lift/fill was performed.
    """
    _TOP_ZONE = 1_285_875  # 25% of standard slide height 5143500 EMU

    title_sp = None
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "") in ("title", "ctrTitle"):
            title_sp = sp
            break

    def _has_text(sp) -> bool:
        return bool("".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip())

    # Case B: title ph exists but is empty — try to fill it from a free shape
    if title_sp is not None:
        if _has_text(title_sp):
            return False  # already has content, nothing to do
        # Fall through to find a donor shape

    # Build candidate pool (same logic for both Case A and Case B)
    # Pass 1: free shapes in top zone
    free_top = [
        sp for sp in slide_root.iter(f"{{{NS_P}}}sp")
        if sp.find(f".//{{{NS_P}}}ph") is None
        and _sp_top_emu(sp) < _TOP_ZONE
        and _has_text(sp)
    ]
    # Pass 2: body-ph shapes (likely mis-typed title).
    # No zone filter — body phs rarely have an explicit xfrm.
    body_top = [
        sp for sp in slide_root.iter(f"{{{NS_P}}}sp")
        if sp.find(f".//{{{NS_P}}}ph") is not None
        and sp.find(f".//{{{NS_P}}}ph").get("type", "body") == "body"
        and _has_text(sp)
    ]
    # Pass 3: any free shape anywhere
    free_any = [
        sp for sp in slide_root.iter(f"{{{NS_P}}}sp")
        if sp.find(f".//{{{NS_P}}}ph") is None and _has_text(sp)
    ]

    pool = free_top or body_top or free_any
    if not pool:
        return False

    def _title_score(sp) -> tuple:
        text = "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()
        short_bonus = 1 if len(text) < 80 else 0
        max_sz = 0
        for rPr in sp.iter(f"{{{NS_A}}}rPr"):
            try:
                max_sz = max(max_sz, int(rPr.get("sz", 0)))
            except (ValueError, TypeError):
                pass
        y = _sp_top_emu(sp)
        return (short_bonus, max_sz, -y)

    best = max(pool, key=_title_score)

    if title_sp is not None:
        # Case B: copy donor text into the existing empty title ph txBody
        src_txBody = best.find(f".//{{{NS_P}}}txBody")
        if src_txBody is None:
            src_txBody = best.find(f".//{{{NS_A}}}txBody")
        dst_txBody = title_sp.find(f".//{{{NS_P}}}txBody")
        if src_txBody is None or dst_txBody is None:
            return False
        # Replace all paragraphs in dst with src paragraphs
        for old_p in dst_txBody.findall(f"{{{NS_A}}}p"):
            dst_txBody.remove(old_p)
        for p in src_txBody.findall(f"{{{NS_A}}}p"):
            dst_txBody.append(p)
        # Remove the donor free shape so content isn't duplicated
        spTree = slide_root.find(f".//{{{NS_P}}}spTree")
        if spTree is not None and best in list(spTree):
            spTree.remove(best)
        return True

    # Case A: promote best candidate to type="title"
    nvSpPr = best.find(f"{{{NS_P}}}nvSpPr")
    if nvSpPr is None:
        return False
    nvPr = nvSpPr.find(f"{{{NS_P}}}nvPr")
    if nvPr is None:
        nvPr = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")

    for old_ph in nvPr.findall(f"{{{NS_P}}}ph"):
        nvPr.remove(old_ph)

    ph_el = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
    ph_el.set("type", "title")
    return True


# ── Overflow detection ────────────────────────────────────────────────────────

def _detect_overflow(slide_root: etree._Element) -> bool:
    """
    Simple heuristic: any txBody with bodyPr normAutofit overflowing is a proxy.
    We check for <a:bodyPr> with no autofit child (PowerPoint default = no-overflow guard).
    Returns True if any text body likely overflows its placeholder.
    Searches both NS_P and NS_A namespaces (slide XML uses NS_P, layout XML uses NS_A).
    """
    txbodies = list(slide_root.iter(f"{{{NS_P}}}txBody")) + \
               list(slide_root.iter(f"{{{NS_A}}}txBody"))
    for txBody in txbodies:
        bodyPr = txBody.find(f"{{{NS_A}}}bodyPr")
        if bodyPr is None:
            continue
        has_autofit = (
            bodyPr.find(f"{{{NS_A}}}normAutofit") is not None
            or bodyPr.find(f"{{{NS_A}}}spAutoFit") is not None
        )
        if not has_autofit:
            text = "".join(t.text or "" for t in txBody.iter(f"{{{NS_A}}}t"))
            if len(text) > 900:
                return True
    return False


# ── Public API ────────────────────────────────────────────────────────────────

def render_deck(
    source_pptx: Path,
    schema_decisions: list[dict],
    master_pptx: Path,
    output_pptx: Path,
    output_log: Optional[Path] = None,
) -> list[dict]:
    """
    Apply schema decisions to the source deck and produce an output PPTX.

    Args:
      source_pptx:      Path to the source PPTX to convert
      schema_decisions: list[SchemaDecision] from run_schema_engine()
      master_pptx:      Path to FulcrumQ master PPTX (source of layouts)
      output_pptx:      Destination path for the converted PPTX
      output_log:       Optional path to write render_log.json

    Returns list[RenderResult].
    """
    source_pptx = Path(source_pptx)
    master_pptx = Path(master_pptx)
    output_pptx = Path(output_pptx)

    source_fm = _load_pptx(source_pptx)
    master_fm = _load_pptx(master_pptx)

    # Load font map — from layout_rules.json if present, else built-in defaults
    _rules_path = Path(__file__).parent / "layout_rules.json"
    if _rules_path.exists():
        _rules_cfg = json.loads(_rules_path.read_text())
        font_map = _rules_cfg.get("font_map", _DEFAULT_FONT_MAP)
    else:
        font_map = _DEFAULT_FONT_MAP

    # Build a working copy of the source
    out_fm: dict[str, bytes] = dict(source_fm)

    # Collect slide paths in order
    slide_paths = sorted(
        [k for k in source_fm if re.match(r"ppt/slides/slide\d+\.xml$", k)],
        key=lambda p: int(re.search(r"slide(\d+)\.xml$", p).group(1)),
    )

    decisions_by_id = {d["slide_id"]: d for d in schema_decisions}

    render_log: list[dict] = []
    layout_offset_counter = 900  # start layouts at a high index to avoid collisions

    for i, slide_path in enumerate(slide_paths):
        slide_id = i + 1
        decision = decisions_by_id.get(slide_id)

        if decision is None:
            render_log.append({
                "slide_id":            slide_id,
                "layout_used":         "UNKNOWN",
                "content_mapped":      {},
                "overflow_detected":   False,
                "missing_placeholders": [],
                "warnings":            ["no_schema_decision"],
            })
            continue

        layout_name = decision["selected_layout"]

        # Mutate slide XML: lift title + remap fonts + strip label bullets + remove footer shapes
        slide_root = _xml(source_fm[slide_path])
        lifted     = _lift_title_placeholder(slide_root)
        fonts_remapped = _remap_fonts(slide_root, font_map)
        _strip_bullets(slide_root)
        _strip_footer_shapes(slide_root, layout_name)
        # Write modified slide back so _remap_slide_layout sees updated content
        source_fm = dict(source_fm)  # shallow copy so we don't mutate the original
        source_fm[slide_path] = _ser(slide_root)

        content_mapped, missing_phs, remap_warnings = _remap_slide_layout(
            slide_path, layout_name, source_fm, master_fm, out_fm, layout_offset_counter
        )
        layout_offset_counter += 1

        overflow = _detect_overflow(slide_root)
        entry_warnings = list(decision.get("warnings", [])) + remap_warnings
        if lifted:
            entry_warnings.append("title_lifted_from_free_shape")
        if fonts_remapped:
            entry_warnings.append(f"fonts_remapped:{fonts_remapped}")

        if overflow:
            entry_warnings.append("overflow_risk")
        if missing_phs:
            entry_warnings.append(f"missing_placeholders:{','.join(missing_phs)}")

        render_log.append({
            "slide_id":             slide_id,
            "layout_used":          layout_name,
            "content_mapped": {
                "title":   content_mapped.get("title", []),
                "bullets": content_mapped.get("body", []),
                "subtitle": content_mapped.get("subtitle", []),
                "tables":  _has_table(slide_root),
                "images":  _has_image(slide_root),
            },
            "overflow_detected":    overflow,
            "missing_placeholders": missing_phs,
            "warnings":             entry_warnings,
        })

    _save_pptx(out_fm, output_pptx)

    if output_log is not None:
        Path(output_log).write_text(json.dumps(render_log, indent=2))

    return render_log


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 5:
        print("Usage: python renderer.py source.pptx schema_decisions.json master.pptx output.pptx [render_log.json]")
        sys.exit(1)

    source_pptx      = Path(sys.argv[1])
    schema_decisions = json.loads(Path(sys.argv[2]).read_text())
    master_pptx      = Path(sys.argv[3])
    output_pptx      = Path(sys.argv[4])
    out_log          = Path(sys.argv[5]) if len(sys.argv) > 5 else output_pptx.with_suffix(".render_log.json")

    log = render_deck(source_pptx, schema_decisions, master_pptx, output_pptx, output_log=out_log)

    ok  = sum(1 for r in log if not r["warnings"])
    warn = sum(1 for r in log if r["warnings"])
    print(f"Rendered {len(log)} slides: {ok} clean, {warn} with warnings")
    for r in log:
        if r["warnings"]:
            print(f"  slide {r['slide_id']:2d}  [{r['layout_used']:<22}]  ⚑ {', '.join(r['warnings'])}")
    print(f"\nOutput  → {output_pptx}")
    print(f"Log     → {out_log}")
