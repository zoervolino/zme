"""
qa_validator.py  —  Layer 4: QA Validator
==========================================
FulcrumQ deck conversion pipeline.

Validates the rendered output PPTX against the render log and classifier output.
Produces qa_report.json.

CONTRACT:
  Input:  rendered PPTX path
          list[RenderResult]      (from renderer.py)
          list[ClassifierResult]  (from classifier.py)
  Output: list[QAResult]         + optional qa_report.json

LAYER RULES — this module MUST NOT:
  - Modify any slide XML
  - Write any PPTX file
  - Make layout decisions
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_LAYOUT_CT  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
SLIDE_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"

# ── Known valid FQ layout names ───────────────────────────────────────────────

_VALID_FQ_LAYOUTS = {
    "Cover_Dark",
    "Cover_Light",
    "Cover_Photo",
    "Divider_Dark",
    "Divider_Crystal",
    "Ending_PivotPurple",
    "Ending_Dark",
    "Content_Dark_Sub",
    "Content_Light_Sub",
    "1_Content_Dark",
    "Title_Only_Light",
    "Title_Only_Dark",
    "Content_Light",
    "Blank_Dark",
    "Blank_Light",
}

# Minimum acceptable body text length to flag "content_lost"
_MIN_BODY_CHARS = 5

# Layout families that intentionally have no body or sldNum placeholder slot.
# Missing body/sldNum from these is expected, not an error.
_NO_BODY_LAYOUTS = {
    "Cover_Dark", "Cover_Light", "Cover_Photo",
    "Divider_Dark", "Divider_Crystal",
    "Ending_PivotPurple", "Ending_Dark",
    "Blank_Dark", "Blank_Light",
    "Title_Only_Light", "Title_Only_Dark",
}
# Placeholder types that are never flagged as missing (FQ uses master-level versions)
_SUPPRESS_MISSING_PH = {"sldNum", "dt", "ftr", "subTitle"}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _load_pptx(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(str(path)) as z:
        return {n: z.read(n) for n in z.namelist()}


def _all_text(root: etree._Element) -> str:
    return "".join(t.text or "" for t in root.iter(f"{{{NS_A}}}t"))


def _slide_layout_name(slide_path: str, fm: dict[str, bytes]) -> Optional[str]:
    """Resolve the layout name actually referenced by the slide in the rendered PPTX."""
    parts = slide_path.rsplit("/", 1)
    rels_path = f"{parts[0]}/_rels/{parts[1]}.rels" if len(parts) == 2 else f"_rels/{slide_path}.rels"
    if rels_path not in fm:
        return None
    rels_root = etree.fromstring(fm[rels_path])
    for rel in rels_root:
        if rel.get("Type") == SLIDE_LAYOUT_REL:
            target = rel.get("Target", "")
            layout_path = "/".join(
                ("ppt/slides/" + target).replace("\\", "/").split("/")
            )
            # Normalise: resolve ../ segments
            parts_list = []
            for seg in layout_path.split("/"):
                if seg == "..":
                    if parts_list:
                        parts_list.pop()
                elif seg and seg != ".":
                    parts_list.append(seg)
            resolved = "/".join(parts_list)
            if resolved in fm:
                layout_root = etree.fromstring(fm[resolved])
                cSld = layout_root.find(f"{{{NS_P}}}cSld")
                return (cSld.get("name") if cSld is not None else None) or Path(resolved).stem
    return None


# ── Per-slide checks ──────────────────────────────────────────────────────────

def _check_slide(
    slide_num: int,
    slide_root: etree._Element,
    render_entry: dict,
    classifier_entry: Optional[dict],
    fm: dict[str, bytes],
    slide_path: str,
) -> dict:
    """
    Run all QA checks for a single slide.
    Returns a QAResult dict.
    """
    issues: list[str] = []

    # ── 1. Layout is a known FQ layout ───────────────────────────────────────
    actual_layout = _slide_layout_name(slide_path, fm)
    expected_layout = render_entry.get("layout_used", "")

    if actual_layout is None:
        issues.append("layout_rel_missing")
    elif actual_layout not in _VALID_FQ_LAYOUTS:
        issues.append(f"unknown_layout:{actual_layout}")
    elif actual_layout != expected_layout:
        issues.append(f"layout_mismatch:expected={expected_layout},actual={actual_layout}")

    # ── 2. Title placeholder present and non-empty ────────────────────────────
    title_texts = [
        _all_text(sp)
        for sp in slide_root.iter(f"{{{NS_P}}}sp")
        if any(
            ph.get("type", "") in ("title", "ctrTitle")
            for ph in sp.iter(f"{{{NS_P}}}ph")
        )
    ]
    has_title_ph = bool(title_texts)
    title_empty  = has_title_ph and all(not t.strip() for t in title_texts)

    # Layouts that intentionally have no title ph (purely visual, no text slots)
    _NO_TITLE_LAYOUTS = {"Ending_PivotPurple", "Ending_Dark", "Blank_Dark", "Blank_Light"}
    if not has_title_ph and expected_layout not in _NO_TITLE_LAYOUTS:
        # Suppress on slides whose only text is in non-content placeholders (sldNum, dt, ftr)
        _NOISE_PH = {"sldNum", "dt", "ftr"}
        meaningful_text = "".join(
            _all_text(sp)
            for sp in slide_root.iter(f"{{{NS_P}}}sp")
            if not any(ph.get("type", "") in _NOISE_PH for ph in sp.iter(f"{{{NS_P}}}ph"))
        ).strip()
        if meaningful_text:
            issues.append("title_placeholder_missing")
    elif title_empty:
        issues.append("title_empty")

    # ── 3. Content not lost: if classifier saw body text, it should still exist
    if classifier_entry:
        src_density = classifier_entry.get("metadata", {}).get("text_density", "low")
        rendered_text = _all_text(slide_root)
        if src_density in ("medium", "high") and len(rendered_text.strip()) < _MIN_BODY_CHARS:
            issues.append("content_lost")

    # ── 4. Overflow risk propagated from render log ───────────────────────────
    # Only flag as review-worthy on high-density slides; medium density is informational.
    if render_entry.get("overflow_detected"):
        src_density = (classifier_entry or {}).get("metadata", {}).get("text_density", "high")
        if src_density == "high":
            issues.append("overflow_risk")
        else:
            issues.append("overflow_info")  # informational only, won't affect status

    # ── 5. Missing placeholder propagated from render log ────────────────────
    _COVER_LAYOUTS = {"Cover_Dark", "Cover_Light", "Cover_Photo"}
    for mp in render_entry.get("missing_placeholders", []):
        if mp in _SUPPRESS_MISSING_PH:
            continue
        if mp == "body" and expected_layout in _NO_BODY_LAYOUTS:
            continue
        # subTitle is optional on cover layouts — source decks often omit it
        if mp == "subTitle" and expected_layout in _COVER_LAYOUTS:
            continue
        issues.append(f"missing_placeholder:{mp}")

    # ── 6. Orphan shapes (sp with explicit font/color overrides still present)
    #    — proxy: count shapes with hardcoded latin font that isn't an FQ font
    _FQ_FONTS = {"Aileron", "Switzer", "AileronBlack", "AileronBold"}
    orphan_fonts: set[str] = set()
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        for rPr in sp.iter(f"{{{NS_A}}}rPr"):
            latin = rPr.find(f"{{{NS_A}}}latin")
            if latin is not None:
                face = latin.get("typeface", "")
                if face and face not in _FQ_FONTS and not face.startswith("+"):
                    orphan_fonts.add(face)
    if orphan_fonts:
        issues.append(f"non_fq_fonts:{','.join(sorted(orphan_fonts)[:3])}")

    # ── Determine status ──────────────────────────────────────────────────────
    critical    = {"content_lost", "layout_rel_missing", "title_placeholder_missing"}
    review      = {"overflow_risk", "layout_mismatch", "title_empty", "missing_placeholder", "non_fq_fonts", "unknown_layout"}
    info_only   = {"overflow_info"}  # never affects status

    status_issues = [i for i in issues if not any(i.startswith(x) for x in info_only)]

    if any(any(i.startswith(c) for c in critical) for i in status_issues):
        status = "fail"
    elif any(any(i.startswith(r) for r in review) for i in status_issues):
        status = "review"
    else:
        status = "pass"

    return {
        "slide_id": slide_num,
        "status":   status,
        "issues":   issues,
    }


# ── Global checks ─────────────────────────────────────────────────────────────

def _global_checks(
    render_log: list[dict],
    classifier_results: list[dict],
    fm: dict[str, bytes],
) -> list[str]:
    """
    Deck-level checks — return list of issue strings.
    """
    issues: list[str] = []

    # Slide count consistency
    slide_paths = [k for k in fm if re.match(r"ppt/slides/slide\d+\.xml$", k)]
    if len(slide_paths) != len(render_log):
        issues.append(
            f"slide_count_mismatch:rendered={len(slide_paths)},log={len(render_log)}"
        )

    # Classifier and render log must cover same slide IDs
    cls_ids    = {r["slide_id"] for r in classifier_results}
    render_ids = {r["slide_id"] for r in render_log}
    if cls_ids != render_ids:
        issues.append("slide_id_mismatch_between_layers")

    # Cover slide (slide 1) must be a Cover layout
    first = next((r for r in render_log if r["slide_id"] == 1), None)
    if first:
        if not first["layout_used"].startswith("Cover"):
            issues.append(f"slide_1_not_cover_layout:{first['layout_used']}")

    return issues


# ── Public API ────────────────────────────────────────────────────────────────

def validate_render(
    rendered_pptx: Path,
    render_log: list[dict],
    classifier_results: list[dict],
    output_path: Optional[Path] = None,
) -> dict:
    """
    Validate a rendered PPTX against its render log and classifier output.

    Args:
      rendered_pptx:     Path to the rendered output PPTX
      render_log:        list[RenderResult] from render_deck()
      classifier_results: list[ClassifierResult] from classify_deck()
      output_path:       Optional path to write qa_report.json

    Returns qa_report dict:
      {
        "summary": {"total": int, "pass": int, "review": int, "fail": int},
        "global_issues": list[str],
        "slides": list[QAResult],
      }
    """
    rendered_pptx = Path(rendered_pptx)
    fm = _load_pptx(rendered_pptx)

    slide_paths = sorted(
        [k for k in fm if re.match(r"ppt/slides/slide\d+\.xml$", k)],
        key=lambda p: int(re.search(r"slide(\d+)\.xml$", p).group(1)),
    )

    render_by_id      = {r["slide_id"]: r for r in render_log}
    classifier_by_id  = {r["slide_id"]: r for r in classifier_results}

    qa_slides: list[dict] = []
    for i, slide_path in enumerate(slide_paths):
        slide_num    = i + 1
        slide_root   = etree.fromstring(fm[slide_path])
        render_entry = render_by_id.get(slide_num, {})
        cls_entry    = classifier_by_id.get(slide_num)

        qa_slides.append(
            _check_slide(slide_num, slide_root, render_entry, cls_entry, fm, slide_path)
        )

    global_issues = _global_checks(render_log, classifier_results, fm)

    summary = {
        "total":  len(qa_slides),
        "pass":   sum(1 for s in qa_slides if s["status"] == "pass"),
        "review": sum(1 for s in qa_slides if s["status"] == "review"),
        "fail":   sum(1 for s in qa_slides if s["status"] == "fail"),
    }

    report = {
        "summary":       summary,
        "global_issues": global_issues,
        "slides":        qa_slides,
    }

    if output_path is not None:
        Path(output_path).write_text(json.dumps(report, indent=2))

    return report


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 4:
        print("Usage: python qa_validator.py rendered.pptx render_log.json classifier_output.json [qa_report.json]")
        sys.exit(1)

    rendered_pptx      = Path(sys.argv[1])
    render_log         = json.loads(Path(sys.argv[2]).read_text())
    classifier_results = json.loads(Path(sys.argv[3]).read_text())
    out = Path(sys.argv[4]) if len(sys.argv) > 4 else rendered_pptx.with_suffix(".qa_report.json")

    report = validate_render(rendered_pptx, render_log, classifier_results, output_path=out)
    s = report["summary"]
    print(f"QA: {s['total']} slides — {s['pass']} pass / {s['review']} review / {s['fail']} fail")
    if report["global_issues"]:
        print(f"Global issues: {', '.join(report['global_issues'])}")
    for slide in report["slides"]:
        if slide["status"] != "pass":
            mark = "✗" if slide["status"] == "fail" else "△"
            print(f"  {mark} slide {slide['slide_id']:2d}  {slide['status'].upper():<6}  {'; '.join(slide['issues'])}")
    print(f"\nWrote → {out}")
