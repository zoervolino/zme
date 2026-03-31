"""
classifier.py  —  Layer 1: Slide Classification
================================================
FulcrumQ deck conversion pipeline.

Classifies each slide as: cover | content | divider | ending.
Extracts per-slide metadata used by the Schema Engine.

CONTRACT:
  Input:  PPTX file path
  Output: list[ClassifierResult] + optional classifier_output.json

LAYER RULES — this module MUST NOT:
  - Apply styling, fonts, or colors
  - Make layout template decisions
  - Modify any slide XML
  - Produce output PPTX
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

CONFIDENCE_THRESHOLD = 0.75

# Closing-slide trigger words
_ENDING_WORDS = {"thank", "question", "contact", "get in touch", "next step", "reach out"}


# ── Metadata extractors ───────────────────────────────────────────────────────

def _all_text(root) -> str:
    return "".join(t.text or "" for t in root.iter(f"{{{NS_A}}}t"))


def _text_density(root) -> str:
    n = len(_all_text(root).strip())
    if n < 100:
        return "low"
    if n < 500:
        return "medium"
    return "high"


def _has_chart(root) -> bool:
    return bool(list(root.iter(f"{{{NS_P}}}graphicFrame")))


def _image_count(root) -> int:
    return len(list(root.iter(f"{{{NS_P}}}pic")))


def _has_logo(root, slide_w: int, slide_h: int) -> bool:
    """True if any image is < 5 % of slide area (corner logo heuristic)."""
    area = slide_w * slide_h
    for pic in root.iter(f"{{{NS_P}}}pic"):
        spPr = pic.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        xfrm = spPr.find(f"{{{NS_A}}}xfrm")
        if xfrm is None:
            continue
        ext = xfrm.find(f"{{{NS_A}}}ext")
        if ext is None:
            continue
        try:
            if int(ext.get("cx", 0)) * int(ext.get("cy", 0)) < 0.05 * area:
                return True
        except ValueError:
            pass
    return False


def _bg_is_dark(root) -> bool:
    bg = root.find(f"{{{NS_P}}}cSld/{{{NS_P}}}bg")
    if bg is None:
        return False
    for srgb in bg.iter(f"{{{NS_A}}}srgbClr"):
        v = srgb.get("val", "").upper()
        try:
            r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
            if 0.299 * r + 0.587 * g + 0.114 * b < 128:
                return True
        except Exception:
            pass
    return False


def _has_subtitle(root) -> bool:
    """True if slide already has a subtitle placeholder OR title suggests a split."""
    for ph in root.iter(f"{{{NS_P}}}ph"):
        t = ph.get("type", "")
        if t == "subTitle":
            return True
        if t == "body" and ph.get("idx") == "12":
            return True
    # Soft-return or pipe in title text
    for ph in root.iter(f"{{{NS_P}}}ph"):
        if ph.get("type", "") not in ("title", "ctrTitle"):
            continue
        ancestor = ph
        for _ in range(3):
            ancestor = ancestor.getparent()
            if ancestor is None:
                break
        if ancestor is None:
            continue
        if ancestor.find(f".//{{{NS_A}}}br") is not None:
            return True
        txt = "".join(t.text or "" for t in ancestor.iter(f"{{{NS_A}}}t"))
        if " | " in txt:
            return True
    return False


# ── Heuristic classifier ──────────────────────────────────────────────────────

def _classify_heuristic(root, slide_num: int, total: int) -> tuple[str, float, str]:
    """Return (type, confidence, reasoning)."""
    phs       = {ph.get("type", "body") for ph in root.iter(f"{{{NS_P}}}ph")}
    all_sp    = list(root.iter(f"{{{NS_P}}}sp"))
    ph_sp     = [sp for sp in all_sp if sp.find(f".//{{{NS_P}}}ph") is not None]
    free_sp   = [sp for sp in all_sp if sp.find(f".//{{{NS_P}}}ph") is None]
    txt       = _all_text(root).strip()
    chars     = len(txt)
    has_body  = "body" in phs
    bg_dark   = _bg_is_dark(root)
    has_grfx  = _has_chart(root)

    # ── Cover ────────────────────────────────────────────────────────────────
    if slide_num == 1:
        if not has_body and ("ctrTitle" in phs or "title" in phs):
            return "cover", 0.90, "Slide 1, title placeholder, no body → cover"
        if chars < 200:
            return "cover", 0.80, f"Slide 1, {chars} chars, sparse → cover"
        return "cover", 0.62, "Slide 1 → cover by position; has body content (low confidence)"

    # ── Ending ───────────────────────────────────────────────────────────────
    if slide_num == total:
        if any(w in txt.lower() for w in _ENDING_WORDS):
            return "ending", 0.92, "Last slide with closing language → ending"
        if chars < 100:
            return "ending", 0.78, "Last slide, sparse text → ending"

    # ── Divider ──────────────────────────────────────────────────────────────
    text_shapes = [
        sp for sp in all_sp
        if "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()
    ]
    if len(text_shapes) == 1:
        sp = text_shapes[0]
        for rPr in sp.iter(f"{{{NS_A}}}rPr"):
            try:
                sz = int(rPr.get("sz", 0))
                if sz >= 3600:
                    return "divider", 0.88, f"Single text shape at {sz//100}pt → divider"
            except ValueError:
                pass
        if not has_body:
            return "divider", 0.82, "Single text shape, no body placeholder → divider"

    if "title" in phs and not has_body and chars < 150:
        if bg_dark:
            return "divider", 0.86, "Title-only, dark background → divider"
        if free_sp:
            return "divider", 0.76, "Title-only with decorative shapes → divider"

    # ── Content ──────────────────────────────────────────────────────────────
    extra = len(free_sp)
    conf = 0.68 if extra > 10 else (0.80 if has_body else 0.72)
    parts = [f"{_text_density(root)} density"]
    if extra > 0:
        parts.append(f"{extra} extra shapes")
    if has_grfx:
        parts.append("chart/table")
    return "content", conf, f"CONTENT ({', '.join(parts)})"


# ── Public API ────────────────────────────────────────────────────────────────

def classify_slide(
    root,
    slide_num: int,
    total: int,
    slide_w: int = 9144000,
    slide_h: int = 5143500,
) -> dict:
    """Classify a single slide and return a ClassifierResult dict."""
    slide_type, confidence, reasoning = _classify_heuristic(root, slide_num, total)

    warnings: list[str] = []
    if confidence < CONFIDENCE_THRESHOLD:
        warnings.append(f"low_confidence:{confidence:.2f}")
    if slide_num == 1 and slide_type != "cover":
        warnings.append("slide_1_not_cover")

    return {
        "slide_id":   slide_num,
        "type":       slide_type,
        "confidence": round(confidence, 3),
        "reasoning":  reasoning,
        "metadata": {
            "text_density":  _text_density(root),
            "has_chart":     _has_chart(root),
            "image_count":   _image_count(root),
            "has_logo":      _has_logo(root, slide_w, slide_h),
            "bg_is_dark":    _bg_is_dark(root),
            "has_subtitle":  _has_subtitle(root),
        },
        "warnings": warnings,
    }


def classify_deck(pptx_path, output_path: Optional[Path] = None) -> list[dict]:
    """
    Classify every slide in pptx_path.
    Optionally writes classifier_output.json to output_path.
    Returns list of ClassifierResult dicts.
    """
    pptx_path = Path(pptx_path)
    with zipfile.ZipFile(str(pptx_path)) as z:
        file_map = {n: z.read(n) for n in z.namelist()}

    prs = etree.fromstring(file_map["ppt/presentation.xml"])
    NS_P2 = NS_P
    sldSz   = prs.find(f"{{{NS_P2}}}sldSz")
    slide_w = int(sldSz.get("cx", 9144000)) if sldSz is not None else 9144000
    slide_h = int(sldSz.get("cy", 5143500)) if sldSz is not None else 5143500

    slide_paths = sorted(
        [k for k in file_map
         if re.match(r"ppt/slides/slide\d+\.xml$", k)],
        key=lambda p: int(re.search(r"slide(\d+)\.xml$", p).group(1)),
    )
    total = len(slide_paths)

    results = [
        classify_slide(
            etree.fromstring(file_map[sp]),
            i + 1,
            total,
            slide_w,
            slide_h,
        )
        for i, sp in enumerate(slide_paths)
    ]

    if output_path is not None:
        Path(output_path).write_text(json.dumps(results, indent=2))

    return results


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python classifier.py deck.pptx [output.json]")
        sys.exit(1)
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(sys.argv[1]).with_suffix(".classifier.json")
    results = classify_deck(sys.argv[1], output_path=out)
    for r in results:
        flag = " ⚑" if r["warnings"] else ""
        print(f"  slide {r['slide_id']:2d}  [{r['type']:<8}]  {r['confidence']:.2f}{flag}  {r['reasoning'][:60]}")
    print(f"\nWrote → {out}")
