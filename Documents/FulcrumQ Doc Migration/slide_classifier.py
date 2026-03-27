"""
slide_classifier.py  —  Phase 1: Slide Classifier

Classifies each slide as one of:  cover | content | divider | ending

For each slide, the classifier:
  1. Extracts structured metadata from the slide XML
  2. Scores each type using weighted evidence signals
  3. Picks the winner; computes a confidence value (0.0 – 1.0)
  4. Flags slides below CONFIDENCE_THRESHOLD for manual review

Output:  classifier_output.json  (written alongside the source file)
Schema per slide:
  {
    "slide":      int,        # 1-based
    "type":       str,        # cover | content | divider | ending
    "layout":     str,        # recommended FulcrumQ layout name
    "confidence": float,      # 0.0 – 1.0
    "review":     bool,       # true if confidence < CONFIDENCE_THRESHOLD
    "signals":    [str, ...], # human-readable evidence that fired
    "reasoning":  str         # one-line summary
  }

Usage:
    python3 slide_classifier.py Org_Efficiency_fq.pptx

Or import and call from convert.py:
    from slide_classifier import classify_deck, write_classifier_output
"""

import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

# ── Namespaces ────────────────────────────────────────────────────────────────
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ── Thresholds ────────────────────────────────────────────────────────────────
CONFIDENCE_THRESHOLD = 0.75   # slides below this get review=True
SPARSE_CHARS         = 80     # total chars below this → sparse / title-only
HEAVY_BODY_CHARS     = 120    # body chars above this → content-heavy
HEAVY_EXTRA_SHAPES   = 4      # non-ph shapes above this → content-heavy
SLIDE_W              = 12_192_000
SLIDE_H              = 6_858_000

# ── Type → default layout mapping ────────────────────────────────────────────
# Refined further by dark-bg detection and content density inside classify_slide.
DEFAULT_LAYOUTS = {
    "cover":   "Cover_Light",
    "divider": "Divider_Dark",
    "ending":  "Ending_PivotPurple",
    "content": "Content_Light",
}

ENDING_KEYWORDS = ("thank", "question", "contact", "connect", "reach out",
                   "let's talk", "lets talk", "get in touch", "next step")
DIVIDER_KEYWORDS = ("agenda", "section", "chapter", "part ", "module",
                    "overview", "introduction", "conclusion")


# ── Metadata extraction ───────────────────────────────────────────────────────

def extract_metadata(slide_root, slide_idx: int, total_slides: int,
                     source_layout_name: str = "") -> dict:
    """
    Pull all classification signals from a slide's XML into a flat metadata dict.
    Nothing here makes a classification decision — that happens in classify_slide().
    """
    n_ph_title = n_ph_body = body_chars = total_chars = n_extra = n_pics = 0
    has_subtitle_ph = False

    for el in slide_root.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag

        if tag == "sp":
            ph    = el.find(f".//{{{NS_P}}}ph")
            texts = "".join(t.text or "" for t in el.iter(f"{{{NS_A}}}t")).strip()
            total_chars += len(texts)

            if ph is not None:
                pt = ph.get("type", "body")
                if pt in ("title", "ctrTitle"):
                    n_ph_title += 1
                elif pt in ("body", "obj"):
                    n_ph_body  += 1
                    body_chars += len(texts)
                elif pt == "subTitle":
                    has_subtitle_ph = True
            else:
                xfrm = el.find(f".//{{{NS_A}}}xfrm")
                ext  = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
                if ext is not None:
                    try:
                        area = int(ext.get("cx", 0)) * int(ext.get("cy", 0))
                        if area < SLIDE_W * SLIDE_H * 0.80:
                            n_extra += 1
                    except (ValueError, TypeError):
                        n_extra += 1
                else:
                    n_extra += 1

        elif tag == "pic":
            n_pics += 1
            n_extra += 1
        elif tag == "grpSp":
            n_extra += 1

    slide_text_lc = " ".join(
        t.text or "" for t in slide_root.iter(f"{{{NS_A}}}t")
    ).lower()

    has_ending_kw  = any(kw in slide_text_lc for kw in ENDING_KEYWORDS)
    has_divider_kw = any(kw in slide_text_lc for kw in DIVIDER_KEYWORDS)

    # Background fill detection: does any large non-ph shape have a dark solid fill?
    has_dark_bg_shape = _has_large_dark_fill(slide_root)

    # Colored block pattern (section divider): large filled non-ph shape with text
    has_colored_block = _has_colored_block_title(slide_root)

    # Source layout name signals (lowercase for matching)
    src = source_layout_name.lower()

    return dict(
        slide_idx        = slide_idx,
        total_slides     = total_slides,
        source_layout    = source_layout_name,
        src_lc           = src,
        n_ph_title       = n_ph_title,
        n_ph_body        = n_ph_body,
        has_subtitle_ph  = has_subtitle_ph,
        body_chars       = body_chars,
        total_chars      = total_chars,
        n_extra          = n_extra,
        n_pics           = n_pics,
        has_ending_kw    = has_ending_kw,
        has_divider_kw   = has_divider_kw,
        has_dark_bg      = has_dark_bg_shape,
        has_colored_block= has_colored_block,
        is_first         = (slide_idx == 1),
        is_last          = (slide_idx == total_slides),
        is_sparse        = (total_chars < SPARSE_CHARS and n_ph_body == 0 and n_extra == 0),
        is_heavy         = (body_chars > HEAVY_BODY_CHARS or n_extra > HEAVY_EXTRA_SHAPES),
    )


def _has_large_dark_fill(slide_root) -> bool:
    """True if a non-ph shape covers ≥30% of the slide with a dark fill (luminance < 80)."""
    MIN_AREA = SLIDE_W * SLIDE_H * 0.30
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        if sp.find(f".//{{{NS_P}}}ph") is not None:
            continue
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        sf   = spPr.find(f".//{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
        if sf is None:
            continue
        try:
            h = sf.get("val", "FFFFFF")
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            lum = 0.299 * r + 0.587 * g + 0.114 * b
            if lum >= 80:
                continue
        except Exception:
            continue
        xfrm = spPr.find(f".//{{{NS_A}}}xfrm")
        ext  = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
        if ext is None:
            continue
        try:
            if int(ext.get("cx", 0)) * int(ext.get("cy", 0)) >= MIN_AREA:
                return True
        except (ValueError, TypeError):
            continue
    return False


def _has_colored_block_title(slide_root) -> bool:
    """True if a large non-ph, non-transparent shape (≥8% slide area) contains text.

    Fill detection is broad: shape is treated as filled unless it has an explicit
    <a:noFill> as a direct child of spPr.  This catches shapes whose fill comes
    from p:style/fillRef or theme inheritance (no solidFill/gradFill in spPr).
    Also detects Pattern B/C: large bg shape (no text) + title ph + no body ph.
    """
    MIN_AREA = SLIDE_W * SLIDE_H * 0.08

    has_large_bg = False
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        if sp.find(f".//{{{NS_P}}}ph") is not None:
            continue
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        # Skip explicitly transparent shapes
        if spPr.find(f"{{{NS_A}}}noFill") is not None:
            continue
        xfrm = spPr.find(f".//{{{NS_A}}}xfrm")
        ext  = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
        if ext is None:
            continue
        try:
            area = int(ext.get("cx", 0)) * int(ext.get("cy", 0))
        except (ValueError, TypeError):
            continue
        if area < MIN_AREA:
            continue
        # Pattern A: shape has text → divider
        if "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip():
            return True
        has_large_bg = True

    # Pattern B: large bg shape + any text (title ph OR non-ph text box) + no body ph
    if has_large_bg:
        has_text    = False
        has_body_ph = False
        for sp in slide_root.iter(f"{{{NS_P}}}sp"):
            ph = sp.find(f".//{{{NS_P}}}ph")
            if ph is not None:
                pt = ph.get("type", "body")
                if pt in ("title", "ctrTitle"):
                    if "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip():
                        has_text = True
                elif pt in ("body", "obj"):
                    has_body_ph = True
            else:
                if "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip():
                    has_text = True
        if has_text and not has_body_ph:
            return True

    # Pattern D: layout-background divider — single large-text shape, no other
    # content. Background comes from the slide layout (nothing to detect on the
    # slide XML itself).  Example: one body ph, 48pt "TALENT to VALUE".
    all_shapes = list(slide_root.iter(f"{{{NS_P}}}sp"))
    text_shapes = [
        sp for sp in all_shapes
        if "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()
    ]
    if len(text_shapes) == 1:
        sp = text_shapes[0]
        ph = sp.find(f".//{{{NS_P}}}ph")
        ph_type = ph.get("type", "body") if ph is not None else None
        if ph_type not in ("sldNum", "dt", "ftr"):
            # Check font size from first rPr sz attribute (hundredths of a point)
            sz_hpt = 0
            for rPr in sp.iter(f"{{{NS_A}}}rPr"):
                try:
                    sz_hpt = int(rPr.get("sz", 0))
                    break
                except (ValueError, TypeError):
                    pass
            if sz_hpt >= 3600:  # ≥36pt
                return True

    return False


# ── Scoring ───────────────────────────────────────────────────────────────────

def _score(m: dict) -> dict[str, tuple[float, list[str]]]:
    """
    Score each slide type independently.
    Returns {type: (raw_score, [signal_strings])}.
    Scores are not normalized here — classify_slide() picks the winner and
    converts to confidence.
    """
    src = m["src_lc"]
    scores = {t: [0.0, []] for t in ("cover", "content", "divider", "ending")}

    def add(t, pts, reason):
        scores[t][0] += pts
        scores[t][1].append(reason)

    # ── COVER signals ─────────────────────────────────────────────────────────
    if m["is_first"]:
        add("cover", 0.40, "slide 1")
    if any(kw in src for kw in ("cover", "title sl", "title slide")):
        add("cover", 0.30, f"source layout '{m['source_layout']}'")
    if m["is_sparse"] and m["is_first"]:
        add("cover", 0.20, "sparse content on slide 1")
    if m["n_ph_body"] == 0 and m["n_extra"] == 0:
        add("cover", 0.15, "no body or extra shapes")
    if m["has_subtitle_ph"] and m["is_first"]:
        add("cover", 0.10, "subtitle placeholder present")
    if m["n_pics"] > 0 and m["is_first"]:
        add("cover", 0.10, "picture on slide 1")
    # Penalise if heavy content — unlikely to be a cover
    if m["is_heavy"]:
        add("cover", -0.30, "heavy content (penalty)")

    # ── DIVIDER signals ───────────────────────────────────────────────────────
    if any(kw in src for kw in ("section", "divider", "chapter", "agenda")):
        add("divider", 0.40, f"source layout '{m['source_layout']}'")
    if m["has_colored_block"]:
        add("divider", 0.35, "large colored block with text")
    if m["has_divider_kw"] and not m["is_heavy"]:
        add("divider", 0.20, "divider keyword in text")
    if m["is_sparse"] and not m["is_first"] and not m["is_last"]:
        add("divider", 0.10, "sparse mid-deck slide")
    if m["n_ph_body"] == 0 and not m["is_first"] and not m["is_last"]:
        add("divider", 0.10, "no body placeholder, not first/last")
    if m["has_dark_bg"] and m["is_sparse"]:
        add("divider", 0.15, "dark background + sparse content")
    if m["is_heavy"]:
        add("divider", -0.25, "heavy content (penalty)")

    # ── ENDING signals ────────────────────────────────────────────────────────
    if m["is_last"]:
        add("ending", 0.35, "last slide")
    if m["has_ending_kw"]:
        add("ending", 0.40, "ending keyword in text")
    if any(kw in src for kw in ("ending", "thank", "closing", "contact")):
        add("ending", 0.30, f"source layout '{m['source_layout']}'")
    if m["is_sparse"] and m["is_last"]:
        add("ending", 0.15, "sparse last slide")

    # ── CONTENT signals (default / residual) ──────────────────────────────────
    if m["n_ph_body"] > 0:
        add("content", 0.40, "body placeholder present")
    if m["body_chars"] > HEAVY_BODY_CHARS:
        add("content", 0.25, f"{m['body_chars']} body chars")
    if m["n_extra"] > HEAVY_EXTRA_SHAPES:
        add("content", 0.20, f"{m['n_extra']} extra shapes")
    if m["is_heavy"]:
        add("content", 0.15, "heavy content")
    # Small baseline so content always has a floor
    add("content", 0.10, "content baseline")

    return {t: (max(0.0, v[0]), v[1]) for t, v in scores.items()}


# ── Classification ────────────────────────────────────────────────────────────

def classify_slide(m: dict) -> dict:
    """
    Given metadata dict from extract_metadata(), return a classification record.
    """
    raw = _score(m)
    # Sort by score descending; on ties prefer content
    ranked = sorted(raw.items(), key=lambda kv: kv[1][0], reverse=True)
    winner_type, (winner_score, winner_signals) = ranked[0]
    second_score = ranked[1][1][0] if len(ranked) > 1 else 0.0

    # Confidence: winner's share of the top-2 combined score, clamped to [0, 1]
    total_top2 = winner_score + second_score
    if total_top2 == 0:
        confidence = 0.50
    else:
        confidence = round(min(1.0, winner_score / total_top2 * (winner_score / max(winner_score, 1.0))), 3)
        # Simpler: ratio approach capped at 0.97
        confidence = round(min(0.97, winner_score / (winner_score + second_score + 0.01)), 3)

    # Hard overrides for very strong single signals — prevents low confidence
    # on obvious cases
    if winner_type == "cover"   and m["is_first"] and not m["is_heavy"]:
        confidence = max(confidence, 0.85)
    if winner_type == "ending"  and m["is_last"]  and m["has_ending_kw"]:
        confidence = max(confidence, 0.90)
    if winner_type == "divider" and m["has_colored_block"]:
        confidence = max(confidence, 0.80)

    # Layout recommendation
    layout = DEFAULT_LAYOUTS[winner_type]
    if winner_type == "content":
        if m["has_dark_bg"]:
            layout = "Content_Dark"
        elif m["has_subtitle_ph"]:
            layout = "Content_Light_Sub"
    elif winner_type == "cover" and m["n_pics"] > 0:
        layout = "Cover_Photo"

    reasoning = (
        f"{winner_type.upper()} ({confidence:.0%}) — "
        + ", ".join(winner_signals[:3])
        + (f" [+{len(winner_signals)-3} more]" if len(winner_signals) > 3 else "")
    )

    return {
        "slide":      m["slide_idx"],
        "type":       winner_type,
        "layout":     layout,
        "confidence": confidence,
        "review":     confidence < CONFIDENCE_THRESHOLD,
        "signals":    winner_signals,
        "reasoning":  reasoning,
    }


# ── Deck-level classification ─────────────────────────────────────────────────

def classify_deck(file_map: dict, slides: list[str],
                  old_layout_names: dict | None = None) -> list[dict]:
    """
    Classify all slides in a file_map.

    Parameters
    ----------
    file_map         : {zip_path: bytes}  — full PPTX file map
    slides           : ordered list of slide paths, e.g. ['ppt/slides/slide1.xml', ...]
    old_layout_names : optional {layout_path: layout_name} for source layout hints

    Returns list of classification records (one per slide), in slide order.
    """
    old_layout_names = old_layout_names or {}
    total = len(slides)
    results = []

    for i, spath in enumerate(slides, 1):
        slide_bytes = file_map.get(spath, b"")
        if not slide_bytes:
            results.append({
                "slide": i, "type": "content", "layout": "Content_Light",
                "confidence": 0.0, "review": True,
                "signals": ["slide XML missing"],
                "reasoning": "CONTENT (0%) — slide XML missing",
            })
            continue

        slide_root = etree.fromstring(slide_bytes)

        # Resolve source layout name if provided
        rels_path = re.sub(
            r"ppt/slides/(slide\d+\.xml)$",
            r"ppt/slides/_rels/\1.rels",
            spath,
        )
        src_layout_name = ""
        rels_bytes = file_map.get(rels_path, b"")
        if rels_bytes:
            rels_root = etree.fromstring(rels_bytes)
            for rel in rels_root:
                tgt = rel.get("Target", "")
                if "slideLayout" in tgt:
                    lpath = "ppt/slideLayouts/" + tgt.split("/")[-1]
                    src_layout_name = old_layout_names.get(lpath, "")
                    break

        m = extract_metadata(slide_root, i, total, src_layout_name)
        results.append(classify_slide(m))

    # ── Pass 2: deck-level reconciliation ─────────────────────────────────────
    results = _reconcile(results)
    return results


def _reconcile(results: list[dict]) -> list[dict]:
    """
    Holistic pass over all slide classifications:
    - If no cover found, promote slide 1 to cover if it isn't already
    - If ending not last, demote to content + flag for review
    - Adjacent dividers: demote the second to content + review
    - Ensure at most one cover
    """
    n = len(results)
    if n == 0:
        return results

    # Promote slide 1 to cover if no cover exists
    has_cover = any(r["type"] == "cover" for r in results)
    if not has_cover and results[0]["type"] != "cover":
        r = results[0]
        r["type"]      = "cover"
        r["layout"]    = DEFAULT_LAYOUTS["cover"]
        r["review"]    = True
        r["signals"].append("promoted to cover — no cover found (reconcile)")
        r["reasoning"] = "COVER (reconcile) — no cover found, slide 1 promoted"

    # At most one cover — demote extras beyond slide 1
    cover_indices = [i for i, r in enumerate(results) if r["type"] == "cover"]
    for ci in cover_indices[1:]:
        r = results[ci]
        r["type"]      = "content"
        r["layout"]    = DEFAULT_LAYOUTS["content"]
        r["review"]    = True
        r["signals"].append("demoted from cover — only 1 cover allowed (reconcile)")
        r["reasoning"] = "CONTENT (reconcile) — extra cover demoted"

    # Ending must be last — demote any mid-deck endings
    for i, r in enumerate(results[:-1]):
        if r["type"] == "ending":
            r["type"]      = "content"
            r["layout"]    = DEFAULT_LAYOUTS["content"]
            r["review"]    = True
            r["signals"].append("demoted from ending — not last slide (reconcile)")
            r["reasoning"] = "CONTENT (reconcile) — ending slide not last"

    # Adjacent dividers — demote second of each pair
    for i in range(n - 1):
        if results[i]["type"] == "divider" and results[i+1]["type"] == "divider":
            r = results[i+1]
            r["type"]      = "content"
            r["layout"]    = DEFAULT_LAYOUTS["content"]
            r["review"]    = True
            r["signals"].append("demoted from divider — adjacent divider (reconcile)")
            r["reasoning"] = "CONTENT (reconcile) — adjacent divider"

    return results


# ── Output ────────────────────────────────────────────────────────────────────

def write_classifier_output(results: list[dict], output_path: Path) -> Path:
    """Write classifier_output.json alongside output_path."""
    out = output_path.with_suffix(".classifier.json")
    out.write_text(json.dumps(results, indent=2))
    return out


def print_classifier_summary(results: list[dict]) -> None:
    review_count = sum(1 for r in results if r["review"])
    type_counts  = {}
    for r in results:
        type_counts[r["type"]] = type_counts.get(r["type"], 0) + 1

    print(f"\n{'='*55}")
    print("SLIDE CLASSIFIER")
    for t in ("cover", "divider", "content", "ending"):
        n = type_counts.get(t, 0)
        if n:
            print(f"  {t:<10}: {n} slide(s)")
    if review_count:
        print(f"\n  Needs review ({review_count} slide(s)):")
        for r in results:
            if r["review"]:
                print(f"    slide {r['slide']:2d}  [{r['type']:<8}]  "
                      f"conf={r['confidence']:.0%}  {r['reasoning']}")
    else:
        print("  All slides classified with high confidence.")
    print(f"{'='*55}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def _load_zip(path: Path) -> dict:
    with zipfile.ZipFile(str(path), "r") as z:
        return {n: z.read(n) for n in z.namelist()}


def _ordered_slides(file_map: dict) -> list[str]:
    keys = [k for k in file_map if re.match(r"ppt/slides/slide\d+\.xml$", k)]
    return sorted(keys, key=lambda k: int(re.search(r"(\d+)\.xml$", k).group(1)))


if __name__ == "__main__":
    import argparse as _ap
    p = _ap.ArgumentParser(description="FulcrumQ slide classifier")
    p.add_argument("pptx", help="Source .pptx file")
    p.add_argument("--vision", action="store_true",
                   help="Run AI vision pass (requires ANTHROPIC_API_KEY + LibreOffice)")
    p.add_argument("--keep-images", action="store_true",
                   help="Keep rasterized slide PNGs (with --vision)")
    args = p.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        print(f"Not found: {src}")
        sys.exit(1)

    file_map = _load_zip(src)
    slides   = _ordered_slides(file_map)
    results  = classify_deck(file_map, slides)

    if args.vision:
        from slide_vision import classify_deck_vision, print_vision_summary
        vision_results = classify_deck_vision(
            src, xml_hints=results, keep_images=args.keep_images
        )
        print_vision_summary(vision_results)
        out = write_classifier_output(vision_results, src)
    else:
        print_classifier_summary(results)
        out = write_classifier_output(results, src)

    print(f"\n  Output → {out.name}\n")
