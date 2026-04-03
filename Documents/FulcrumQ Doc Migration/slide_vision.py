"""
slide_vision.py
AI vision layer for the FulcrumQ slide classifier.

Two-pass classification:
  Pass 1 — Claude vision classifies each rasterized slide individually,
            with slide-position hints (is_first, is_last).
  Pass 2 — Holistic reconciliation: Claude reviews the full deck structure
            and corrects any misclassifications given overall flow.

XML heuristics from slide_classifier.py can optionally seed Pass 1 as hints,
improving accuracy on low-confidence vision calls.

Usage (CLI):
    python3 slide_vision.py deck.pptx
    python3 slide_vision.py deck.pptx --keep-images --out-dir ./slides_out

Usage (API):
    from slide_vision import classify_deck_vision
    results = classify_deck_vision("deck.pptx", xml_hints=xml_results)
"""

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

try:
    import anthropic
    _HAS_ANTHROPIC = True
except ImportError:
    _HAS_ANTHROPIC = False

# ── Constants ─────────────────────────────────────────────────────────────────

SOFFICE_CANDIDATES = [
    "/opt/homebrew/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "soffice",
    "libreoffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/libreoffice",
]

VISION_MODEL = "claude-haiku-4-5-20251001"   # fast + cheap for per-slide work
RECON_MODEL  = "claude-sonnet-4-6"            # stronger model for deck-level reasoning

SLIDE_TYPES = ("cover", "content", "divider", "ending")

CONFIDENCE_THRESHOLD = 0.75   # below this → flagged for manual review

# ── Prompts ───────────────────────────────────────────────────────────────────

_PASS1_SYSTEM = """\
You are a slide classification engine for a professional consulting deck converter.
Your job is to classify each slide into exactly one category.
Always respond with valid JSON and nothing else."""

_PASS1_USER = """\
Classify this slide. Slide {slide_num} of {total_slides}.{hint_line}

Categories:
  cover    — Opening slide. Large title + optional subtitle, no body bullets or data.
             Branded background, logo, or hero image. Almost always slide 1.
  content  — Regular slide with substantial information: bullets, charts, tables,
             icons, diagrams. Has a title + real body content.
  divider  — Section separator. One large bold heading, strong visual treatment
             (solid colour block, full-bleed background). Minimal text — just a
             label to introduce the next section.
  ending   — Closing slide. "Thank you", "Questions?", "Get in touch", or similar.
             Very sparse. Usually the last slide.

Respond with JSON only:
{{"type": "cover|content|divider|ending", "confidence": 0.0–1.0, "reasoning": "one sentence"}}"""

_PASS2_SYSTEM = """\
You are reviewing the slide-type classification of a complete deck.
Correct any structural errors. Respond with valid JSON and nothing else."""

_PASS2_USER = """\
{total_slides}-slide deck — per-slide classifications from Pass 1:

{summary}

Check for structural issues:
  1. Exactly one cover slide, and it must be slide 1.
  2. No two divider slides adjacent to each other.
  3. Ending slide (if present) must be the last slide.
  4. Any slide that looks misclassified given the overall flow?

Return a list of overrides — only slides you want to change.
Empty list if everything is correct.

[{{"slide": N, "type": "cover|content|divider|ending", "confidence": 0.0–1.0, "reasoning": "why"}}]"""


# ── LibreOffice rasterization ─────────────────────────────────────────────────

def _find_soffice() -> Optional[str]:
    for path in SOFFICE_CANDIDATES:
        try:
            r = subprocess.run([path, "--version"], capture_output=True, timeout=5)
            if r.returncode == 0:
                return path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None


def rasterize_pptx(pptx_path: Path, out_dir: Path) -> list[Path]:
    """Convert PPTX → one PNG per slide.

    Strategy:
      1. LibreOffice converts PPTX → PDF (all pages, reliable on macOS).
      2. pymupdf renders each PDF page → slide_NNN.png in out_dir.

    Returns a list of PNG paths sorted by slide number.
    Raises RuntimeError if LibreOffice or pymupdf is not available.
    """
    soffice = _find_soffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice not found. Install with: brew install --cask libreoffice\n"
            "Expected path: /Applications/LibreOffice.app"
        )

    try:
        import fitz  # pymupdf
    except ImportError:
        raise RuntimeError(
            "pymupdf not found. Install with: pip install pymupdf"
        )

    pptx_path = Path(pptx_path).resolve()
    out_dir   = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: PPTX → PDF
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(pptx_path)],
        capture_output=True, text=True, timeout=180,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice PDF conversion failed:\n{result.stderr}")

    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"Expected PDF not found: {pdf_path}")

    # Step 2: PDF → per-page PNGs via pymupdf
    doc  = fitz.open(str(pdf_path))
    pngs = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        pix  = page.get_pixmap(dpi=96)
        out_path = out_dir / f"slide_{page_num + 1:03d}.png"
        pix.save(str(out_path))
        pngs.append(out_path)
    doc.close()

    return pngs


# ── Pass 1: per-slide vision ──────────────────────────────────────────────────

def _encode_image(path: Path) -> str:
    return base64.standard_b64encode(path.read_bytes()).decode("utf-8")


def _parse_json_response(text: str) -> dict:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    return {}


def _classify_slide_vision(
    client,
    img_path: Path,
    slide_num: int,
    total_slides: int,
    xml_hint: Optional[dict] = None,
) -> dict:
    """Call Claude vision to classify a single slide image."""
    hint_line = ""
    if xml_hint:
        hint_line = (
            f"\nXML heuristic hint: {xml_hint['type']} "
            f"(confidence {xml_hint.get('confidence', 0):.2f})"
        )

    msg = client.messages.create(
        model=VISION_MODEL,
        max_tokens=200,
        system=_PASS1_SYSTEM,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": _encode_image(img_path),
                    },
                },
                {
                    "type": "text",
                    "text": _PASS1_USER.format(
                        slide_num=slide_num,
                        total_slides=total_slides,
                        hint_line=hint_line,
                    ),
                },
            ],
        }],
    )

    data = _parse_json_response(msg.content[0].text)
    slide_type = data.get("type", "content")
    if slide_type not in SLIDE_TYPES:
        slide_type = "content"

    return {
        "slide":      slide_num,
        "type":       slide_type,
        "confidence": float(data.get("confidence", 0.5)),
        "reasoning":  data.get("reasoning", ""),
        "changed":    False,
        "source":     "vision_pass1",
    }


# ── Pass 2: holistic reconciliation ──────────────────────────────────────────

def _reconcile_vision(client, pass1: list[dict]) -> list[dict]:
    """Send all Pass 1 results to Claude for deck-level structural review."""
    summary_lines = [
        f"  Slide {r['slide']:2d}: {r['type']:<10} "
        f"conf={r['confidence']:.2f}  {r['reasoning']}"
        for r in pass1
    ]
    summary = "\n".join(summary_lines)

    msg = client.messages.create(
        model=RECON_MODEL,
        max_tokens=1024,
        system=_PASS2_SYSTEM,
        messages=[{
            "role": "user",
            "content": _PASS2_USER.format(
                total_slides=len(pass1),
                summary=summary,
            ),
        }],
    )

    raw = msg.content[0].text.strip()
    try:
        overrides = json.loads(raw)
        if not isinstance(overrides, list):
            overrides = []
    except json.JSONDecodeError:
        m = re.search(r'\[.*\]', raw, re.DOTALL)
        try:
            overrides = json.loads(m.group(0)) if m else []
        except json.JSONDecodeError:
            overrides = []

    # Apply overrides
    results  = [dict(r) for r in pass1]
    idx_map  = {r["slide"]: i for i, r in enumerate(results)}
    for ov in overrides:
        sn = ov.get("slide")
        if sn not in idx_map:
            continue
        i        = idx_map[sn]
        new_type = ov.get("type", results[i]["type"])
        if new_type not in SLIDE_TYPES:
            continue
        if new_type != results[i]["type"]:
            results[i]["type"]       = new_type
            results[i]["confidence"] = float(ov.get("confidence", results[i]["confidence"]))
            results[i]["reasoning"] += f"  [P2: {ov.get('reasoning', '')}]"
            results[i]["changed"]    = True
            results[i]["source"]     = "vision_pass2"

    return results


# ── XML hint merge ────────────────────────────────────────────────────────────

def _merge_xml_hints(vision_results: list[dict], xml_hints: list[dict]) -> list[dict]:
    """For slides where vision confidence < threshold, fall back to XML heuristic
    if the XML result is confident."""
    xml_map = {r["slide"]: r for r in (xml_hints or [])}
    results = [dict(r) for r in vision_results]
    for r in results:
        if r["confidence"] >= CONFIDENCE_THRESHOLD:
            continue
        xml = xml_map.get(r["slide"])
        if xml and xml.get("confidence", 0) >= CONFIDENCE_THRESHOLD:
            r["reasoning"] += (
                f"  [XML fallback: {xml['type']} @ {xml['confidence']:.2f}]"
            )
            r["type"]       = xml["type"]
            r["confidence"] = xml["confidence"]
            r["source"]     = "xml_fallback"
    return results


# ── Public API ────────────────────────────────────────────────────────────────

def classify_deck_vision(
    pptx_path,
    xml_hints: Optional[list[dict]] = None,
    out_dir: Optional[Path] = None,
    keep_images: bool = False,
) -> list[dict]:
    """Full two-pass AI vision classification pipeline.

    Args:
        pptx_path:   Path to source .pptx file.
        xml_hints:   Optional list of XML-heuristic results from slide_classifier.py
                     (used as fallback for low-confidence vision calls).
        out_dir:     Where to write slide images. If None, uses a temp dir that is
                     cleaned up unless keep_images=True.
        keep_images: If True, slide PNGs are kept after classification.

    Returns:
        List of dicts: {slide, type, confidence, reasoning, changed, source}
    """
    if not _HAS_ANTHROPIC:
        raise ImportError("anthropic package required: pip install anthropic")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY environment variable not set.\n"
            "Set it with: export ANTHROPIC_API_KEY=sk-ant-..."
        )

    client   = anthropic.Anthropic(api_key=api_key)
    tmp_dir  = Path(out_dir) if out_dir else Path(tempfile.mkdtemp(prefix="fq_slides_"))
    owns_tmp = out_dir is None

    try:
        print(f"Rasterizing {Path(pptx_path).name} …")
        pngs  = rasterize_pptx(pptx_path, tmp_dir)
        total = len(pngs)
        print(f"  {total} slide(s) → {tmp_dir}")

        # ── Pass 1 ──────────────────────────────────────────────────────────
        print("Pass 1: classifying slides …")
        xml_map = {r["slide"]: r for r in (xml_hints or [])}
        pass1   = []
        for i, img_path in enumerate(pngs):
            sn     = i + 1
            result = _classify_slide_vision(
                client, img_path, sn, total, xml_hint=xml_map.get(sn)
            )
            flag = "  ⚑" if result["confidence"] < CONFIDENCE_THRESHOLD else ""
            print(
                f"  {sn:2d}/{total}  {result['type']:<10}"
                f"  {result['confidence']:.2f}{flag}  {result['reasoning'][:55]}"
            )
            pass1.append(result)

        # ── Pass 2 ──────────────────────────────────────────────────────────
        print("Pass 2: holistic reconciliation …")
        final   = _reconcile_vision(client, pass1)
        changed = [r for r in final if r["changed"]]
        if changed:
            print(f"  {len(changed)} slide(s) reclassified:")
            for r in changed:
                print(f"    Slide {r['slide']}: → {r['type']}  {r['reasoning'][-60:]}")
        else:
            print("  No changes.")

        # ── XML fallback for low-confidence slides ───────────────────────────
        if xml_hints:
            final = _merge_xml_hints(final, xml_hints)

    finally:
        if owns_tmp and not keep_images:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    return final


def print_vision_summary(results: list[dict]) -> None:
    w = 60
    print(f"\n{'─' * w}")
    print(f"  Vision Classification  ({len(results)} slides)")
    print(f"{'─' * w}")
    for r in results:
        flag = " ⚑" if r["confidence"] < CONFIDENCE_THRESHOLD else "  "
        chg  = " ↺" if r.get("changed") else "  "
        src  = f"[{r.get('source','?')[:12]}]"
        print(
            f"  {r['slide']:2d}  {r['type']:<10}  {r['confidence']:.2f}"
            f"{flag}{chg}  {src:<16}  {r['reasoning'][:42]}"
        )
    low = [r for r in results if r["confidence"] < CONFIDENCE_THRESHOLD]
    if low:
        print(
            f"\n  ⚑ {len(low)} slide(s) below {CONFIDENCE_THRESHOLD} — "
            "flag for manual review"
        )
    print(f"{'─' * w}\n")


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="AI vision slide classifier")
    parser.add_argument("pptx", help="Source .pptx file")
    parser.add_argument("--out-dir", help="Directory to write slide images")
    parser.add_argument("--keep-images", action="store_true",
                        help="Keep slide PNG files after classification")
    parser.add_argument("--xml-hints", help="JSON file from slide_classifier.py")
    args = parser.parse_args()

    xml_hints = None
    if args.xml_hints:
        xml_hints = json.loads(Path(args.xml_hints).read_text())

    out_dir = Path(args.out_dir) if args.out_dir else None
    results = classify_deck_vision(
        Path(args.pptx),
        xml_hints=xml_hints,
        out_dir=out_dir,
        keep_images=args.keep_images,
    )

    print_vision_summary(results)

    out_json = Path(args.pptx).with_suffix(".vision.json")
    out_json.write_text(json.dumps(results, indent=2))
    print(f"Results → {out_json}")
