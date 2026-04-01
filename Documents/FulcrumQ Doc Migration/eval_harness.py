"""
eval_harness.py  —  Evaluation Harness
=======================================
Runs the FulcrumQ 4-layer pipeline on a batch of source decks,
rasterizes before/after slides, and produces a per-deck scorecard
plus a browseable HTML report.

Usage:
    python eval_harness.py \\
        --source-dir ./ \\
        --master FQ_PPTX_Theme_vF.pptx \\
        --output-dir eval_output/

    # Or point at explicit decks:
    python eval_harness.py \\
        --decks "deck1.pptx" "deck2.pptx" \\
        --master FQ_PPTX_Theme_vF.pptx \\
        --output-dir eval_output/

Output per deck (eval_output/{deck_stem}/):
    classifier_output.json
    schema_decisions.json
    render_log.json
    qa_report.json
    {deck_stem}_FQ.pptx
    before/slide_NNN.png
    after/slide_NNN.png

Summary:
    eval_output/scorecard.json
    eval_output/report.html
"""

from __future__ import annotations

import argparse
import base64
import json
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Optional

# ── Pipeline imports ───────────────────────────────────────────────────────────
REPO = Path(__file__).parent
sys.path.insert(0, str(REPO))

from classifier    import classify_deck
from schema_engine import run_schema_engine
from renderer      import render_deck
from qa_validator  import validate_render

LAYOUT_RULES = REPO / "layout_rules.json"

# ── Layer attribution ─────────────────────────────────────────────────────────
# Maps issue-prefix → responsible layer
_LAYER_PREFIXES = {
    "L1_classifier": {
        "low_confidence", "slide_1_not_cover",
    },
    "L2_schema": {
        "layout_fallback_used", "no_schema_decision",
        "low_classifier_confidence",
    },
    "L3_renderer": {
        "overflow_risk", "missing_placeholder", "layout_not_found",
        "fallback_layout_also_missing",
    },
    "L4_qa": {
        "content_lost", "title_placeholder_missing", "title_empty",
        "layout_mismatch", "non_fq_fonts", "layout_rel_missing",
        "unknown_layout",
    },
    "global": {
        "slide_count_mismatch", "slide_id_mismatch", "slide_1_not_cover_layout",
    },
}

def _issue_to_layer(issue: str) -> str:
    for layer, prefixes in _LAYER_PREFIXES.items():
        if any(issue.startswith(p) for p in prefixes):
            return layer
    return "L4_qa"  # default: treat as output quality problem


def _root_cause_layer(layer_issues: dict[str, list]) -> str:
    """Return the layer with the most issues (weighted: L4 > L3 > L2 > L1)."""
    weights = {"L4_qa": 4, "L3_renderer": 3, "L2_schema": 2, "L1_classifier": 1, "global": 5}
    scored = {
        layer: len(issues) * weights.get(layer, 1)
        for layer, issues in layer_issues.items()
        if issues
    }
    if not scored:
        return "none"
    return max(scored, key=scored.get)


# ── Rasterization ─────────────────────────────────────────────────────────────

_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def _pptx_to_pngs(pptx_path: Path, out_dir: Path, dpi: int = 100) -> list[Path]:
    """
    Convert a PPTX to per-slide PNGs using LibreOffice → PDF → fitz.
    Returns list of PNG paths in slide order. Returns [] if unavailable.
    """
    if _SOFFICE is None:
        return []

    try:
        import fitz  # pymupdf
    except ImportError:
        return []

    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        # Convert PPTX → PDF via LibreOffice
        result = subprocess.run(
            [_SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(tmp), str(pptx_path)],
            capture_output=True, timeout=120,
        )
        if result.returncode != 0:
            return []

        pdf_files = list(tmp.glob("*.pdf"))
        if not pdf_files:
            return []

        pdf_path = pdf_files[0]
        doc = fitz.open(str(pdf_path))
        fitz.TOOLS.mupdf_warnings()  # flush/discard accumulated warnings
        pngs: list[Path] = []
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        for i, page in enumerate(doc):
            png_path = out_dir / f"slide_{i+1:03d}.png"
            pix = page.get_pixmap(matrix=mat)
            pix.save(str(png_path))
            pngs.append(png_path)

        doc.close()

    return pngs


# ── Per-deck evaluation ────────────────────────────────────────────────────────

def eval_deck(
    source_pptx: Path,
    master_pptx: Path,
    deck_out: Path,
    rasterize: bool = True,
) -> dict:
    """
    Run the full pipeline on one deck and return its scorecard dict.
    Never raises — errors are captured and reflected in the scorecard.
    """
    source_pptx = Path(source_pptx)
    deck_out.mkdir(parents=True, exist_ok=True)
    stem = source_pptx.stem

    scorecard: dict = {
        "deck":             stem,
        "source_path":      str(source_pptx),
        "status":           "error",
        "root_cause_layer": "unknown",
        "slide_count":      0,
        "qa_summary":       {"pass": 0, "review": 0, "fail": 0},
        "layer_issues": {
            "L1_classifier": [],
            "L2_schema":     [],
            "L3_renderer":   [],
            "L4_qa":         [],
            "global":        [],
        },
        "artifacts": {},
        "error":            None,
        "before_pngs":      [],
        "after_pngs":       [],
    }

    try:
        # ── Artifact paths ────────────────────────────────────────────────────
        cls_json    = deck_out / "classifier_output.json"
        dec_json    = deck_out / "schema_decisions.json"
        ren_json    = deck_out / "render_log.json"
        qa_json     = deck_out / "qa_report.json"
        out_pptx    = deck_out / f"{stem}_FQ.pptx"

        scorecard["artifacts"] = {
            "classifier_output": str(cls_json),
            "schema_decisions":  str(dec_json),
            "render_log":        str(ren_json),
            "qa_report":         str(qa_json),
            "output_pptx":       str(out_pptx),
        }

        # ── Layer 1 ───────────────────────────────────────────────────────────
        cls_results = classify_deck(source_pptx, output_path=cls_json)
        scorecard["slide_count"] = len(cls_results)

        # Collect L1 issues
        for r in cls_results:
            for w in r.get("warnings", []):
                scorecard["layer_issues"]["L1_classifier"].append(
                    f"slide {r['slide_id']}: {w}"
                )

        # ── Layer 2 ───────────────────────────────────────────────────────────
        decisions = run_schema_engine(cls_results, LAYOUT_RULES, output_path=dec_json)

        for d in decisions:
            for w in d.get("warnings", []):
                layer = _issue_to_layer(w)
                scorecard["layer_issues"][layer].append(
                    f"slide {d['slide_id']}: {w}"
                )

        # ── Layer 3 ───────────────────────────────────────────────────────────
        ren_log = render_deck(source_pptx, decisions, master_pptx, out_pptx, output_log=ren_json)

        for r in ren_log:
            for w in r.get("warnings", []):
                layer = _issue_to_layer(w)
                scorecard["layer_issues"][layer].append(
                    f"slide {r['slide_id']}: {w}"
                )

        # ── Layer 4 ───────────────────────────────────────────────────────────
        qa_report = validate_render(out_pptx, ren_log, cls_results, output_path=qa_json)
        scorecard["qa_summary"] = qa_report["summary"]

        for issue in qa_report.get("global_issues", []):
            scorecard["layer_issues"]["global"].append(issue)

        for slide in qa_report["slides"]:
            if slide["status"] != "pass":
                for iss in slide["issues"]:
                    layer = _issue_to_layer(iss)
                    scorecard["layer_issues"][layer].append(
                        f"slide {slide['slide_id']}: {iss}"
                    )

        # ── Scorecard status ──────────────────────────────────────────────────
        s = qa_report["summary"]
        if s["fail"] > 0 or qa_report["global_issues"]:
            scorecard["status"] = "fail"
        elif s["review"] > 0:
            scorecard["status"] = "review"
        else:
            scorecard["status"] = "pass"

        scorecard["root_cause_layer"] = _root_cause_layer(scorecard["layer_issues"])

        # ── Rasterize ─────────────────────────────────────────────────────────
        if rasterize:
            before_dir = deck_out / "before"
            after_dir  = deck_out / "after"
            scorecard["before_pngs"] = [str(p) for p in _pptx_to_pngs(source_pptx, before_dir)]
            if out_pptx.exists():
                scorecard["after_pngs"]  = [str(p) for p in _pptx_to_pngs(out_pptx,  after_dir)]

    except Exception as e:
        scorecard["error"] = traceback.format_exc()
        scorecard["status"] = "error"

    return scorecard


# ── HTML report ───────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "pass":   "#2ecc71",
    "review": "#f39c12",
    "fail":   "#e74c3c",
    "error":  "#8e44ad",
}

_LAYER_LABEL = {
    "L1_classifier": "L1 Classifier",
    "L2_schema":     "L2 Schema",
    "L3_renderer":   "L3 Renderer",
    "L4_qa":         "L4 QA",
    "global":        "Global",
    "none":          "—",
    "unknown":       "Unknown",
}


def _b64_img(path: str) -> str:
    try:
        data = Path(path).read_bytes()
        return "data:image/png;base64," + base64.b64encode(data).decode()
    except Exception:
        return ""


def _issue_list_html(issues: list[str], max_show: int = 6) -> str:
    if not issues:
        return "<em>none</em>"
    shown = issues[:max_show]
    rest = len(issues) - max_show
    html = "<ul class='issues'>" + "".join(f"<li>{i}</li>" for i in shown) + "</ul>"
    if rest > 0:
        html += f"<p class='more'>…and {rest} more</p>"
    return html


def _slide_strip_html(before_pngs: list[str], after_pngs: list[str], max_slides: int = 6) -> str:
    """Side-by-side before/after thumbnails for up to max_slides."""
    if not before_pngs and not after_pngs:
        return "<p><em>No slide images (LibreOffice not available or rasterization skipped)</em></p>"

    n = min(max_slides, max(len(before_pngs), len(after_pngs)))
    rows = []
    for i in range(n):
        b_src = _b64_img(before_pngs[i]) if i < len(before_pngs) else ""
        a_src = _b64_img(after_pngs[i])  if i < len(after_pngs)  else ""
        b_img = f'<img src="{b_src}" loading="lazy">' if b_src else "<div class=\'no-img\'>no image</div>"
        a_img = f'<img src="{a_src}" loading="lazy">' if a_src else "<div class=\'no-img\'>no image</div>"
        rows.append(
            f"<div class='slide-pair'>"
            f"<div class='label'>Before — slide {i+1}</div>{b_img}"
            f"<div class='label'>After</div>{a_img}"
            f"</div>"
        )
    total = max(len(before_pngs), len(after_pngs))
    more  = total - n
    strip = "".join(rows)
    if more > 0:
        strip += f"<p class='more'>…{more} more slides not shown</p>"
    return strip


def build_html_report(scorecards: list[dict], output_path: Path) -> None:
    rows = []
    for sc in scorecards:
        color = _STATUS_COLOR.get(sc["status"], "#999")
        qs = sc["qa_summary"]
        rc = _LAYER_LABEL.get(sc["root_cause_layer"], sc["root_cause_layer"])
        issue_count = sum(len(v) for v in sc["layer_issues"].values())
        deck_name = sc["deck"]
        rows.append(
            f"<tr style='cursor:pointer'>"
            f"<td><a href='#{deck_name}'>{deck_name}</a></td>"
            f"<td style='color:{color};font-weight:bold'>{sc['status'].upper()}</td>"
            f"<td>{sc['slide_count']}</td>"
            f"<td>{qs['pass']}</td><td>{qs['review']}</td><td>{qs['fail']}</td>"
            f"<td>{rc}</td><td>{issue_count}</td>"
            f"</tr>"
        )

    deck_sections = []
    for sc in scorecards:
        color = _STATUS_COLOR.get(sc["status"], "#999")
        layer_html = ""
        for layer_key in ("L1_classifier", "L2_schema", "L3_renderer", "L4_qa", "global"):
            issues = sc["layer_issues"].get(layer_key, [])
            label  = _LAYER_LABEL[layer_key]
            layer_html += (
                f"<div class='layer-block'>"
                f"<h4>{label} <span class='badge'>{len(issues)}</span></h4>"
                f"{_issue_list_html(issues)}"
                f"</div>"
            )

        slide_strip = _slide_strip_html(sc.get("before_pngs", []), sc.get("after_pngs", []))
        err_block = ""
        if sc.get("error"):
            err_block = f"<pre class='error-block'>{sc['error']}</pre>"

        deck_sections.append(f"""
<section id="{sc['deck']}" class="deck-section">
  <h2><span class="status-dot" style="background:{color}"></span>{sc['deck']}</h2>
  <p class="meta">Source: <code>{sc['source_path']}</code> &nbsp;|&nbsp; Slides: {sc['slide_count']}
     &nbsp;|&nbsp; Status: <strong style="color:{color}">{sc['status'].upper()}</strong>
     &nbsp;|&nbsp; Root cause: <strong>{_LAYER_LABEL.get(sc['root_cause_layer'], sc['root_cause_layer'])}</strong>
  </p>
  {err_block}
  <div class="layers-grid">{layer_html}</div>
  <div class="slide-strip">{slide_strip}</div>
</section>""")

    total = len(scorecards)
    n_pass   = sum(1 for s in scorecards if s["status"] == "pass")
    n_review = sum(1 for s in scorecards if s["status"] == "review")
    n_fail   = sum(1 for s in scorecards if s["status"] in ("fail", "error"))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>FulcrumQ Pipeline Eval Report</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
          background: #f5f5f5; color: #222; line-height: 1.5; }}
  header {{ background: #1a1a2e; color: #fff; padding: 24px 32px; }}
  header h1 {{ font-size: 1.6rem; }}
  header p {{ opacity: 0.7; font-size: 0.9rem; margin-top: 4px; }}
  .summary-bar {{ display: flex; gap: 24px; padding: 16px 32px;
                  background: #fff; border-bottom: 1px solid #ddd; }}
  .stat {{ text-align: center; }}
  .stat .n {{ font-size: 2rem; font-weight: 700; }}
  .stat .label {{ font-size: 0.75rem; text-transform: uppercase; opacity: 0.6; }}
  .pass-n {{ color: #2ecc71; }} .review-n {{ color: #f39c12; }}
  .fail-n {{ color: #e74c3c; }}
  main {{ max-width: 1200px; margin: 32px auto; padding: 0 16px; }}
  table {{ width: 100%; border-collapse: collapse; background: #fff;
           border-radius: 8px; overflow: hidden;
           box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 40px; }}
  th {{ background: #1a1a2e; color: #fff; padding: 10px 14px; text-align: left;
        font-size: 0.8rem; text-transform: uppercase; }}
  td {{ padding: 10px 14px; border-bottom: 1px solid #eee; font-size: 0.9rem; }}
  tr:hover {{ background: #f0f4ff; }}
  .deck-section {{ background: #fff; border-radius: 8px; padding: 24px;
                   margin-bottom: 32px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .deck-section h2 {{ font-size: 1.1rem; display: flex; align-items: center; gap: 10px;
                      margin-bottom: 8px; }}
  .status-dot {{ width: 12px; height: 12px; border-radius: 50%; flex-shrink: 0; }}
  .meta {{ font-size: 0.82rem; opacity: 0.7; margin-bottom: 16px; }}
  .layers-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
                  gap: 12px; margin-bottom: 20px; }}
  .layer-block {{ background: #f9f9f9; border-radius: 6px; padding: 12px; }}
  .layer-block h4 {{ font-size: 0.82rem; text-transform: uppercase; color: #555;
                     margin-bottom: 6px; }}
  .badge {{ background: #e0e0e0; border-radius: 10px; padding: 1px 7px;
            font-size: 0.75rem; font-weight: 700; }}
  ul.issues {{ padding-left: 16px; font-size: 0.8rem; }}
  ul.issues li {{ margin-bottom: 2px; word-break: break-word; }}
  .more {{ font-size: 0.75rem; opacity: 0.6; margin-top: 4px; }}
  .slide-strip {{ display: flex; gap: 12px; flex-wrap: wrap; margin-top: 8px; }}
  .slide-pair {{ display: flex; flex-direction: column; gap: 4px; }}
  .slide-pair img {{ width: 220px; border: 1px solid #ddd; border-radius: 4px; }}
  .label {{ font-size: 0.7rem; opacity: 0.6; }}
  .no-img {{ width: 220px; height: 124px; background: #eee; border-radius: 4px;
             display: flex; align-items: center; justify-content: center;
             font-size: 0.75rem; color: #aaa; }}
  .error-block {{ background: #fff5f5; border: 1px solid #fcc; border-radius: 6px;
                  padding: 12px; font-size: 0.78rem; white-space: pre-wrap;
                  margin-bottom: 16px; max-height: 200px; overflow: auto; }}
  a {{ color: #3a6ff7; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
</style>
</head>
<body>
<header>
  <h1>FulcrumQ Pipeline Evaluation Report</h1>
  <p>{total} decks evaluated</p>
</header>
<div class="summary-bar">
  <div class="stat"><div class="n pass-n">{n_pass}</div><div class="label">Pass</div></div>
  <div class="stat"><div class="n review-n">{n_review}</div><div class="label">Review</div></div>
  <div class="stat"><div class="n fail-n">{n_fail}</div><div class="label">Fail / Error</div></div>
  <div class="stat"><div class="n">{total}</div><div class="label">Total</div></div>
</div>
<main>
<h2 style="margin-bottom:12px;font-size:1rem;opacity:.7">Summary</h2>
<table>
<thead>
  <tr>
    <th>Deck</th><th>Status</th><th>Slides</th>
    <th>Pass</th><th>Review</th><th>Fail</th>
    <th>Root Cause</th><th>Issues</th>
  </tr>
</thead>
<tbody>{''.join(rows)}</tbody>
</table>

{''.join(deck_sections)}
</main>
</body>
</html>"""

    output_path.write_text(html, encoding="utf-8")


# ── Source deck discovery ─────────────────────────────────────────────────────

_SKIP_PATTERNS = re.compile(
    r"(^~\$|_fq\.|FulcrumQ|FQ_|LatestMaster|icons|Brand|Theme|Template|Canonical|Default)",
    re.IGNORECASE,
)


def _discover_decks(source_dir: Path) -> list[Path]:
    """Find source PPTX files in source_dir that look like client decks."""
    decks = []
    for p in sorted(source_dir.glob("*.pptx")):
        if _SKIP_PATTERNS.search(p.name):
            continue
        decks.append(p)
    return decks


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="FulcrumQ pipeline evaluation harness")
    parser.add_argument("--master", required=True, help="FulcrumQ master PPTX")
    parser.add_argument("--output-dir", required=True, help="Output directory")
    parser.add_argument("--source-dir", default=None, help="Scan this dir for source decks")
    parser.add_argument("--decks", nargs="*", default=[], help="Explicit deck paths")
    parser.add_argument("--no-rasterize", action="store_true", help="Skip slide image generation")
    parser.add_argument("--max-slides-shown", type=int, default=6,
                        help="Max before/after slides shown per deck in HTML (default 6)")
    args = parser.parse_args()

    master  = Path(args.master)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Collect decks
    decks: list[Path] = [Path(d) for d in args.decks]
    if args.source_dir:
        decks += _discover_decks(Path(args.source_dir))
    decks = list(dict.fromkeys(decks))  # deduplicate, preserve order

    if not decks:
        print("No decks found. Use --decks or --source-dir.")
        sys.exit(1)

    if not master.exists():
        print(f"Master PPTX not found: {master}")
        sys.exit(1)

    rasterize = not args.no_rasterize
    if rasterize and _SOFFICE is None:
        print("Warning: LibreOffice not found — skipping rasterization.")
        rasterize = False

    scorecards: list[dict] = []
    print(f"Evaluating {len(decks)} deck(s) → {out_dir}\n")

    for i, deck_path in enumerate(decks, 1):
        print(f"[{i}/{len(decks)}] {deck_path.name}")
        deck_out = out_dir / deck_path.stem
        sc = eval_deck(deck_path, master, deck_out, rasterize=rasterize)
        scorecards.append(sc)

        color_map = {"pass": "\033[32m", "review": "\033[33m",
                     "fail": "\033[31m", "error": "\033[35m"}
        reset = "\033[0m"
        c = color_map.get(sc["status"], "")
        qa = sc["qa_summary"]
        print(
            f"  {c}{sc['status'].upper()}{reset}  "
            f"{qa['pass']}✓ {qa['review']}△ {qa['fail']}✗  "
            f"root→{sc['root_cause_layer']}"
        )
        if sc.get("error"):
            print(f"  ERROR: {sc['error'].splitlines()[-1]}")

    # Write scorecard.json
    sc_path = out_dir / "scorecard.json"
    sc_path.write_text(json.dumps(scorecards, indent=2))

    # Write report.html
    report_path = out_dir / "report.html"
    build_html_report(scorecards, report_path)

    # Summary
    n_pass   = sum(1 for s in scorecards if s["status"] == "pass")
    n_review = sum(1 for s in scorecards if s["status"] == "review")
    n_fail   = sum(1 for s in scorecards if s["status"] in ("fail", "error"))
    print(f"\n{'─'*50}")
    print(f"  {len(decks)} decks: {n_pass} pass / {n_review} review / {n_fail} fail")
    print(f"  Scorecard → {sc_path}")
    print(f"  Report    → {report_path}")


if __name__ == "__main__":
    main()
