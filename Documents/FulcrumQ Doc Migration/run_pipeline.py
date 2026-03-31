"""
run_pipeline.py  —  Top-level Orchestrator
===========================================
FulcrumQ deck conversion pipeline.

Runs all four layers in sequence and produces all four artifacts:
  1. classifier_output.json
  2. schema_decisions.json
  3. render_log.json  +  output PPTX
  4. qa_report.json

Usage:
    python run_pipeline.py source.pptx master.pptx [output_dir]

All JSON artifacts are written to output_dir (default: same dir as source).
Fails fast if any layer produces a FAIL status in QA.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from classifier    import classify_deck
from schema_engine import run_schema_engine
from renderer      import render_deck
from qa_validator  import validate_render

LAYOUT_RULES = Path(__file__).parent / "layout_rules.json"


def run_pipeline(
    source_pptx: Path,
    master_pptx: Path,
    output_dir: Path,
) -> dict:
    """
    Run the full 4-layer pipeline.

    Returns the QA report dict. Raises RuntimeError if QA reports any FAIL slides
    or global issues that indicate a broken conversion.
    """
    source_pptx = Path(source_pptx)
    master_pptx = Path(master_pptx)
    output_dir  = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = source_pptx.stem

    # ── Artifact paths ────────────────────────────────────────────────────────
    classifier_json  = output_dir / f"{stem}.classifier_output.json"
    decisions_json   = output_dir / f"{stem}.schema_decisions.json"
    render_log_json  = output_dir / f"{stem}.render_log.json"
    output_pptx      = output_dir / f"{stem}_FQ.pptx"
    qa_json          = output_dir / f"{stem}.qa_report.json"

    # ── Layer 1: Classify ─────────────────────────────────────────────────────
    print(f"[1/4] Classifying {source_pptx.name} …")
    classifier_results = classify_deck(source_pptx, output_path=classifier_json)
    _report_layer("Classifier", classifier_results, "type")
    print(f"      → {classifier_json.name}")

    # ── Layer 2: Schema decisions ─────────────────────────────────────────────
    print(f"[2/4] Running schema engine …")
    schema_decisions = run_schema_engine(
        classifier_results,
        rules_path=LAYOUT_RULES,
        output_path=decisions_json,
    )
    _report_layer("Schema", schema_decisions, "selected_layout")
    print(f"      → {decisions_json.name}")

    # ── Layer 3: Render ───────────────────────────────────────────────────────
    print(f"[3/4] Rendering …")
    render_log = render_deck(
        source_pptx,
        schema_decisions,
        master_pptx,
        output_pptx,
        output_log=render_log_json,
    )
    warn_count = sum(1 for r in render_log if r["warnings"])
    print(f"      {len(render_log)} slides rendered, {warn_count} with warnings")
    print(f"      → {output_pptx.name}")
    print(f"      → {render_log_json.name}")

    # ── Layer 4: QA ───────────────────────────────────────────────────────────
    print(f"[4/4] Running QA …")
    qa_report = validate_render(
        output_pptx,
        render_log,
        classifier_results,
        output_path=qa_json,
    )
    s = qa_report["summary"]
    print(f"      QA: {s['total']} slides — {s['pass']} pass / {s['review']} review / {s['fail']} fail")
    print(f"      → {qa_json.name}")

    if qa_report["global_issues"]:
        print(f"  ⚑  Global issues: {', '.join(qa_report['global_issues'])}")

    for slide in qa_report["slides"]:
        if slide["status"] != "pass":
            mark = "✗" if slide["status"] == "fail" else "△"
            print(f"  {mark} slide {slide['slide_id']:2d}  {slide['status'].upper():<6}  {'; '.join(slide['issues'])}")

    fail_count = s["fail"]
    if fail_count > 0:
        raise RuntimeError(
            f"Pipeline completed with {fail_count} FAIL slide(s). "
            f"Check {qa_json} for details."
        )

    return qa_report


def _report_layer(label: str, results: list[dict], key: str) -> None:
    counts: dict[str, int] = {}
    for r in results:
        v = r.get(key, "?")
        counts[v] = counts.get(v, 0) + 1
    summary = ", ".join(f"{v}×{n}" for v, n in sorted(counts.items()))
    print(f"      {label}: {summary}")


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python run_pipeline.py source.pptx master.pptx [output_dir]")
        sys.exit(1)

    source  = Path(sys.argv[1])
    master  = Path(sys.argv[2])
    out_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else source.parent / "pipeline_output"

    try:
        report = run_pipeline(source, master, out_dir)
        print("\nPipeline complete.")
    except RuntimeError as e:
        print(f"\n✗ {e}")
        sys.exit(2)
