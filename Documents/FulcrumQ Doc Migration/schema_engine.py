"""
schema_engine.py  —  Layer 2: Layout Decision Engine
======================================================
FulcrumQ deck conversion pipeline.

Reads classifier_output.json + layout_rules.json.
Produces per-slide layout decisions (schema_decisions.json).

CONTRACT:
  Input:  list[ClassifierResult]  (from classifier.py)
          layout_rules.json       (config file)
  Output: list[SchemaDecision]   + optional schema_decisions.json

LAYER RULES — this module MUST NOT:
  - Touch any PPTX XML
  - Apply fonts, colors, or shapes
  - Read from the source PPTX directly
  - Produce output PPTX
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Optional


# ── Rule evaluation ────────────────────────────────────────────────────────────

def _eval_condition(condition: dict, meta: dict) -> bool:
    """
    Evaluate one rule's conditions dict against a merged metadata dict.

    Supported condition keys:
      type           — exact string match to slide type
      bg_is_dark     — bool match
      has_subtitle   — bool match
      has_chart      — bool match
      text_density   — exact string match ("low" | "medium" | "high")
      image_count_gt — int: metadata["image_count"] > value
    """
    slide_type = meta.get("type", "")
    metadata   = meta.get("metadata", {})

    for key, expected in condition.items():
        if key == "type":
            if slide_type != expected:
                return False
        elif key == "bg_is_dark":
            if bool(metadata.get("bg_is_dark", False)) != bool(expected):
                return False
        elif key == "has_subtitle":
            if bool(metadata.get("has_subtitle", False)) != bool(expected):
                return False
        elif key == "has_chart":
            if bool(metadata.get("has_chart", False)) != bool(expected):
                return False
        elif key == "text_density":
            if metadata.get("text_density", "") != expected:
                return False
        elif key == "image_count_gt":
            if int(metadata.get("image_count", 0)) <= int(expected):
                return False
        else:
            # Unknown condition key — skip (forward compatibility)
            pass
    return True


def _select_layout(classifier_result: dict, rules: list[dict], fallback: str) -> tuple[str, list[str], bool]:
    """
    Walk rules (highest priority first) and return (layout, decision_path, fallback_used).
    """
    for rule in sorted(rules, key=lambda r: r["priority"], reverse=True):
        if _eval_condition(rule["conditions"], classifier_result):
            return rule["layout"], rule["decision_path"], False

    return fallback, [f"no rule matched → fallback={fallback}"], True


def _apply_subtitle_upgrade(layout: str, classifier_result: dict, upgrade_map: dict) -> tuple[str, bool]:
    """
    If the slide has a subtitle and the selected layout has an upgrade available, apply it.
    Returns (upgraded_layout, was_upgraded).
    """
    has_subtitle = classifier_result.get("metadata", {}).get("has_subtitle", False)
    if not has_subtitle:
        return layout, False
    upgraded = upgrade_map.get(layout)
    if upgraded:
        return upgraded, True
    return layout, False


# ── Public API ─────────────────────────────────────────────────────────────────

def decide_layout(
    classifier_result: dict,
    rules: list[dict],
    fallback_layout: str,
    subtitle_upgrade_map: dict,
) -> dict:
    """
    Produce one SchemaDecision from one ClassifierResult.

    Returns:
      {
        "slide_id":       int,
        "selected_layout": str,
        "decision_path":  list[str],
        "fallback_used":  bool,
        "warnings":       list[str],
      }
    """
    slide_id = classifier_result["slide_id"]
    warnings: list[str] = list(classifier_result.get("warnings", []))

    layout, decision_path, fallback_used = _select_layout(
        classifier_result, rules, fallback_layout
    )

    # Subtitle upgrade pass (only for content-family layouts)
    upgraded_layout, was_upgraded = _apply_subtitle_upgrade(
        layout, classifier_result, subtitle_upgrade_map
    )
    if was_upgraded:
        decision_path = decision_path + [f"subtitle_upgrade → {upgraded_layout}"]
        layout = upgraded_layout

    if fallback_used:
        warnings.append("layout_fallback_used")

    # Propagate low-confidence from classifier
    if classifier_result.get("confidence", 1.0) < 0.75:
        warnings.append(f"low_classifier_confidence:{classifier_result['confidence']:.2f}")

    return {
        "slide_id":        slide_id,
        "selected_layout": layout,
        "decision_path":   decision_path,
        "fallback_used":   fallback_used,
        "warnings":        warnings,
    }


def run_schema_engine(
    classifier_results: list[dict],
    rules_path: Path,
    output_path: Optional[Path] = None,
) -> list[dict]:
    """
    Run Layer 2 over all classifier results.

    Args:
      classifier_results: list[ClassifierResult] from classify_deck()
      rules_path:         Path to layout_rules.json
      output_path:        Optional path to write schema_decisions.json

    Returns list[SchemaDecision].
    """
    rules_path = Path(rules_path)
    config = json.loads(rules_path.read_text())

    rules              = config["layout_rules"]
    fallback_layout    = config["fallback_layout"]
    subtitle_upgrade   = config.get("subtitle_upgrade_map", {})

    decisions = [
        decide_layout(r, rules, fallback_layout, subtitle_upgrade)
        for r in classifier_results
    ]

    if output_path is not None:
        Path(output_path).write_text(json.dumps(decisions, indent=2))

    return decisions


# ── CLI ────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Usage: python schema_engine.py classifier_output.json layout_rules.json [decisions.json]")
        sys.exit(1)

    classifier_results = json.loads(Path(sys.argv[1]).read_text())
    rules_path         = Path(sys.argv[2])
    out = Path(sys.argv[3]) if len(sys.argv) > 3 else Path(sys.argv[1]).with_suffix(".decisions.json")

    decisions = run_schema_engine(classifier_results, rules_path, output_path=out)
    for d in decisions:
        flag    = " ⚑" if d["warnings"] else ""
        fb_mark = " [FALLBACK]" if d["fallback_used"] else ""
        print(f"  slide {d['slide_id']:2d}  [{d['selected_layout']:<22}]{fb_mark}{flag}")
        if d["warnings"]:
            print(f"           warnings: {', '.join(d['warnings'])}")
    print(f"\nWrote → {out}")
