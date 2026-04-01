"""
services/schema_to_decisions.py — Semantic Schema → SchemaDecision Translator
==============================================================================
Converts the output of the Semantic Ingest pipeline (semantic_schema.json)
into the list[SchemaDecision] format that renderer.py expects.

This is the only bridge between the semantic pipeline and the render pipeline.
It does NOT call any API, modify any PPTX, or touch the renderer internals.

CONTRACT:
    schema_to_decisions(schema: dict) -> list[dict]
    schema_to_classifier_stubs(schema: dict) -> list[dict]

SchemaDecision shape (must match schema_engine.decide_layout output exactly):
    {
        "slide_id":       int,
        "selected_layout": str,
        "decision_path":  list[str],
        "fallback_used":  bool,
        "warnings":       list[str],
    }

ClassifierResult stub shape (for validate_render — only fields it reads):
    {
        "slide_id":  int,
        "type":      str,
        "metadata":  {"text_density": str},
        "confidence": float,
        "warnings":  list[str],
    }
"""

from __future__ import annotations

import logging
from typing import Any

logger = logging.getLogger(__name__)

# ── Layout mapping ────────────────────────────────────────────────────────────
# Maps (component_type, render_strategy) → selected_layout.
# Layout names must exactly match those in layout_rules.json / master PPTX.

_FALLBACK_LAYOUT = "Content_Light"

_LAYOUT_MAP: dict[tuple[str, str], str] = {
    # Cover
    ("cover",            "standard"):          "Cover_Light",
    # Divider
    ("divider",          "standard"):          "Divider_Crystal",
    # Ending
    ("ending",           "standard"):          "Ending_PivotPurple",
    # Plain text content
    ("content_text",     "standard"):          "Content_Light",
    # Native table
    ("content_table",    "table_extract"):     "Content_Light",
    ("content_table",    "standard"):          "Content_Light",
    # Mixed (text + visuals)
    ("content_mixed",    "standard"):          "Content_Light",
    # Charts — pass through as content, no layout switch
    ("chart",            "image_passthrough"): "Content_Light",
    ("chart",            "standard"):          "Content_Light",
    # Screenshot table — attempt table layout; flag for review
    ("screenshot_table", "table_extract"):     "Content_Light",
    ("screenshot_table", "manual"):            "Content_Light",
    # Screenshot other — image passthrough, safe layout
    ("screenshot_other", "image_passthrough"): "Content_Light",
    ("screenshot_other", "manual"):            "Content_Light",
    # Empty
    ("empty",            "skip"):              "Content_Light",
}

# component_type → classifier "type" (what QA validator uses for its checks)
_PURPOSE_MAP: dict[str, str] = {
    "cover":            "cover",
    "divider":          "divider",
    "ending":           "ending",
    "content_text":     "content",
    "content_table":    "content",
    "content_mixed":    "content",
    "chart":            "content",
    "screenshot_table": "content",
    "screenshot_other": "content",
    "empty":            "content",
    "unknown":          "content",
}

# component_type → approximate text_density (used by QA overflow/content checks)
_DENSITY_MAP: dict[str, str] = {
    "cover":            "low",
    "divider":          "low",
    "ending":           "low",
    "content_text":     "medium",
    "content_table":    "medium",
    "content_mixed":    "medium",
    "chart":            "low",
    "screenshot_table": "low",   # actual text not extractable from image
    "screenshot_other": "low",
    "empty":            "low",
    "unknown":          "medium",
}


# ── Public API ────────────────────────────────────────────────────────────────

def schema_to_decisions(schema: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Convert a semantic_schema dict into a list of SchemaDecision dicts
    ready to be passed to renderer.render_deck().

    Args:
        schema: dict loaded from semantic_schema.json.
                Must contain a "slides" list.

    Returns:
        list[SchemaDecision], one per slide, in slide_id order.

    Raises:
        ValueError: if "slides" key is missing or empty.
    """
    slides = schema.get("slides")
    if not slides:
        raise ValueError("semantic_schema has no 'slides' list")

    decisions: list[dict[str, Any]] = []
    counts: dict[str, int] = {}

    for slide in sorted(slides, key=lambda s: s.get("slide_id", 0)):
        decision = _slide_to_decision(slide)
        decisions.append(decision)

        ct = slide.get("component_type", "unknown")
        counts[ct] = counts.get(ct, 0) + 1

    # Summary log — one line, not per-slide
    logger.info(
        "schema_to_decisions: %d slides — %s",
        len(decisions),
        ", ".join(f"{k}×{v}" for k, v in sorted(counts.items())),
    )

    rs_counts: dict[str, int] = {}
    for s in slides:
        rs = s.get("render_strategy", "unknown")
        rs_counts[rs] = rs_counts.get(rs, 0) + 1
    logger.info(
        "render_strategy breakdown: %s",
        ", ".join(f"{k}×{v}" for k, v in sorted(rs_counts.items())),
    )

    return decisions


def schema_to_classifier_stubs(schema: dict[str, Any]) -> list[dict[str, Any]]:
    """
    Produce minimal ClassifierResult stubs from the semantic schema.
    Used to satisfy validate_render()'s classifier_results argument.

    Only populates the fields that qa_validator.py actually reads:
        slide_id, type, metadata.text_density, confidence, warnings
    """
    slides = schema.get("slides", [])
    stubs = []
    for slide in sorted(slides, key=lambda s: s.get("slide_id", 0)):
        ct = slide.get("component_type", "unknown")
        stubs.append({
            "slide_id":  slide.get("slide_id"),
            "type":      _PURPOSE_MAP.get(ct, "content"),
            "metadata":  {
                "text_density": _DENSITY_MAP.get(ct, "medium"),
                "bg_is_dark":   False,
                "has_subtitle": False,
                "has_chart":    ct == "chart",
                "image_count":  1 if ct in ("screenshot_table", "screenshot_other") else 0,
            },
            "confidence": slide.get("confidence", 0.5),
            "warnings":   list(slide.get("warnings", [])),
        })
    return stubs


def validate_schema_vs_pptx(schema: dict[str, Any], pptx_slide_count: int) -> None:
    """
    Raise ValueError if schema slide count does not match PPTX slide count.
    Call this before schema_to_decisions to get a clear early error.
    """
    schema_count = len(schema.get("slides", []))
    if schema_count != pptx_slide_count:
        raise ValueError(
            f"Slide count mismatch: semantic_schema has {schema_count} slides "
            f"but the PPTX has {pptx_slide_count}. "
            f"Ensure the schema was generated from the same deck."
        )


# ── Per-slide translation ─────────────────────────────────────────────────────

def _slide_to_decision(slide: dict[str, Any]) -> dict[str, Any]:
    """Translate one semantic slide dict into one SchemaDecision dict."""
    slide_id        = slide.get("slide_id")
    component_type  = slide.get("component_type", "unknown")
    render_strategy = slide.get("render_strategy", "standard")
    recoverability  = slide.get("recoverability", "auto")
    warnings: list[str] = list(slide.get("warnings", []))

    # Look up layout — try (component_type, render_strategy) first,
    # then (component_type, "standard") as a secondary fallback.
    layout = (
        _LAYOUT_MAP.get((component_type, render_strategy))
        or _LAYOUT_MAP.get((component_type, "standard"))
    )
    fallback_used = layout is None
    if fallback_used:
        layout = _FALLBACK_LAYOUT
        warnings.append(f"no_layout_mapping:component_type={component_type}")
        logger.warning(
            "Slide %s: no layout mapping for (%s, %s) — using fallback %s",
            slide_id, component_type, render_strategy, _FALLBACK_LAYOUT,
        )

    # Build decision_path for traceability
    decision_path = [
        f"semantic:component_type={component_type}",
        f"semantic:render_strategy={render_strategy}",
        f"semantic:recoverability={recoverability}",
        f"→ {layout}" + (" [fallback]" if fallback_used else ""),
    ]

    # Propagate low-confidence as a warning
    if slide.get("confidence", 1.0) < 0.5:
        warnings.append(f"low_semantic_confidence:{slide.get('confidence', 0):.2f}")

    # Flag manual-review slides
    if recoverability in ("manual", "review"):
        warnings.append(f"recoverability:{recoverability}")

    return {
        "slide_id":        slide_id,
        "selected_layout": layout,
        "decision_path":   decision_path,
        "fallback_used":   fallback_used,
        "warnings":        warnings,
    }
