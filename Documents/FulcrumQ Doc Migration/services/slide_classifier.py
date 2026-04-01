"""
services/slide_classifier.py — Rule-Based Slide Router
=======================================================
First-pass, zero-cost classifier that decides whether a slide
needs an AgentRelay API call or can be resolved locally.

CONTRACT:
  classify_slide(slide, slide_idx, total_slides) -> RouteDecision

RouteDecision fields:
  route              "local" | "api" | "fallback"
  component_type     specific slide component label
  local_purpose      detected_purpose for local route (cover/content/divider/ending/unknown)
  confidence         0.0–1.0 heuristic confidence
  reason             human-readable explanation of the routing decision
  render_strategy    "standard" | "table_extract" | "image_passthrough" | "skip" | "manual"
  recoverability     "auto" | "review" | "manual"

RULES:
  - No I/O, no network, no imports from other services modules
  - Pure deterministic functions only
  - Never raises — always returns a RouteDecision
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# ── Data class ────────────────────────────────────────────────────────────────

@dataclass
class RouteDecision:
    route:           str          # "local" | "api" | "fallback"
    component_type:  str          # see _COMPONENT_TYPES below
    local_purpose:   str          # cover | content | divider | ending | unknown
    confidence:      float        # 0.0–1.0
    reason:          str          # free-text explanation
    render_strategy: str          # standard | table_extract | image_passthrough | skip | manual
    recoverability:  str          # auto | review | manual
    warnings:        list[str] = field(default_factory=list)


# ── Valid labels (for documentation) ─────────────────────────────────────────

_COMPONENT_TYPES = frozenset({
    "cover", "content_text", "content_table", "content_mixed",
    "divider", "ending", "empty", "chart",
    "screenshot_table", "screenshot_other", "unknown",
})

# Keywords that hint a slide title or text is describing tabular data
_TABLE_KEYWORDS = frozenset({
    "table", "tracker", "dashboard", "scorecard", "matrix",
    "comparison", "breakdown", "summary", "metrics", "kpi",
    "data", "results", "report", "roster", "schedule", "plan",
    "budget", "forecast", "pipeline", "status",
})


# ── Main entry point ──────────────────────────────────────────────────────────

def classify_slide(
    slide: dict[str, Any],
    slide_idx: int,
    total_slides: int,
) -> RouteDecision:
    """
    Classify a raw slide dict and return a RouteDecision.

    Args:
        slide:        raw slide dict from pptx_extract.extract_raw_deck
        slide_idx:    0-based index of this slide in the deck
        total_slides: total number of slides in the deck

    Returns:
        RouteDecision — never raises
    """
    try:
        return _classify(slide, slide_idx, total_slides)
    except Exception as exc:  # noqa: BLE001
        return RouteDecision(
            route="fallback",
            component_type="unknown",
            local_purpose="unknown",
            confidence=0.0,
            reason=f"classifier error: {exc}",
            render_strategy="manual",
            recoverability="manual",
            warnings=[f"Classifier raised unexpectedly: {exc}"],
        )


def _classify(
    slide: dict[str, Any],
    slide_idx: int,
    total_slides: int,
) -> RouteDecision:
    raw_text = slide.get("raw_text", [])
    shapes   = slide.get("shapes", [])
    tables   = slide.get("tables", [])
    layout   = _parse_layout_sig(slide.get("layout_sig", ""))

    density, char_count = _text_density(raw_text)
    sc = _shape_counts(shapes)
    is_first = slide_idx == 0
    is_last  = slide_idx == total_slides - 1

    n_pictures = sc["picture"] + sc["picture_shape"]
    n_charts   = sc["chart"]

    # ── 1. Empty slide ────────────────────────────────────────────────────────
    if not shapes and density == "none":
        return RouteDecision(
            route="local", component_type="empty", local_purpose="unknown",
            confidence=1.0, reason="No shapes and no text",
            render_strategy="skip", recoverability="auto",
        )

    # ── 2. Cover: first slide with title-class placeholder ────────────────────
    if is_first and (_has_ph(shapes, "ctrTitle") or _has_ph(shapes, "subTitle")):
        return RouteDecision(
            route="local", component_type="cover", local_purpose="cover",
            confidence=0.95, reason="First slide with ctrTitle/subTitle placeholder",
            render_strategy="standard", recoverability="auto",
        )

    # Cover fallback: first slide, very low text, no images, no body
    if is_first and density in ("none", "low") and n_pictures == 0 and not _has_ph(shapes, "body"):
        return RouteDecision(
            route="local", component_type="cover", local_purpose="cover",
            confidence=0.80, reason="First slide, low text density, no body placeholder",
            render_strategy="standard", recoverability="auto",
        )

    # ── 3. Ending: last slide, low text, no images ────────────────────────────
    if is_last and density in ("none", "low") and n_pictures == 0 and not tables:
        return RouteDecision(
            route="local", component_type="ending", local_purpose="ending",
            confidence=0.90, reason="Last slide with low text density",
            render_strategy="standard", recoverability="auto",
        )

    # ── 4. Divider: short text only, no body placeholder, no images/tables ────
    if (density in ("none", "low")
            and n_pictures == 0
            and not tables
            and not _has_ph(shapes, "body")
            and not is_first
            and not is_last):
        return RouteDecision(
            route="local", component_type="divider", local_purpose="divider",
            confidence=0.85, reason="Short text, no body or images — section divider",
            render_strategy="standard", recoverability="auto",
        )

    # ── 5. Native table (real extracted table data, not screenshot) ───────────
    if tables and n_pictures == 0:
        return RouteDecision(
            route="local", component_type="content_table", local_purpose="content",
            confidence=0.92, reason=f"Native table data present ({len(tables)} table(s))",
            render_strategy="table_extract", recoverability="auto",
        )

    # ── 6. Chart slide ────────────────────────────────────────────────────────
    if n_charts > 0 and n_pictures == 0:
        return RouteDecision(
            route="local", component_type="chart", local_purpose="content",
            confidence=0.90, reason=f"Chart shape detected ({n_charts} chart(s))",
            render_strategy="image_passthrough", recoverability="auto",
        )

    # ── 7. Simple text content: title + body placeholders, no images ──────────
    if (_has_ph(shapes, "title")
            and _has_ph(shapes, "body")
            and n_pictures == 0
            and not tables
            and density != "none"):
        confidence = 0.88 if density == "medium" else 0.75
        return RouteDecision(
            route="local", component_type="content_text", local_purpose="content",
            confidence=confidence,
            reason="Title + body placeholders, no images — standard text content",
            render_strategy="standard", recoverability="auto",
        )

    # ── 8. Image / screenshot handling ───────────────────────────────────────
    if n_pictures > 0:
        # Check if any title text hints at tabular content
        title_text = _title_text(shapes, raw_text).lower()
        hints_table = any(kw in title_text for kw in _TABLE_KEYWORDS)

        if hints_table:
            # Potentially recoverable table screenshot → send to API
            return RouteDecision(
                route="api", component_type="screenshot_table", local_purpose="content",
                confidence=0.55,
                reason=f"Picture shapes + title hints at table content: '{title_text[:60]}'",
                render_strategy="manual", recoverability="review",
                warnings=["screenshot_table: API will attempt structure recovery"],
            )
        else:
            # Image-only, no table hint → not recoverable without OCR
            return RouteDecision(
                route="fallback", component_type="screenshot_other", local_purpose="unknown",
                confidence=0.70,
                reason="Image shapes with no table-related keywords — not recoverable",
                render_strategy="image_passthrough", recoverability="manual",
                warnings=["screenshot_other: manual review required; no OCR available"],
            )

    # ── 9. Mixed text + other content → ambiguous, send to API ───────────────
    if density in ("medium", "high") and (n_pictures > 0 or n_charts > 0 or sc["freeform"] > 2):
        return RouteDecision(
            route="api", component_type="content_mixed", local_purpose="content",
            confidence=0.50, reason="Mixed content — text + visual elements",
            render_strategy="standard", recoverability="review",
        )

    # ── 10. High text density with no clear structure → API ──────────────────
    if density == "high":
        return RouteDecision(
            route="api", component_type="unknown", local_purpose="unknown",
            confidence=0.45, reason="High text density without clear structural signals",
            render_strategy="standard", recoverability="review",
        )

    # ── 11. Catch-all fallback for anything we can't confidently classify ─────
    return RouteDecision(
        route="api", component_type="unknown", local_purpose="unknown",
        confidence=0.40, reason="No deterministic rule matched — routing to API",
        render_strategy="standard", recoverability="review",
    )


# ── Signal helpers ────────────────────────────────────────────────────────────

def _text_density(raw_text: list[str]) -> tuple[str, int]:
    """Return (density_label, total_char_count)."""
    total = sum(len(t) for t in raw_text)
    if total == 0:
        return "none", 0
    if total < 50:
        return "low", total
    if total < 350:
        return "medium", total
    return "high", total


def _shape_counts(shapes: list[dict[str, Any]]) -> dict[str, int]:
    """Count shapes by shape_type."""
    counts: dict[str, int] = {
        "placeholder": 0, "freeform": 0,
        "picture": 0, "picture_shape": 0,
        "table": 0, "chart": 0,
    }
    for s in shapes:
        t = s.get("shape_type", "unknown")
        counts[t] = counts.get(t, 0) + 1
    return counts


def _has_ph(shapes: list[dict[str, Any]], ph_type: str) -> bool:
    """Return True if any shape has the given placeholder_type."""
    return any(s.get("placeholder_type") == ph_type for s in shapes)


def _parse_layout_sig(sig: str) -> dict[str, Any]:
    """Parse layout signature string into component flags."""
    parts = sig.split("|") if sig else []
    n_shapes  = int("".join(filter(str.isdigit, parts[0])) or "0") if parts else 0
    has_table = len(parts) > 1 and parts[1] == "T"
    has_image = len(parts) > 2 and parts[2] == "I"
    has_chart = len(parts) > 3 and parts[3] == "C"
    return {"n_shapes": n_shapes, "has_table": has_table,
            "has_image": has_image, "has_chart": has_chart}


def _title_text(shapes: list[dict[str, Any]], raw_text: list[str]) -> str:
    """
    Return the title text of a slide.
    Prefer placeholder title/ctrTitle; fall back to first raw_text entry.
    """
    for s in shapes:
        if s.get("placeholder_type") in ("title", "ctrTitle") and s.get("text"):
            return s["text"]
    return raw_text[0] if raw_text else ""
