"""
services/semantic_pipeline.py — Semantic Ingest Orchestrator
=============================================================
Pipeline architecture:
  source deck
  → raw extraction          (pptx_extract)
  → rule-based routing      (slide_classifier)  ← NEW
  → API for ambiguous only  (agentrelay_client)
  → unified schema          (normalize into one shape)
  → JSON artifacts

CONTRACT:
  run_semantic_ingest(pptx_path, output_dir, retries) -> dict

ROUTING DECISIONS (per slide):
  local      — deterministic heuristic is confident; no API call
  api        — ambiguous; send to AgentRelay for interpretation
  fallback   — not recoverable or classifier error; safe defaults

RULES:
  - No UI code
  - Always writes both artifacts (even on partial failure)
  - Obvious slides never call the API
  - Both local and API paths normalize to the same schema shape
"""

from __future__ import annotations

import json
import logging
import traceback
from pathlib import Path
from typing import Any

from services.pptx_extract      import extract_raw_deck
from services.agentrelay_client import call_semantic_api, normalize_semantic_response
from services.slide_classifier  import classify_slide, RouteDecision

logger = logging.getLogger(__name__)

# ── Artifact paths ────────────────────────────────────────────────────────────

_RAW_DIR    = "raw_extract"
_SCHEMA_DIR = "semantic_schema"


# ── Pipeline ──────────────────────────────────────────────────────────────────

def run_semantic_ingest(
    pptx_path: str | Path,
    output_dir: str | Path | None = None,
    retries: int | None = None,
) -> dict[str, Any]:
    """
    Run the full semantic ingestion pipeline.

    Steps:
      1. Extract raw signals  (pptx_extract)
      2. Save raw_extract.json artifact
      3. Route each slide:
           local    → build schema from heuristic (no API)
           api      → call AgentRelay, normalize response
           fallback → safe defaults, flag for review
      4. Combine into unified schema
      5. Save semantic_schema.json artifact

    Returns unified schema dict — see _SCHEMA_SHAPE for field reference.
    """
    pptx_path = Path(pptx_path)
    deck_name = pptx_path.stem

    if output_dir is None:
        output_dir = pptx_path.parent / "artifacts"
    output_dir = Path(output_dir)

    raw_dir    = output_dir / _RAW_DIR
    schema_dir = output_dir / _SCHEMA_DIR
    raw_dir.mkdir(parents=True, exist_ok=True)
    schema_dir.mkdir(parents=True, exist_ok=True)

    # ── Step 1: Extract ───────────────────────────────────────────────────────
    logger.info("Extracting raw signals: %s", pptx_path)
    raw_deck = extract_raw_deck(pptx_path)

    # ── Step 2: Save raw_extract ──────────────────────────────────────────────
    raw_extract_path = raw_dir / f"{deck_name}.json"
    _write_json(raw_extract_path, raw_deck)
    logger.info("Saved raw_extract: %s", raw_extract_path)

    # ── Step 3: Route + classify each slide ───────────────────────────────────
    slides_in = raw_deck["slides"]
    total     = len(slides_in)

    semantic_slides: list[dict[str, Any]] = []
    failures:        list[dict[str, Any]] = []
    route_log:       list[dict[str, Any]] = []

    api_kwargs: dict[str, Any] = {}
    if retries is not None:
        api_kwargs["retries"] = retries

    for idx, slide in enumerate(slides_in):
        slide_id = slide["slide_id"]
        decision = classify_slide(slide, slide_idx=idx, total_slides=total)

        route_log.append({
            "slide_id":       slide_id,
            "route":          decision.route,
            "component_type": decision.component_type,
            "confidence":     decision.confidence,
            "reason":         decision.reason,
        })
        logger.debug(
            "Slide %d → route=%s  type=%s  conf=%.2f  [%s]",
            slide_id, decision.route, decision.component_type,
            decision.confidence, decision.reason,
        )

        if decision.route == "local":
            result = _build_local_schema(slide, decision)
            semantic_slides.append(result)

        elif decision.route == "api":
            result = _call_api_with_fallback(
                slide, decision, slide_id, failures, api_kwargs
            )
            semantic_slides.append(result)

        else:  # fallback
            result = _build_fallback_schema(slide, decision)
            semantic_slides.append(result)
            logger.info(
                "Slide %d: fallback route — %s", slide_id, decision.reason
            )

    if failures:
        logger.warning(
            "%d/%d slides had API failures", len(failures), total
        )

    # ── Step 4: Combine ───────────────────────────────────────────────────────
    schema_path = schema_dir / f"{deck_name}.json"
    semantic_schema: dict[str, Any] = {
        "source":      raw_deck["source"],
        "slide_count": raw_deck["slide_count"],
        "slides":      semantic_slides,
        "failures":    failures,
        "route_log":   route_log,
        "artifacts": {
            "raw_extract":     str(raw_extract_path.resolve()),
            "semantic_schema": str(schema_path.resolve()),
        },
    }

    # ── Step 5: Save semantic_schema ──────────────────────────────────────────
    _write_json(schema_path, semantic_schema)
    logger.info("Saved semantic_schema: %s", schema_path)

    return semantic_schema


# ── Per-route schema builders ─────────────────────────────────────────────────

def _build_local_schema(
    slide: dict[str, Any],
    decision: RouteDecision,
) -> dict[str, Any]:
    """
    Build the full normalized schema for a locally-classified slide.
    No API call. Extracts summary from raw_text directly.
    """
    raw_text = slide.get("raw_text", [])
    tables   = slide.get("tables", [])

    # Derive a lightweight content_summary from text
    summary = ""
    if raw_text:
        summary = " | ".join(raw_text[:3])
        if len(summary) > 200:
            summary = summary[:197] + "…"

    # key_entities: first few meaningful text fragments
    entities = [t for t in raw_text if len(t) > 2][:5]

    # structured_fields: include table preview if native table
    structured: dict[str, Any] = {}
    if tables:
        structured["table_row_count"]  = sum(len(t) for t in tables)
        structured["table_count"]      = len(tables)
        if tables[0]:
            structured["table_headers"] = tables[0][0]

    result = _base_schema(slide["slide_id"], decision)
    result.update({
        "detected_purpose":  decision.local_purpose,
        "content_summary":   summary,
        "key_entities":      entities,
        "structured_fields": structured,
        "confidence":        decision.confidence,
    })
    return result


def _build_fallback_schema(
    slide: dict[str, Any],
    decision: RouteDecision,
) -> dict[str, Any]:
    """Build a safe fallback schema — no API, no guarantees."""
    result = _base_schema(slide["slide_id"], decision)
    result.update({
        "detected_purpose":  "unknown",
        "content_summary":   "",
        "key_entities":      [],
        "structured_fields": {},
        "confidence":        decision.confidence,
    })
    return result


def _call_api_with_fallback(
    slide: dict[str, Any],
    decision: RouteDecision,
    slide_id: int,
    failures: list[dict[str, Any]],
    api_kwargs: dict[str, Any],
) -> dict[str, Any]:
    """
    Call AgentRelay, normalize the response, and inject routing fields.
    On any error, record failure and return a fallback schema.
    """
    try:
        logger.debug("Calling AgentRelay for slide %d", slide_id)
        raw_response = call_semantic_api(slide, **api_kwargs)
        normalized   = normalize_semantic_response(raw_response, slide_id=slide_id)

        # Inject routing fields from the decision
        _inject_routing_fields(normalized, decision)
        _warn_if_degraded(normalized, slide_id)
        return normalized

    except ConnectionError as exc:
        msg = f"AgentRelay unreachable for slide {slide_id}: {exc}"
        logger.error(msg)
        failures.append({"slide_id": slide_id, "error": msg, "type": "connection"})
        fb = _build_fallback_schema(slide, decision)
        fb["warnings"].append(f"API call failed (connection): {exc}")
        return fb

    except ValueError as exc:
        msg = f"Malformed AgentRelay response for slide {slide_id}: {exc}"
        logger.error(msg)
        failures.append({"slide_id": slide_id, "error": msg, "type": "malformed"})
        fb = _build_fallback_schema(slide, decision)
        fb["warnings"].append(f"API call failed (malformed): {exc}")
        return fb

    except Exception as exc:  # noqa: BLE001
        msg = f"Unexpected error for slide {slide_id}: {exc}"
        logger.error(msg)
        logger.debug(traceback.format_exc())
        failures.append({"slide_id": slide_id, "error": msg, "type": "unexpected"})
        fb = _build_fallback_schema(slide, decision)
        fb["warnings"].append(f"API call failed (unexpected): {exc}")
        return fb


# ── Schema helpers ────────────────────────────────────────────────────────────

def _base_schema(slide_id: int, decision: RouteDecision) -> dict[str, Any]:
    """Return a skeleton schema dict populated with routing metadata."""
    return {
        "slide_id":             slide_id,
        "component_type":       decision.component_type,
        "classification_source": (
            "rules"    if decision.route == "local"    else
            "api"      if decision.route == "api"      else
            "fallback"
        ),
        "render_strategy":      decision.render_strategy,
        "recoverability":       decision.recoverability,
        "warnings":             list(decision.warnings),
        # filled by caller:
        "detected_purpose":     "unknown",
        "content_summary":      "",
        "key_entities":         [],
        "structured_fields":    {},
        "confidence":           0.0,
    }


def _inject_routing_fields(
    normalized: dict[str, Any],
    decision: RouteDecision,
) -> None:
    """
    Add routing metadata into an already-normalized API response in-place.
    Uses decision values as defaults; respects any fields the API already set.
    """
    normalized.setdefault("component_type",        decision.component_type)
    normalized.setdefault("classification_source", "api")
    normalized.setdefault("render_strategy",       decision.render_strategy)
    normalized.setdefault("recoverability",        decision.recoverability)

    # Merge decision warnings without duplicating
    existing = normalized.get("warnings", [])
    for w in decision.warnings:
        if w not in existing:
            existing.append(w)
    normalized["warnings"] = existing


def _warn_if_degraded(normalized: dict[str, Any], slide_id: int) -> None:
    """Log warnings for low-quality API results."""
    if normalized.get("detected_purpose") in (None, "", "unknown"):
        logger.warning("Slide %d: detected_purpose is unknown after API", slide_id)
    if not normalized.get("content_summary"):
        logger.warning("Slide %d: content_summary is empty after API", slide_id)
    if normalized.get("confidence", 0.0) < 0.4:
        logger.warning("Slide %d: low confidence %.2f from API", slide_id,
                       normalized.get("confidence", 0.0))


# ── I/O helpers ──────────────────────────────────────────────────────────────

def _write_json(path: Path, data: Any, indent: int = 2) -> None:
    path.write_text(json.dumps(data, indent=indent, ensure_ascii=False), encoding="utf-8")


def _fallback_slide(slide_id: int, error_msg: str) -> dict[str, Any]:
    """Legacy helper kept for test compatibility."""
    return {
        "slide_id":              slide_id,
        "detected_purpose":      "unknown",
        "component_type":        "unknown",
        "classification_source": "fallback",
        "render_strategy":       "manual",
        "recoverability":        "manual",
        "content_summary":       "",
        "key_entities":          [],
        "structured_fields":     {},
        "confidence":            0.0,
        "warnings":              [f"API call failed: {error_msg}"],
    }
