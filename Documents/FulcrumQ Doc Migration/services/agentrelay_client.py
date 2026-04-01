"""
services/agentrelay_client.py — AnyQuest / AgentRelay Semantic Client
======================================================================
Backend transport for semantic slide interpretation via the AgentRelay
webhook-relay pattern (Method 2: Full Agent Relay).

Integration pattern mirrors the existing _aq_get_response() used throughout
the data/app.py codebase. See ANYQUEST_API_INTEGRATION_GUIDE for full spec.

Flow per call:
  1. Generate unique session ID (UUID)
  2. Open WebSocket to relay (before submit, to avoid race condition)
  3. POST prompt to /submit with agentId + Prompt (capital P) + webhookId
  4. Block on ws.recv() until response arrives
  5. Parse JSON from response["content"]
  6. Return parsed dict

CONTRACT:
  - call_semantic_api(payload, **kwargs) -> dict
  - normalize_semantic_response(response, slide_id) -> dict

RULES:
  - No UI code
  - Retries with backoff on transient errors
  - Raises on unrecoverable errors (auth, permanent failures)
"""

from __future__ import annotations

import base64
import json
import logging
import time
import uuid
from typing import Any

import requests

try:
    import websocket
except ImportError:
    websocket = None  # type: ignore[assignment]

logger = logging.getLogger(__name__)

# ── AgentRelay endpoints (matches existing codebase pattern) ──────────────────

_SUBMIT_URL = "https://anyquest-webhook-relay-production-863f.up.railway.app/submit"
_WS_URL     = "wss://anyquest-webhook-relay-production-863f.up.railway.app/ws?id={}"
_AGENT_ID   = "generic-prompt-agent"

# ── Timeouts ──────────────────────────────────────────────────────────────────

WS_CONNECT_TIMEOUT  = 20    # seconds to establish WebSocket
HTTP_SUBMIT_TIMEOUT = 15    # seconds for POST /submit
WS_RECV_TIMEOUT     = 90    # seconds to wait for LLM response
DEFAULT_RETRIES     = 3
RETRY_BACKOFF       = 2.0   # seconds; doubles each attempt

# ── Normalized response defaults ──────────────────────────────────────────────

_DEFAULT_NORMALIZED: dict[str, Any] = {
    "slide_id":        None,
    "detected_purpose": "unknown",
    "content_summary":  "",
    "key_entities":    [],
    "structured_fields": {},
    "confidence":      0.0,
    "warnings":        [],
}

_SYSTEM_INSTRUCTIONS = (
    "You are a slide content analyzer for a PowerPoint deck conversion pipeline. "
    "Respond ONLY with a valid JSON object — no markdown fences, no commentary. "
    "The JSON must have exactly these fields:\n"
    "  slide_id: integer\n"
    "  detected_purpose: one of cover|content|divider|ending|unknown\n"
    "  content_summary: 1-2 sentence plain-English description\n"
    "  key_entities: array of important nouns/names/metrics\n"
    "  structured_fields: object with any structured data you can extract "
    "(title, subtitle, metrics, table_rows, section_name, etc.)\n"
    "  confidence: float 0.0–1.0"
)


# ── Main API call ─────────────────────────────────────────────────────────────

def call_semantic_api(
    payload: dict[str, Any],
    retries: int = DEFAULT_RETRIES,
    **_ignored,
) -> dict[str, Any]:
    """
    Send a slide payload to AnyQuest via AgentRelay and return parsed JSON.

    Follows the exact synchronous pattern from data/app.py _aq_get_response():
      1. Open WebSocket first (avoid race where response arrives before WS ready)
      2. POST to /submit with agentId, Prompt (capital P), webhookId
      3. Block on ws.recv() for the LLM response
      4. Parse response["content"] as JSON

    Args:
        payload:  slide-level dict from pptx_extract
        retries:  retry attempts on transient errors

    Returns:
        Parsed JSON dict from the LLM (via AgentRelay content field).

    Raises:
        ImportError:     if websocket-client is not installed
        ConnectionError: all retries exhausted on transient failures
        ValueError:      LLM returned non-JSON content (not retried)
    """
    if websocket is None:
        raise ImportError(
            "websocket-client not installed. Run: pip install websocket-client"
        )

    prompt   = _build_prompt(payload)
    last_exc: Exception | None = None
    delay    = RETRY_BACKOFF

    for attempt in range(1, retries + 1):
        session_id = str(uuid.uuid4())
        ws_uri     = _WS_URL.format(session_id)
        ws         = websocket.WebSocket()

        # Step 1: Open WebSocket before submitting (prevents race condition)
        try:
            ws.connect(ws_uri, timeout=WS_CONNECT_TIMEOUT)
        except Exception as exc:
            last_exc = exc
            logger.warning(
                "WebSocket connect failed (attempt %d/%d): %s", attempt, retries, exc
            )
            if attempt < retries:
                time.sleep(delay)
                delay *= 2
            continue

        # Step 2: POST to /submit
        try:
            resp = requests.post(
                _SUBMIT_URL,
                files={
                    "agentId":   (None, _AGENT_ID),
                    "Prompt":    (None, prompt),   # capital P — required by agent config
                    "webhookId": (None, session_id),
                },
                timeout=HTTP_SUBMIT_TIMEOUT,
            )
            resp.raise_for_status()
            submit_data = resp.json()
            logger.debug(
                "Slide %s submitted — jobId=%s requestId=%s",
                payload.get("slide_id"),
                submit_data.get("jobId"),
                submit_data.get("requestId"),
            )
        except Exception as exc:
            ws.close()
            last_exc = exc
            logger.warning(
                "POST /submit failed (attempt %d/%d): %s", attempt, retries, exc
            )
            if attempt < retries:
                time.sleep(delay)
                delay *= 2
            continue

        # Step 3: Wait for response on WebSocket
        try:
            ws.settimeout(WS_RECV_TIMEOUT)
            raw = ws.recv()
        except Exception as exc:
            ws.close()
            last_exc = exc
            logger.warning(
                "ws.recv() failed (attempt %d/%d): %s", attempt, retries, exc
            )
            if attempt < retries:
                time.sleep(delay)
                delay *= 2
            continue
        finally:
            ws.close()

        if not raw:
            last_exc = RuntimeError("Empty WebSocket message received")
            logger.warning("Empty WS message (attempt %d/%d)", attempt, retries)
            if attempt < retries:
                time.sleep(delay)
                delay *= 2
            continue

        # Step 4: Parse relay envelope, then parse inner JSON from LLM
        return _parse_relay_response(raw)

    raise ConnectionError(
        f"AgentRelay call failed after {retries} attempts: {last_exc}"
    )


# ── Prompt construction ───────────────────────────────────────────────────────

def _build_prompt(payload: dict[str, Any]) -> str:
    """Build the full prompt string sent to the agent via AgentRelay."""
    slide_id = payload.get("slide_id", "?")
    raw_text = payload.get("raw_text", [])
    shapes   = payload.get("shapes", [])
    tables   = payload.get("tables", [])
    colors   = payload.get("colors", [])
    layout   = payload.get("layout_sig", "")

    lines = [
        _SYSTEM_INSTRUCTIONS,
        "",
        f"Analyze this PowerPoint slide and return the JSON schema described above.",
        "",
        f"Slide {slide_id}",
        f"Layout signature: {layout}",
    ]

    if raw_text:
        lines.append("Text content:")
        for t in raw_text:
            lines.append(f"  - {t}")

    if shapes:
        lines.append(f"Shapes ({len(shapes)}):")
        for s in shapes:
            ph = f" [ph={s.get('placeholder_type')}]" if s.get("placeholder_type") else ""
            lines.append(f"  {s.get('shape_type','?')}{ph}: {s.get('text','')[:120]}")

    if tables:
        lines.append(f"Tables: {len(tables)}")
        for i, tbl in enumerate(tables[:2]):
            lines.append(f"  Table {i+1}: {len(tbl)} rows × {len(tbl[0]) if tbl else 0} cols")
            for row in tbl[:3]:    # first 3 rows as sample
                lines.append(f"    {' | '.join(row)}")

    if colors:
        lines.append(f"Colors seen: {', '.join(colors[:8])}")

    return "\n".join(lines)


# ── Response parsing ──────────────────────────────────────────────────────────

def _parse_relay_response(raw: str) -> dict[str, Any]:
    """
    Parse the AgentRelay WebSocket envelope and extract LLM JSON content.

    AgentRelay message format:
      {"id": "<webhookId>", "content": "<LLM text response>"}

    The LLM text response itself must be valid JSON (per our prompt instructions).

    Raises:
        ValueError: if outer envelope or inner content is not parseable JSON
    """
    # Parse outer relay envelope
    try:
        envelope = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"AgentRelay returned non-JSON envelope: {exc}\nRaw: {raw[:300]}") from exc

    content = envelope.get("content", "")
    if not content:
        # Envelope parsed but content is empty — return envelope itself and let normalizer handle
        logger.warning("AgentRelay envelope has empty content field")
        return envelope

    # Strip accidental markdown fences from LLM output
    text = content.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        text = "\n".join(l for l in lines if not l.startswith("```")).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"LLM response inside AgentRelay envelope is not JSON: {exc}\n"
            f"Content: {content[:300]}"
        ) from exc


# ── Normalization ─────────────────────────────────────────────────────────────

def normalize_semantic_response(
    response: dict[str, Any],
    slide_id: int | None = None,
) -> dict[str, Any]:
    """
    Normalize an AgentRelay/LLM response to the stable semantic schema.

    Fills missing fields with safe defaults.
    Coerces types where possible.
    Records normalization warnings.

    Args:
        response:  raw parsed dict from call_semantic_api
        slide_id:  expected slide_id (used to fill/verify the field)

    Returns:
        Normalized dict with all required fields guaranteed present.
    """
    normalized = dict(_DEFAULT_NORMALIZED)
    warnings: list[str] = []

    # slide_id
    resp_id = response.get("slide_id")
    if resp_id is not None:
        try:
            normalized["slide_id"] = int(resp_id)
        except (TypeError, ValueError):
            warnings.append(f"slide_id coercion failed: {resp_id!r}")
            normalized["slide_id"] = slide_id
    else:
        if slide_id is not None:
            warnings.append("slide_id missing from response; filled from context")
        normalized["slide_id"] = slide_id

    # detected_purpose
    valid_purposes = {"cover", "content", "divider", "ending", "unknown"}
    purpose = str(response.get("detected_purpose", "")).strip().lower()
    if purpose not in valid_purposes:
        warnings.append(
            f"detected_purpose '{purpose}' not in valid set" if purpose
            else "detected_purpose missing or empty"
        )
        purpose = "unknown"
    normalized["detected_purpose"] = purpose

    # content_summary
    normalized["content_summary"] = str(response.get("content_summary", "")).strip()
    if not normalized["content_summary"]:
        warnings.append("content_summary is empty")

    # key_entities
    entities = response.get("key_entities", [])
    if isinstance(entities, list):
        normalized["key_entities"] = [str(e) for e in entities]
    elif isinstance(entities, str):
        normalized["key_entities"] = [entities] if entities else []
        warnings.append("key_entities was string; wrapped in list")
    else:
        normalized["key_entities"] = []
        warnings.append(f"key_entities unexpected type: {type(entities).__name__}")

    # structured_fields
    fields = response.get("structured_fields", {})
    if isinstance(fields, dict):
        normalized["structured_fields"] = fields
    else:
        normalized["structured_fields"] = {}
        warnings.append(f"structured_fields unexpected type: {type(fields).__name__}")

    # confidence
    try:
        normalized["confidence"] = float(response.get("confidence", 0.0))
    except (TypeError, ValueError):
        normalized["confidence"] = 0.0
        warnings.append(f"confidence coercion failed: {response.get('confidence')!r}")

    # Pass through any extra fields
    known = set(_DEFAULT_NORMALIZED.keys())
    extra = {k: v for k, v in response.items() if k not in known}
    if extra:
        normalized["extra"] = extra

    existing = response.get("warnings", [])
    normalized["warnings"] = (existing if isinstance(existing, list) else []) + warnings

    return normalized


# ── Backwards-compat alias (used by semantic_pipeline.py) ────────────────────
call_local_api = call_semantic_api
