"""
agent_color_pass.py
Semantic color classification for slides via AgentRelay.
Called before the main color pipeline (style_slide) runs in convert_deck.py.
"""

_AGENT_RELAY_URL = "https://anyquest-webhook-relay-production-863f.up.railway.app"

_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Standard widescreen EMU dimensions used for normalizing shape positions
_SLIDE_W = 9144000
_SLIDE_H = 5143500

_PROMPT = """\
Look at this presentation slide image carefully. Focus only on colored elements \
— specifically any red, amber/orange, green, pink/magenta, or teal colored shapes, \
dots, circles, or gradient elements.

Red is semantic-only in this brand system. Treat red as meaningful only when it clearly \
signals negative / at-risk / low-score status. Never use decorative or generic emphasis \
as a reason to preserve or introduce red.

For each colored element or group of related colored elements, determine its semantic purpose:

- "rating_scale": colored dots or shapes in a column/row next to items in a list, \
indicating status, readiness, or rating per item
- "rag_indicator": explicit red/amber/green traffic light system, usually 3 shapes \
in a row or group
- "legend": color swatches with labels explaining what colors mean
- "gradient_slider": a horizontal or vertical bar that transitions through colors \
(dark → amber → green, or similar spectrum)
- "brand_decorative": a colored shape that is purely visual/structural with no semantic \
meaning (background rectangle, divider line, accent shape)

Return ONLY this JSON with no preamble:
{
  "semantic_colors": [
    {
      "description": "brief description of where this element is on the slide",
      "semantic": "rating_scale|rag_indicator|legend|gradient_slider|brand_decorative",
      "action": "preserve_original|remap_to_rag|remap_to_purple|interpolate_gradient",
      "confidence": 0.0-1.0
    }
  ],
  "has_semantic_color_system": true|false
}

Action rules:
- rating_scale → "preserve_original" (these dots mean something, do not remap)
- rag_indicator → "remap_to_rag" (remap to EA3323/FFB547/00C27A brand equivalents)
- legend → "preserve_original" (legend swatches must match whatever they are labeling)
- gradient_slider → "interpolate_gradient" (preserve endpoints, interpolate middle stops)
- brand_decorative → "remap_to_purple" (normal converter behavior)

Do NOT classify charts, headers, section bars, or decorative accents as rag_indicator \
unless they clearly encode evaluative status.

If you see no semantic color elements, return has_semantic_color_system: false and an \
empty semantic_colors array.\
"""


def agent_color_prepass(slide_image_path: str, slide_root) -> dict:
    """
    Send slide image to AgentRelay, get back semantic color classification.
    Returns a dict the converter uses before style_slide() runs.
    Returns empty dict if call fails or times out — caller falls back to heuristics.
    """
    try:
        import base64 as _b64
        import io as _io
        import json as _json
        import requests as _req
        import websocket as _ws
        from PIL import Image as _Img
    except ImportError:
        return {}

    try:
        img = _Img.open(slide_image_path)
        if img.width > 640:
            _h = int(img.height * 640 / img.width)
            img = img.resize((640, _h), _Img.LANCZOS)
        buf = _io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=70)
        b64 = _b64.b64encode(buf.getvalue()).decode()

        prompt = _PROMPT + f"\n\nSlide image (base64 JPEG):\n{b64}"

        resp = _req.post(
            f"{_AGENT_RELAY_URL}/submit",
            files={"agentId": (None, "generic-prompt-agent"),
                   "Prompt":  (None, prompt)},
            timeout=30,
        )
        if not resp.ok:
            print(f"  [color-agent] /submit {resp.status_code}: {resp.text[:200]}", flush=True)
            return {}

        result = resp.json()
        if not result.get("success"):
            return {}

        conn = _ws.create_connection(
            f"wss://anyquest-webhook-relay-production-863f.up.railway.app/ws?id={result['webhookId']}",
            timeout=60,
        )
        msg = conn.recv()
        conn.close()
        content = _json.loads(msg)["content"]
        if "{" in content:
            content = content[content.index("{"):]
        agent_data = _json.loads(content)

    except Exception as e:
        print(f"  [color-agent] error: {e}", flush=True)
        return {}

    try:
        semantic_colors = agent_data.get("semantic_colors", [])
        has_semantic = agent_data.get("has_semantic_color_system", False)

        if not has_semantic or not semantic_colors:
            return {}

        shapes = _get_shape_positions(slide_root)

        preserve_original_shapes = set()
        remap_to_rag_shapes = set()
        interpolate_gradient_shapes = set()
        confidences = []

        for item in semantic_colors:
            action = item.get("action", "preserve_original")
            description = item.get("description", "")
            confidence = item.get("confidence", 0.5)
            confidences.append(confidence)

            matched_ids, certain = _match_shapes_to_description(shapes, description)

            if not certain:
                # Uncertain match — always err toward preserve
                preserve_original_shapes.update(matched_ids)
            elif action == "preserve_original":
                preserve_original_shapes.update(matched_ids)
            elif action == "remap_to_rag":
                remap_to_rag_shapes.update(matched_ids)
            elif action == "interpolate_gradient":
                interpolate_gradient_shapes.update(matched_ids)
            # "remap_to_purple" — normal converter behavior, no special handling needed

        avg_confidence = sum(confidences) / len(confidences) if confidences else 0.0

        return {
            "has_semantic_color_system": True,
            "preserve_original_shapes": preserve_original_shapes,
            "remap_to_rag_shapes": remap_to_rag_shapes,
            "interpolate_gradient_shapes": interpolate_gradient_shapes,
            "raw_response": semantic_colors,
            "confidence": avg_confidence,
        }

    except Exception as e:
        print(f"  [color-agent] error: {e}", flush=True)
        return {}


def interpolate_gradient_stops(start_hex: str, end_hex: str, steps: int) -> list:
    """
    Generate intermediate hex colors as linear RGB interpolation between two endpoints.
    Used for gradient sliders where middle stops should be mathematically calculated
    rather than independently snapped to brand colors.
    """
    def hex_to_rgb(h):
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    def rgb_to_hex(r, g, b):
        return f"{int(r):02X}{int(g):02X}{int(b):02X}"

    r1, g1, b1 = hex_to_rgb(start_hex)
    r2, g2, b2 = hex_to_rgb(end_hex)

    return [
        rgb_to_hex(
            r1 + (r2 - r1) * i / (steps - 1),
            g1 + (g2 - g1) * i / (steps - 1),
            b1 + (b2 - b1) * i / (steps - 1),
        )
        for i in range(steps)
    ]


# ── Internal helpers ──────────────────────────────────────────────────────────

def _get_shape_positions(slide_root) -> list:
    """
    Extract shape IDs and normalized bounding boxes from slide XML.
    Returns list of dicts with keys: id, x, y, cx, cy, cx_center, cy_center.
    """
    shapes = []
    for sp in slide_root.iter(f"{{{_NS_P}}}sp"):
        cNvPr = sp.find(f".//{{{_NS_P}}}cNvPr")
        sp_id = cNvPr.get("id") if cNvPr is not None else None

        xfrm = sp.find(f".//{{{_NS_A}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{_NS_A}}}off")
        ext = xfrm.find(f"{{{_NS_A}}}ext")
        if off is None or ext is None:
            continue

        x  = int(off.get("x",  0))
        y  = int(off.get("y",  0))
        cx = int(ext.get("cx", 0))
        cy = int(ext.get("cy", 0))

        x_n  = x  / _SLIDE_W
        y_n  = y  / _SLIDE_H
        cx_n = cx / _SLIDE_W
        cy_n = cy / _SLIDE_H

        shapes.append({
            "id":        sp_id,
            "x":         x_n,
            "y":         y_n,
            "cx":        cx_n,
            "cy":        cy_n,
            "cx_center": x_n + cx_n / 2,
            "cy_center": y_n + cy_n / 2,
        })
    return shapes


def _match_shapes_to_description(shapes: list, description: str):
    """
    Match shapes to a text description by positional keywords.
    Returns (set_of_ids, is_certain).
    If no positional keywords are found, returns (all_ids, False) — caller
    should treat uncertain matches as preserve_original.
    """
    desc = description.lower()

    # Horizontal region filter
    h_filter = None
    if "left" in desc:
        h_filter = lambda cx: cx < 0.40
    elif "right" in desc:
        h_filter = lambda cx: cx > 0.60
    elif "center" in desc or ("middle" in desc and "top" not in desc and "bottom" not in desc):
        h_filter = lambda cx: 0.30 <= cx <= 0.70

    # Vertical region filter
    v_filter = None
    if "top" in desc or "upper" in desc:
        v_filter = lambda cy: cy < 0.40
    elif "bottom" in desc or "lower" in desc:
        v_filter = lambda cy: cy > 0.60

    all_ids = {s["id"] for s in shapes if s["id"] is not None}

    # No positional signal — uncertain
    if h_filter is None and v_filter is None:
        return all_ids, False

    matched = [
        s["id"] for s in shapes
        if s["id"] is not None
        and (h_filter(s["cx_center"]) if h_filter else True)
        and (v_filter(s["cy_center"]) if v_filter else True)
    ]

    # If more than half the shapes match, the region is too broad — uncertain
    if len(matched) > len(shapes) / 2:
        return all_ids, False

    return set(matched), True
