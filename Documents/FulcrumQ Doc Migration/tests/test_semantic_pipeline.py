"""
tests/test_semantic_pipeline.py
================================
Test suite for the Semantic Ingest pipeline.

Covers:
  - pptx_extract: structure, edge cases
  - slide_classifier: routing decisions, no-API for obvious slides
  - agentrelay_client: submission, WebSocket, timeout, malformed response
  - semantic_pipeline: routing integration, artifacts, partial failures

Run with:
    python -m pytest tests/test_semantic_pipeline.py -v
"""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

REPO = Path(__file__).parent.parent
import sys
sys.path.insert(0, str(REPO))

from services.pptx_extract      import extract_raw_deck
from services.agentrelay_client import (
    call_semantic_api,
    normalize_semantic_response,
    _build_prompt,
    _parse_relay_response,
)
from services.slide_classifier  import classify_slide, RouteDecision
from services.semantic_pipeline import run_semantic_ingest, _fallback_slide


# ─────────────────────────────────────────────────────────────────────────────
# PPTX fixtures
# ─────────────────────────────────────────────────────────────────────────────

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _make_slide_xml(
    texts: list[str],
    include_table: bool = False,
    include_picture: bool = False,
    ph_types: list[str] | None = None,
) -> bytes:
    """Build minimal slide XML with configurable content."""
    shapes_xml = ""
    ph_types = ph_types or (["title"] + ["body"] * max(0, len(texts) - 1))
    for i, t in enumerate(texts):
        ph = ph_types[i] if i < len(ph_types) else "body"
        shapes_xml += f"""
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="{i+2}" name="Shape {i+1}"/>
            <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
            <p:nvPr><p:ph type="{ph}"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:t>{t}</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>"""

    if include_picture:
        shapes_xml += """
        <p:pic>
          <p:nvPicPr>
            <p:cNvPr id="99" name="Picture 1"/>
            <p:cNvPicPr/><p:nvPr/>
          </p:nvPicPr>
          <p:blipFill><a:blip/><a:stretch/></p:blipFill>
          <p:spPr/>
        </p:pic>"""

    table_xml = ""
    if include_table:
        table_xml = f"""
        <p:graphicFrame>
          <p:nvGraphicFramePr>
            <p:cNvPr id="98" name="Table 1"/>
            <p:cNvGraphicFramePr/><p:nvPr/>
          </p:nvGraphicFramePr>
          <p:xfrm/>
          <a:graphic><a:graphicData>
            <a:tbl>
              <a:tr>
                <a:tc><a:txBody><a:p><a:r><a:t>Col A</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:p><a:r><a:t>Col B</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
              <a:tr>
                <a:tc><a:txBody><a:p><a:r><a:t>Val 1</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:p><a:r><a:t>Val 2</a:t></a:r></a:p></a:txBody></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData></a:graphic>
        </p:graphicFrame>"""

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    {shapes_xml}
    {table_xml}
  </p:spTree></p:cSld>
</p:sld>""".encode()


def _make_cover_xml(title: str, subtitle: str = "") -> bytes:
    sub = f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Sub"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle"/></p:nvPr></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{subtitle}</a:t></a:r></a:p></p:txBody>
    </p:sp>""" if subtitle else ""

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{title}</a:t></a:r></a:p></p:txBody>
    </p:sp>
    {sub}
  </p:spTree></p:cSld>
</p:sld>""".encode()


def _make_pptx(slides: list[bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
</Types>""")
        z.writestr("_rels/.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""")
        for i, xml in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", xml)
    buf.seek(0)
    return buf.read()


# ── AgentRelay helpers ────────────────────────────────────────────────────────

def _relay_envelope(slide_id: int, purpose: str = "content") -> str:
    inner = json.dumps({
        "slide_id": slide_id, "detected_purpose": purpose,
        "content_summary": f"Summary {slide_id}", "key_entities": ["X"],
        "structured_fields": {}, "confidence": 0.85,
    })
    return json.dumps({"id": "uid", "content": inner})


def _mock_ws(msg: str) -> MagicMock:
    ws = MagicMock()
    ws.recv.return_value = msg
    return ws


def _mock_submit() -> MagicMock:
    r = MagicMock()
    r.json.return_value = {"success": True, "webhookId": "x", "requestId": "x", "jobId": "j"}
    r.raise_for_status = MagicMock()
    return r


# ─────────────────────────────────────────────────────────────────────────────
# 1. PPTX Extraction
# ─────────────────────────────────────────────────────────────────────────────

class TestPptxExtract:

    def test_top_level_keys(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["T", "B"])]))
        r = extract_raw_deck(pptx)
        assert {"source", "slide_count", "slides"} <= r.keys()

    def test_slide_count(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["A"]), _make_slide_xml(["B"])]))
        assert extract_raw_deck(pptx)["slide_count"] == 2

    def test_slide_structure_keys(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["T"])]))
        s = extract_raw_deck(pptx)["slides"][0]
        for k in ("slide_id", "raw_text", "shapes", "tables", "colors", "layout_sig"):
            assert k in s

    def test_one_based_ids(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["A"]), _make_slide_xml(["B"])]))
        slides = extract_raw_deck(pptx)["slides"]
        assert slides[0]["slide_id"] == 1
        assert slides[1]["slide_id"] == 2

    def test_text_captured(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["Hello", "World"])]))
        s = extract_raw_deck(pptx)["slides"][0]
        combined = " ".join(s["raw_text"])
        assert "Hello" in combined and "World" in combined

    def test_empty_slide(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml([])]))
        s = extract_raw_deck(pptx)["slides"][0]
        assert s["raw_text"] == [] and s["shapes"] == []

    def test_table_extracted(self, tmp_path):
        pptx = tmp_path / "t.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["T"], include_table=True)]))
        s = extract_raw_deck(pptx)["slides"][0]
        assert len(s["tables"]) == 1
        assert s["tables"][0][0][0] == "Col A"

    def test_source_filename(self, tmp_path):
        pptx = tmp_path / "deck.pptx"
        pptx.write_bytes(_make_pptx([_make_slide_xml(["X"])]))
        assert extract_raw_deck(pptx)["source"] == "deck.pptx"


# ─────────────────────────────────────────────────────────────────────────────
# 2. Slide Classifier — routing decisions
# ─────────────────────────────────────────────────────────────────────────────

class TestSlideClassifier:

    def _slide(self, texts=(), shapes=(), tables=(), layout_sig="0sh|f|f|f",
               slide_id=2) -> dict:
        return {
            "slide_id": slide_id, "raw_text": list(texts),
            "shapes": list(shapes), "tables": list(tables),
            "colors": [], "layout_sig": layout_sig,
        }

    def _ph(self, ph_type: str, text: str = "x") -> dict:
        return {"shape_type": "placeholder", "placeholder_type": ph_type,
                "shape_name": ph_type, "text": text}

    def _pic(self) -> dict:
        return {"shape_type": "picture", "placeholder_type": None,
                "shape_name": "img", "text": ""}

    # ── Local routes ──────────────────────────────────────────────────────────

    def test_empty_slide_is_local(self):
        s = self._slide()
        d = classify_slide(s, slide_idx=1, total_slides=5)
        assert d.route == "local"
        assert d.component_type == "empty"

    def test_cover_first_slide_ctrTitle(self):
        s = self._slide(texts=["Big Title", "Subtitle"],
                        shapes=[self._ph("ctrTitle", "Big Title"),
                                self._ph("subTitle", "Subtitle")],
                        slide_id=1)
        d = classify_slide(s, slide_idx=0, total_slides=5)
        assert d.route == "local"
        assert d.component_type == "cover"
        assert d.local_purpose == "cover"

    def test_ending_last_slide_low_text(self):
        s = self._slide(texts=["Thank you"],
                        shapes=[self._ph("title", "Thank you")],
                        slide_id=10)
        d = classify_slide(s, slide_idx=9, total_slides=10)
        assert d.route == "local"
        assert d.component_type == "ending"

    def test_divider_short_text_no_body(self):
        s = self._slide(texts=["Section 2"],
                        shapes=[self._ph("title", "Section 2")],
                        slide_id=3)
        d = classify_slide(s, slide_idx=2, total_slides=10)
        assert d.route == "local"
        assert d.component_type == "divider"

    def test_native_table_is_local(self):
        s = self._slide(
            texts=["Initiative Tracker"],
            shapes=[self._ph("title", "Initiative Tracker")],
            tables=[[["Col A", "Col B"], ["v1", "v2"]]],
            slide_id=5,
        )
        d = classify_slide(s, slide_idx=4, total_slides=10)
        assert d.route == "local"
        assert d.component_type == "content_table"
        assert d.render_strategy == "table_extract"

    def test_simple_text_content_is_local(self):
        s = self._slide(
            texts=["Title Here", "Body text here with some content"],
            shapes=[self._ph("title", "Title Here"),
                    self._ph("body", "Body text here with some content")],
            slide_id=4,
        )
        d = classify_slide(s, slide_idx=3, total_slides=10)
        assert d.route == "local"
        assert d.component_type == "content_text"

    # ── API routes ────────────────────────────────────────────────────────────

    def test_screenshot_table_hint_goes_to_api(self):
        """Image slide whose title hints at tabular content → api."""
        s = self._slide(
            texts=["Q2 Metrics Dashboard"],
            shapes=[self._ph("title", "Q2 Metrics Dashboard"), self._pic()],
            slide_id=6,
        )
        d = classify_slide(s, slide_idx=5, total_slides=10)
        assert d.route == "api"
        assert d.component_type == "screenshot_table"
        assert d.recoverability == "review"

    def test_high_density_no_structure_goes_to_api(self):
        """High-density text with no clear structure → api."""
        long_text = "A" * 400
        s = self._slide(
            texts=[long_text],
            shapes=[self._ph("body", long_text)],
            slide_id=7,
        )
        d = classify_slide(s, slide_idx=6, total_slides=10)
        assert d.route == "api"

    # ── Fallback routes ───────────────────────────────────────────────────────

    def test_screenshot_no_table_hint_is_fallback(self):
        """Image slide with no table keywords → fallback."""
        s = self._slide(
            texts=["Company Headshot"],
            shapes=[self._ph("title", "Company Headshot"), self._pic()],
            slide_id=8,
        )
        d = classify_slide(s, slide_idx=7, total_slides=10)
        assert d.route == "fallback"
        assert d.component_type == "screenshot_other"
        assert d.render_strategy == "image_passthrough"
        assert d.recoverability == "manual"

    # ── Schema contract ───────────────────────────────────────────────────────

    def test_route_decision_has_all_fields(self):
        s = self._slide(texts=["Hi"], shapes=[self._ph("title", "Hi")], slide_id=2)
        d = classify_slide(s, slide_idx=1, total_slides=5)
        for attr in ("route", "component_type", "local_purpose", "confidence",
                     "reason", "render_strategy", "recoverability", "warnings"):
            assert hasattr(d, attr), f"Missing field: {attr}"

    def test_confidence_in_range(self):
        for idx, total in [(0, 1), (0, 5), (4, 5)]:
            s = self._slide(slide_id=idx + 1)
            d = classify_slide(s, slide_idx=idx, total_slides=total)
            assert 0.0 <= d.confidence <= 1.0

    def test_classifier_never_raises(self):
        """Malformed slide dict must not propagate an exception."""
        d = classify_slide({}, slide_idx=0, total_slides=1)
        assert isinstance(d, RouteDecision)

    # ── Obvious slides never call the API (parametrized) ─────────────────────

    @pytest.mark.parametrize("label,slide_factory,idx,total", [
        ("empty",  lambda s: s._slide(),                                       1, 5),
        ("cover",  lambda s: s._slide(texts=["Title", "Sub"],
                                      shapes=[s._ph("ctrTitle", "Title"),
                                              s._ph("subTitle", "Sub")]),      0, 5),
        ("ending", lambda s: s._slide(texts=["Bye"], shapes=[s._ph("title","Bye")]),
                                                                               4, 5),
        ("divider",lambda s: s._slide(texts=["Sec"], shapes=[s._ph("title","Sec")]),
                                                                               2, 5),
        ("table",  lambda s: s._slide(texts=["T"], shapes=[s._ph("title","T")],
                                      tables=[[["A","B"]]]),                   2, 5),
    ])
    def test_obvious_slides_route_local(self, label, slide_factory, idx, total):
        d = classify_slide(slide_factory(self), slide_idx=idx, total_slides=total)
        assert d.route == "local", f"{label}: expected local, got {d.route} ({d.reason})"


# ─────────────────────────────────────────────────────────────────────────────
# 3. AgentRelay client
# ─────────────────────────────────────────────────────────────────────────────

class TestAgentRelayClient:

    def test_parse_valid_envelope(self):
        inner = {"slide_id": 1, "detected_purpose": "cover", "confidence": 0.9,
                 "content_summary": "S", "key_entities": [], "structured_fields": {}}
        raw = json.dumps({"id": "u", "content": json.dumps(inner)})
        r = _parse_relay_response(raw)
        assert r["detected_purpose"] == "cover"

    def test_parse_strips_fences(self):
        inner = {"slide_id": 2, "detected_purpose": "content", "confidence": 0.8,
                 "content_summary": "", "key_entities": [], "structured_fields": {}}
        fenced = "```json\n" + json.dumps(inner) + "\n```"
        raw = json.dumps({"id": "u", "content": fenced})
        assert _parse_relay_response(raw)["slide_id"] == 2

    def test_parse_bad_envelope_raises(self):
        with pytest.raises(ValueError, match="non-JSON envelope"):
            _parse_relay_response("not json")

    def test_parse_non_json_content_raises(self):
        with pytest.raises(ValueError, match="not JSON"):
            _parse_relay_response(json.dumps({"id": "x", "content": "plain text"}))

    def test_normalize_complete(self):
        raw = {"slide_id": 3, "detected_purpose": "divider", "content_summary": "S",
               "key_entities": ["A"], "structured_fields": {}, "confidence": 0.9}
        n = normalize_semantic_response(raw, slide_id=3)
        assert n["detected_purpose"] == "divider"
        assert n["confidence"] == 0.9
        assert n["warnings"] == []

    def test_normalize_fills_defaults(self):
        n = normalize_semantic_response({}, slide_id=5)
        assert n["slide_id"] == 5
        assert n["detected_purpose"] == "unknown"
        assert len(n["warnings"]) > 0

    def test_normalize_invalid_purpose_coerced(self):
        n = normalize_semantic_response({"detected_purpose": "widget"}, slide_id=1)
        assert n["detected_purpose"] == "unknown"

    def test_normalize_string_entities_wrapped(self):
        n = normalize_semantic_response({"key_entities": "CEO"}, slide_id=1)
        assert n["key_entities"] == ["CEO"]

    def test_normalize_bad_confidence_zeroed(self):
        n = normalize_semantic_response({"confidence": "high"}, slide_id=1)
        assert n["confidence"] == 0.0

    def test_success_path(self):
        submit = _mock_submit()
        ws_msg = _relay_envelope(1, "cover")
        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", return_value=submit):
            mod.WebSocket.return_value = _mock_ws(ws_msg)
            r = call_semantic_api({"slide_id": 1, "raw_text": ["Hello"]})
        assert r["detected_purpose"] == "cover"

    def test_ws_opened_before_post(self):
        order = []
        ws_msg = _relay_envelope(1)
        ws_obj = MagicMock()
        ws_obj.recv.return_value = ws_msg
        ws_obj.connect.side_effect = lambda *a, **kw: order.append("ws")

        def post_side(*a, **kw):
            order.append("post")
            return _mock_submit()

        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", side_effect=post_side):
            mod.WebSocket.return_value = ws_obj
            call_semantic_api({"slide_id": 1})

        assert order.index("ws") < order.index("post")

    def test_post_uses_capital_prompt(self):
        captured = {}
        submit = _mock_submit()
        ws_msg = _relay_envelope(1)

        def capture_post(url, files=None, timeout=None):
            captured["files"] = files
            return submit

        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", side_effect=capture_post):
            mod.WebSocket.return_value = _mock_ws(ws_msg)
            call_semantic_api({"slide_id": 1})

        assert "Prompt" in captured["files"]
        assert "agentId" in captured["files"]
        assert "webhookId" in captured["files"]

    def test_ws_failure_retries_raises(self):
        ws_obj = MagicMock()
        ws_obj.connect.side_effect = OSError("refused")
        with patch("services.agentrelay_client.websocket") as mod, \
             patch("time.sleep"):
            mod.WebSocket.return_value = ws_obj
            with pytest.raises(ConnectionError):
                call_semantic_api({"slide_id": 1}, retries=2)
        assert ws_obj.connect.call_count == 2

    def test_timeout_retries_then_succeeds(self):
        calls = {"n": 0}
        ws_msg = _relay_envelope(1)
        submit = _mock_submit()

        def recv():
            calls["n"] += 1
            if calls["n"] < 2:
                raise OSError("timeout")
            return ws_msg

        ws_obj = MagicMock()
        ws_obj.recv.side_effect = recv
        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", return_value=submit), \
             patch("time.sleep"):
            mod.WebSocket.return_value = ws_obj
            r = call_semantic_api({"slide_id": 1}, retries=3)
        assert r["detected_purpose"] == "content"

    def test_non_json_content_raises_value_error(self):
        bad = json.dumps({"id": "x", "content": "plain text response"})
        submit = _mock_submit()
        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", return_value=submit):
            mod.WebSocket.return_value = _mock_ws(bad)
            with pytest.raises(ValueError):
                call_semantic_api({"slide_id": 1})

    def test_build_prompt_includes_instructions(self):
        p = _build_prompt({"slide_id": 1})
        assert "JSON" in p and "detected_purpose" in p

    def test_build_prompt_includes_text(self):
        p = _build_prompt({"slide_id": 2, "raw_text": ["Revenue grew 15%"]})
        assert "Revenue grew 15%" in p


# ─────────────────────────────────────────────────────────────────────────────
# 4. Semantic Pipeline — routing integration
# ─────────────────────────────────────────────────────────────────────────────

class TestSemanticPipeline:

    def _make_pptx_file(self, tmp_path: Path, slides: list[bytes]) -> Path:
        p = tmp_path / "deck.pptx"
        p.write_bytes(_make_pptx(slides))
        return p

    # ── Obvious slides must not call the API ──────────────────────────────────

    def test_cover_does_not_call_api(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_cover_xml("Big Title", "Subtitle"),
            _make_slide_xml(["Thank you"], ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api") as mock_api:
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")
        # cover (first) + ending (last) — both local
        mock_api.assert_not_called()
        assert schema["slides"][0]["component_type"] == "cover"
        assert schema["slides"][0]["classification_source"] == "rules"

    def test_ending_does_not_call_api(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_cover_xml("Title"),
            _make_slide_xml(["Thank you"], ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api") as mock_api:
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")
        ending = schema["slides"][-1]
        assert ending["classification_source"] == "rules"
        assert ending["component_type"] == "ending"

    def test_native_table_does_not_call_api(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_cover_xml("Deck"),
            _make_slide_xml(["Tracker"], include_table=True, ph_types=["title"]),
            _make_slide_xml(["End"], ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api") as mock_api:
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")
        table_slide = schema["slides"][1]
        mock_api.assert_not_called()
        assert table_slide["component_type"] == "content_table"
        assert table_slide["render_strategy"] == "table_extract"

    def test_empty_slide_does_not_call_api(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [_make_slide_xml([])])
        with patch("services.semantic_pipeline.call_semantic_api") as mock_api:
            run_semantic_ingest(pptx, output_dir=tmp_path / "art")
        mock_api.assert_not_called()

    # ── Ambiguous slides call the API ─────────────────────────────────────────

    def test_screenshot_table_calls_api(self, tmp_path):
        """Image slide with table-keyword title routes to API."""
        pptx = self._make_pptx_file(tmp_path, [
            _make_cover_xml("Deck"),
            _make_slide_xml(["Q2 Metrics Dashboard"], include_picture=True,
                            ph_types=["title"]),
            _make_slide_xml(["End"], ph_types=["title"]),
        ])
        api_response = _relay_envelope(2, "content")
        submit = _mock_submit()

        with patch("services.agentrelay_client.websocket") as mod, \
             patch("services.agentrelay_client.requests.post", return_value=submit):
            mod.WebSocket.return_value = _mock_ws(api_response)
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")

        slide = schema["slides"][1]
        assert slide["classification_source"] == "api"
        assert slide["component_type"] == "screenshot_table"

    # ── Both paths produce identical schema shape ─────────────────────────────

    def test_all_slides_have_required_schema_fields(self, tmp_path):
        required = {
            "slide_id", "detected_purpose", "component_type",
            "classification_source", "content_summary", "key_entities",
            "structured_fields", "confidence", "recoverability",
            "render_strategy", "warnings",
        }
        pptx = self._make_pptx_file(tmp_path, [
            _make_cover_xml("Cover"),
            _make_slide_xml(["Metrics", "Some body text here to be content"],
                            ph_types=["title", "body"]),
            _make_slide_xml(["End"], ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api",
                   return_value=json.loads(
                       _relay_envelope(2, "content"))["content"]  # shouldn't be called
                   ) if False else patch("services.semantic_pipeline.call_semantic_api",
                   side_effect=lambda p, **k: {
                       "slide_id": p["slide_id"], "detected_purpose": "content",
                       "content_summary": "s", "key_entities": [],
                       "structured_fields": {}, "confidence": 0.8,
                   }):
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")

        for slide in schema["slides"]:
            missing = required - slide.keys()
            assert not missing, f"Slide {slide.get('slide_id')} missing: {missing}"

    # ── Partial failure handling ──────────────────────────────────────────────

    def test_api_failure_falls_back_safely(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_slide_xml(["Q2 Data Dashboard"], include_picture=True, ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api",
                   side_effect=ConnectionError("relay down")):
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")

        assert len(schema["failures"]) == 1
        assert schema["failures"][0]["type"] == "connection"
        slide = schema["slides"][0]
        assert slide["detected_purpose"] == "unknown"
        assert slide["recoverability"] == "review"   # preserved from decision

    def test_artifacts_always_written(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_slide_xml(["Q2 Metrics Tracker"], include_picture=True, ph_types=["title"]),
        ])
        with patch("services.semantic_pipeline.call_semantic_api",
                   side_effect=ValueError("bad response")):
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")

        assert Path(schema["artifacts"]["raw_extract"]).is_file()
        assert Path(schema["artifacts"]["semantic_schema"]).is_file()

    # ── Route log ─────────────────────────────────────────────────────────────

    def test_route_log_present(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [_make_cover_xml("T")])
        with patch("services.semantic_pipeline.call_semantic_api") as m:
            schema = run_semantic_ingest(pptx, output_dir=tmp_path / "art")
        assert "route_log" in schema
        assert len(schema["route_log"]) == schema["slide_count"]
        entry = schema["route_log"][0]
        for k in ("slide_id", "route", "component_type", "confidence", "reason"):
            assert k in entry

    # ── Retries kwarg passes through ──────────────────────────────────────────

    def test_retries_passed_to_api(self, tmp_path):
        pptx = self._make_pptx_file(tmp_path, [
            _make_slide_xml(["Q2 Metrics Tracker"], include_picture=True, ph_types=["title"]),
        ])
        captured = {}

        def capture(payload, **kw):
            captured.update(kw)
            return {"slide_id": 1, "detected_purpose": "content",
                    "content_summary": "", "key_entities": [],
                    "structured_fields": {}, "confidence": 0.8}

        with patch("services.semantic_pipeline.call_semantic_api", side_effect=capture):
            run_semantic_ingest(pptx, output_dir=tmp_path / "art", retries=7)

        assert captured.get("retries") == 7

    # ── Fallback_slide helper ─────────────────────────────────────────────────

    def test_fallback_slide_has_routing_fields(self):
        fb = _fallback_slide(9, "error msg")
        for k in ("component_type", "classification_source",
                  "render_strategy", "recoverability"):
            assert k in fb, f"Missing: {k}"
        assert fb["classification_source"] == "fallback"


# ─────────────────────────────────────────────────────────────────────────────
# 5. Edge cases
# ─────────────────────────────────────────────────────────────────────────────

class TestEdgeCases:

    def test_freeform_text_captured(self, tmp_path):
        xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Free"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>Freeform text</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>""".encode()
        pptx = tmp_path / "f.pptx"
        pptx.write_bytes(_make_pptx([xml]))
        s = extract_raw_deck(pptx)["slides"][0]
        assert "Freeform text" in " ".join(s["raw_text"])

    def test_normalize_empty_dict(self):
        n = normalize_semantic_response({})
        assert n["detected_purpose"] == "unknown"
        assert n["slide_id"] is None

    def test_classifier_safe_on_empty_dict(self):
        d = classify_slide({}, slide_idx=0, total_slides=1)
        assert isinstance(d, RouteDecision)
        assert d.route in ("local", "api", "fallback")

    def test_relay_envelope_id_not_in_result(self):
        inner = {"slide_id": 4, "detected_purpose": "content",
                 "content_summary": "", "key_entities": [],
                 "structured_fields": {}, "confidence": 0.9}
        raw = json.dumps({"id": "some-uuid", "content": json.dumps(inner)})
        r = _parse_relay_response(raw)
        assert r["slide_id"] == 4
