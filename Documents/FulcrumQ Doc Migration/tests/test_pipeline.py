"""
tests/test_pipeline.py
======================
Comprehensive test suite for the FulcrumQ 4-layer conversion pipeline.

Covers:
  - Layer 1: classifier.py
  - Layer 2: schema_engine.py
  - Layer 3: renderer.py
  - Layer 4: qa_validator.py
  - Integration: run_pipeline orchestration

Run with:
    python -m pytest tests/test_pipeline.py -v
"""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import pytest
from lxml import etree

# ── Path setup ────────────────────────────────────────────────────────────────

REPO = Path(__file__).parent.parent
import sys
sys.path.insert(0, str(REPO))

from classifier    import classify_slide, classify_deck
from schema_engine import decide_layout, run_schema_engine, _eval_condition, _select_layout
from renderer      import (
    render_deck, _extract_placeholder_content, _detect_overflow,
    _load_pptx, _save_pptx
)
from qa_validator  import validate_render, _check_slide, _global_checks


LAYOUT_RULES = REPO / "layout_rules.json"

# ─────────────────────────────────────────────────────────────────────────────
# Fixtures / helpers
# ─────────────────────────────────────────────────────────────────────────────

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _make_slide_xml(
    ph_type: str = "title",
    text: str = "Hello World",
    bg_hex: str = None,
    extra_text_shapes: int = 0,
    font_size_hundredths: int = 2400,
    has_body: bool = False,
    ph_idx: str = None,
) -> bytes:
    """Build a minimal slide XML bytes for testing."""
    nsmap = {
        "p": NS_P,
        "a": NS_A,
    }

    # Build <p:sld> root manually via string for simplicity
    body_ph = ""
    if has_body:
        body_ph = """
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="3" name="Body"/>
            <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
            <p:nvPr><p:ph type="body"/></p:nvPr>
          </p:nvSpPr>
          <p:spPr/>
          <p:txBody>
            <a:bodyPr/>
            <a:p><a:r><a:t>Body text here for testing overflow detection and content mapping</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>"""

    extra_shapes = ""
    for i in range(extra_text_shapes):
        extra_shapes += f"""
        <p:sp>
          <p:nvSpPr>
            <p:cNvPr id="{10+i}" name="Extra{i}"/>
            <p:cNvSpPr/>
            <p:nvPr/>
          </p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="100" y="100"/><a:ext cx="500000" cy="200000"/></a:xfrm>
          </p:spPr>
          <p:txBody>
            <a:bodyPr/>
            <a:p><a:r><a:rPr sz="{font_size_hundredths}"/><a:t>Extra shape {i}</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>"""

    idx_attr = f' idx="{ph_idx}"' if ph_idx else ""
    type_attr = f' type="{ph_type}"' if ph_type else ""

    bg_xml = ""
    if bg_hex:
        bg_xml = f"""
        <p:bg>
          <p:bgPr>
            <a:solidFill><a:srgbClr val="{bg_hex}"/></a:solidFill>
          </p:bgPr>
        </p:bg>"""

    xml_str = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="{NS_P}" xmlns:a="{NS_A}">
  <p:cSld>{bg_xml}
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph{type_attr}{idx_attr}/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:rPr sz="{font_size_hundredths}"/><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      {body_ph}
      {extra_shapes}
    </p:spTree>
  </p:cSld>
</p:sld>"""
    return xml_str.encode("utf-8")


def _make_minimal_pptx(slides: list[bytes], include_layout: bool = True) -> bytes:
    """
    Build an in-memory PPTX (ZIP) with the given slide XMLs.
    Returns bytes suitable for writing to a file or passing to _load_pptx via tmp.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml
        slide_overrides = "\n".join(
            f'  <Override PartName="/ppt/slides/slide{i+1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            for i in range(len(slides))
        )
        layout_override = ""
        if include_layout:
            layout_override = """
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
   ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>"""

        ct_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
   ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
{slide_overrides}
{layout_override}
</Types>"""
        z.writestr("[Content_Types].xml", ct_xml)

        # _rels/.rels
        z.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>""")

        # ppt/presentation.xml
        sld_ids = "\n".join(
            f'    <p:sldId id="{256+i}" r:id="rId{i+1}"/>'
            for i in range(len(slides))
        )
        z.writestr("ppt/presentation.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="{NS_P}"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
{sld_ids}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
</p:presentation>""")

        # ppt/_rels/presentation.xml.rels
        rels = "\n".join(
            f'  <Relationship Id="rId{i+1}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i+1}.xml"/>'
            for i in range(len(slides))
        )
        z.writestr("ppt/_rels/presentation.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}
</Relationships>""")

        # Slides
        for i, slide_xml in enumerate(slides):
            z.writestr(f"ppt/slides/slide{i+1}.xml", slide_xml)
            z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
   Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>""")

        # Minimal layout
        if include_layout:
            z.writestr("ppt/slideLayouts/slideLayout1.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="{NS_P}" xmlns:a="{NS_A}">
  <p:cSld name="Content_Light">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/><p:txBody><a:bodyPr/><a:p/></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:spPr/><p:txBody><a:bodyPr/><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sldLayout>""")
            z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""")

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Layer 1: Classifier
# ─────────────────────────────────────────────────────────────────────────────

class TestClassifier:

    def test_slide_1_cover_title_only(self):
        xml = _make_slide_xml(ph_type="ctrTitle", text="Q3 Strategy", has_body=False)
        root = etree.fromstring(xml)
        result = classify_slide(root, 1, 5)
        assert result["type"] == "cover"
        assert result["confidence"] >= 0.85
        assert result["slide_id"] == 1

    def test_slide_1_cover_sparse(self):
        xml = _make_slide_xml(ph_type="title", text="Short", has_body=False)
        root = etree.fromstring(xml)
        result = classify_slide(root, 1, 5)
        assert result["type"] == "cover"
        assert result["slide_id"] == 1

    def test_slide_1_cover_with_body_low_confidence(self):
        long_text = "A" * 300
        xml = _make_slide_xml(ph_type="title", text=long_text, has_body=True)
        root = etree.fromstring(xml)
        result = classify_slide(root, 1, 5)
        assert result["type"] == "cover"
        assert result["confidence"] < 0.75
        assert "low_confidence" in " ".join(result["warnings"])

    def test_last_slide_ending_keywords(self):
        xml = _make_slide_xml(ph_type="title", text="Thank you! Questions?")
        root = etree.fromstring(xml)
        result = classify_slide(root, 5, 5)
        assert result["type"] == "ending"
        assert result["confidence"] >= 0.88

    def test_last_slide_ending_sparse(self):
        xml = _make_slide_xml(ph_type="title", text="QA")
        root = etree.fromstring(xml)
        result = classify_slide(root, 5, 5)
        assert result["type"] == "ending"

    def test_divider_single_large_text(self):
        xml = _make_slide_xml(ph_type="", text="Section 1", font_size_hundredths=4400, has_body=False)
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["type"] == "divider"

    def test_divider_title_only_dark_bg(self):
        xml = _make_slide_xml(ph_type="title", text="Section 2", bg_hex="1A1A2E", has_body=False)
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["type"] == "divider"

    def test_content_default(self):
        xml = _make_slide_xml(ph_type="title", text="Analysis", has_body=True)
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["type"] == "content"

    def test_metadata_keys_present(self):
        xml = _make_slide_xml(ph_type="title", text="Test")
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        required_keys = {"text_density", "has_chart", "image_count", "has_logo", "bg_is_dark", "has_subtitle"}
        assert required_keys.issubset(set(result["metadata"].keys()))

    def test_classify_deck_returns_list(self, tmp_path):
        slides = [
            _make_slide_xml(ph_type="ctrTitle", text="Cover"),
            _make_slide_xml(ph_type="title", text="Content Slide", has_body=True),
            _make_slide_xml(ph_type="title", text="Thank you"),
        ]
        pptx_bytes = _make_minimal_pptx(slides)
        pptx_path = tmp_path / "test.pptx"
        pptx_path.write_bytes(pptx_bytes)

        results = classify_deck(pptx_path)
        assert len(results) == 3
        assert results[0]["type"] == "cover"
        assert results[2]["type"] == "ending"

    def test_classify_deck_writes_json(self, tmp_path):
        slides = [_make_slide_xml(ph_type="ctrTitle", text="Cover")]
        pptx_bytes = _make_minimal_pptx(slides)
        pptx_path = tmp_path / "test.pptx"
        pptx_path.write_bytes(pptx_bytes)

        out_path = tmp_path / "out.json"
        classify_deck(pptx_path, output_path=out_path)
        assert out_path.exists()
        data = json.loads(out_path.read_text())
        assert isinstance(data, list)

    def test_slide_1_not_cover_warning(self):
        # Force a situation where slide 1 is classified not-cover (hard to do, but
        # verify the warning infrastructure works by calling classify_slide directly
        # with a patched result)
        xml = _make_slide_xml(ph_type="title", text="A" * 50, has_body=True)
        root = etree.fromstring(xml)
        # Any slide 1 should always be "cover" in our heuristic — verify no slide_1_not_cover
        result = classify_slide(root, 1, 5)
        assert "slide_1_not_cover" not in result["warnings"]

    def test_text_density_low(self):
        xml = _make_slide_xml(ph_type="title", text="Short")
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["metadata"]["text_density"] == "low"

    def test_text_density_high(self):
        xml = _make_slide_xml(ph_type="title", text="X" * 600, has_body=True)
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["metadata"]["text_density"] == "high"

    def test_bg_is_dark_detected(self):
        xml = _make_slide_xml(ph_type="title", text="Dark", bg_hex="0D0D0D")
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["metadata"]["bg_is_dark"] is True

    def test_bg_is_dark_light(self):
        xml = _make_slide_xml(ph_type="title", text="Light", bg_hex="FFFFFF")
        root = etree.fromstring(xml)
        result = classify_slide(root, 2, 5)
        assert result["metadata"]["bg_is_dark"] is False

    def test_has_subtitle_body_idx_12(self):
        xml = _make_slide_xml(ph_type="body", text="Subtitle text", ph_idx="12")
        root = etree.fromstring(xml)
        result = classify_slide(root, 1, 5)
        assert result["metadata"]["has_subtitle"] is True


# ─────────────────────────────────────────────────────────────────────────────
# Layer 2: Schema Engine
# ─────────────────────────────────────────────────────────────────────────────

class TestSchemaEngine:

    def _cover_result(self, bg_dark=False) -> dict:
        return {
            "slide_id": 1,
            "type": "cover",
            "confidence": 0.90,
            "warnings": [],
            "metadata": {
                "bg_is_dark": bg_dark,
                "has_subtitle": False,
                "has_chart": False,
                "image_count": 0,
                "text_density": "low",
            },
        }

    def _content_result(self, bg_dark=False, has_subtitle=False,
                        text_density="medium", has_chart=False, image_count=0) -> dict:
        return {
            "slide_id": 2,
            "type": "content",
            "confidence": 0.80,
            "warnings": [],
            "metadata": {
                "bg_is_dark": bg_dark,
                "has_subtitle": has_subtitle,
                "has_chart": has_chart,
                "image_count": image_count,
                "text_density": text_density,
            },
        }

    def _load_rules(self):
        config = json.loads(LAYOUT_RULES.read_text())
        return config["layout_rules"], config["fallback_layout"], config.get("subtitle_upgrade_map", {})

    def test_cover_dark(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._cover_result(bg_dark=True), rules, fallback, upgrade)
        assert d["selected_layout"] == "Cover_Dark"
        assert not d["fallback_used"]

    def test_cover_light(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._cover_result(bg_dark=False), rules, fallback, upgrade)
        assert d["selected_layout"] == "Cover_Light"

    def test_content_dark(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._content_result(bg_dark=True), rules, fallback, upgrade)
        assert d["selected_layout"] == "1_Content_Dark"

    def test_content_default(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._content_result(), rules, fallback, upgrade)
        assert d["selected_layout"] == "Content_Light"

    def test_content_subtitle_light(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._content_result(has_subtitle=True), rules, fallback, upgrade)
        assert d["selected_layout"] == "Content_Light_Sub"

    def test_content_subtitle_dark(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._content_result(bg_dark=True, has_subtitle=True), rules, fallback, upgrade)
        assert d["selected_layout"] == "Content_Dark_Sub"

    def test_content_sparse_visual(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(
            self._content_result(text_density="low", has_chart=False, image_count=2),
            rules, fallback, upgrade
        )
        assert d["selected_layout"] == "Title_Only_Light"

    def test_ending_layout(self):
        rules, fallback, upgrade = self._load_rules()
        ending = {
            "slide_id": 5, "type": "ending", "confidence": 0.92, "warnings": [],
            "metadata": {"bg_is_dark": False, "has_subtitle": False, "has_chart": False,
                         "image_count": 0, "text_density": "low"},
        }
        d = decide_layout(ending, rules, fallback, upgrade)
        assert d["selected_layout"] == "Ending_PivotPurple"

    def test_divider_dark(self):
        rules, fallback, upgrade = self._load_rules()
        divider = {
            "slide_id": 3, "type": "divider", "confidence": 0.88, "warnings": [],
            "metadata": {"bg_is_dark": True, "has_subtitle": False, "has_chart": False,
                         "image_count": 0, "text_density": "low"},
        }
        d = decide_layout(divider, rules, fallback, upgrade)
        assert d["selected_layout"] == "Divider_Dark"

    def test_divider_light(self):
        rules, fallback, upgrade = self._load_rules()
        divider = {
            "slide_id": 3, "type": "divider", "confidence": 0.88, "warnings": [],
            "metadata": {"bg_is_dark": False, "has_subtitle": False, "has_chart": False,
                         "image_count": 0, "text_density": "low"},
        }
        d = decide_layout(divider, rules, fallback, upgrade)
        assert d["selected_layout"] == "Divider_Crystal"

    def test_fallback_for_unknown_type(self):
        rules, fallback, upgrade = self._load_rules()
        unknown = {
            "slide_id": 9, "type": "alien", "confidence": 0.50, "warnings": [],
            "metadata": {},
        }
        d = decide_layout(unknown, rules, fallback, upgrade)
        assert d["fallback_used"] is True
        assert d["selected_layout"] == fallback
        assert "layout_fallback_used" in d["warnings"]

    def test_low_confidence_propagated(self):
        rules, fallback, upgrade = self._load_rules()
        result = self._cover_result()
        result["confidence"] = 0.62
        d = decide_layout(result, rules, fallback, upgrade)
        assert any("low_classifier_confidence" in w for w in d["warnings"])

    def test_decision_path_non_empty(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._cover_result(), rules, fallback, upgrade)
        assert len(d["decision_path"]) >= 1

    def test_decision_path_contains_layout(self):
        rules, fallback, upgrade = self._load_rules()
        d = decide_layout(self._cover_result(), rules, fallback, upgrade)
        assert any("Cover_Light" in step for step in d["decision_path"])

    def test_run_schema_engine_writes_json(self, tmp_path):
        classifier_results = [
            {"slide_id": 1, "type": "cover", "confidence": 0.90, "warnings": [],
             "metadata": {"bg_is_dark": False, "has_subtitle": False, "has_chart": False,
                          "image_count": 0, "text_density": "low"}},
            {"slide_id": 2, "type": "content", "confidence": 0.80, "warnings": [],
             "metadata": {"bg_is_dark": False, "has_subtitle": False, "has_chart": False,
                          "image_count": 0, "text_density": "medium"}},
        ]
        out = tmp_path / "decisions.json"
        decisions = run_schema_engine(classifier_results, LAYOUT_RULES, output_path=out)
        assert out.exists()
        assert len(decisions) == 2
        assert decisions[0]["selected_layout"] == "Cover_Light"
        assert decisions[1]["selected_layout"] == "Content_Light"

    def test_eval_condition_type_match(self):
        assert _eval_condition({"type": "cover"}, {"type": "cover", "metadata": {}}) is True

    def test_eval_condition_type_no_match(self):
        assert _eval_condition({"type": "cover"}, {"type": "content", "metadata": {}}) is False

    def test_eval_condition_bg_dark(self):
        meta = {"type": "content", "metadata": {"bg_is_dark": True}}
        assert _eval_condition({"bg_is_dark": True}, meta) is True
        assert _eval_condition({"bg_is_dark": False}, meta) is False

    def test_eval_condition_image_count_gt(self):
        meta = {"type": "content", "metadata": {"image_count": 3, "text_density": "low", "has_chart": False}}
        assert _eval_condition({"image_count_gt": 0}, meta) is True
        assert _eval_condition({"image_count_gt": 5}, meta) is False

    def test_subtitle_upgrade_applied(self):
        # content-subtitle-light rule (priority 759) fires directly → Content_Light_Sub
        rules, fallback, upgrade = self._load_rules()
        result = self._content_result(has_subtitle=True)
        d = decide_layout(result, rules, fallback, upgrade)
        assert d["selected_layout"] == "Content_Light_Sub"
        assert not d["fallback_used"]


# ─────────────────────────────────────────────────────────────────────────────
# Layer 3: Renderer
# ─────────────────────────────────────────────────────────────────────────────

class TestRenderer:

    def _make_pptx_file(self, tmp_path, slides) -> Path:
        p = tmp_path / "source.pptx"
        p.write_bytes(_make_minimal_pptx(slides))
        return p

    def test_extract_placeholder_title(self):
        xml = _make_slide_xml(ph_type="title", text="My Title")
        root = etree.fromstring(xml)
        content = _extract_placeholder_content(root)
        assert "title" in content
        assert "My Title" in content["title"]

    def test_extract_placeholder_body(self):
        xml = _make_slide_xml(ph_type="title", text="Title", has_body=True)
        root = etree.fromstring(xml)
        content = _extract_placeholder_content(root)
        assert "body" in content

    def test_extract_placeholder_subtitle_idx12(self):
        xml = _make_slide_xml(ph_type="body", text="Subtitle text", ph_idx="12")
        root = etree.fromstring(xml)
        content = _extract_placeholder_content(root)
        assert "subtitle" in content
        assert "Subtitle text" in content["subtitle"]

    def test_overflow_detection_long_text(self):
        long_text = "word " * 200
        xml = _make_slide_xml(ph_type="body", text=long_text)
        root = etree.fromstring(xml)
        # Only triggers if > 600 chars and no autofit — our helper has no autofit
        assert _detect_overflow(root) is True

    def test_overflow_detection_short_text(self):
        xml = _make_slide_xml(ph_type="body", text="Short content")
        root = etree.fromstring(xml)
        assert _detect_overflow(root) is False

    def test_render_deck_produces_pptx(self, tmp_path):
        slides = [
            _make_slide_xml(ph_type="ctrTitle", text="Cover"),
            _make_slide_xml(ph_type="title", text="Content", has_body=True),
            _make_slide_xml(ph_type="title", text="Thank you"),
        ]
        source = self._make_pptx_file(tmp_path, slides)
        master = source  # use same file as master (minimal test)

        decisions = [
            {"slide_id": 1, "selected_layout": "Content_Light", "decision_path": [], "fallback_used": False, "warnings": []},
            {"slide_id": 2, "selected_layout": "Content_Light", "decision_path": [], "fallback_used": False, "warnings": []},
            {"slide_id": 3, "selected_layout": "Content_Light", "decision_path": [], "fallback_used": False, "warnings": []},
        ]
        out_pptx = tmp_path / "out.pptx"
        log = render_deck(source, decisions, master, out_pptx)

        assert out_pptx.exists()
        assert len(log) == 3

    def test_render_log_keys(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        source = self._make_pptx_file(tmp_path, slides)
        decisions = [{"slide_id": 1, "selected_layout": "Content_Light",
                      "decision_path": [], "fallback_used": False, "warnings": []}]
        out_pptx = tmp_path / "out.pptx"
        log = render_deck(source, decisions, source, out_pptx)

        required = {"slide_id", "layout_used", "content_mapped",
                    "overflow_detected", "missing_placeholders", "warnings"}
        assert required.issubset(set(log[0].keys()))

    def test_render_log_written_to_file(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        source = self._make_pptx_file(tmp_path, slides)
        decisions = [{"slide_id": 1, "selected_layout": "Content_Light",
                      "decision_path": [], "fallback_used": False, "warnings": []}]
        out_pptx = tmp_path / "out.pptx"
        log_path = tmp_path / "render.log.json"
        render_deck(source, decisions, source, out_pptx, output_log=log_path)
        assert log_path.exists()
        data = json.loads(log_path.read_text())
        assert isinstance(data, list)

    def test_render_missing_decision_gets_warning(self, tmp_path):
        slides = [
            _make_slide_xml(ph_type="title", text="Slide 1"),
            _make_slide_xml(ph_type="title", text="Slide 2"),
        ]
        source = self._make_pptx_file(tmp_path, slides)
        # Only decision for slide 1 — slide 2 has no decision
        decisions = [{"slide_id": 1, "selected_layout": "Content_Light",
                      "decision_path": [], "fallback_used": False, "warnings": []}]
        out_pptx = tmp_path / "out.pptx"
        log = render_deck(source, decisions, source, out_pptx)
        slide2_entry = next(e for e in log if e["slide_id"] == 2)
        assert "no_schema_decision" in slide2_entry["warnings"]

    def test_load_save_roundtrip(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Roundtrip")]
        pptx_bytes = _make_minimal_pptx(slides)
        p = tmp_path / "orig.pptx"
        p.write_bytes(pptx_bytes)

        fm = _load_pptx(p)
        out = tmp_path / "copy.pptx"
        _save_pptx(fm, out)
        fm2 = _load_pptx(out)
        assert set(fm.keys()) == set(fm2.keys())


# ─────────────────────────────────────────────────────────────────────────────
# Layer 4: QA Validator
# ─────────────────────────────────────────────────────────────────────────────

class TestQAValidator:

    def _render_entry(self, slide_id=1, layout="Content_Light",
                      overflow=False, missing=None, warnings=None) -> dict:
        return {
            "slide_id": slide_id,
            "layout_used": layout,
            "content_mapped": {"title": ["Title"], "bullets": [], "subtitle": [], "tables": False, "images": False},
            "overflow_detected": overflow,
            "missing_placeholders": missing or [],
            "warnings": warnings or [],
        }

    def _classifier_entry(self, slide_id=1, density="medium") -> dict:
        return {
            "slide_id": slide_id,
            "type": "content",
            "confidence": 0.80,
            "warnings": [],
            "metadata": {
                "text_density": density,
                "bg_is_dark": False,
                "has_subtitle": False,
                "has_chart": False,
                "image_count": 0,
            },
        }

    def test_pass_on_clean_slide(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Clean Slide")]
        pptx_bytes = _make_minimal_pptx(slides)
        pptx_path = tmp_path / "rendered.pptx"
        pptx_path.write_bytes(pptx_bytes)

        render_log    = [self._render_entry(1)]
        classifier    = [self._classifier_entry(1, density="low")]

        report = validate_render(pptx_path, render_log, classifier)
        assert report["summary"]["total"] == 1
        # Content_Light is a valid FQ layout in our validator — should be fine
        # (layout mismatch may be flagged, status may be review — but not fail)
        assert report["summary"]["fail"] == 0

    def test_overflow_flagged_as_review_high_density(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))

        render_log = [self._render_entry(1, overflow=True)]
        # high density → overflow_risk (review-worthy)
        classifier = [self._classifier_entry(1, density="high")]
        report = validate_render(pptx_path, render_log, classifier)
        slide = report["slides"][0]
        assert "overflow_risk" in slide["issues"]
        assert slide["status"] in ("review", "fail")

    def test_overflow_informational_medium_density(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))

        render_log = [self._render_entry(1, overflow=True)]
        # medium density → overflow_info only (not review)
        classifier = [self._classifier_entry(1, density="medium")]
        report = validate_render(pptx_path, render_log, classifier)
        slide = report["slides"][0]
        assert "overflow_info" in slide["issues"]
        assert "overflow_risk" not in slide["issues"]

    def test_missing_placeholder_issue(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))

        render_log = [self._render_entry(1, missing=["body"])]
        classifier = [self._classifier_entry(1)]
        report = validate_render(pptx_path, render_log, classifier)
        slide = report["slides"][0]
        assert any("missing_placeholder" in i for i in slide["issues"])

    def test_global_slide_count_mismatch(self, tmp_path):
        slides = [
            _make_slide_xml(ph_type="title", text="S1"),
            _make_slide_xml(ph_type="title", text="S2"),
        ]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))

        render_log = [self._render_entry(1)]  # only 1 entry for 2-slide deck
        classifier = [self._classifier_entry(1)]
        report = validate_render(pptx_path, render_log, classifier)
        assert any("slide_count_mismatch" in i for i in report["global_issues"])

    def test_report_summary_keys(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))
        render_log = [self._render_entry(1)]
        classifier = [self._classifier_entry(1)]
        report = validate_render(pptx_path, render_log, classifier)
        assert {"total", "pass", "review", "fail"}.issubset(report["summary"].keys())

    def test_report_written_to_file(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))
        render_log = [self._render_entry(1)]
        classifier = [self._classifier_entry(1)]
        out = tmp_path / "qa.json"
        validate_render(pptx_path, render_log, classifier, output_path=out)
        assert out.exists()
        data = json.loads(out.read_text())
        assert "summary" in data and "slides" in data

    def test_global_cover_layout_check(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide 1")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))
        # Slide 1 assigned a non-Cover layout
        render_log = [self._render_entry(1, layout="Content_Light")]
        classifier = [self._classifier_entry(1)]
        report = validate_render(pptx_path, render_log, classifier)
        assert any("slide_1_not_cover" in i for i in report["global_issues"])

    def test_slide_id_mismatch_flagged(self, tmp_path):
        slides = [_make_slide_xml(ph_type="title", text="Slide")]
        pptx_path = tmp_path / "r.pptx"
        pptx_path.write_bytes(_make_minimal_pptx(slides))
        render_log = [self._render_entry(1)]
        classifier = [self._classifier_entry(99)]  # wrong slide_id
        report = validate_render(pptx_path, render_log, classifier)
        assert any("slide_id_mismatch" in i for i in report["global_issues"])


# ─────────────────────────────────────────────────────────────────────────────
# Integration: layout_rules.json integrity
# ─────────────────────────────────────────────────────────────────────────────

class TestLayoutRulesJson:

    def test_file_exists(self):
        assert LAYOUT_RULES.exists(), "layout_rules.json not found"

    def test_valid_json(self):
        data = json.loads(LAYOUT_RULES.read_text())
        assert isinstance(data, dict)

    def test_required_top_level_keys(self):
        data = json.loads(LAYOUT_RULES.read_text())
        assert "layout_rules" in data
        assert "fallback_layout" in data

    def test_rules_are_prioritized(self):
        data = json.loads(LAYOUT_RULES.read_text())
        rules = data["layout_rules"]
        priorities = [r["priority"] for r in rules]
        # All priorities distinct
        assert len(priorities) == len(set(priorities))

    def test_all_rules_have_required_fields(self):
        data = json.loads(LAYOUT_RULES.read_text())
        for rule in data["layout_rules"]:
            assert "id" in rule
            assert "priority" in rule
            assert "conditions" in rule
            assert "layout" in rule

    def test_cover_dark_highest_priority(self):
        data = json.loads(LAYOUT_RULES.read_text())
        rules = data["layout_rules"]
        by_id = {r["id"]: r for r in rules}
        assert by_id["cover-dark"]["priority"] > by_id["cover-light"]["priority"]
        assert by_id["cover-dark"]["priority"] > by_id["content-default"]["priority"]

    def test_fallback_layout_exists_in_rules(self):
        data = json.loads(LAYOUT_RULES.read_text())
        layouts = {r["layout"] for r in data["layout_rules"]}
        assert data["fallback_layout"] in layouts

    def test_subtitle_upgrade_map_valid(self):
        data = json.loads(LAYOUT_RULES.read_text())
        upgrade_map = data.get("subtitle_upgrade_map", {})
        for src, dst in upgrade_map.items():
            assert isinstance(src, str) and isinstance(dst, str)
