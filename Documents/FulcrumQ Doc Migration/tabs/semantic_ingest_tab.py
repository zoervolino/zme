"""
tabs/semantic_ingest_tab.py — Semantic Ingest UI Tab
=====================================================
Streamlit UI tab for the Semantic Ingest pipeline.

RESPONSIBILITIES (UI ONLY):
  - Upload or select a PPTX file
  - Trigger semantic ingestion via run_semantic_ingest()
  - Show per-slide progress
  - Display slide_id, detected_purpose, structured content preview
  - Allow user to proceed to converter

RULES:
  - No extraction logic
  - No API logic
  - No AgentRelay / WebSocket / webhook code
  - All pipeline calls go through services/semantic_pipeline.py
"""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

import streamlit as st

# Ensure the project root is on sys.path so services.* is importable
# regardless of the working directory Streamlit was launched from.
_PROJECT_ROOT = Path(__file__).parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))


def _load_pipeline():
    try:
        from services.semantic_pipeline import run_semantic_ingest
        return run_semantic_ingest, None
    except Exception as exc:          # catch ImportError AND any init-time errors
        return None, str(exc)


# ── Constants ─────────────────────────────────────────────────────────────────

_TAB_KEY = "semantic_ingest"


# ── Main render function ──────────────────────────────────────────────────────

def render_semantic_ingest_tab() -> None:
    """Render the Semantic Ingest tab. Call this from app.py."""
    st.subheader("Semantic Ingest")
    st.caption(
        "Extract slide content, send it to AnyQuest via AgentRelay, "
        "and produce a structured schema for the deck converter pipeline."
    )

    run_semantic_ingest, import_err = _load_pipeline()
    if import_err:
        st.error(f"Pipeline unavailable — check that `services/` dependencies are installed.")
        with st.expander("Error detail"):
            st.code(import_err)
        return

    # ── Run settings ──────────────────────────────────────────────────────────
    with st.expander("Settings", expanded=False):
        output_dir_override = st.text_input(
            "Artifact output directory (leave blank for default)",
            value="",
            key=f"{_TAB_KEY}_output_dir",
            help="Defaults to artifacts/ alongside the PPTX. Override if needed.",
        )
        retries_input = st.number_input(
            "AgentRelay retries per slide",
            min_value=1,
            max_value=10,
            value=3,
            key=f"{_TAB_KEY}_retries",
            help="Number of retry attempts if AgentRelay is slow to respond.",
        )

    # ── File selection ────────────────────────────────────────────────────────
    st.markdown("#### Select PPTX")
    upload_mode = st.radio(
        "Source",
        ["Upload file", "Use path"],
        horizontal=True,
        key=f"{_TAB_KEY}_upload_mode",
    )

    pptx_path: Path | None = None
    uploaded_name: str = ""

    if upload_mode == "Upload file":
        uploaded = st.file_uploader(
            "Upload PPTX",
            type=["pptx"],
            key=f"{_TAB_KEY}_uploader",
        )
        if uploaded is not None:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
            tmp.write(uploaded.read())
            tmp.flush()
            pptx_path     = Path(tmp.name)
            uploaded_name = uploaded.name
            st.info(f"Loaded: **{uploaded.name}** ({uploaded.size:,} bytes)")
    else:
        path_input = st.text_input(
            "PPTX file path",
            placeholder="/path/to/deck.pptx",
            key=f"{_TAB_KEY}_path_input",
        )
        if path_input:
            candidate = Path(path_input.strip())
            if candidate.is_file() and candidate.suffix.lower() == ".pptx":
                pptx_path     = candidate
                uploaded_name = candidate.name
                st.success(f"File found: {candidate.name}")
            else:
                st.warning("File not found or not a .pptx")

    # ── Run button ────────────────────────────────────────────────────────────
    if st.button("Run Semantic Ingest", type="primary", key=f"{_TAB_KEY}_run_btn"):
        if pptx_path is None:
            st.warning("Please upload a PPTX file or enter a valid file path above.")
        else:
            _run_pipeline(
                run_semantic_ingest=run_semantic_ingest,
                pptx_path=pptx_path,
                output_dir=output_dir_override.strip() or None,
                retries=int(retries_input),
            )

    # ── Show cached results ───────────────────────────────────────────────────
    result_key = f"{_TAB_KEY}_last_result"
    if result_key in st.session_state:
        _render_results(st.session_state[result_key])


# ── Pipeline execution ────────────────────────────────────────────────────────

def _run_pipeline(
    run_semantic_ingest,
    pptx_path: Path,
    output_dir: str | None,
    retries: int,
) -> None:
    """Execute the pipeline and store results in session state."""
    result_key  = f"{_TAB_KEY}_last_result"
    progress_ph = st.empty()
    status_ph   = st.empty()

    progress_ph.progress(0, text="Starting semantic ingestion via AgentRelay…")

    try:
        kwargs: dict = {"retries": retries}
        if output_dir:
            kwargs["output_dir"] = output_dir

        with st.spinner("Submitting slides to AgentRelay and waiting for responses…"):
            schema = run_semantic_ingest(pptx_path, **kwargs)

        progress_ph.progress(100, text="Done.")
        n_slides   = schema.get("slide_count", 0)
        n_failures = len(schema.get("failures", []))

        if n_failures == 0:
            status_ph.success(f"Ingested {n_slides} slides — no failures.")
        else:
            status_ph.warning(
                f"Ingested {n_slides} slides with {n_failures} failure(s). "
                "See Results below."
            )

        st.session_state[result_key] = schema

    except Exception as exc:  # noqa: BLE001
        progress_ph.empty()
        status_ph.error(f"Pipeline error: {exc}")
        st.exception(exc)


# ── Results rendering ─────────────────────────────────────────────────────────

def _render_results(schema: dict) -> None:
    """Render the semantic schema results."""
    st.divider()
    st.markdown("### Results")

    artifacts = schema.get("artifacts", {})
    col1, col2 = st.columns(2)
    with col1:
        raw_path = artifacts.get("raw_extract", "")
        st.metric("Raw extract", Path(raw_path).name if raw_path else "—")
        if raw_path:
            st.caption(raw_path)
    with col2:
        sem_path = artifacts.get("semantic_schema", "")
        st.metric("Semantic schema", Path(sem_path).name if sem_path else "—")
        if sem_path:
            st.caption(sem_path)

    failures = schema.get("failures", [])
    if failures:
        with st.expander(f"Failures ({len(failures)})", expanded=True):
            for f in failures:
                st.error(
                    f"Slide {f.get('slide_id')}: [{f.get('type','?')}] {f.get('error','')}"
                )

    slides = schema.get("slides", [])
    if not slides:
        st.info("No slide results to display.")
        return

    st.markdown(f"#### Slide Results ({len(slides)} slides)")

    for slide in slides:
        slide_id = slide.get("slide_id", "?")
        purpose  = slide.get("detected_purpose", "unknown")
        summary  = slide.get("content_summary", "")
        entities = slide.get("key_entities", [])
        fields   = slide.get("structured_fields", {})
        conf     = slide.get("confidence", 0.0)
        warnings = slide.get("warnings", [])

        conf_color = "green" if conf >= 0.7 else ("orange" if conf >= 0.4 else "red")

        with st.expander(
            f"Slide {slide_id} — {purpose} (conf: {conf:.0%})",
            expanded=False,
        ):
            c1, c2 = st.columns([3, 1])
            with c1:
                st.markdown(f"**Purpose:** {purpose}")
                if summary:
                    st.markdown(f"**Summary:** {summary}")
                if entities:
                    st.markdown("**Key entities:** " + ", ".join(str(e) for e in entities))
            with c2:
                st.markdown(f"**Confidence:** :{conf_color}[{conf:.0%}]")

            if fields:
                st.markdown("**Structured fields:**")
                st.json(fields, expanded=False)

            if warnings:
                for w in warnings:
                    st.warning(w)

    st.divider()
    schema_path = artifacts.get("semantic_schema", "")
    if schema_path and Path(schema_path).is_file():
        st.markdown("#### Proceed to Converter")
        st.info(
            f"Schema saved to `{schema_path}`. "
            "You can now use this schema in the Deck Converter tab."
        )
        if st.button("Show schema path", key=f"{_TAB_KEY}_copy_path"):
            st.code(schema_path)

    st.markdown("#### Download Schema")
    schema_json = json.dumps(schema, indent=2, ensure_ascii=False)
    st.download_button(
        label="Download semantic_schema.json",
        data=schema_json,
        file_name=f"{schema.get('source', 'deck')}_semantic_schema.json",
        mime="application/json",
        key=f"{_TAB_KEY}_dl_btn",
    )
