"""
app.py — FulcrumQ Deck Converter
Streamlit UI for convert_deck.py

Run:
    streamlit run app.py
"""

import base64
import contextlib
import io
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FulcrumQ Deck Converter",
    page_icon=str(Path(__file__).parent / "logo package" / "PNG" / "monogram_default.png"),
    layout="wide",
)

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))
import convert_deck as cd

import os as _os
# Load .env from project dir if present (so ANTHROPIC_API_KEY doesn't need to be set manually)
_env_file = BASE_DIR / ".env"
if _env_file.exists():
    for _line in _env_file.read_text().splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            _os.environ.setdefault(_k.strip(), _v.strip())
try:
    from slide_vision import classify_deck_vision, CONFIDENCE_THRESHOLD as _VIS_THRESHOLD
    _HAS_VISION = True
except ImportError:
    _HAS_VISION = False

from lxml import etree as _ET

# ── Audit helpers ─────────────────────────────────────────────────────────────
_BRAND_COLORS = {
    "765FFF","917FFF","AD9FFF","C8BFFF","E9E4FF",
    "1D1D1D","3B3B3B","585858","767676","7A828D",
    "B2BBCA","D0D7DF","FFFFFF","00C27A","FF2E88",
    "FFB547","60BDBC","281A42",
}
_BRAND_FONTS = {"Segoe UI", "Arial"}
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _audit_source(pptx_bytes: bytes) -> dict:
    """Scan source PPTX and return colors + fonts not covered by any mapping."""
    unmapped_colors: dict[str, int] = {}
    unmapped_fonts:  dict[str, int] = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        slide_keys = [k for k in z.namelist()
                      if re.match(r"ppt/slides/slide\d+\.xml$", k)]
        for key in slide_keys:
            root = _ET.fromstring(z.read(key))
            for srgb in root.iter(f"{{{_NS_A}}}srgbClr"):
                v = srgb.get("val", "").upper()
                if (v and v not in cd.COLOR_REMAP
                        and not cd._is_red_hue(v)
                        and v not in _BRAND_COLORS):
                    unmapped_colors[v] = unmapped_colors.get(v, 0) + 1
            for latin in root.iter(f"{{{_NS_A}}}latin"):
                face = latin.get("typeface", "")
                if (face and face not in ("+mj-lt", "+mn-lt", "")
                        and face not in cd.FONT_MAP
                        and face not in _BRAND_FONTS):
                    unmapped_fonts[face] = unmapped_fonts.get(face, 0) + 1
    return {"colors": unmapped_colors, "fonts": unmapped_fonts}


def _brand_check(pptx_bytes: bytes) -> list[dict]:
    """Scan a PPTX per-slide for off-brand colors, fonts, and CEO Works text.
    Returns list of {slide, colors: {hex: count}, fonts: {face: count}, names: int}.
    """
    issues = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        slide_keys = sorted(
            [k for k in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", k)],
            key=lambda k: int(re.search(r"\d+", k.split("/")[-1]).group()),
        )
        for i, key in enumerate(slide_keys, 1):
            root = _ET.fromstring(z.read(key))
            bad_colors: dict[str, int] = {}
            bad_fonts:  dict[str, int] = {}
            name_count = 0

            for srgb in root.iter(f"{{{_NS_A}}}srgbClr"):
                v = srgb.get("val", "").upper()
                if v and v not in _BRAND_COLORS:
                    bad_colors[v] = bad_colors.get(v, 0) + 1

            for latin in root.iter(f"{{{_NS_A}}}latin"):
                face = latin.get("typeface", "")
                if face and face not in ("+mj-lt", "+mn-lt", "") and face not in _BRAND_FONTS:
                    bad_fonts[face] = bad_fonts.get(face, 0) + 1

            for t_el in root.iter(f"{{{_NS_A}}}t"):
                if t_el.text and cd._CEO_WORKS_RE.search(t_el.text):
                    name_count += 1

            if bad_colors or bad_fonts or name_count:
                issues.append({
                    "slide":  i,
                    "colors": bad_colors,
                    "fonts":  bad_fonts,
                    "names":  name_count,
                })
    return issues


def _render_slide_thumbs(pptx_path: Path, width_px: int = 320) -> list[bytes]:
    """Render PPTX slides → PNG thumbnails via LibreOffice + PyMuPDF.
    Returns [] silently if LibreOffice times out or is unavailable."""
    import fitz
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf",
                 "--outdir", str(tmp_path), str(pptx_path)],
                capture_output=True, timeout=240,
            )
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return []
        pdf_path = tmp_path / (pptx_path.stem + ".pdf")
        if not pdf_path.exists():
            return []
        doc = fitz.open(str(pdf_path))
        thumbs = []
        for page in doc:
            scale = width_px / page.rect.width
            mat   = fitz.Matrix(scale, scale)
            pix   = page.get_pixmap(matrix=mat, alpha=False)
            thumbs.append(pix.tobytes("png"))
        doc.close()
        return thumbs


# ── Brand CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Segoe UI (Windows system font) → Inter as the closest web-safe substitute
   on macOS. Weights loaded: 400 Regular, 600 SemiBold, 700 Bold, 900 Black.
   Arial is the brand body font — always available as a system font. */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');

/* ── Brand tokens — sourced directly from FulcrumQ color guide ── */
:root {
    /* Typography — mirrors master PPTX font assignment:
       Headlines : Segoe UI  (Inter on macOS, exact metric match)
       Body      : Arial     (system font, no substitute needed)
       Mono      : Courier New / Fira Code for code/log blocks        */
    --fq-font-head: 'Segoe UI', 'Inter', Arial, sans-serif;
    --fq-font-body: Arial, sans-serif;
    --fq-font-mono: 'Courier New', 'Fira Code', monospace;
    /* Pivot Purple + tints (15-20% usage) */
    --fq-purple:        #765FFF;
    --fq-tint1:         #917FFF;
    --fq-tint2:         #AD9FFF;
    --fq-tint3:         #C8BFFF;

    /* Guiding Grey + shades (75-80% base, dark UI) */
    --fq-grey:          #1D1D1D;
    --fq-grey1:         #3B3B3B;
    --fq-grey2:         #585858;
    --fq-grey3:         #767676;

    /* Named greys */
    --fq-grey-mid:      #7A828D;
    --fq-light-grey:    #B2BBCA;
    --fq-soft-grey:     #D0D7DF;

    /* Secondary accents (3-5% combined MAX) */
    --fq-vector-dp:     #281A42;
    --fq-green:         #00C27A;
    --fq-magenta:       #FF2E88;
    --fq-amber:         #FFB547;
    --fq-teal:          #60BDBC;

    /* White */
    --fq-white:         #FFFFFF;

    /* Utility */
    --fq-panel:     rgba(255,255,255,0.03);
    --fq-border:    rgba(118,95,255,0.22);
}

/* ── Base ── */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] > div {
    background: var(--fq-grey) !important;
    color: #e8e8f0;
    font-family: var(--fq-font-body);
    font-size: 12px;
    font-weight: 400;
}
[data-testid="stHeader"],
[data-testid="stToolbar"] { display: none !important; }
#MainMenu, footer { visibility: hidden !important; }

.block-container {
    padding-top: 1.5rem !important;
    max-width: 1080px;
}

/* ── Keyframes ── */
@keyframes gradient-pan {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
@keyframes fold-in {
    0% {
        opacity: 0;
        transform: perspective(700px) rotateX(-18deg) translateY(-16px) scale(0.97);
    }
    100% {
        opacity: 1;
        transform: perspective(700px) rotateX(0deg) translateY(0) scale(1);
    }
}
@keyframes fold-in-left {
    0% {
        opacity: 0;
        transform: perspective(700px) rotateY(20deg) translateX(-20px) scale(0.97);
    }
    100% {
        opacity: 1;
        transform: perspective(700px) rotateY(0deg) translateX(0) scale(1);
    }
}
@keyframes border-pulse {
    0%, 100% {
        border-color: rgba(118,95,255,0.35);
        box-shadow: 0 0 0 0 rgba(118,95,255,0);
    }
    50% {
        border-color: rgba(118,95,255,0.8);
        box-shadow: 0 0 22px 4px rgba(118,95,255,0.18);
    }
}
@keyframes fade-up {
    from { opacity: 0; transform: translateY(14px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes ripple-out {
    0%   { transform: scale(0); opacity: 0.5; }
    100% { transform: scale(4.5); opacity: 0; }
}
@keyframes shimmer {
    0%   { background-position: -400px 0; }
    100% { background-position: 400px 0; }
}
@keyframes glow-pulse {
    /* Pivot Purple #765FFF → Signal Magenta #FF2E88 */
    0%, 100% { filter: drop-shadow(0 0 8px rgba(118,95,255,0.6)); }
    50%       { filter: drop-shadow(0 0 18px rgba(255,46,136,0.7)); }
}

/* ── Hero header ── */
.fq-header {
    background: linear-gradient(135deg,
        var(--fq-grey)   0%,
        var(--fq-grey1)  45%,
        var(--fq-purple) 100%
    );
    background-size: 300% 300%;
    animation: gradient-pan 8s ease infinite;
    border-radius: 16px;
    padding: 2rem 2.5rem 1.75rem;
    margin-bottom: 0.25rem;
    border: 1px solid var(--fq-border);
    position: relative;
    overflow: hidden;
}
.fq-header::before {
    content: '';
    position: absolute;
    top: -30%; right: -5%;
    width: 320px; height: 320px;
    /* Signal Magenta #FF2E88 soft glow */
    background: radial-gradient(circle, rgba(255,46,136,0.10) 0%, transparent 68%);
    pointer-events: none;
}
.fq-header::after {
    content: '';
    position: absolute;
    bottom: -40%; left: 20%;
    width: 240px; height: 240px;
    background: radial-gradient(circle, rgba(0,194,122,0.08) 0%, transparent 68%);
    pointer-events: none;
}
.fq-header h1,
.fq-header h1 * {
    font-family: 'Segoe UI', 'Inter', Arial, sans-serif !important;
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 0.35rem;
    letter-spacing: -0.025em;
    background: linear-gradient(90deg, var(--fq-white) 25%, var(--fq-tint2) 75%, var(--fq-tint1) 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
/* Hide Streamlit auto-injected heading anchor */
.fq-header h1 a, .fq-header h1 svg { display: none !important; }

.fq-header p {
    font-family: var(--fq-font-body);
    font-size: 0.82rem;
    font-weight: 400;
    color: rgba(255,255,255,0.48);
    margin: 0;
    letter-spacing: 0.01em;
}
.fq-steps {
    display: flex;
    gap: 0.5rem;
    margin-top: 1rem;
    flex-wrap: wrap;
}
.fq-step-pill {
    font-family: var(--fq-font-head);
    font-weight: 700;
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 20px;
    padding: 3px 10px;
    font-size: 10px;
    color: rgba(255,255,255,0.52);
    letter-spacing: 0.06em;
    transition: background 0.2s, color 0.2s, border-color 0.2s;
    cursor: default;
}
.fq-step-pill:hover {
    background: rgba(118,95,255,0.18);
    border-color: rgba(118,95,255,0.45);
    color: var(--fq-tint3);
}

/* ── Tabs ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid var(--fq-border) !important;
    gap: 4px;
}
[data-testid="stTabs"] [data-baseweb="tab"],
[data-testid="stTabs"] [data-baseweb="tab"] *,
[data-testid="stTabs"] [data-baseweb="tab-list"] button,
[data-testid="stTabs"] [data-baseweb="tab-list"] button * {
    background: transparent !important;
    color: var(--fq-grey3) !important;
    font-family: Arial, sans-serif !important;
    font-size: 0.78rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    padding: 0.55rem 1.3rem !important;
    border-radius: 8px 8px 0 0 !important;
    border: none !important;
    transition: color 0.2s, background 0.2s !important;
    position: relative;
}
[data-testid="stTabs"] [data-baseweb="tab"]::after {
    content: '';
    position: absolute;
    bottom: 0; left: 50%; right: 50%;
    height: 2px;
    background: var(--fq-purple);
    transition: left 0.25s, right 0.25s;
    border-radius: 2px 2px 0 0;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover::after {
    left: 20%; right: 20%;
}
[data-testid="stTabs"] [aria-selected="true"]::after {
    left: 0; right: 0;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover {
    color: rgba(255,255,255,0.85) !important;
    background: rgba(118,95,255,0.08) !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    color: var(--fq-tint2) !important;
    background: rgba(118,95,255,0.1) !important;
}

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
    animation: fold-in 0.45s cubic-bezier(0.23,1,0.32,1) both;
}
[data-testid="stFileUploader"] section {
    background: var(--fq-panel) !important;
    border: 2px dashed rgba(118,95,255,0.38) !important;
    border-radius: 14px !important;
    padding: 2.2rem 1.5rem !important;
    animation: border-pulse 4s ease-in-out infinite !important;
    transition: background 0.25s, transform 0.25s, border-color 0.25s !important;
}
[data-testid="stFileUploader"] section:hover {
    background: rgba(118,95,255,0.05) !important;
    transform: scale(1.012) !important;
    border-color: rgba(118,95,255,0.7) !important;
    animation: none !important;
    box-shadow: 0 0 28px rgba(118,95,255,0.15) !important;
}
[data-testid="stFileUploaderDropzone"] {
    color: var(--fq-tint2) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: rgba(255,255,255,0.5) !important;
    font-size: 0.82rem !important;
}

/* ── Primary button — Guiding Grey → Pivot Purple ── */
[data-testid="stBaseButton-primary"],
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--fq-grey) 0%, var(--fq-purple) 100%) !important;
    border: none !important;
    border-radius: 9px !important;
    color: var(--fq-white) !important;
    font-family: Arial, sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
    padding: 0.6rem 1.6rem !important;
    position: relative;
    overflow: hidden;
    transition: transform 0.18s, box-shadow 0.18s !important;
    will-change: transform;
}
[data-testid="stBaseButton-primary"]:hover,
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 10px 30px rgba(118,95,255,0.45) !important;
}
[data-testid="stBaseButton-primary"]:active,
.stButton > button[kind="primary"]:active {
    transform: translateY(0) scale(0.96) !important;
    box-shadow: 0 4px 12px rgba(118,95,255,0.3) !important;
}

/* ── Download button — Anchor Green #00C27A ── */
.stDownloadButton > button {
    background: transparent !important;
    border: 1.5px solid var(--fq-green) !important;
    color: var(--fq-green) !important;
    border-radius: 9px !important;
    font-family: Arial, sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
    transition: background 0.22s, transform 0.18s, box-shadow 0.22s !important;
}
.stDownloadButton > button * {
    font-family: Arial, sans-serif !important;
    color: inherit !important;
}
.stDownloadButton > button:hover {
    background: rgba(0,194,122,0.08) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 24px rgba(0,194,122,0.25) !important;
}
.stDownloadButton > button:active {
    transform: translateY(0) scale(0.96) !important;
}

/* ── Alerts ── */
[data-testid="stAlert"][data-baseweb="notification"] {
    background: rgba(0,194,122,0.06) !important;
    border: 1px solid rgba(0,194,122,0.26) !important;
    border-radius: 10px !important;
    animation: fold-in 0.38s ease both !important;
}
[data-testid="stAlert"][kind="error"] {
    background: rgba(255,46,136,0.06) !important;
    border: 1px solid rgba(255,46,136,0.26) !important;
}

/* ── Log terminal ── */
.log-box {
    background: #0a0a12;
    border: 1px solid rgba(118,95,255,0.18);
    border-radius: 12px;
    padding: 1.1rem 1.3rem;
    font-family: var(--fq-font-mono);
    font-size: 11px;
    white-space: pre-wrap;
    max-height: 360px;
    overflow-y: auto;
    color: #b8b8cc;
    line-height: 1.65;
    animation: fold-in 0.45s cubic-bezier(0.23,1,0.32,1) both;
    position: relative;
}
.log-box::before {
    content: '● ● ●';
    display: block;
    color: rgba(118,95,255,0.4);
    font-size: 10px;
    letter-spacing: 3px;
    margin-bottom: 0.6rem;
    font-family: monospace;
}
.log-box::-webkit-scrollbar { width: 5px; }
.log-box::-webkit-scrollbar-track { background: transparent; }
.log-box::-webkit-scrollbar-thumb {
    background: rgba(118,95,255,0.3);
    border-radius: 3px;
}

/* ── Section label ── */
.section-label {
    font-family: var(--fq-font-head);
    font-size: 9px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: var(--fq-grey3);
    margin: 1.5rem 0 0.5rem;
}

/* ── Color swatches ── */
.swatch-row {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 3px 0;
    font-family: var(--fq-font-mono);
    font-size: 11px;
    color: rgba(255,255,255,0.62);
    padding: 5px 8px;
    border-radius: 7px;
    transition: background 0.22s, transform 0.22s, color 0.22s;
    cursor: default;
}
.swatch-row:hover {
    background: rgba(118,95,255,0.1);
    transform: translateX(6px);
    color: rgba(255,255,255,0.88);
}
.swatch {
    width: 19px; height: 19px;
    border-radius: 50%;
    display: inline-block;
    flex-shrink: 0;
    border: 1px solid rgba(255,255,255,0.1);
    transition: transform 0.22s, box-shadow 0.22s;
}
.swatch-row:hover .swatch {
    transform: scale(1.35);
    box-shadow: 0 2px 10px rgba(0,0,0,0.5);
}
.arrow {
    color: rgba(118,95,255,0.55);
    font-weight: 700;
    transition: color 0.22s, transform 0.22s;
    display: inline-block;
}
.swatch-row:hover .arrow {
    color: var(--fq-tint1);
    transform: translateX(2px);
}

/* ── Master badge ── */
.master-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 14px;
    border-radius: 20px;
    font-family: var(--fq-font-head);
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.03em;
    animation: fold-in-left 0.4s ease both;
    transition: box-shadow 0.2s;
    cursor: default;
}
.master-ok {
    background: rgba(0,194,122,0.08);
    border: 1px solid rgba(0,194,122,0.28);
    color: var(--fq-green);
}
.master-ok:hover { box-shadow: 0 0 16px rgba(0,194,122,0.18); }
.master-err {
    background: rgba(255,46,136,0.07);
    border: 1px solid rgba(255,46,136,0.28);
    color: var(--fq-magenta);
}

/* ── Horizontal rule ── */
hr { border-color: rgba(118,95,255,0.12) !important; }

/* ── Spinner ── */
[data-testid="stSpinner"] svg { color: var(--fq-purple) !important; }

/* ── Scheme fill table ── */
.scheme-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 4px 0;
    font-family: var(--fq-font-mono);
    font-size: 11px;
    color: rgba(255,255,255,0.62);
    padding: 5px 8px;
    border-radius: 7px;
    transition: background 0.22s, transform 0.22s;
    cursor: default;
}
.scheme-row:hover {
    background: rgba(118,95,255,0.1);
    transform: translateX(6px);
    color: rgba(255,255,255,0.9);
}

/* ── Hero download box ── */
.fq-hero-dl {
    display: inline-flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 4px;
    padding: 14px 18px;
    border: 1.5px solid rgba(0,194,122,0.55);
    border-radius: 12px;
    background: rgba(0,194,122,0.07);
    text-decoration: none;
    transition: background 0.2s, border-color 0.2s, box-shadow 0.2s;
    cursor: pointer;
    white-space: nowrap;
    box-shadow: 0 0 0 3px rgba(0,194,122,0.12), 0 0 28px rgba(0,194,122,0.18);
}
.fq-hero-dl:hover {
    background: rgba(0,194,122,0.13);
    border-color: rgba(0,194,122,0.85);
    box-shadow: 0 0 0 3px rgba(0,194,122,0.2), 0 0 40px rgba(0,194,122,0.3);
}
.fq-hero-dl-label {
    font-family: Arial, sans-serif;
    font-size: 9px;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #00C27A;
}
.fq-hero-dl-name {
    font-family: 'Segoe UI', 'Inter', Arial, sans-serif;
    font-size: 1rem;
    font-weight: 700;
    color: #ffffff;
    line-height: 1.2;
}
.fq-hero-dl-ext {
    font-family: Arial, sans-serif;
    font-size: 9px;
    font-weight: 400;
    color: rgba(255,255,255,0.45);
    letter-spacing: 0.06em;
}

/* ── Pivot line — white → Pivot Purple ── */
.fq-pivot-line {
    height: 3px;
    background: linear-gradient(90deg,
        rgba(255,255,255,0.9) 0%,
        rgba(255,255,255,0.6) 30%,
        var(--fq-purple)     66%,
        var(--fq-tint1)      100%
    );
    border-radius: 2px;
    margin: 1.25rem 0;
}

/* ── Icon grid cells ── */
.icon-cell {
    background: rgba(118,95,255,0.04);
    border: 1px solid rgba(118,95,255,0.14);
    border-radius: 10px;
    padding: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.18s, border-color 0.18s, box-shadow 0.18s;
    cursor: default;
}
.icon-cell:hover {
    background: rgba(118,95,255,0.10);
    border-color: rgba(118,95,255,0.40);
    box-shadow: 0 0 14px rgba(118,95,255,0.18);
}

/* ── Unmapped audit table ── */
.audit-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: monospace; }
.audit-table th {
    font-family: Arial, sans-serif; font-size: 9px; font-weight: 700;
    letter-spacing: 0.12em; text-transform: uppercase;
    color: var(--fq-grey3); padding: 4px 8px; text-align: left;
    border-bottom: 1px solid var(--fq-border);
}
.audit-table td { padding: 4px 8px; color: rgba(255,255,255,0.65); }
.audit-table tr:hover td { background: rgba(118,95,255,0.06); }
.audit-dot { display:inline-block; width:12px; height:12px; border-radius:50%;
             vertical-align:middle; margin-right:6px; border:1px solid rgba(255,255,255,0.15); }

/* ── Slide preview thumbnails ── */
.slide-thumb {
    border-radius: 8px;
    border: 1px solid var(--fq-border);
    overflow: hidden;
    transition: box-shadow 0.18s, transform 0.18s;
    cursor: default;
}
.slide-thumb:hover {
    transform: translateY(-3px);
    box-shadow: 0 10px 28px rgba(0,0,0,0.5);
}
.slide-thumb img { display: block; width: 100%; }
.slide-thumb-num {
    font-family: Arial, sans-serif; font-size: 9px;
    color: rgba(255,255,255,0.3); text-align: center; padding: 3px 0 4px;
    background: rgba(0,0,0,0.25);
}

/* ── Typography specimen ── */
.type-row {
    display: flex; align-items: baseline; gap: 16px;
    padding: 14px 0; border-bottom: 1px solid rgba(118,95,255,0.10);
}
.type-meta {
    width: 200px; flex-shrink: 0;
    font-family: Arial, sans-serif; font-size: 9px;
    color: var(--fq-grey3); line-height: 1.8;
}
.type-meta strong { display: block; color: rgba(255,255,255,0.6); font-size: 10px; margin-bottom: 2px; }
.type-sample { flex: 1; line-height: 1.2; }

/* ── Logo tile ── */
.logo-tile {
    border-radius: 14px;
    padding: 28px 20px 16px;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 12px;
    border: 1px solid rgba(255,255,255,0.07);
    transition: transform 0.2s cubic-bezier(0.23,1,0.32,1),
                box-shadow 0.2s cubic-bezier(0.23,1,0.32,1);
}
.logo-tile:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 32px rgba(0,0,0,0.4);
}
.logo-variant-label {
    font-family: var(--fq-font-head);
    font-size: 9px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-top: 4px;
}

/* ── Expanders — force dark background so content is readable ── */
[data-testid="stExpander"] {
    background: #111118 !important;
    border: 1px solid var(--fq-border) !important;
    border-radius: 10px !important;
}
[data-testid="stExpander"] summary {
    color: rgba(255,255,255,0.75) !important;
    font-family: Arial, sans-serif !important;
    font-size: 0.82rem !important;
}
[data-testid="stExpander"] summary:hover {
    color: #ffffff !important;
}
[data-testid="stExpander"] > div[data-testid="stExpanderDetails"] {
    background: #111118 !important;
}

/* ── Kill Streamlit red — radio dot only ── */
[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
    background-color: #FF2E88 !important;
}
/* Target only the inner filled dot of the selected radio button */
[data-baseweb="radio"] [data-checked="true"] > div > div,
[data-baseweb="radio"] [role="radio"][aria-checked="true"] > div > div {
    background-color: #FF2E88 !important;
    border-color: #FF2E88 !important;
}
</style>
""", unsafe_allow_html=True)

# ── JS: ripple, tilt, fold ────────────────────────────────────────────────────
components.html("""
<script>
(function() {
    const doc = window.parent.document;

    function attachRipple(btn) {
        if (btn._fqRipple) return;
        btn._fqRipple = true;
        btn.style.position = 'relative';
        btn.style.overflow = 'hidden';
        btn.addEventListener('pointerdown', function(e) {
            const rect = btn.getBoundingClientRect();
            const d = Math.max(btn.offsetWidth, btn.offsetHeight);
            const r = doc.createElement('span');
            r.style.cssText = [
                'position:absolute','border-radius:50%',
                `width:${d}px`,`height:${d}px`,
                'background:rgba(255,255,255,0.22)',
                'transform:scale(0)','animation:ripple-out 0.58s linear',
                'pointer-events:none',
                `left:${e.clientX - rect.left - d/2}px`,
                `top:${e.clientY - rect.top - d/2}px`,
            ].join(';');
            btn.appendChild(r);
            setTimeout(() => r && r.remove(), 620);
        });
    }

    function attachTilt(el) {
        if (el._fqTilt) return;
        el._fqTilt = true;
        el.addEventListener('mousemove', function(e) {
            const rect = el.getBoundingClientRect();
            const cx = rect.left + rect.width / 2;
            const cy = rect.top  + rect.height / 2;
            const dx = (e.clientX - cx) / (rect.width  / 2);
            const dy = (e.clientY - cy) / (rect.height / 2);
            el.style.transform = `perspective(700px) rotateY(${dx * 4}deg) rotateX(${-dy * 4}deg) scale(1.012)`;
            el.style.transition = 'transform 0.08s linear';
        });
        el.addEventListener('mouseleave', function() {
            el.style.transform = 'perspective(700px) rotateY(0deg) rotateX(0deg) scale(1)';
            el.style.transition = 'transform 0.4s cubic-bezier(0.23,1,0.32,1)';
        });
    }

    function attachFoldIn(el) {
        if (el._fqFold) return;
        el._fqFold = true;
        el.style.animation = 'fold-in 0.45s cubic-bezier(0.23,1,0.32,1) both';
    }

    const mo = new MutationObserver(function() {
        doc.querySelectorAll(
            '[data-testid="stBaseButton-primary"], .stButton > button[kind="primary"]'
        ).forEach(attachRipple);
        doc.querySelectorAll('.stDownloadButton > button').forEach(attachRipple);
        doc.querySelectorAll('[data-testid="stFileUploader"] section').forEach(attachTilt);
        doc.querySelectorAll('[data-testid="stAlert"]').forEach(attachFoldIn);
    });
    mo.observe(doc.body, { childList: true, subtree: true });
})();
</script>
""", height=0)

# ── Hero header ───────────────────────────────────────────────────────────────
_logo_path = BASE_DIR / "logo package" / "PNG" / "primary logo + tagline_reverse.png"
_logo_b64  = base64.b64encode(_logo_path.read_bytes()).decode() if _logo_path.exists() else ""
_logo_img  = (
    f'<img src="data:image/png;base64,{_logo_b64}" '
    f'style="height:38px;width:auto;object-fit:contain;margin-bottom:1rem;display:block;">'
    if _logo_b64 else ""
)

_potx_path   = BASE_DIR / "FulcrumQ_Theme_vF.potx"
_potx_b64    = base64.b64encode(_potx_path.read_bytes()).decode() if _potx_path.exists() else ""
_potx_button = (
    f'<a class="fq-hero-dl" '
    f'href="data:application/vnd.openxmlformats-officedocument.presentationml.template;base64,{_potx_b64}" '
    f'download="FulcrumQ_Template.potx">'
    f'<span class="fq-hero-dl-label">Download template</span>'
    f'<span class="fq-hero-dl-name">FulcrumQ Master</span>'
    f'<span class="fq-hero-dl-ext">.POTX</span>'
    f'</a>'
    if _potx_b64 else ""
)

st.markdown(f"""
<div class="fq-header" style="display:flex;justify-content:space-between;align-items:flex-start;">
    <div>
        {_logo_img}
        <h1>Deck Converter</h1>
        <p>Transform any source deck to FulcrumQ brand in seconds.</p>
        <div class="fq-steps">
            <span class="fq-step-pill">① Master swap</span>
            <span class="fq-step-pill">② Layout remap</span>
            <span class="fq-step-pill">③ Vestige cleanup</span>
            <span class="fq-step-pill">④ Brand style pass</span>
        </div>
    </div>
    {_potx_button}
</div>
<div class="fq-pivot-line"></div>
""", unsafe_allow_html=True)

tab_convert, tab_palette, tab_logos, tab_icons, tab_type = st.tabs(
    ["  Convert  ", "  Color Palette  ", "  Logo Suite  ", "  Icons  ", "  Typography  "]
)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — CONVERT
# ══════════════════════════════════════════════════════════════════════════════
with tab_convert:
    _cmode = st.radio(
        "_cmode",
        ["Convert", "Brand Check"],
        horizontal=True,
        label_visibility="collapsed",
    )
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # ── Convert mode ──────────────────────────────────────────────────────────
    if _cmode == "Convert":
        master_exists = cd.MASTER_X.exists() or cd.MASTER_PPTX.exists()

        uploaded = st.file_uploader(
            "Upload source PPTX",
            type=["pptx"],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        if uploaded:
            st.markdown("<br>", unsafe_allow_html=True)

            # AI vision toggle
            _api_key_set = bool(_os.environ.get("ANTHROPIC_API_KEY"))
            _vision_help = (
                "Runs a two-pass AI vision classifier on each slide after conversion. "
                "Requires ANTHROPIC_API_KEY + LibreOffice."
                if _HAS_VISION and _api_key_set
                else "Set ANTHROPIC_API_KEY to enable AI classification."
            )
            use_vision = st.toggle(
                "AI slide classification",
                value=False,
                disabled=not (_HAS_VISION and _api_key_set),
                help=_vision_help,
            )

            col_btn, _ = st.columns([1, 4])
            with col_btn:
                run = st.button(
                    f"▶  Convert  ({len(uploaded)} file{'s' if len(uploaded) != 1 else ''})",
                    type="primary",
                    disabled=not master_exists,
                    use_container_width=True,
                )

            if run:
                with tempfile.TemporaryDirectory() as tmp:
                    tmp_path = Path(tmp)
                    results: list[tuple] = []

                    for uf in uploaded:
                        src_bytes = uf.getvalue()
                        src_path  = tmp_path / uf.name
                        src_path.write_bytes(src_bytes)

                        audit = _audit_source(src_bytes)

                        log_buf = io.StringIO()
                        with st.spinner(f"Converting {uf.name}…"):
                            with contextlib.redirect_stdout(log_buf):
                                try:
                                    out_path = cd.convert(src_path)
                                except Exception as exc:
                                    results.append((uf.name, None, str(exc), audit, [], None))
                                    continue

                        with st.spinner(f"Rendering preview…"):
                            thumbs = _render_slide_thumbs(out_path)

                        # AI vision classification (runs on source file)
                        vis_results = None
                        if use_vision:
                            with st.spinner(f"AI classifying {uf.name}…"):
                                try:
                                    vis_results = classify_deck_vision(src_path)
                                except Exception as _ve:
                                    st.warning(f"Vision classifier failed: {_ve}")

                        results.append((uf.name, out_path, None, audit, thumbs, vis_results))

                        logs = log_buf.getvalue()
                        with st.expander(f"Log — {uf.name}", expanded=False):
                            st.markdown(
                                f'<div class="log-box">{logs}</div>',
                                unsafe_allow_html=True,
                            )

                    n_ok  = sum(1 for _, o, *_ in results if o and o.exists())
                    n_err = len(results) - n_ok
                    if n_ok:
                        st.success(f"✓  {n_ok}/{len(results)} converted")
                    if n_err:
                        st.error(f"✗  {n_err} failed")

                    # Badge colours per slide type
                    _VIS_BADGE = {
                        "cover":   ("#765FFF", "#fff"),
                        "divider": ("#281A42", "#fff"),
                        "ending":  ("#00C27A", "#fff"),
                        "content": ("#E9E4FF", "#281A42"),
                    }

                    for orig_name, out_path, err, audit, thumbs, vis_results in results:
                        if err or not (out_path and out_path.exists()):
                            st.error(f"{orig_name}: {err}")
                            continue

                        st.download_button(
                            label=f"⬇  {out_path.name}",
                            data=out_path.read_bytes(),
                            file_name=out_path.name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"dl_{out_path.name}",
                        )

                        # ── Slide preview ──────────────────────────────────
                        if thumbs:
                            st.markdown(
                                '<div class="section-label" style="margin-top:1rem;">Slide preview</div>',
                                unsafe_allow_html=True,
                            )
                            # Build vision lookup {slide_num: result}
                            _vis_map = {r["slide"]: r for r in (vis_results or [])}
                            _THUMB_COLS = 6
                            for _ri in range(0, len(thumbs), _THUMB_COLS):
                                _row = thumbs[_ri: _ri + _THUMB_COLS]
                                _cells = ""
                                for _ti, _tb in enumerate(_row, _ri + 1):
                                    _tb64 = base64.b64encode(_tb).decode()
                                    _vr   = _vis_map.get(_ti)
                                    if _vr:
                                        _bg, _fg = _VIS_BADGE.get(_vr["type"], ("#ccc", "#000"))
                                        _flag    = " ⚑" if _vr["confidence"] < _VIS_THRESHOLD else ""
                                        _badge   = (
                                            f'<div style="font-size:9px;font-weight:600;'
                                            f'background:{_bg};color:{_fg};border-radius:3px;'
                                            f'padding:1px 5px;margin-top:3px;text-align:center;">'
                                            f'{_vr["type"].upper()}{_flag}</div>'
                                        )
                                    else:
                                        _badge = ""
                                    _cells += (
                                        f'<div class="slide-thumb">'
                                        f'<img src="data:image/png;base64,{_tb64}">'
                                        f'<div class="slide-thumb-num">{_ti}</div>'
                                        f'{_badge}'
                                        f'</div>'
                                    )
                                st.markdown(
                                    f'<div style="display:grid;grid-template-columns:repeat({_THUMB_COLS},1fr);'
                                    f'gap:8px;margin-bottom:8px;">{_cells}</div>',
                                    unsafe_allow_html=True,
                                )

                        # ── Vision classification detail ────────────────────
                        if vis_results:
                            with st.expander("AI classification detail", expanded=False):
                                _low = [r for r in vis_results if r["confidence"] < _VIS_THRESHOLD]
                                if _low:
                                    st.warning(f"{len(_low)} slide(s) flagged for review (confidence < {_VIS_THRESHOLD})")
                                _rows = "".join(
                                    f"<tr>"
                                    f"<td>{r['slide']}</td>"
                                    f"<td><b>{r['type']}</b></td>"
                                    f"<td>{r['confidence']:.2f}</td>"
                                    f"<td>{'↺' if r.get('changed') else ''}</td>"
                                    f"<td style='color:#888;font-size:12px'>{r['reasoning'][:80]}</td>"
                                    f"</tr>"
                                    for r in vis_results
                                )
                                st.markdown(
                                    f"<table style='width:100%;font-size:13px;border-collapse:collapse'>"
                                    f"<thead><tr><th>#</th><th>Type</th><th>Conf</th><th></th><th>Reasoning</th></tr></thead>"
                                    f"<tbody>{_rows}</tbody></table>",
                                    unsafe_allow_html=True,
                                )

                        # ── Unmapped report ────────────────────────────────
                        uc  = audit["colors"]
                        uf2 = audit["fonts"]
                        if uc or uf2:
                            with st.expander(
                                f"Unmapped items — {len(uc)} color(s), {len(uf2)} font(s)",
                                expanded=True,
                            ):
                                col_c, col_f = st.columns(2)
                                with col_c:
                                    st.markdown(
                                        '<div class="section-label">Colors not mapped</div>',
                                        unsafe_allow_html=True,
                                    )
                                    if uc:
                                        rows = "".join(
                                            f'<tr><td><span class="audit-dot" style="background:#{h};"></span>#{h}</td>'
                                            f'<td style="color:rgba(255,255,255,0.35);">×{cnt}</td></tr>'
                                            for h, cnt in sorted(uc.items(), key=lambda x: -x[1])
                                        )
                                        st.markdown(
                                            f'<table class="audit-table"><thead><tr><th>Hex</th><th>Uses</th></tr>'
                                            f'</thead><tbody>{rows}</tbody></table>',
                                            unsafe_allow_html=True,
                                        )
                                    else:
                                        st.markdown(
                                            '<span style="font-size:11px;color:var(--fq-green);">All colors mapped ✓</span>',
                                            unsafe_allow_html=True,
                                        )
                                with col_f:
                                    st.markdown(
                                        '<div class="section-label">Fonts not mapped</div>',
                                        unsafe_allow_html=True,
                                    )
                                    if uf2:
                                        rows = "".join(
                                            f'<tr><td style="font-family:\'{f}\',sans-serif;">{f}</td>'
                                            f'<td style="color:rgba(255,255,255,0.35);">×{cnt}</td></tr>'
                                            for f, cnt in sorted(uf2.items(), key=lambda x: -x[1])
                                        )
                                        st.markdown(
                                            f'<table class="audit-table"><thead><tr><th>Typeface</th><th>Uses</th></tr>'
                                            f'</thead><tbody>{rows}</tbody></table>',
                                            unsafe_allow_html=True,
                                        )
                                    else:
                                        st.markdown(
                                            '<span style="font-size:11px;color:var(--fq-green);">All fonts mapped ✓</span>',
                                            unsafe_allow_html=True,
                                        )

    # ── Brand Check mode ──────────────────────────────────────────────────────
    else:
        st.markdown(
            '<div class="section-label" style="margin-bottom:0.5rem;">'
            'Upload a deck to flag off-brand colors, fonts, and CEO Works references per slide</div>',
            unsafe_allow_html=True,
        )
        bc_uploaded = st.file_uploader(
            "Upload PPTX for brand check",
            type=["pptx"],
            accept_multiple_files=False,
            label_visibility="collapsed",
            key="bc_upload",
        )
        if bc_uploaded:
            col_bc, _ = st.columns([1, 4])
            with col_bc:
                bc_run = st.button(
                    "Check Brand Compliance",
                    type="primary",
                    use_container_width=True,
                    key="bc_run",
                )
            if bc_run:
                with st.spinner("Scanning slides…"):
                    bc_issues = _brand_check(bc_uploaded.getvalue())

                with zipfile.ZipFile(io.BytesIO(bc_uploaded.getvalue())) as _z:
                    total_slides = sum(
                        1 for k in _z.namelist()
                        if re.match(r"ppt/slides/slide\d+\.xml$", k)
                    )

                clean = total_slides - len(bc_issues)
                st.markdown(
                    f'<div style="display:flex;gap:12px;margin:0.8rem 0;">'
                    f'<span class="master-badge master-ok">✓ {clean} slide{"s" if clean != 1 else ""} clean</span>'
                    f'<span class="master-badge master-err">✗ {len(bc_issues)} slide{"s" if len(bc_issues) != 1 else ""} flagged</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                if not bc_issues:
                    st.success("All slides pass brand check.")
                else:
                    for issue in bc_issues:
                        snum = issue["slide"]
                        n_c  = len(issue["colors"])
                        n_f  = len(issue["fonts"])
                        n_n  = issue["names"]
                        parts = []
                        if n_c: parts.append(f"{n_c} off-brand color{'s' if n_c != 1 else ''}")
                        if n_f: parts.append(f"{n_f} off-brand font{'s' if n_f != 1 else ''}")
                        if n_n: parts.append(f"{n_n} 'CEO Works' ref{'s' if n_n != 1 else ''}")
                        label = f"Slide {snum} — {', '.join(parts)}"

                        with st.expander(label, expanded=False):
                            col_c2, col_f2, col_n2 = st.columns(3)
                            with col_c2:
                                st.markdown(
                                    '<div class="section-label">Off-brand colors</div>',
                                    unsafe_allow_html=True,
                                )
                                if issue["colors"]:
                                    rows_c = "".join(
                                        f'<tr><td><span class="audit-dot" style="background:#{h};"></span>#{h}</td>'
                                        f'<td style="color:rgba(255,255,255,0.35);">×{cnt}</td></tr>'
                                        for h, cnt in sorted(issue["colors"].items(), key=lambda x: -x[1])
                                    )
                                    st.markdown(
                                        f'<table class="audit-table"><thead><tr><th>Hex</th><th>Uses</th></tr>'
                                        f'</thead><tbody>{rows_c}</tbody></table>',
                                        unsafe_allow_html=True,
                                    )
                                else:
                                    st.markdown(
                                        '<span style="font-size:11px;color:var(--fq-green);">None ✓</span>',
                                        unsafe_allow_html=True,
                                    )
                            with col_f2:
                                st.markdown(
                                    '<div class="section-label">Off-brand fonts</div>',
                                    unsafe_allow_html=True,
                                )
                                if issue["fonts"]:
                                    rows_f = "".join(
                                        f'<tr><td style="font-family:\'{f}\',sans-serif;">{f}</td>'
                                        f'<td style="color:rgba(255,255,255,0.35);">×{cnt}</td></tr>'
                                        for f, cnt in sorted(issue["fonts"].items(), key=lambda x: -x[1])
                                    )
                                    st.markdown(
                                        f'<table class="audit-table"><thead><tr><th>Typeface</th><th>Uses</th></tr>'
                                        f'</thead><tbody>{rows_f}</tbody></table>',
                                        unsafe_allow_html=True,
                                    )
                                else:
                                    st.markdown(
                                        '<span style="font-size:11px;color:var(--fq-green);">None ✓</span>',
                                        unsafe_allow_html=True,
                                    )
                            with col_n2:
                                st.markdown(
                                    '<div class="section-label">Brand name issues</div>',
                                    unsafe_allow_html=True,
                                )
                                if n_n:
                                    st.markdown(
                                        f'<span style="font-size:11px;color:var(--fq-magenta);">'
                                        f'{n_n}x "CEO Works" — run Convert to fix</span>',
                                        unsafe_allow_html=True,
                                    )
                                else:
                                    st.markdown(
                                        '<span style="font-size:11px;color:var(--fq-green);">None ✓</span>',
                                        unsafe_allow_html=True,
                                    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — COLOR PALETTE
# ══════════════════════════════════════════════════════════════════════════════
with tab_palette:

    pal_view = st.radio(
        "palette_view",
        ["Brand Palette", "Conversion Map"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if pal_view == "Brand Palette":

        BRAND_GROUPS = [
            ("Pivot Purple", [
                ("765FFF", "Pivot Purple"),
                ("917FFF", "Tint 1"),
                ("AD9FFF", "Tint 2"),
                ("C8BFFF", "Tint 3"),
                ("E9E4FF", "Light Lavender"),
            ]),
            ("Accents", [
                ("00C27A", "Anchor Green"),
                ("FFB547", "Amber"),
                ("60BDBC", "Teal"),
                ("FF2E88", "Signal Magenta"),
                ("C8BFFF", "Pivot Purple Tint 3"),
            ]),
            ("Dark", [
                ("281A42", "Vector Dark Purple"),
                ("1D1D1D", "Guiding Grey"),
                ("3B3B3B", "Grey 1"),
                ("585858", "Grey 2"),
                ("767676", "Grey 3"),
            ]),
            ("Mid / Light", [
                ("7A828D", "Grey Mid"),
                ("B2BBCA", "Light Grey"),
                ("D0D7DF", "Soft Grey"),
                ("E9E4FF", "Light Lavender"),
                ("FFFFFF", "White"),
            ]),
        ]

        def _text_color(hex_val: str) -> str:
            try:
                r = int(hex_val[0:2], 16)
                g = int(hex_val[2:4], 16)
                b = int(hex_val[4:6], 16)
                lum = 0.299 * r + 0.587 * g + 0.114 * b
                return "#1D1D1D" if lum > 140 else "#FFFFFF"
            except Exception:
                return "#FFFFFF"

        st.markdown("""
<style>
.pal-group-label {
    font-family: var(--fq-font-head);
    font-size: 9px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.14em;
    color: var(--fq-grey3); margin: 1.4rem 0 0.6rem;
}
.pal-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 8px;
}
.pal-tile {
    border-radius: 12px; height: 96px;
    display: flex; flex-direction: column;
    justify-content: flex-end; padding: 8px 10px;
    cursor: crosshair; border: 1px solid rgba(255,255,255,0.07);
    transition: transform 0.18s cubic-bezier(0.23,1,0.32,1),
                box-shadow 0.18s cubic-bezier(0.23,1,0.32,1);
    position: relative; overflow: hidden;
}
.pal-tile:hover {
    transform: translateY(-4px) scale(1.04);
    box-shadow: 0 12px 32px rgba(0,0,0,0.45);
    z-index: 2; border-color: rgba(255,255,255,0.22);
}
.pal-tile-hex {
    font-family: var(--fq-font-mono); font-size: 10px;
    font-weight: 700; letter-spacing: 0.04em; line-height: 1.2;
}
.pal-tile-name {
    font-family: var(--fq-font-body); font-size: 9px;
    opacity: 0.72; line-height: 1.2; margin-top: 1px;
}
</style>
""", unsafe_allow_html=True)

        seen = set()
        for group_label, colors in BRAND_GROUPS:
            st.markdown(
                f'<div class="pal-group-label">{group_label}</div>',
                unsafe_allow_html=True,
            )
            tiles = ""
            for hex_val, name in colors:
                if hex_val in seen:
                    continue
                seen.add(hex_val)
                tc = _text_color(hex_val)
                border = "1px solid rgba(0,0,0,0.15)" if hex_val == "FFFFFF" else "1px solid rgba(255,255,255,0.07)"
                tiles += (
                    f'<div class="pal-tile" style="background:#{hex_val};border:{border};">'
                    f'<span class="pal-tile-hex" style="color:{tc};">#{hex_val}</span>'
                    f'<span class="pal-tile-name" style="color:{tc};">{name}</span>'
                    f'</div>'
                )
            st.markdown(f'<div class="pal-grid">{tiles}</div>', unsafe_allow_html=True)

    else:
        # Group every COLOR_REMAP entry by its target brand color so the map
        # is always in sync with convert_deck.py — no hardcoded lists needed.
        _TARGET_ORDER = [
            ("1D1D1D", "→ Guiding Grey"),
            ("3B3B3B", "→ Grey 1"),
            ("585858", "→ Grey 2"),
            ("767676", "→ Grey 3"),
            ("7A828D", "→ Grey Mid"),
            ("B2BBCA", "→ Light Grey"),
            ("D0D7DF", "→ Soft Grey"),
            ("FFFFFF",  "→ White"),
            ("281A42", "→ Vector Dark Purple"),
            ("765FFF", "→ Pivot Purple"),
            ("917FFF", "→ Purple Tint 1"),
            ("AD9FFF", "→ Purple Tint 2"),
            ("C8BFFF", "→ Purple Tint 3"),
            ("E9E4FF", "→ Shift Lavender"),
            ("00C27A", "→ Anchor Green"),
            ("60BDBC", "→ Teal"),
            ("FFB547", "→ Amber"),
            ("FF2E88", "→ Signal Magenta"),
        ]

        def swatch_row(src_hex: str, tgt_hex: str) -> str:
            return (
                f'<div class="swatch-row">'
                f'<span class="swatch" style="background:#{src_hex};"></span>'
                f'<code>#{src_hex}</code>'
                f'<span class="arrow">→</span>'
                f'<span class="swatch" style="background:#{tgt_hex};"></span>'
                f'<code>#{tgt_hex}</code>'
                f'</div>'
            )

        # Build target → [sources] index
        _by_target: dict[str, list[str]] = {t: [] for t, _ in _TARGET_ORDER}
        for src, tgt in cd.COLOR_REMAP.items():
            if tgt in _by_target:
                _by_target[tgt].append(src)

        cols = st.columns(3)
        for _ci, (tgt, label) in enumerate(_TARGET_ORDER):
            srcs = _by_target.get(tgt, [])
            if not srcs:
                continue
            with cols[_ci % 3]:
                st.markdown(
                    f'<div class="section-label">{label}</div>',
                    unsafe_allow_html=True,
                )
                st.markdown(
                    "".join(swatch_row(s, tgt) for s in sorted(srcs)),
                    unsafe_allow_html=True,
                )

        st.divider()
        st.markdown(
            '<div class="section-label">Scheme fill overrides</div>',
            unsafe_allow_html=True,
        )
        scheme_html = "".join(
            f'<div class="scheme-row">'
            f'<code>{k}</code>'
            f'<span class="arrow">→</span>'
            f'<span class="swatch" style="background:#{v};"></span>'
            f'<code>#{v}</code>'
            f'</div>'
            for k, v in cd.SCHEME_FILL_MAP.items()
        )
        st.markdown(scheme_html, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — LOGO SUITE
# ══════════════════════════════════════════════════════════════════════════════
with tab_logos:
    LOGO_DIR = BASE_DIR / "logo package"

    VARIANT_BG = {
        "default":   "#FFFFFF",
        "alternate": "#281A42",
        "reverse":   "#1D1D1D",
    }
    VARIANT_LABEL = {
        "default":   "Default",
        "alternate": "Alternate",
        "reverse":   "Reverse",
    }
    VARIANT_LABEL_COLOR = {
        "default":   "#281A42",
        "alternate": "rgba(255,255,255,0.6)",
        "reverse":   "rgba(255,255,255,0.6)",
    }

    FAMILIES = [
        ("Primary Logo + Tagline", "primary logo + tagline"),
        ("Primary Logo",           "primary logo"),
        ("Monogram",               "monogram"),
        ("Triangle",               "triangle"),
    ]
    VARIANTS = ["default", "alternate", "reverse"]
    DL_FMTS  = [("PNG", "png"), ("SVG", "svg"), ("PDF", "pdf")]

    def _b64(path: Path) -> str:
        return base64.b64encode(path.read_bytes()).decode()

    def _dl_button(path: Path, fmt: str, key: str):
        mime_map = {
            "png": "image/png",
            "svg": "image/svg+xml",
            "pdf": "application/pdf",
        }
        st.download_button(
            label=fmt,
            data=path.read_bytes(),
            file_name=path.name,
            mime=mime_map.get(fmt.lower(), "application/octet-stream"),
            key=key,
            use_container_width=True,
        )

    for family_label, stem in FAMILIES:
        with st.expander(family_label, expanded=(stem == "primary logo + tagline")):
            cols = st.columns(3)
            for col, variant in zip(cols, VARIANTS):
                png_path = LOGO_DIR / "PNG" / f"{stem}_{variant}.png"
                bg  = VARIANT_BG[variant]
                ltc = VARIANT_LABEL_COLOR[variant]

                with col:
                    if png_path.exists():
                        img_b64 = _b64(png_path)
                        st.markdown(
                            f'<div class="logo-tile" style="background:{bg};">'
                            f'<img src="data:image/png;base64,{img_b64}" '
                            f'style="max-width:100%;max-height:120px;object-fit:contain;">'
                            f'<span class="logo-variant-label" style="color:{ltc};">'
                            f'{VARIANT_LABEL[variant]}</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown(
                            f'<div class="logo-tile" style="background:{bg};min-height:120px;">'
                            f'<span style="color:rgba(255,255,255,0.2);font-size:11px;">not found</span>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )

                    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
                    dl_cols = st.columns(len(DL_FMTS))
                    for dc, (fmt_label, fmt_ext) in zip(dl_cols, DL_FMTS):
                        dl_path = LOGO_DIR / fmt_ext.upper() / f"{stem}_{variant}.{fmt_ext}"
                        with dc:
                            if dl_path.exists():
                                _dl_button(dl_path, fmt_label, f"dl_{stem}_{variant}_{fmt_ext}")
                            else:
                                st.markdown(
                                    f'<div style="text-align:center;font-size:9px;'
                                    f'color:rgba(255,255,255,0.18);padding:6px 0;">—</div>',
                                    unsafe_allow_html=True,
                                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ICONS
# ══════════════════════════════════════════════════════════════════════════════
with tab_icons:
    import io as _io
    import numpy as _np
    from pptx import Presentation as _Prs
    from pptx.enum.shapes import MSO_SHAPE_TYPE as _MSO
    from PIL import Image as _Image

    _ICONS_PATH = BASE_DIR / "icons_fq.pptx"
    _DUO_DARK  = _np.array([0x76, 0x5F, 0xFF], dtype=_np.float32)
    _DUO_LIGHT = _np.array([0xFF, 0xFF, 0xFF], dtype=_np.float32)

    def _apply_duotone(img: "_Image.Image") -> "_Image.Image":
        rgba = _np.array(img.convert("RGBA"), dtype=_np.float32)
        gray = (rgba[:, :, 0] * 0.299 + rgba[:, :, 1] * 0.587 + rgba[:, :, 2] * 0.114) / 255.0
        t = gray[:, :, _np.newaxis]
        rgb_out = (_DUO_DARK * (1 - t) + _DUO_LIGHT * t).clip(0, 255).astype(_np.uint8)
        result = _np.dstack([rgb_out, rgba[:, :, 3].astype(_np.uint8)])
        return _Image.fromarray(result, "RGBA")

    _WIDE_SHAPES = {"Picture 26"}
    _EXCLUDED    = {"Picture 31"}
    _ICON_SIZE   = 120

    def _fit_to_square(img: "_Image.Image", size: int) -> "_Image.Image":
        img.thumbnail((size, size), _Image.LANCZOS)
        canvas = _Image.new("RGBA", (size, size), (0, 0, 0, 0))
        x = (size - img.width) // 2
        y = (size - img.height) // 2
        canvas.paste(img, (x, y))
        return canvas

    def _fit_to_height(img: "_Image.Image", h: int) -> "_Image.Image":
        ratio = h / img.height
        return img.resize((max(1, int(img.width * ratio)), h), _Image.LANCZOS)

    @st.cache_data(show_spinner=False)
    def _load_icons(path_str: str):
        """Return (sections, wide_img).
        sections = [{"title": str, "icons": [PIL.Image]}]
        wide_img  = PIL.Image | None  — rendered separately at the end.
        Text shapes within the slide act as section dividers (sorted by position).
        Blob-hash deduplication prevents identical images appearing twice."""
        import hashlib
        prs = _Prs(path_str)
        seen: set[str] = set()
        sections: list[dict] = []
        cur_title = ""
        cur_icons: list = []
        wide_img = None

        for slide in prs.slides:
            for shape in sorted(slide.shapes, key=lambda s: (s.top or 0, s.left or 0)):
                # ── Text shape → start a new section ─────────────────────────
                if shape.has_text_frame and shape.shape_type != _MSO.PICTURE:
                    text = shape.text_frame.text.strip()
                    if len(text) > 1:           # skip page numbers / single chars
                        if cur_icons:
                            sections.append({"title": cur_title, "icons": cur_icons})
                            cur_icons = []
                        cur_title = text
                    continue

                if shape.shape_type != _MSO.PICTURE:
                    continue
                if shape.name in _EXCLUDED:
                    continue

                try:
                    blob = shape.image.blob
                    bh   = hashlib.md5(blob).hexdigest()
                    if bh in seen:
                        continue
                    seen.add(bh)

                    raw = _Image.open(_io.BytesIO(blob)).convert("RGBA")
                    img = _apply_duotone(raw)

                    if shape.name in _WIDE_SHAPES:
                        wide_img = _fit_to_height(img, _ICON_SIZE)
                    else:
                        cur_icons.append(_fit_to_square(img, _ICON_SIZE))
                except Exception:
                    pass

        if cur_icons:
            sections.append({"title": cur_title, "icons": cur_icons})

        return sections, wide_img

    st.markdown("### Icon Library")

    if not _ICONS_PATH.exists():
        st.error(f"File not found: `{_ICONS_PATH}`")
    else:
        _sections, _wide_img = _load_icons(str(_ICONS_PATH))
        _COLS_PER_ROW = 10

        for _sec in _sections:
            if _sec["title"]:
                st.markdown(
                    f'<div class="section-label" style="margin-top:1.4rem;">{_sec["title"]}</div>',
                    unsafe_allow_html=True,
                )
            if _sec["icons"]:
                _cells = ""
                for _img in _sec["icons"]:
                    _buf = _io.BytesIO()
                    _img.save(_buf, format="PNG")
                    _ib64 = base64.b64encode(_buf.getvalue()).decode()
                    _cells += (
                        f'<div class="icon-cell">'
                        f'<img src="data:image/png;base64,{_ib64}" '
                        f'width="{_ICON_SIZE}" height="{_ICON_SIZE}" style="display:block;">'
                        f'</div>'
                    )
                st.markdown(
                    f'<div style="display:grid;grid-template-columns:repeat({_COLS_PER_ROW},1fr);'
                    f'gap:6px;margin-bottom:4px;">{_cells}</div>',
                    unsafe_allow_html=True,
                )

        # ── Wide composite — own full-width row at the end ────────────────────
        if _wide_img is not None:
            st.markdown(
                '<div class="section-label" style="margin-top:1.4rem;">Click process</div>',
                unsafe_allow_html=True,
            )
            _buf = _io.BytesIO()
            _wide_img.save(_buf, format="PNG")
            _wb64 = base64.b64encode(_buf.getvalue()).decode()
            st.markdown(
                f'<div class="icon-cell" style="display:inline-flex;align-items:center;'
                f'padding:10px 14px;height:{_ICON_SIZE}px;">'
                f'<img src="data:image/png;base64,{_wb64}" '
                f'style="height:{_ICON_SIZE - 20}px;width:auto;display:block;">'
                f'</div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — TYPOGRAPHY
# ══════════════════════════════════════════════════════════════════════════════
with tab_type:
    import json as _json

    # Type scale scraped from style_guide.json + master PPTX spec
    # color_bg = the dark UI background color we render the sample against
    TYPE_SCALE = [
        {
            "name": "Cover Title",
            "font": "Segoe UI", "size_pt": 48, "weight": 700,
            "color": "#FFFFFF", "bg": "#1D1D1D",
            "sample": "The Science of Talent to Value",
            "notes": "Segoe UI Bold · 48pt · White · Cover slides only",
        },
        {
            "name": "Cover Subtitle",
            "font": "Segoe UI", "size_pt": 28, "weight": 400,
            "color": "#765FFF", "bg": "#1D1D1D",
            "sample": "Transform decks at the speed of strategy",
            "notes": "Segoe UI Regular · 28pt · Pivot Purple",
        },
        {
            "name": "Section Title",
            "font": "Segoe UI", "size_pt": 32, "weight": 700,
            "color": "#FFFFFF", "bg": "#281A42",
            "sample": "Section 01 — Talent Architecture",
            "notes": "Segoe UI Bold · 32pt · White · Divider slides",
        },
        {
            "name": "Slide Title",
            "font": "Segoe UI", "size_pt": 28, "weight": 700,
            "color": "#1D1D1D", "bg": "#FFFFFF",
            "sample": "Organizational Efficiency",
            "notes": "Segoe UI Bold · 28pt · Guiding Grey",
        },
        {
            "name": "Slide Subtitle",
            "font": "Segoe UI", "size_pt": 18, "weight": 700,
            "color": "#765FFF", "bg": "#FFFFFF",
            "sample": "Key findings across markets",
            "notes": "Segoe UI Bold · 18pt · Pivot Purple",
        },
        {
            "name": "Body",
            "font": "Arial", "size_pt": 12, "weight": 400,
            "color": "#1D1D1D", "bg": "#FFFFFF",
            "sample": "Supporting detail that frames the core insight with context and precision across the organization.",
            "notes": "Arial Regular · 12pt · Guiding Grey",
        },
        {
            "name": "Body Small",
            "font": "Arial", "size_pt": 10, "weight": 400,
            "color": "#3B3B3B", "bg": "#FFFFFF",
            "sample": "Secondary annotation or footnote text used below a data point or chart.",
            "notes": "Arial Regular · 10pt · Grey 1",
        },
        {
            "name": "Caption",
            "font": "Arial", "size_pt": 8, "weight": 400,
            "color": "#585858", "bg": "#FFFFFF",
            "sample": "Source: FulcrumQ internal analysis, FY25 · All figures USD",
            "notes": "Arial Regular · 8pt · Grey 2",
        },
        {
            "name": "Table Header",
            "font": "Segoe UI", "size_pt": 10, "weight": 600,
            "color": "#FFFFFF", "bg": "#765FFF",
            "sample": "CATEGORY  ·  OWNER  ·  STATUS  ·  TIMELINE",
            "notes": "Segoe UI SemiBold · 10pt · White on Pivot Purple",
        },
        {
            "name": "Table Body",
            "font": "Arial", "size_pt": 9, "weight": 400,
            "color": "#1D1D1D", "bg": "#FFFFFF",
            "sample": "Talent acquisition  ·  Sarah Chen  ·  On track  ·  Q2 FY25",
            "notes": "Arial Regular · 9pt · Guiding Grey",
        },
        {
            "name": "Label",
            "font": "Segoe UI", "size_pt": 9, "weight": 700,
            "color": "#765FFF", "bg": "#1D1D1D",
            "sample": "INITIATIVE  ·  PRIORITY  ·  IMPACT  ·  EFFORT",
            "notes": "Segoe UI Bold · 9pt · Pivot Purple",
        },
    ]

    st.markdown(
        '<div class="section-label" style="margin-bottom:1rem;">Master type scale — sourced from LatestMaster_FQ</div>',
        unsafe_allow_html=True,
    )

    for entry in TYPE_SCALE:
        # Scale pt to approximate px for web display (cap at 3rem)
        px = min(entry["size_pt"] * 1.2, 56)
        fw = entry["weight"]
        ff = f"'{entry['font']}', 'Inter', Arial, sans-serif"
        fc = entry["color"]
        bg = entry["bg"]

        # Render sample on its authentic background color
        sample_style = (
            f"font-family:{ff};"
            f"font-size:{px}px;"
            f"font-weight:{fw};"
            f"color:{fc};"
            f"background:{bg};"
            f"padding:8px 14px;"
            f"border-radius:8px;"
            f"line-height:1.2;"
            f"display:block;"
        )

        st.markdown(
            f'<div class="type-row">'
            f'<div class="type-meta"><strong>{entry["name"]}</strong>{entry["notes"]}</div>'
            f'<div class="type-sample"><span style="{sample_style}">{entry["sample"]}</span></div>'
            f'</div>',
            unsafe_allow_html=True,
        )

