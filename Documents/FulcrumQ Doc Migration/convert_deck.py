"""
convert_deck.py
Converts any source PPTX to FulcrumQ brand by:

  STEP 1  — Master swap (destructive on old master)
              Remove source master, all its layouts, and its theme
              Inject v7 master + layouts + theme in their place
              Update presentation.xml master ref

  STEP 2  — Layout remapping
              For each slide, resolve best-fit v7 layout by name or semantic match
              Rewrite slide rels to point at the new layout

  STEP 3  — Vestige cleanup
              Remove shapes that are artifacts of the old master:
                · Arrow auto-shapes (prst contains "Arrow")
                · Zero-or-near-zero-size shapes with no text
                · Full-slide background rectangles with no text content

  STEP 4  — Brand style pass
              Fonts  → Aileron (bold headings) / Switzer (body)
              Colors → map old palette to FulcrumQ brand colors

Usage:
    python3 convert_deck.py Org_Efficiency.pptx
"""

import colorsys
import re
import sys
import uuid
import zipfile
from copy import deepcopy
from pathlib import Path
from lxml import etree

BASE_DIR    = Path(__file__).parent
MASTER_PPTX = BASE_DIR / "LatestMaster_FQ.potx"   # legacy fallback

# Dual-master strategy
# X = default (horizontal footer at bottom)
MASTER_X = BASE_DIR / "FQ_PPTX_Theme_vF.pptx"

NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS  = "http://schemas.openxmlformats.org/package/2006/content-types"

SLIDE_MASTER_CT = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
SLIDE_LAYOUT_CT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
THEME_CT        = "application/vnd.openxmlformats-officedocument.theme+xml"
SLIDE_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
SLIDE_LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
THEME_REL        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

# ── Sentence-case rules (ported from copyedit.py) ────────────────────────────
# Words that must keep their exact casing regardless of position
PRESERVE_WORDS = {
    # Brand names / product terms
    "FulcrumQ", "T2V", "JTBD", "JTBDs",
    # C-suite / titles
    "CEO", "CFO", "COO", "CIO", "CTO", "CHRO", "CLO", "CMO", "CSO",
    "CDO", "CPO", "CRO", "CAO", "CCO", "CGO", "CXO",
    "SVP", "EVP", "VP", "AVP", "GM", "MD", "Board",
    "CEOs", "CFOs", "COOs", "CTOs", "CHROs", "CMOs",
    # Business / finance
    "P&L", "EBITDA", "EBIT", "EBITA", "ROIC", "ROE", "ROA", "ROI",
    "KPI", "KPIs", "OKR", "OKRs", "NPS", "MBO",
    "M&A", "IPO", "SPV", "LBO", "DCF", "NPV", "IRR", "WACC",
    "PE", "VC", "GP", "LP",
    "SaaS", "ERP", "CRM", "HCM", "SCM", "WMS", "TMS",
    "OPEX", "OpEx", "CAPEX", "CapEx", "G&A", "S&M", "COGS",
    "ARR", "MRR", "ACV", "TCV", "CAC", "LTV", "ARPU", "GMV",
    "CAGR", "YoY", "QoQ", "MoM",
    # Strategy / consulting
    "GTM", "TAM", "SAM", "SOM", "MVP", "POC", "RFP", "RFI", "SOW",
    "PMO", "PMI", "RACI", "OKR", "OKRs", "VSM", "BPO", "SSC",
    "S&OP", "SKU", "SKUs",
    # HR / org
    "HiPo", "HiPos", "D&I", "DEI", "L&D", "PIP", "HC",
    "FTE", "FTEs", "RIF", "HRIS",
    # Tech / data
    "AI", "ML", "NLP", "LLM", "API", "APIs", "SaaS", "PaaS", "IaaS",
    "IT", "HR", "R&D", "B2B", "B2C", "D2C", "SLA", "SLAs",
    "ETL", "ELT", "BI", "SQL", "NoSQL", "CX", "UX", "UI",
    # Time / fiscal
    "Q1", "Q2", "Q3", "Q4",
    "FY24", "FY25", "FY26", "FY27", "FY28",
    # Geography
    "US", "U.S.", "UK", "EU", "EMEA", "APAC", "LATAM", "NYC", "NA",
}
_PRESERVE_MAP = {w.lower(): w for w in PRESERVE_WORDS}


def _is_special_word(word: str) -> bool:
    """Return True if word should NOT be lowercased (acronym, proper noun, number, etc.).

    NOTE: We do NOT treat generic all-caps as an acronym here — source PPTX titles
    are often styled all-caps (AGENDA, TALENT, OBJECTIVES) and must be sentence-cased.
    Only words explicitly in PRESERVE_WORDS or matching structural patterns are kept."""
    clean = word.strip(".,;:!?()\"'-\u2013\u2014")
    if not clean:
        return False
    # Known preserve word (checked first — canonical list is the authority)
    if clean.lower() in _PRESERVE_MAP:
        return True
    # Contains & or / between letters: S&OP, AI/ML, P&L
    if re.search(r"[A-Za-z][&/][A-Za-z]", clean):
        return True
    # Mixed case with internal cap: OpEx, CapEx, AmerisourceBergen
    if re.search(r"[a-z][A-Z]", clean):
        return True
    # Starts uppercase + digit: FY25, Q3, B2B
    if re.search(r"^[A-Z].*\d", clean):
        return True
    # Starts with $, ~, or digit: $19B, ~1,200
    if clean[0] in "$~" or clean[0].isdigit():
        return True
    # Single uppercase letter (e.g. section labels)
    if len(clean) == 1 and clean.isupper():
        return True
    return False


# ── Chicago Manual of Style title case ───────────────────────────────────────
# Minor words: always lowercase unless first/last word or after : — /
_CMOS_MINOR = {
    # Articles
    "a", "an", "the",
    # Coordinating conjunctions
    "and", "but", "or", "nor", "for", "so", "yet",
    # Prepositions (CMOS lowercases all prepositions regardless of length)
    "about", "above", "across", "after", "against", "along", "amid", "among",
    "around", "as", "at", "atop", "before", "behind", "below", "beneath",
    "beside", "between", "beyond", "by", "despite", "down", "during",
    "except", "following", "from", "in", "inside", "into", "like", "minus",
    "near", "next", "of", "off", "on", "onto", "opposite", "out", "outside",
    "over", "past", "per", "plus", "since", "than", "through", "throughout",
    "till", "to", "toward", "towards", "under", "unlike", "until", "up",
    "upon", "versus", "via", "with", "within", "without",
}


def _to_title_case(text: str) -> str:
    """Apply Chicago Manual of Style title case to a string.

    Rules:
      - First and last word always capitalized.
      - All words capitalized except articles, coordinating conjunctions, and
        prepositions (CMOS_MINOR set above) — unless they are first or last.
      - Word immediately after : — / always capitalized.
      - Preserved words (PRESERVE_WORDS) keep their canonical casing.
      - All-caps source words that aren't in PRESERVE_WORDS get title-cased
        (e.g. "TALENT" → "Talent") — avoids shouting.
    """
    if not text or not text.strip():
        return text

    words = text.strip().split()
    n     = len(words)
    result = []
    capitalize_next = False  # force-cap after : — /

    for i, w in enumerate(words):
        is_first = (i == 0)
        is_last  = (i == n - 1)
        punct    = w.strip(".,;:!?()\"'\u2013\u2014/")
        lower_p  = punct.lower()

        # Preserved word — canonical casing
        if lower_p in _PRESERVE_MAP:
            result.append(_PRESERVE_MAP[lower_p])
            capitalize_next = w.endswith((":", "\u2013", "\u2014", "/"))
            continue

        # Other special words (mixed-case proper nouns, acronyms, numbers) — keep as-is
        if _is_special_word(punct):
            result.append(w)
            capitalize_next = w.endswith((":", "\u2013", "\u2014", "/"))
            continue

        # Determine whether this word should be capitalized
        force_cap = is_first or is_last or capitalize_next
        minor     = lower_p in _CMOS_MINOR

        if force_cap or not minor:
            # Capitalize: upper first letter, lower the rest
            # (handles all-caps source like "TALENT" → "Talent")
            cased = w[0].upper() + w[1:].lower() if w else w
        else:
            cased = w.lower()

        result.append(cased)
        capitalize_next = w.rstrip(".,;!?()\"'").endswith((":", "\u2013", "\u2014", "/"))

    return " ".join(result)


def _to_sentence_case(text: str) -> str:
    """Sentence-case a string — used for subtitles and body labels."""
    if not text or not text.strip():
        return text
    words = text.strip().split()
    result = []
    for i, w in enumerate(words):
        clean = w.strip(".,;:!?()\"'-\u2013\u2014")
        if _is_special_word(clean):
            result.append(_PRESERVE_MAP.get(clean.lower(), w) if clean.lower() in _PRESERVE_MAP else w)
        elif i == 0:
            result.append(w[0].upper() + w[1:].lower() if w else w)
        else:
            result.append(w.lower())
    return " ".join(result)


# ── Terminal color helpers ────────────────────────────────────────────────────
def dot(hex_str: str) -> str:
    """Return a true-color ANSI dot ● for the given 6-char hex color."""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"\x1b[38;2;{r};{g};{b}m●\x1b[0m"


def print_palette():
    """Print COLOR_REMAP with colored dots so you can visually inspect the mapping."""
    # Group by semantic category using comment markers from the dict order
    categories = [
        ("Reds",            ["C00000","B4000C","BA0A2F","BA0B2F","EC0000","FF0000","FF0001"]),
        ("Blues",           ["0073BE","125E7F","1D3C6D","467886","288FC2","4472C4","2E75B6"]),
        ("Yellows/Ambers",  ["FFFF00","FCB913","FFC000","FEAA12","FF6600","E36C09","FFA400"]),
        ("Off-brand purples",["461E96","7030A0","8B7CC8","B8ACE0","211359","96607D"]),
        ("Greens/Teals",    ["00B050","8CC53F","70AD47","548235","91CF4F","92D050","00B6C2","B3EAEE"]),
        ("Blacks",          ["000000","1E1E1E","0D0D0D"]),
        ("Dark greys",      ["333333","404040","595959","555555"]),
        ("Medium greys",    ["696158","7F7F7F","808080","919291","A6A6A6"]),
        ("Light greys",     ["C0C0C0","BFBFBF","DADADA","D9D9D9","E9EAE9","D0D2D0","DCECF5"]),
    ]
    print("\n── COLOR_REMAP palette ──────────────────────────────")
    for label, keys in categories:
        print(f"\n  {label}")
        for src_hex in keys:
            tgt_hex = COLOR_REMAP.get(src_hex)
            if tgt_hex:
                print(f"    {dot(src_hex)} #{src_hex}  →  {dot(tgt_hex)} #{tgt_hex}")

    print("\n── Brand accent palette ─────────────────────────────")
    ACCENT_LABELS = [
        ("765FFF", "Pivot Purple"),
        ("917FFF", "Pivot Purple Tint 1"),
        ("AD9FFF", "Pivot Purple Tint 2"),
        ("C8BFFF", "Pivot Purple Tint 3"),
        ("E9E4FF", "Light Lavender (neutral)"),
        ("00C27A", "Anchor Green"),
        ("FF2E88", "Signal Magenta"),
        ("FFB547", "Amber"),
        ("60BDBC", "Teal"),
        ("C8BFFF", "Sage Green"),
        ("FF2E88", "Hot Pink (accent backup)"),
    ]
    for hex_val, label in ACCENT_LABELS:
        print(f"  {dot(hex_val)} #{hex_val}  {label}")
    print()


# ── Brand font rules ──────────────────────────────────────────────────────────
FONT_FALLBACKS = {"Segoe UI": (34, 0), "Arial": (34, 0)}

# Character spacing (hundredths of a point; negative = tighter)
TITLE_CHAR_SPC = -25   # -0.25 pt on Segoe UI Bold titles
BODY_CHAR_SPC  = -10   # -0.10 pt on Arial body text

# Non-brand typefaces → brand equivalent
# All map to Arial by default; style_slide overrides to Segoe UI for
# shapes in the title region (top 1.5 in of slide) via position check.
FONT_MAP = {
    "Aileron Bold":                   ("Segoe UI", True),
    "Aileron Light":                  ("Segoe UI", False),
    "Helvetica Neue":                 ("Arial",  None),
    "Helvetica Neue Light":           ("Arial",  False),
    "Helvetica Neue Medium":          ("Arial",  None),
    "Helvetica Neue Thin":            ("Arial",  False),
    "Helvetica Neue Condensed":       ("Arial",  None),
    "Helvetica Neue Condensed Black": ("Arial",  None),
    "Helvetica":                      ("Arial",  None),
    "Arial":                          ("Arial",  None),
    "Arial Black":                    ("Segoe UI", True),
    "Arial Narrow":                   ("Arial",  None),
    "Calibri":                        ("Arial",  None),
    "Calibri Light":                  ("Arial",  False),
    "Inter":                          ("Arial",  None),
    "Inter Medium":                   ("Arial",  None),
    "Lato":                           ("Arial",  None),
    "Lato Light":                     ("Arial",  False),
    "Lato Medium":                    ("Arial",  None),
    "Lato Semibold":                  ("Arial",  None),
    "Lato Regular":                   ("Arial",  None),
    "Ltao":                           ("Arial",  None),   # common typo for Lato
    "Open Sans":                      ("Arial",  None),
    "Roboto":                         ("Arial",  None),
    "Roboto Light":                   ("Arial",  False),
    "Roboto Medium":                  ("Arial",  None),
    "Roboto Thin":                    ("Arial",  False),
    "Roboto Condensed Medium":        ("Arial",  None),
    "SoDo Sans":                      ("Arial",  None),
    "SoDo Sans Light":                ("Arial",  False),
    "SoDo Sans Narrow":               ("Arial",  None),
    "Source Sans Pro":                ("Arial",  None),
    "Source Sans Pro Black":          ("Arial",  None),
    "Aptos":                          ("Arial",  None),
    "Aptos Narrow":                   ("Arial",  None),
    "Avenir Next":                    ("Arial",  None),
    "Bierstadt":                      ("Arial",  None),
    "Century Gothic":                 ("Arial",  None),
    "Comic Sans MS":                  ("Arial",  None),
    "Georgia":                        ("Arial",  None),
    "Times New Roman":                ("Arial",  None),
}

# Bold non-brand → upgrade to Aileron
BOLD_UPGRADE = {
    "Helvetica Neue", "Arial", "Calibri", "Open Sans",
    "Roboto", "Lato", "Inter",
}

# ── Brand color remap ─────────────────────────────────────────────────────────
COLOR_REMAP = {
    # Reds → dark neutral (preserves distinction from RAG red)
    "C00000": "281A42", "B4000C": "281A42",
    "BA0A2F": "281A42", "BA0B2F": "281A42",
    "B02418": "281A42", "AA2634": "281A42",  # additional source reds
    "EC0000": "281A42", "FF0000": "281A42", "FF0001": "281A42",
    # Blues → Pivot Purple (primary) or dark navy / blue-grey (muted)
    "0073BE": "765FFF",
    "125E7F": "281A42", "1D3C6D": "281A42", "467886": "60BDBC",
    "288FC2": "60BDBC",
    "4472C4": "765FFF", "2E75B6": "765FFF",   # common PPT blues → Pivot Purple
    # Yellows → Sage Green (light); ambers/oranges → Amber
    "FFFF00": "FFB547", "FCB913": "FFB547",
    "FFC000": "FFB547", "FEAA12": "FFB547",
    "FF6600": "FFB547", "E36C09": "FFB547", "FFA400": "FFB547",
    # Off-brand purples → Pivot Purple family (by intensity)
    "461E96": "281A42",                        # strong dark purple → Vector Dark Purple
    "7030A0": "917FFF",                        # medium purple → Tint 1
    "8B7CC8": "AD9FFF",                        # light medium → Tint 2
    "B8ACE0": "C8BFFF",                        # very light → Tint 3
    "211359": "281A42",                        # near-black purple → Vector Dark Purple
    "96607D": "C8BFFF",                        # muted mauve → Tint 3
    "C14BFF": "FF2E88",                        # Hot Lilac → Signal Magenta
    # Greens → Anchor Green; light teals → Sage Green
    "00B050": "00C27A", "8CC53F": "00C27A",
    "70AD47": "00C27A", "548235": "00C27A",
    "91CF4F": "00C27A", "92D050": "00C27A",
    "00B6C2": "60BDBC", "00B6C3": "60BDBC", "B3EAEE": "C8BFFF",
    # Light / pale cyan-turquoise → Light Lavender
    "C3FDFE": "E9E4FF", "A0FCFE": "E9E4FF",
    "BFFAFE": "E9E4FF", "D0FEFF": "E9E4FF", "CCFFFF": "E9E4FF",
    "E0FFFF": "E9E4FF", "E5FFFF": "E9E4FF", "F0FFFF": "E9E4FF",
    # Blacks / near-blacks → dark
    "000000": "1D1D1D", "1E1E1E": "1D1D1D",
    "0D0D0D": "1D1D1D",
    # Dark greys
    "333333": "3B3B3B", "404040": "3B3B3B",
    "595959": "585858", "555555": "585858",
    # Medium greys
    "696158": "767676", "7F7F7F": "7A828D",
    "808080": "7A828D", "919291": "7A828D",
    "A6A6A6": "7A828D",
    # Light greys
    "C0C0C0": "B2BBCA", "BFBFBF": "B2BBCA",
    "DADADA": "D0D7DF", "D9D9D9": "D0D7DF", "E9EAE9": "D0D7DF",
    "D0D2D0": "E9E4FF", "DCECF5": "E9E4FF",

    # ── Near-miss brand colors (off-by-one rounding artifacts) ───────────────
    "2B3A42": "281A42",   # ≈ Vector Dark Purple (R channel off)
    "E9E4FE": "E9E4FF",   # ≈ Light Lavender (B channel off by 1)
    "FCB546": "FFB547",   # ≈ Momentum Amber (slight drift)
    "FCB547": "FFB547",
    "FFB546": "FFB547",

    # ── Deck-specific additions from audit ────────────────────────────────────

    # Dark brand backgrounds / near-blacks → Guiding Grey
    "1E3832": "1D1D1D", "1D3831": "1D1D1D", "1D3A31": "1D1D1D",
    "1D383B": "1D1D1D", "1D3931": "1D1D1D", "00341F": "281A42",
    "212121": "1D1D1D", "231F20": "1D1D1D", "020203": "1D1D1D",
    "030405": "1D1D1D", "0E0E0E": "1D1D1D", "030304": "1D1D1D",
    "111111": "1D1D1D", "1B1B1B": "1D1D1D", "171D1A": "1D1D1D",
    "1C1917": "1D1D1D", "161D1A": "1D1D1D",

    # Dark greys → Grey 1 / Grey 2
    "37373A": "3B3B3B", "3F3F3F": "3B3B3B", "424242": "3B3B3B",
    "40403F": "3B3B3B", "444444": "3B3B3B",
    "4C4C4C": "585858", "515151": "585858",

    # Medium greys → Grey 2 / Grey 3 / Grey Mid
    "5E5E5E": "585858",
    "818181": "767676", "6D6E71": "767676", "7E7E7E": "767676",
    "6E747A": "7A828D", "9E9E9F": "7A828D", "9A9A9A": "7A828D",
    "969696": "7A828D", "929292": "7A828D", "A5A5A5": "7A828D",

    # Light greys / near-whites → Lightest Grey (F2F2F2) or Soft Grey (D0D7DF)
    # Very near-white → F2F2F2 (Lightest Grey brand color)
    "F5F5F5": "F2F2F2", "F4F4F4": "F2F2F2",
    "F1F1F1": "F2F2F2", "F6F6F6": "F2F2F2", "F9F9F9": "F2F2F2",
    "F1F2F2": "F2F2F2", "FBFFFF": "FFFFFF",  "FEFFFE": "FFFFFF",
    # Slightly darker near-whites → Soft Grey (D0D7DF)
    "E8E8E8": "D0D7DF", "EDEBE9": "D0D7DF",
    "E6E6E6": "D0D7DF", "CACACA": "B2BBCA", "B7B7B7": "B2BBCA",
    "A2CEBD": "B2BBCA",

    # Light blues / blue-greys → Soft Grey
    "DDE8F7": "D0D7DF", "C9E8FF": "D0D7DF", "D9E1F2": "D0D7DF",
    "B7E8FE": "D0D7DF", "B3EBF4": "D0D7DF", "B4EBF4": "D0D7DF",

    # Light greens → Soft Grey
    "D3E8E0": "D0D7DF", "D5E9E2": "D0D7DF", "DFF2EC": "D0D7DF",
    "D0EAE2": "D0D7DF", "D5E8E1": "D0D7DF", "E0F3EC": "D0D7DF",
    "E0EEE8": "D0D7DF",

    # Light creams / pale yellows → Soft Grey
    "FEF1D0": "D0D7DF", "EFD09F": "D0D7DF", "FFF8BA": "D0D7DF",

    # Light purples / lavendars → Tint 3 / Light Lavender
    "D883FF": "C8BFFF", "D984FF": "C8BFFF", "D7C8F4": "C8BFFF",
    "865ADE": "917FFF",
    # Light cyans / aquas → Light Lavender (closest brand color perceptually)
    "E5FFFF": "E9E4FF", "E0FFFF": "E9E4FF", "F0FFFF": "E9E4FF",
    "CCFFFF": "C8BFFF", "B3FFFF": "C8BFFF",

    # Greens → Anchor Green
    "29BA74": "00C27A", "01744B": "00C27A", "015738": "00C27A",
    "007042": "00C27A", "01B336": "00C27A", "92D051": "00C27A",
    "007F50": "00C27A", "05B050": "00C27A", "00704A": "00C27A",
    "0C8C43": "00C27A", "006141": "00C27A", "006C47": "00C27A",
    "73B442": "00C27A", "BADC8C": "00C27A",

    # Sage / muted greens → Teal
    "5FAB8E": "60BDBC", "5EAB8E": "60BDBC",

    # Teals / cyans → Teal
    "1D95A4": "60BDBC", "34CED8": "60BDBC", "34CED9": "60BDBC",
    "2FCCD6": "60BDBC", "2FCCD7": "60BDBC", "30CDD7": "60BDBC",
    "10ADBE": "60BDBC", "40C4F4": "60BDBC", "09AFDC": "60BDBC",
    "00FFFF": "60BDBC", "039ACA": "60BDBC", "145E7F": "60BDBC",
    "00B0F0": "60BDBC", "009FDF": "60BDBC",

    # Blues → Pivot Purple
    "0070C0": "765FFF", "4E2BC5": "765FFF", "6280C2": "765FFF",
    "4372C4": "765FFF", "497DBA": "765FFF", "4F81BC": "765FFF",
    "385D89": "765FFF", "2056AE": "765FFF", "0000FF": "765FFF",
    "0C5F99": "765FFF", "245A88": "765FFF",
    "8EA9DB": "AD9FFF",   # light blue → Tint 2 (closest perceptual)

    # Dark navies → Vector Dark Purple
    "213F6F": "281A42", "002060": "281A42", "004681": "281A42",
    "183D6A": "281A42", "00446B": "281A42", "1F497D": "281A42",
    "071949": "281A42", "1B3969": "281A42", "44546A": "281A42",
    "001F60": "281A42", "173D60": "281A42", "461D96": "281A42",

    # Ambers / oranges → Momentum Amber
    "FBB020": "FFB547", "FFCB00": "FFB547", "FDB813": "FFB547",
    "FCB020": "FFB547", "FFD753": "FFB547", "FFAB12": "FFB547",
    "EE910C": "FFB547", "F97C00": "FFB547", "FF6B00": "FFB547",

    # Additional off-brand purples
    "8C5EB9": "917FFF", "955BBD": "917FFF",  # medium purples → Tint 1
    "57307F": "281A42", "57307E": "281A42",  # dark purples → Vector Dark Purple
    "6752EF": "765FFF",                       # bright purple → Pivot Purple
    "C6AFDC": "C8BFFF", "C2A8DA": "C8BFFF", "C1A8DA": "C8BFFF",  # light purples → Tint 3
    "151731": "281A42",                       # very dark navy → Vector Dark Purple
    "F7F2FA": "E9E4FF", "F6F1FA": "E9E4FF", # very light lavender → Shift Lavender
    "E0DDE1": "D0D7DF",                       # light grey-purple → Soft Grey
}

# Target color for any custom red not already in COLOR_REMAP
RED_REMAP_TARGET = "765FFF"

def _is_red_hue(hex_val: str, min_sat: float = 0.55, min_val: float = 0.25) -> bool:
    """Return True if color is a saturated red shade (HSV hue 0-20° or 340-360°).
    Excludes pastels (low saturation) and near-blacks (low value) so salmon pinks
    and very dark maroons are left alone."""
    try:
        r = int(hex_val[0:2], 16) / 255
        g = int(hex_val[2:4], 16) / 255
        b = int(hex_val[4:6], 16) / 255
        v = max(r, g, b)
        if v < min_val:
            return False
        delta = v - min(r, g, b)
        if v == 0 or (delta / v) < min_sat:   # HSV saturation check
            return False
        if v == r:
            hue = 60 * (((g - b) / delta) % 6)
        elif v == g:
            hue = 60 * ((b - r) / delta + 2)
        else:
            hue = 60 * ((r - g) / delta + 4)
        return hue <= 20 or hue >= 340
    except Exception:
        return False


def _hex_to_lab(h: str) -> tuple[float, float, float]:
    """Convert a 6-char hex string to CIE Lab (D65). No external deps."""
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    # sRGB → linear
    r = r / 12.92 if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = g / 12.92 if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = b / 12.92 if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    # linear RGB → XYZ (D65)
    x = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047
    y = (r * 0.2126 + g * 0.7152 + b * 0.0722) / 1.00000
    z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883
    # XYZ → Lab
    def f(t): return t ** (1/3) if t > 0.008856 else 7.787 * t + 16/116
    return 116 * f(y) - 16, 500 * (f(x) - f(y)), 200 * (f(y) - f(z))


# Precompute Lab for every approved brand output color (remap targets).
# Must match slide 4 palette exactly — no non-brand colors here.
_BRAND_LAB: list[tuple[str, tuple]] = [
    (h, _hex_to_lab(h)) for h in {
        # Pivot Purple family
        "765FFF","917FFF","AD9FFF","C8BFFF","E9E4FF",
        # Guiding Grey scale
        "1D1D1D","3B3B3B","585858","767676","7A828D",
        "B2BBCA","D0D7DF","F2F2F2","FFFFFF",
        # Brand darks
        "281A42",
        # Accents
        "00C27A","60BDBC","FFB547","FF2E88",
    }
]


def _nearest_brand_color(h: str) -> str:
    """Return the perceptually closest approved brand color using CIE76 ΔE.

    Achromatic guard: if the source color has low chroma (sqrt(a²+b²) < 10),
    restrict candidates to the grey/neutral brand colors only — prevents
    near-grey inputs from snapping to E9E4FF or other tinted brand colors.
    """
    _ACHROMATIC_BRAND = {
        "1D1D1D","3B3B3B","585858","767676","7A828D",
        "B2BBCA","D0D7DF","F2F2F2","FFFFFF","281A42",
    }
    try:
        lab = _hex_to_lab(h)
    except Exception:
        return "1D1D1D"
    src_chroma = (lab[1] ** 2 + lab[2] ** 2) ** 0.5
    candidates = (
        [(bh, bl) for bh, bl in _BRAND_LAB if bh in _ACHROMATIC_BRAND]
        if src_chroma < 10 else _BRAND_LAB
    )
    best, best_d = "1D1D1D", float("inf")
    for brand_h, brand_lab in candidates:
        d = sum((a - b) ** 2 for a, b in zip(lab, brand_lab)) ** 0.5
        if d < best_d:
            best_d, best = d, brand_h
    return best


def _apply_scheme_transforms(base_hex: str, transforms) -> str:
    """Apply OOXML tint/shade/lumMod/lumOff child transforms to a brand base color,
    then snap the result to the nearest approved brand color via CIE76 ΔE.

    This preserves the LUMINANCE INTENT of tinted scheme colors.  A source shape
    that used `dk2 + tint(20%)` (appearing very light) must not collapse to the
    raw dark brand equivalent of dk2 — it should land on a light brand color.
    """
    r = float(int(base_hex[0:2], 16))
    g = float(int(base_hex[2:4], 16))
    b = float(int(base_hex[4:6], 16))

    for t in transforms:
        tag = t.tag.split("}")[-1]
        try:
            v = int(t.get("val", "100000")) / 100000.0
        except (ValueError, TypeError):
            continue
        if tag == "tint":
            # tint(v): v=0 → white, v=1 → original color
            r = r * v + 255.0 * (1.0 - v)
            g = g * v + 255.0 * (1.0 - v)
            b = b * v + 255.0 * (1.0 - v)
        elif tag == "shade":
            # shade(v): v=0 → black, v=1 → original color
            r *= v; g *= v; b *= v
        elif tag in ("lumMod", "lumOff"):
            rn, gn, bn = r / 255.0, g / 255.0, b / 255.0
            h, l, s = colorsys.rgb_to_hls(rn, gn, bn)
            l = (l * v) if tag == "lumMod" else (l + v)
            l = max(0.0, min(1.0, l))
            rn, gn, bn = colorsys.hls_to_rgb(h, l, s)
            r, g, b = rn * 255.0, gn * 255.0, bn * 255.0
        r = max(0.0, min(255.0, r))
        g = max(0.0, min(255.0, g))
        b = max(0.0, min(255.0, b))

    result = f"{int(round(r)):02X}{int(round(g)):02X}{int(round(b)):02X}"
    return _nearest_brand_color(result)


def _resolve_color(v: str, rag_preserve: set) -> str | None:
    """Resolve a source hex to a brand target color.
    Priority: RAG preserve → explicit COLOR_REMAP → red hue → nearest Lab match.
    Never returns None — every color gets a brand equivalent."""
    if v in rag_preserve:
        return None
    if v in COLOR_REMAP:
        return COLOR_REMAP[v]
    if _is_red_hue(v):
        return RED_REMAP_TARGET
    # Fallback: nearest perceptual match across all approved brand colors
    try:
        return _nearest_brand_color(v)
    except Exception:
        return "1D1D1D"


def _in_table(element):
    """Return True if element is a descendant of an <a:tbl> node."""
    node = element.getparent()
    while node is not None:
        if node.tag == f"{{{NS_A}}}tbl":
            return True
        node = node.getparent()
    return False


def _hue_bucket(hex_val: str):
    """Classify a hex color as 'R' (red), 'A' (amber), 'G' (green), or None.
    Rejects near-greys (saturation < 0.30) and near-blacks (value < 0.20).
    """
    try:
        r = int(hex_val[0:2], 16) / 255
        g = int(hex_val[2:4], 16) / 255
        b = int(hex_val[4:6], 16) / 255
        h, s, v = colorsys.rgb_to_hsv(r, g, b)
    except Exception:
        return None
    if s < 0.30 or v < 0.20:
        return None
    hue = h * 360
    if hue <= 25 or hue >= 335:
        return "R"
    if 26 <= hue <= 65:
        return "A"
    if 90 <= hue <= 170:
        return "G"
    return None


def _grpsp_direct_sp_fills(grpSp):
    """Return [(sp_elem, hex_fill, (x, y, cx, cy))] for direct <p:sp> children
    of a grpSp that have an explicit solidFill and a readable xfrm."""
    results = []
    for child in grpSp:
        if child.tag.split("}")[-1] != "sp":
            continue
        spPr = child.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        sf = spPr.find(f"{{{NS_A}}}solidFill")
        if sf is None:
            continue
        srgb = sf.find(f"{{{NS_A}}}srgbClr")
        if srgb is None:
            continue
        hex_val = srgb.get("val", "").upper()
        xfrm = spPr.find(f"{{{NS_A}}}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{{{NS_A}}}off")
        ext = xfrm.find(f"{{{NS_A}}}ext")
        if off is None or ext is None:
            continue
        try:
            x, y   = int(off.get("x", 0)), int(off.get("y", 0))
            cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
        except ValueError:
            continue
        results.append((child, hex_val, (x, y, cx, cy)))
    return results


def _shapes_uniform_and_linear(xfrms):
    """Return True if a list of (x, y, cx, cy) tuples are:
    1. Similar in size — each cx/cy within 50% of the median.
    2. Linearly arranged — all x-centers collinear (horizontal) OR
       all y-centers collinear (vertical), within 1× median dimension.
    3. Tightly and evenly spaced — gaps ≤ 2× median dimension on the
       primary axis, and no gap more than 3× any other gap.
    """
    if len(xfrms) < 2:
        return False

    cxs = sorted(t[2] for t in xfrms)
    cys = sorted(t[3] for t in xfrms)
    med_cx = cxs[len(cxs) // 2]
    med_cy = cys[len(cys) // 2]
    if med_cx == 0 or med_cy == 0:
        return False

    # Size uniformity
    if any(abs(c - med_cx) / med_cx > 0.50 for c in cxs):
        return False
    if any(abs(c - med_cy) / med_cy > 0.50 for c in cys):
        return False

    cx_list = [t[0] + t[2] // 2 for t in xfrms]  # center-x of each shape
    cy_list = [t[1] + t[3] // 2 for t in xfrms]  # center-y of each shape

    is_horiz = (max(cy_list) - min(cy_list)) <= med_cy
    is_vert  = (max(cx_list) - min(cx_list)) <= med_cx
    if not (is_horiz or is_vert):
        return False

    # Spacing uniformity along the primary axis
    if is_horiz:
        ordered = sorted(xfrms, key=lambda t: t[0])
        gaps = [ordered[i][0] - (ordered[i-1][0] + ordered[i-1][2])
                for i in range(1, len(ordered))]
        axis_dim = med_cx
    else:
        ordered = sorted(xfrms, key=lambda t: t[1])
        gaps = [ordered[i][1] - (ordered[i-1][1] + ordered[i-1][3])
                for i in range(1, len(ordered))]
        axis_dim = med_cy

    if gaps:
        if max(gaps) > 2 * axis_dim:
            return False
        pos_gaps = [g for g in gaps if g > 0]
        if len(pos_gaps) >= 2 and max(pos_gaps) > 3 * min(pos_gaps):
            return False

    return True


def _is_rag_gradient_shape(sp):
    """Return True if this shape has a gradFill whose stops span R, A, and G."""
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return False
    gf = spPr.find(f"{{{NS_A}}}gradFill")
    if gf is None:
        return False
    gsLst = gf.find(f"{{{NS_A}}}gsLst")
    if gsLst is None:
        return False
    buckets = set()
    for gs in gsLst.findall(f"{{{NS_A}}}gs"):
        sf = gs.find(f"{{{NS_A}}}solidFill")
        if sf is None:
            continue
        srgb = sf.find(f"{{{NS_A}}}srgbClr")
        if srgb is not None:
            b = _hue_bucket(srgb.get("val", ""))
            if b:
                buckets.add(b)
    return {"R", "A", "G"} <= buckets


def detect_rag_colors(slide_root):
    """Return a set of hex colors to preserve (not remap) on this slide.

    Detection is shape-structure based — not slide-level heuristics.

    (A) Group-level geometric uniformity:
        A <p:grpSp> whose direct <p:sp> children have fills spanning all three
        hue buckets (R/A/G), are similar in size, linearly arranged, and tightly
        spaced → preserve those exact fill hex values.

    (B) Gradient-stop detection:
        A single <p:sp> with a gradFill whose stops span R+A+G → preserve all
        gradient stop hex values on that shape.
    """
    preserve = set()

    # (A) Group-level geometric uniformity
    for grpSp in slide_root.iter(f"{{{NS_P}}}grpSp"):
        sp_fills = _grpsp_direct_sp_fills(grpSp)
        if not (2 <= len(sp_fills) <= 6):
            continue
        buckets = {_hue_bucket(h) for _, h, _ in sp_fills} - {None}
        if not ({"R", "A", "G"} <= buckets):
            continue
        xfrms = [xf for _, _, xf in sp_fills]
        if not _shapes_uniform_and_linear(xfrms):
            continue
        for _, hex_val, _ in sp_fills:
            preserve.add(hex_val)

    # (B) Gradient RAG shapes
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        if not _is_rag_gradient_shape(sp):
            continue
        spPr = sp.find(f"{{{NS_P}}}spPr")
        gf   = spPr.find(f"{{{NS_A}}}gradFill")
        for gs in gf.find(f"{{{NS_A}}}gsLst").findall(f"{{{NS_A}}}gs"):
            sf   = gs.find(f"{{{NS_A}}}solidFill")
            srgb = sf.find(f"{{{NS_A}}}srgbClr") if sf is not None else None
            if srgb is not None:
                preserve.add(srgb.get("val", "").upper())

    return preserve

# ── Layout semantic map ───────────────────────────────────────────────────────
# Maps substrings in old layout names → X/Y master layout names (checked in order).
# "title sl" (Title Slide) → Title_Only_Light, not Cover_Light — those are transitional
# slides with a title + maybe subtitle, not a branded cover.
# Cover_Light is reserved for actual cover slides (slide 1, photo bg, branded).
# Content scan in resolve_layout() can still override based on actual slide content.
LAYOUT_SEMANTIC = [
    ("cover",    "Cover_Light"),
    ("title sl", "Title_Only_Light"),  # "Title Slide" → title-only, not branded cover
    ("title on", "Title_Only_Light"),  # "Title Only"
    ("photo",    "Cover_Photo"),
    ("image",    "Cover_Photo"),
    ("full",     "Cover_Photo"),
    ("section",      "Divider_Dark"),
    ("divider",      "Divider_Dark"),
    ("interstitial", "Divider_Dark"),
    ("transition",   "Divider_Dark"),
    ("chapter",      "Divider_Dark"),
    ("thank",    "Ending_PivotPurple"),
    ("ending",   "Ending_PivotPurple"),
    ("blank",    "Blank_Light"),
    ("subtitle", "Content_Light_Sub"),
    ("two col",  "Content_Light"),
    ("dark",     "1_Content_Dark"),    # vF has no plain Content_Dark
]

# Default layout when no semantic match — general content slide
DEFAULT_LAYOUT = "Content_Light"

# Content-scan thresholds for layout override
_SPARSE_TEXT_THRESHOLD  = 80   # total chars below this = sparse/title-only slide
_HEAVY_SHAPE_COUNT      = 4    # non-ph shapes above this = content-heavy slide
_HEAVY_TEXT_THRESHOLD   = 120  # body placeholder chars above this = content-heavy

EMU = 914400  # 1 inch in EMU

# ── SVG media color remap ─────────────────────────────────────────────────────
SVG_COLOR_REMAP = {
    # Org_Efficiency ring graphic
    "6dcff6": "C8BFFF",   # cyan arc → Pivot Purple Tint 3
    "004a8f": "917FFF",   # dark blue arc → Pivot Purple Tint 1
}

def remap_chart_colors(file_map):
    """Apply COLOR_REMAP to all chart XML files (ppt/charts/*.xml) and
    drawing XML files (ppt/drawings/*.xml — used by OLE/VML objects).
    Also handles chartEx (extended chart type) files."""
    NS_C  = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    count = 0
    for key in list(file_map):
        if not (re.match(r"ppt/charts/.*\.xml$", key) or
                re.match(r"ppt/drawings/.*\.xml$", key)):
            continue
        try:
            root = etree.fromstring(file_map[key])
        except Exception:
            continue
        changed = False
        for srgb in root.iter(f"{{{NS_A}}}srgbClr"):
            v = srgb.get("val", "").upper()
            new_v = _resolve_color(v, set())
            if new_v:
                srgb.set("val", new_v)
                changed = True
                count += 1
        if changed:
            file_map[key] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return count


def remap_svg_colors(file_map):
    """Remap all hex colors in SVG media files to brand equivalents.

    First applies the explicit SVG_COLOR_REMAP lookup (fastest, exact matches),
    then runs every remaining #RRGGBB value through _resolve_color so no
    non-brand color can survive — same guarantee as XML slide processing."""
    count = 0
    for key in list(file_map):
        if not key.endswith(".svg"):
            continue
        svg = file_map[key].decode("utf-8", errors="replace")
        changed = False

        # 1. Explicit overrides (from SVG_COLOR_REMAP)
        for old, new in SVG_COLOR_REMAP.items():
            for variant in (old.upper(), old.lower(), old.capitalize()):
                if variant in svg:
                    svg = svg.replace(variant, new.upper())
                    changed = True

        # 2. Snap every remaining #RRGGBB through _resolve_color
        def _replace_hex(m):
            hex6 = m.group(1).upper()
            mapped = _resolve_color(hex6, set())
            if mapped and mapped.upper() != hex6:
                return "#" + mapped.upper()
            return m.group(0)

        new_svg = re.sub(r"#([0-9A-Fa-f]{6})\b", _replace_hex, svg)
        if new_svg != svg:
            svg = new_svg
            changed = True

        if changed:
            file_map[key] = svg.encode("utf-8")
            count += 1
    return count


# ── STEP 1: Master swap ───────────────────────────────────────────────────────
def swap_master(file_map, v7_map):
    """
    Remove old master + layouts + theme from file_map.
    Inject v7 master + layouts + theme.
    Returns updated file_map and the new v7 layout name→path dict.
    """
    ct_root = etree.fromstring(file_map["[Content_Types].xml"])
    prs_root = etree.fromstring(file_map["ppt/presentation.xml"])
    prs_rels = etree.fromstring(file_map["ppt/_rels/presentation.xml.rels"])

    # ── Collect OLE embeddings referenced by old masters (before deleting) ────
    ole_paths = set()
    for rels_key in [k for k in file_map if re.match(r"ppt/slideMasters/_rels/", k)]:
        try:
            rels_r = etree.fromstring(file_map[rels_key])
            for rel in rels_r:
                tgt = rel.get("Target", "")
                if "oleObject" in tgt or "embedding" in tgt.lower():
                    if "../" in tgt:
                        ole_paths.add("ppt/" + tgt.replace("../", ""))
                    else:
                        ole_paths.add(tgt)
        except Exception:
            pass

    # ── Remove old master / layout / theme files ──────────────────────────────
    old_keys = [k for k in file_map if re.match(
        r"ppt/(slideMasters|slideLayouts|theme)/", k)]
    for k in old_keys:
        del file_map[k]

    # Delete orphaned OLE embedding files so PowerPoint doesn't show a repair dialog
    for ole_path in ole_paths:
        if ole_path in file_map:
            del file_map[ole_path]
            print(f"  Removed orphaned embedding: {ole_path}")

    # Remove Content_Type Overrides for old master / layouts / theme / embeddings
    for el in list(ct_root):
        pn = el.get("PartName", "")
        if re.search(r"/(slideMasters|slideLayouts|theme)/", pn):
            ct_root.remove(el)
        elif pn.lstrip("/") in {p.lstrip("/") for p in ole_paths}:
            ct_root.remove(el)

    # Remove presentation.xml rels pointing to old master or theme
    for rel in list(prs_rels):
        t = rel.get("Type", "")
        if t in (SLIDE_MASTER_REL, THEME_REL):
            prs_rels.remove(rel)

    # ── Build master media rename map to avoid collision with slide media ────
    # Master layouts reference ppt/media/imageN.ext — same names used by slides.
    # Rename all master-referenced media with a "master_" prefix.
    master_media_refs = {}   # old path → new path
    for dkey in v7_map:
        if not re.match(r"ppt/(slideMasters|slideLayouts)/", dkey): continue
        if not dkey.endswith(".xml"): continue
        parts    = dkey.split("/")
        rels_key = "/".join(parts[:-1]) + "/_rels/" + parts[-1] + ".rels"
        if rels_key not in v7_map: continue
        rels_r = etree.fromstring(v7_map[rels_key])
        for rel in rels_r:
            tgt = rel.get("Target", "")
            if "../media/" in tgt:
                old_path = "ppt/" + tgt.replace("../", "")
                if old_path not in master_media_refs:
                    ext      = Path(old_path).suffix
                    new_name = f"ppt/media/master_{Path(old_path).stem}{ext}"
                    master_media_refs[old_path] = new_name

    # ── Inject v7 master / layouts / theme (rewriting media refs in rels) ────
    v7_master_files = [k for k in v7_map if re.match(
        r"ppt/(slideMasters|slideLayouts|theme)/", k)]
    for k in v7_master_files:
        data = v7_map[k]
        if k.endswith(".rels"):
            rels_r  = etree.fromstring(data)
            changed = False
            for rel in rels_r:
                tgt = rel.get("Target", "")
                if "../media/" in tgt:
                    old_path = "ppt/" + tgt.replace("../", "")
                    if old_path in master_media_refs:
                        new_fname = Path(master_media_refs[old_path]).name
                        rel.set("Target", f"../media/{new_fname}")
                        changed = True
            if changed:
                data = etree.tostring(rels_r, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
        file_map[k] = data

    # Copy renamed master media bytes
    for old_path, new_path in master_media_refs.items():
        if old_path in v7_map:
            file_map[new_path] = v7_map[old_path]
    print(f"  Master media: {len(master_media_refs)} file(s) copied with master_ prefix")

    # Copy Content_Types from v7 for master / layouts / theme
    v7_ct = etree.fromstring(v7_map["[Content_Types].xml"])
    existing_parts = {el.get("PartName","") for el in ct_root}
    for el in v7_ct:
        pn = el.get("PartName", "")
        if pn and re.search(r"/(slideMasters|slideLayouts|theme)/", pn):
            if pn not in existing_parts:
                ct_root.append(el.__copy__() if hasattr(el,'__copy__') else
                               etree.fromstring(etree.tostring(el)))

    # Add master relationship to presentation rels
    # Find max rId in prs_rels
    max_n = max(
        (int(m.group(1)) for r in prs_rels
         for m in [re.match(r"rId(\d+)", r.get("Id",""))] if m),
        default=0
    )
    master_rid = f"rId{max_n + 1}"
    rel = etree.SubElement(prs_rels, "Relationship")
    rel.set("Id", master_rid)
    rel.set("Type", SLIDE_MASTER_REL)
    rel.set("Target", "slideMasters/slideMaster1.xml")

    # Update sldMasterIdLst in presentation.xml
    sldMasterIdLst = prs_root.find(f"{{{NS_P}}}sldMasterIdLst")
    if sldMasterIdLst is None:
        sldMasterIdLst = etree.SubElement(prs_root, f"{{{NS_P}}}sldMasterIdLst")
    for el in list(sldMasterIdLst):
        sldMasterIdLst.remove(el)
    mid_el = etree.SubElement(sldMasterIdLst, f"{{{NS_P}}}sldMasterId")
    mid_el.set("id", "2147483648")
    mid_el.set(f"{{{NS_R}}}id", master_rid)

    # Start slide numbering at 1
    prs_root.set("firstSlideNum", "1")

    # Patch theme scheme colors to vF brand palette (slide 4)
    # dk1/lt1 intentionally left untouched — those control default text/background
    # and the vF master already sets them correctly via explicit srgbClr.
    SCHEME_PATCH = {
        "dk2":     "281A42",   # Vector Dark Purple  — secondary headings/dark bg
        "lt2":     "F2F2F2",   # Lightest Grey       — light background tint
        "accent1": "765FFF",   # Pivot Purple        — primary brand
        "accent2": "281A42",   # Vector Dark Purple  — dark complement
        "accent3": "00C27A",   # Anchor Green
        "accent4": "60BDBC",   # Talent Teal
        "accent5": "FFB547",   # Momentum Amber
        "accent6": "FF2E88",   # Signal Magenta
        "hlink":   "765FFF",   # Pivot Purple        — hyperlinks
        "folHlink":"917FFF",   # Tint 1              — visited links
    }
    theme_key = next((k for k in file_map if re.search(r"ppt/theme/theme\d+\.xml$", k)), None)
    if theme_key:
        theme_root = etree.fromstring(file_map[theme_key])
        clrScheme = theme_root.find(f".//{{{NS_A}}}clrScheme")
        if clrScheme is not None:
            for el in clrScheme:
                name = el.tag.split("}")[1]
                if name in SCHEME_PATCH:
                    for child in el:
                        child.set("val", SCHEME_PATCH[name])
        file_map[theme_key] = etree.tostring(
            theme_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── Patch notes master rels: redirect any theme ref to the new theme1.xml ──
    # After the master swap the old theme files are gone; notes masters that still
    # reference them (e.g. theme9.xml) cause a repair dialog on open.
    new_theme_path = next(
        (k for k in file_map if re.match(r"ppt/theme/theme\d+\.xml$", k)), None)
    for nm_rels_key in [k for k in file_map if re.match(r"ppt/notesMasters/_rels/", k)]:
        nm_rels = etree.fromstring(file_map[nm_rels_key])
        changed = False
        for rel in nm_rels:
            if rel.get("Type", "").endswith("/theme"):
                tgt = rel.get("Target", "")
                base = re.sub(r"_rels/[^/]+$", "", nm_rels_key)
                import os as _os
                resolved = _os.path.normpath(base + tgt).lstrip("/")
                if resolved not in file_map and new_theme_path:
                    rel.set("Target", "../theme/" + new_theme_path.split("/")[-1])
                    changed = True
        if changed:
            file_map[nm_rels_key] = etree.tostring(
                nm_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            print(f"  Notes master theme ref updated → {new_theme_path}")

    # Write back prs files
    file_map["ppt/presentation.xml"] = etree.tostring(
        prs_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    file_map["ppt/_rels/presentation.xml.rels"] = etree.tostring(
        prs_rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    file_map["[Content_Types].xml"] = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Return the v7 layout name → path map
    v7_master = etree.fromstring(file_map["ppt/slideMasters/slideMaster1.xml"])
    v7_m_rels = etree.fromstring(file_map["ppt/slideMasters/_rels/slideMaster1.xml.rels"])
    rid_map   = {r.get("Id"): r.get("Target") for r in v7_m_rels}
    idlst     = v7_master.find(f"{{{NS_P}}}sldLayoutIdLst")
    name_map  = {}
    for el in idlst:
        rid = el.get(f"{{{NS_R}}}id")
        tgt = rid_map.get(rid,"").replace("../","ppt/")
        try:
            lx   = etree.fromstring(file_map[tgt])
            name = lx.find(f"{{{NS_P}}}cSld").get("name","")
            name_map[name] = tgt
        except Exception:
            pass
    # ── Sanitize all theme XML files ──────────────────────────────────────────
    # The v7 master PPTX may bundle non-brand themes (e.g. a source deck theme2
    # with 0F9ED5 turquoise that was baked in when the master was authored).
    # Patch every srgbClr in every ppt/theme/*.xml to a brand equivalent so that
    # any schemeClr references that survive in slides resolve to brand colors.
    n_theme_patched = 0
    for key in list(file_map):
        if not re.match(r"ppt/theme/.*\.xml$", key):
            continue
        try:
            t_root = etree.fromstring(file_map[key])
        except Exception:
            continue
        changed = False
        for srgb in t_root.iter(f"{{{NS_A}}}srgbClr"):
            v = srgb.get("val", "").upper()
            mapped = _resolve_color(v, set())
            if mapped and mapped != v:
                srgb.set("val", mapped)
                changed = True
                n_theme_patched += 1
        if changed:
            file_map[key] = etree.tostring(
                t_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if n_theme_patched:
        print(f"  Theme sanitization: {n_theme_patched} non-brand color(s) patched across theme files.")

    print(f"  Master injected. v7 layouts available: {len(name_map)}")
    return name_map




# ── STEP 2: Layout remapping ──────────────────────────────────────────────────
def _shape_text(sp):
    """Plain text content of a shape element."""
    return "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()


def _shape_area(sp):
    """Area of a shape in EMU², or 0 if not determinable."""
    ext = sp.find(f".//{{{NS_A}}}xfrm/{{{NS_A}}}ext")
    if ext is None:
        return 0
    try:
        return int(ext.get("cx", 0)) * int(ext.get("cy", 0))
    except (ValueError, TypeError):
        return 0


def _shape_font_size(sp):
    """Best-guess font size in hundredths-of-pt from the shape's first run."""
    for rPr in sp.iter(f"{{{NS_A}}}rPr"):
        sz = rPr.get("sz")
        if sz:
            try:
                return int(sz)
            except ValueError:
                pass
    for defRPr in sp.iter(f"{{{NS_A}}}defRPr"):
        sz = defRPr.get("sz")
        if sz:
            try:
                return int(sz)
            except ValueError:
                pass
    return 0


def _has_colored_block_title(slide_root):
    """
    Returns True if the slide has the 'section divider' visual pattern.

    Patterns detected:
    (A) A large non-placeholder shape (that is not explicitly transparent)
        which itself contains text.  Covers decks where the colored band IS
        the text container.

    (B) A large non-placeholder shape (not explicitly transparent) with no
        text + a title/ctrTitle placeholder with text + no body placeholder.
        Covers the common case: full-bleed colored bg rectangle as a separate
        shape, with text in a title placeholder on top.

    (C) Slide-level <p:bg> colored fill + title placeholder with text +
        no body placeholder.

    Fill detection is intentionally broad: a shape is treated as "filled"
    unless it has an *explicit* <a:noFill> as a direct child of spPr.
    Shapes that inherit their fill from the theme style (p:style/fillRef) or
    from the master have no explicit fill element in spPr — checking only for
    solidFill/gradFill misses them entirely.
    """
    SLIDE_W, SLIDE_H = 12_192_000, 6_858_000
    MIN_AREA = SLIDE_W * SLIDE_H * 0.08   # block must cover ≥8% of slide

    has_large_colored_bg = False
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        if sp.find(f".//{{{NS_P}}}ph") is not None:
            continue
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        # Skip shapes that are explicitly transparent (noFill as direct child of spPr)
        if spPr.find(f"{{{NS_A}}}noFill") is not None:
            continue
        if _shape_area(sp) < MIN_AREA:
            continue
        # Pattern A: large non-transparent shape contains its own text
        if _shape_text(sp):
            return True
        has_large_colored_bg = True

    # Pattern B: large colored bg shape + title ph with text + no body ph
    # Pattern C: slide-level <p:bg> colored fill + title ph with text + no body ph
    #            (covers full-bleed colored backgrounds set at the slide level,
    #             e.g. a purple "SUCCESSION FOR VALUE" slide where the color is
    #             the slide background, not a shape)
    cSld = slide_root.find(f"{{{NS_P}}}cSld")
    has_slide_bg_fill = False
    if cSld is not None:
        bg = cSld.find(f"{{{NS_P}}}bg")
        if bg is not None:
            bgPr = bg.find(f"{{{NS_P}}}bgPr")
            if bgPr is not None:
                has_slide_bg_fill = (
                    bgPr.find(f".//{{{NS_A}}}solidFill") is not None or
                    bgPr.find(f".//{{{NS_A}}}gradFill")  is not None
                )

    if has_large_colored_bg or has_slide_bg_fill:
        has_text    = False   # any title ph OR non-ph text box
        has_body_ph = False
        for sp in slide_root.iter(f"{{{NS_P}}}sp"):
            ph = sp.find(f".//{{{NS_P}}}ph")
            if ph is not None:
                pt = ph.get("type", "body")
                if pt in ("title", "ctrTitle") and _shape_text(sp):
                    has_text = True
                elif pt in ("body", "obj"):
                    has_body_ph = True
            else:
                # Non-placeholder text box on top of a colored bg → divider label
                if _shape_text(sp):
                    has_text = True
        if has_text and not has_body_ph:
            return True

    # Pattern D: layout-background divider — slide has a single text-bearing shape
    # with large font (≥36pt) and no other content.  The colored background comes
    # entirely from the slide layout, so there are no detectable fill shapes on the
    # slide XML itself.  Example: slide 22 — one body ph, 48pt "TALENT to VALUE",
    # everything else is inherited from the layout.
    all_shapes = list(slide_root.iter(f"{{{NS_P}}}sp"))
    text_shapes = [sp for sp in all_shapes if _shape_text(sp)]
    if len(text_shapes) == 1:
        sp = text_shapes[0]
        ph = sp.find(f".//{{{NS_P}}}ph")
        ph_type = ph.get("type", "body") if ph is not None else None
        # Skip decorative footer-type placeholders
        if ph_type not in ("sldNum", "dt", "ftr"):
            if _shape_font_size(sp) >= 3600:  # ≥36pt
                return True

    return False


def _scan_slide_content(slide_root):
    """
    Returns a dict with content metrics used to validate / override layout picks.
      n_ph_title        — number of title placeholders
      n_ph_body         — number of body/content placeholders
      body_chars        — total characters in body placeholders
      total_chars       — total characters across all text shapes
      n_extra           — non-placeholder, non-background shapes (txBox, pic, group…)
      is_divider_block  — True if a large colored non-ph block with text is present
    """
    n_ph_title = n_ph_body = body_chars = total_chars = n_extra = 0
    SLIDE_W, SLIDE_H = 12_192_000, 6_858_000
    BG_AREA = SLIDE_W * SLIDE_H * 0.80   # shapes covering ≥80% of slide = background

    for el in slide_root.iter():
        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag

        if tag == "sp":
            ph = el.find(f".//{{{NS_P}}}ph")
            texts = "".join(
                t.text or "" for t in el.iter(f"{{{NS_A}}}t")
            ).strip()
            total_chars += len(texts)

            if ph is not None:
                ph_type = ph.get("type", "body")
                if ph_type == "title":
                    n_ph_title += 1
                elif ph_type in ("body", "subTitle", "obj"):
                    n_ph_body  += 1
                    body_chars += len(texts)
            else:
                off = el.find(f".//{{{NS_A}}}xfrm/{{{NS_A}}}off")
                ext = el.find(f".//{{{NS_A}}}xfrm/{{{NS_A}}}ext")
                if off is not None and ext is not None:
                    try:
                        area = int(ext.get("cx", 0)) * int(ext.get("cy", 0))
                        if area < BG_AREA:
                            n_extra += 1
                    except (ValueError, TypeError):
                        n_extra += 1
                else:
                    n_extra += 1

        elif tag in ("pic", "grpSp"):
            n_extra += 1

    return dict(
        n_ph_title=n_ph_title,
        n_ph_body=n_ph_body,
        body_chars=body_chars,
        total_chars=total_chars,
        n_extra=n_extra,
        is_divider_block=_has_colored_block_title(slide_root),
    )


def resolve_layout(old_name, v7_name_map, slide_root=None,
                   slide_idx: int = 0, total_slides: int = 0):
    """
    Pick the best FulcrumQ layout for a slide.

    1. Exact match on old layout name → use as-is (covers already-branded decks)
    2. Semantic match via LAYOUT_SEMANTIC substrings
    3. Default to Content_Light

    Content scan (when slide_root is provided) can override the semantic pick:
    - Slide 1 with no body content → Cover_Light
    - Last slide with thank-you text → Ending_PivotPurple
    - Cover_Light / Title_Only_Light with heavy content → Content_Light
    - Content_Light with sparse content (title-only) → Title_Only_Light
    Dark variants follow the same rules (preserve Dark when explicitly dark).
    """
    # ── Step 1: exact name match ───────────────────────────────────────────────
    if old_name in v7_name_map:
        return v7_name_map[old_name], "exact"

    # ── Step 2: semantic match ─────────────────────────────────────────────────
    matched_layout = None
    matched_label  = None
    for substr, v7_name in LAYOUT_SEMANTIC:
        if substr in old_name.lower() and v7_name in v7_name_map:
            matched_layout = v7_name
            matched_label  = f"→ {v7_name}"
            break

    if matched_layout is None:
        matched_layout = DEFAULT_LAYOUT
        matched_label  = f"→ {DEFAULT_LAYOUT}"

    # ── Step 3: content-scan override ─────────────────────────────────────────
    # If the semantic match already resolved to a Divider, trust it — the source
    # layout name is explicit evidence. Don't let content-scan downgrade it.
    if matched_layout in ("Divider_Dark", "Divider_Light"):
        layout_path = v7_name_map.get(matched_layout)
        if layout_path:
            return layout_path, f"{matched_label} (layout name → divider, preserved)"

    if slide_root is not None:
        m = _scan_slide_content(slide_root)
        is_dark    = "Dark" in matched_layout
        is_heavy   = (m["body_chars"] > _HEAVY_TEXT_THRESHOLD or
                      m["n_extra"]    > _HEAVY_SHAPE_COUNT)
        is_sparse  = (m["total_chars"] < _SPARSE_TEXT_THRESHOLD and
                      m["n_ph_body"] == 0 and m["n_extra"] == 0)

        # ── Cover detection ────────────────────────────────────────────────────
        # Slide 1 with no heavy body content → Cover.
        # Allow subtitle placeholder (subTitle ph counts as n_ph_body in the
        # scan, but is NOT real body content — exclude it from the gate).
        is_already_cover = matched_layout in ("Cover_Light", "Cover_Dark", "Cover_Photo")
        body_without_subtitle = m["body_chars"] < 60  # subtitle text is short
        if (slide_idx == 1 and not is_heavy and not is_already_cover
                and (m["n_ph_body"] == 0 or body_without_subtitle)):
            override = "Cover_Light"
            if override in v7_name_map:
                return v7_name_map[override], f"{matched_label} →cover_detect→ {override}"

        # ── Ending slide detection ─────────────────────────────────────────────
        is_already_ending = matched_layout in ("Ending_PivotPurple",)
        if (total_slides > 0 and slide_idx >= total_slides - 1
                and not is_already_ending):
            slide_text = " ".join(
                t.text or "" for t in slide_root.iter(f"{{{NS_A}}}t")
            ).lower()
            if any(kw in slide_text for kw in ("thank", "question", "contact", "connect")):
                override = "Ending_PivotPurple"
                if override in v7_name_map:
                    return v7_name_map[override], f"{matched_label} →ending_detect→ {override}"

        # Colored block with short text → section divider regardless of layout name
        if m["is_divider_block"] and not is_heavy:
            override = "Divider_Dark"
            if override in v7_name_map:
                return v7_name_map[override], f"{matched_label} →scan→ {override} (colored block)"

        # Cover or Title-only picked, but slide has real body content → content layout
        if matched_layout in ("Cover_Light", "Cover_Dark",
                               "Title_Only_Light", "Title_Only_Dark") and is_heavy:
            override = "1_Content_Dark" if is_dark else "Content_Light"
            if override in v7_name_map:
                return v7_name_map[override], f"{matched_label} →scan→ {override}"

        # Content layout picked, but slide is sparse (just a title) → title-only
        if matched_layout in ("Content_Light", "1_Content_Dark",
                               DEFAULT_LAYOUT) and is_sparse:
            override = "Title_Only_Dark" if is_dark else "Title_Only_Light"
            if override in v7_name_map:
                return v7_name_map[override], f"{matched_label} →scan→ {override}"

    layout_path = v7_name_map.get(matched_layout)
    if layout_path:
        return layout_path, matched_label

    # Last-resort: first layout in the map
    first = next(iter(v7_name_map.values()))
    return first, "fallback"


def remap_slide_layout(file_map, slide_path, v7_layout_path):
    """Rewrite a slide's rels to point to the new v7 layout."""
    rels_path = re.sub(r"ppt/slides/(slide\d+\.xml)$",
                       r"ppt/slides/_rels/\1.rels", slide_path)
    rels_bytes = file_map.get(rels_path, b"")
    if not rels_bytes:
        return

    rels_root = etree.fromstring(rels_bytes)
    layout_num = re.search(r"slideLayout(\d+)\.xml$", v7_layout_path).group(1)

    for rel in rels_root:
        if "slideLayout" in rel.get("Target", ""):
            rel.set("Target", f"../slideLayouts/slideLayout{layout_num}.xml")
            break

    file_map[rels_path] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ── STEP 2b: Title promotion ─────────────────────────────────────────────────
def promote_slide_title(slide_root):
    """
    If a slide has no title/ctrTitle placeholder, find the first content
    placeholder (body type with text) and promote it to title type.
    - Strips its explicit xfrm so it inherits the layout's title position/size.
    - Strips bullet formatting so the title style applies cleanly.
    Returns True if a promotion was made.
    """
    FOOTER_TYPES = {"sldNum", "dt", "ftr"}
    TITLE_TYPES  = {"title", "ctrTitle"}

    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return False

    # Already has a title ph — nothing to do
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "body") in TITLE_TYPES:
            return False

    # Find first body-type ph with actual text content
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None:
            continue
        if ph.get("type", "body") in FOOTER_TYPES | TITLE_TYPES:
            continue
        texts = [t.text for t in sp.iter(f"{{{NS_A}}}t")
                 if t.text and t.text.strip()]
        if not texts:
            continue

        # Promote: set ph type to title, drop idx
        ph.set("type", "title")
        ph.attrib.pop("idx", None)

        # Inherit position/size from layout's title placeholder
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is not None:
            for xfrm in spPr.findall(f"{{{NS_A}}}xfrm"):
                spPr.remove(xfrm)

        # Strip bullet formatting so title txStyle takes over
        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is not None:
            BU_TAGS = {
                f"{{{NS_A}}}buClr", f"{{{NS_A}}}buClrTx",
                f"{{{NS_A}}}buSzPct", f"{{{NS_A}}}buSzPts", f"{{{NS_A}}}buSzTx",
                f"{{{NS_A}}}buFontTx", f"{{{NS_A}}}buFont",
                f"{{{NS_A}}}buChar", f"{{{NS_A}}}buAutoNum", f"{{{NS_A}}}buNone",
            }
            for pPr in txBody.iter(f"{{{NS_A}}}pPr"):
                for tag in BU_TAGS:
                    for el in pPr.findall(tag):
                        pPr.remove(el)
                etree.SubElement(pPr, f"{{{NS_A}}}buNone")

        # ── Subtitle promotion ────────────────────────────────────────────────
        # If a second body-type ph with text exists, promote it to subTitle so
        # it inherits the layout's subtitle position/style rather than floating.
        for sp2 in spTree.findall(f"{{{NS_P}}}sp"):
            if sp2 is sp:
                continue
            ph2 = sp2.find(f".//{{{NS_P}}}ph")
            if ph2 is None:
                continue
            if (ph2.get("type", "body") in FOOTER_TYPES | TITLE_TYPES | {"subTitle"}
                    or ph2.get("idx") == "12"):
                continue
            texts2 = [t.text for t in sp2.iter(f"{{{NS_A}}}t")
                      if t.text and t.text.strip()]
            if not texts2:
                continue
            ph2.set("type", "body")
            ph2.set("idx",  "12")
            spPr2 = sp2.find(f"{{{NS_P}}}spPr")
            if spPr2 is not None:
                for xfrm in spPr2.findall(f"{{{NS_A}}}xfrm"):
                    spPr2.remove(xfrm)
            break

        return True

    return False


def split_title_subtitle(slide_root):
    """
    Detect a Shift+Enter (<a:br>) in a title placeholder.

    If found:
      - Extract all content after the first <a:br> (rest of that paragraph +
        any subsequent paragraphs in the same title shape) into a new
        subTitle placeholder (ph type="subTitle" idx="1").
      - Strip the <a:br> and post-break content from the title shape so it
        contains only the main heading text.

    The new placeholder carries no explicit xfrm — it inherits position/size
    from whatever _Sub layout the caller assigns after this call returns.

    Returns a list with the extracted subtitle text (for logging), or [] if
    no soft return was found.
    """
    TITLE_TYPES = {"title", "ctrTitle"}

    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return []

    # Find the title placeholder
    title_sp = None
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "") in TITLE_TYPES:
            title_sp = sp
            break
    if title_sp is None:
        return []

    txBody = title_sp.find(f"{{{NS_P}}}txBody")
    if txBody is None:
        return []

    paragraphs = txBody.findall(f"{{{NS_A}}}p")

    # Locate the first <a:br> across all paragraphs
    br_para_idx = br_elem_idx = None
    for pi, para in enumerate(paragraphs):
        for ei, elem in enumerate(para):
            if elem.tag == f"{{{NS_A}}}br":
                br_para_idx, br_elem_idx = pi, ei
                break
        if br_para_idx is not None:
            break

    if br_para_idx is None:
        # ── Inline-pipe fallback ─────────────────────────────────────────────────
        # Single paragraph with "Title text | Subtitle text" — pipe as separator.
        # Example: "FUTURE JOBS TO BE DONE & RISKS | CHIEF DIGITAL OFFICER"
        if len(paragraphs) >= 1:
            para0 = paragraphs[0]
            full_text = "".join(t.text or "" for t in para0.iter(f"{{{NS_A}}}t"))
            pipe_idx = None
            for sep in (" | ", "| ", " |"):
                idx_ = full_text.find(sep)
                if idx_ != -1:
                    pipe_idx = idx_
                    sep_len  = len(sep)
                    break
            if pipe_idx is not None:
                title_text_raw  = full_text[:pipe_idx].strip()
                sub_text_raw    = full_text[pipe_idx + sep_len:].strip()
                if title_text_raw and sub_text_raw:
                    # Rewrite title paragraph runs to contain only title text
                    for r in list(para0.findall(f"{{{NS_A}}}r")):
                        para0.remove(r)
                    new_r = etree.SubElement(para0, f"{{{NS_A}}}r")
                    etree.SubElement(new_r, f"{{{NS_A}}}t").text = title_text_raw
                    # Build subtitle shape
                    max_id = max(
                        (int(el.get("id", 0)) for el in slide_root.iter(f"{{{NS_P}}}cNvPr")),
                        default=100,
                    )
                    sub_sp  = etree.SubElement(spTree, f"{{{NS_P}}}sp")
                    nvSpPr  = etree.SubElement(sub_sp, f"{{{NS_P}}}nvSpPr")
                    cNvPr   = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvPr")
                    cNvPr.set("id", str(max_id + 1)); cNvPr.set("name", "SubtitleSplit")
                    cNvSpPr = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvSpPr")
                    etree.SubElement(cNvSpPr, f"{{{NS_A}}}spLocks").set("noGrp", "1")
                    nvPr    = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")
                    ph_el   = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
                    ph_el.set("type", "subTitle"); ph_el.set("idx", "1")
                    spPr2   = etree.SubElement(sub_sp, f"{{{NS_P}}}spPr")
                    etree.SubElement(spPr2, f"{{{NS_A}}}noFill")
                    sub_tx  = etree.SubElement(sub_sp, f"{{{NS_P}}}txBody")
                    etree.SubElement(sub_tx, f"{{{NS_A}}}bodyPr")
                    etree.SubElement(sub_tx, f"{{{NS_A}}}lstStyle")
                    new_p   = etree.SubElement(sub_tx, f"{{{NS_A}}}p")
                    sub_r   = etree.SubElement(new_p, f"{{{NS_A}}}r")
                    etree.SubElement(sub_r, f"{{{NS_A}}}t").text = sub_text_raw
                    etree.SubElement(new_p, f"{{{NS_A}}}endParaRPr").set("lang", "en-US")
                    return [sub_text_raw]

        # ── Multi-paragraph fallback ────────────────────────────────────────────
        # No soft return, but title has ≥2 non-empty paragraphs:
        #   paragraph 0 = title, paragraph 1+ = subtitle
        # Also covers "Title | Subtitle" styled as two separate paragraphs.
        non_empty = [p for p in paragraphs
                     if "".join(t.text or "" for t in p.iter(f"{{{NS_A}}}t")).strip()]
        if len(non_empty) < 2:
            return []
        # Treat paragraphs[1:] as subtitle — same path as the br case
        br_para_idx  = paragraphs.index(non_empty[0])
        br_elem_idx  = len(list(non_empty[0]))  # "break" is at the END of para 0
        br_para      = non_empty[0]
        br_children  = list(br_para)
        post_br_runs = []                       # nothing after the break in para 0
        tail_paras   = paragraphs[br_para_idx + 1:]
        # Strip trailing pipe/colon from the title paragraph
        for r in reversed(list(br_para.iter(f"{{{NS_A}}}t"))):
            if r.text:
                r.text = r.text.rstrip(" |:–—")
                break
        # Remove tail paragraphs from title shape now (skip the later removal block)
        for para in tail_paras:
            txBody.remove(para)
        # Build and inject the subtitle shape, then return
        max_id = max(
            (int(el.get("id", 0)) for el in slide_root.iter(f"{{{NS_P}}}cNvPr")),
            default=100,
        )
        sub_sp  = etree.SubElement(spTree, f"{{{NS_P}}}sp")
        nvSpPr  = etree.SubElement(sub_sp, f"{{{NS_P}}}nvSpPr")
        cNvPr   = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvPr")
        cNvPr.set("id", str(max_id + 1)); cNvPr.set("name", "SubtitleSplit")
        cNvSpPr = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvSpPr")
        etree.SubElement(cNvSpPr, f"{{{NS_A}}}spLocks").set("noGrp", "1")
        nvPr    = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")
        ph_el   = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
        ph_el.set("type", "subTitle"); ph_el.set("idx", "1")
        spPr2   = etree.SubElement(sub_sp, f"{{{NS_P}}}spPr")
        etree.SubElement(spPr2, f"{{{NS_A}}}noFill")
        sub_tx  = etree.SubElement(sub_sp, f"{{{NS_P}}}txBody")
        etree.SubElement(sub_tx, f"{{{NS_A}}}bodyPr")
        etree.SubElement(sub_tx, f"{{{NS_A}}}lstStyle")
        for para in tail_paras:
            new_p = etree.SubElement(sub_tx, f"{{{NS_A}}}p")
            for elem in para:
                new_p.append(deepcopy(elem))
            if new_p.find(f"{{{NS_A}}}endParaRPr") is None:
                etree.SubElement(new_p, f"{{{NS_A}}}endParaRPr").set("lang", "en-US")
        if not sub_tx.findall(f"{{{NS_A}}}p"):
            etree.SubElement(sub_tx, f"{{{NS_A}}}p")
        subtitle_text = " ".join(t.text or "" for t in sub_sp.iter(f"{{{NS_A}}}t")).strip()
        return [subtitle_text or "(subtitle extracted)"]

    br_para = paragraphs[br_para_idx]
    br_children = list(br_para)

    # ── Build subtitle paragraphs from extracted content ──────────────────────
    # First sub-paragraph: runs after the <a:br> in the same paragraph
    post_br_runs = [
        e for e in br_children[br_elem_idx + 1:]
        if e.tag != f"{{{NS_A}}}endParaRPr"
    ]
    # Subsequent full paragraphs (everything after the break paragraph)
    tail_paras = paragraphs[br_para_idx + 1:]

    # ── Trim the title in-place ────────────────────────────────────────────────
    # Remove <a:br> and everything after it from the break paragraph
    for elem in br_children[br_elem_idx:]:
        if elem.tag != f"{{{NS_A}}}endParaRPr":
            br_para.remove(elem)
    # Remove all paragraphs that follow the break paragraph
    for para in tail_paras:
        txBody.remove(para)

    # ── Inject subtitle placeholder ────────────────────────────────────────────
    max_id = max(
        (int(el.get("id", 0)) for el in slide_root.iter(f"{{{NS_P}}}cNvPr")),
        default=100,
    )
    sub_sp   = etree.SubElement(spTree, f"{{{NS_P}}}sp")
    nvSpPr   = etree.SubElement(sub_sp, f"{{{NS_P}}}nvSpPr")
    cNvPr    = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvPr")
    cNvPr.set("id", str(max_id + 1))
    cNvPr.set("name", "SubtitleSplit")
    cNvSpPr  = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvSpPr")
    etree.SubElement(cNvSpPr, f"{{{NS_A}}}spLocks").set("noGrp", "1")
    nvPr     = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")
    ph_el    = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
    ph_el.set("type", "subTitle")
    ph_el.set("idx",  "1")
    spPr     = etree.SubElement(sub_sp, f"{{{NS_P}}}spPr")
    etree.SubElement(spPr, f"{{{NS_A}}}noFill")   # no explicit xfrm — inherit from layout

    sub_tx   = etree.SubElement(sub_sp, f"{{{NS_P}}}txBody")
    etree.SubElement(sub_tx, f"{{{NS_A}}}bodyPr")
    etree.SubElement(sub_tx, f"{{{NS_A}}}lstStyle")

    # First subtitle paragraph (post-break runs)
    if post_br_runs:
        src_pPr = br_para.find(f"{{{NS_A}}}pPr")
        new_p   = etree.SubElement(sub_tx, f"{{{NS_A}}}p")
        if src_pPr is not None:
            new_p.append(deepcopy(src_pPr))
        for run in post_br_runs:
            new_p.append(deepcopy(run))
        etree.SubElement(new_p, f"{{{NS_A}}}endParaRPr").set("lang", "en-US")

    # Remaining paragraphs
    for para in tail_paras:
        new_p = etree.SubElement(sub_tx, f"{{{NS_A}}}p")
        for elem in para:
            new_p.append(deepcopy(elem))
        if new_p.find(f"{{{NS_A}}}endParaRPr") is None:
            etree.SubElement(new_p, f"{{{NS_A}}}endParaRPr").set("lang", "en-US")

    # If nothing landed in the txBody, add an empty paragraph so it's valid
    if not sub_tx.findall(f"{{{NS_A}}}p"):
        etree.SubElement(sub_tx, f"{{{NS_A}}}p")

    subtitle_text = " ".join(
        t.text or "" for t in sub_sp.iter(f"{{{NS_A}}}t")
    ).strip()
    return [subtitle_text or "(subtitle extracted)"]


# Layout upgrade map for subtitle-detected slides
_SUB_LAYOUT_MAP = {
    "Content_Light":    "Content_Light_Sub",
    "Content_Dark":     "Content_Dark_Sub",
    "Title_Only_Light": "Content_Light_Sub",
    "Title_Only_Dark":  "Content_Dark_Sub",
    "1_Content_Dark":   "Content_Dark_Sub",
}


def snap_subtitle_to_layout(slide_root, layout_bytes):
    """
    Autosnap subtitle-vibe placeholders to their v7 layout position.

    A placeholder is "subtitle-vibe" if:
      - It has type="subTitle" (explicit subtitle placeholder), OR
      - It has type="body" with idx=1 and the slide also has a title placeholder
        (first body slot = where subtitle lives on title/content slides)

    For matched placeholders, strip any explicit <a:xfrm> from <p:spPr> so the
    shape inherits its position and size from the v7 layout instead of carrying
    over the old master's coordinates.

    Returns count of placeholders snapped.
    """
    FOOTER_TYPES = {"dt", "ftr", "sldNum"}
    TITLE_TYPES  = {"title", "ctrTitle"}

    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return 0

    # Does this slide have any title placeholder?
    has_title = False
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "") in TITLE_TYPES:
            has_title = True
            break

    # Which (type, idx) pairs exist in the v7 layout?
    layout_root = etree.fromstring(layout_bytes)
    layout_ph_keys = set()
    for sp in layout_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None:
            continue
        pt = ph.get("type", "body")
        pi = ph.get("idx", "0")
        if pt not in FOOTER_TYPES:
            layout_ph_keys.add((pt, pi))

    snapped = 0
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None:
            continue
        ph_type = ph.get("type", "body")
        ph_idx  = ph.get("idx",  "0")

        if ph_type in FOOTER_TYPES | TITLE_TYPES:
            continue

        # Subtitle vibe check
        is_subtitle_vibe = (
            ph_type == "subTitle"
            or (ph_type == "body" and ph_idx == "1" and has_title)
            or (ph_type == "body" and ph_idx == "12")
        )
        if not is_subtitle_vibe:
            continue

        # Only snap if the v7 layout actually has a matching slot to land on.
        # Match by (type, idx) first; fall back to idx-only so subTitle↔body
        # cross-type pairs still snap correctly.
        has_layout_slot = (
            (ph_type, ph_idx) in layout_ph_keys
            or ("subTitle", ph_idx) in layout_ph_keys
            or any(k[1] == ph_idx for k in layout_ph_keys
                   if k[0] not in TITLE_TYPES)
        )
        if not has_layout_slot:
            continue

        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        xfrms = spPr.findall(f"{{{NS_A}}}xfrm")
        for xfrm in xfrms:
            spPr.remove(xfrm)
        if xfrms:
            snapped += 1

    return snapped


NS_R_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
TRIANGLE_MEDIA_KEY = "ppt/media/triangle_bullet.png"

def _ensure_triangle_media(file_map, triangle_bytes):
    """Embed triangle PNG into file_map once; ensure Content-Type exists."""
    file_map[TRIANGLE_MEDIA_KEY] = triangle_bytes
    # Ensure png Default Content-Type
    ct_root = etree.fromstring(file_map["[Content_Types].xml"])
    has_png = any(el.get("Extension") == "png" for el in ct_root)
    if not has_png:
        d = etree.SubElement(ct_root, "Default")
        d.set("Extension", "png")
        d.set("ContentType", "image/png")
        file_map["[Content_Types].xml"] = etree.tostring(
            ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)


def promote_bullet_txbox(slide_root, file_map, slide_path):
    """
    Find the first txBox with buChar bullets. If found:
    - Convert it to a body placeholder (idx=1), keeping its xfrm (position stays).
    - Replace every buChar/buFont with buBlip → triangle_bullet.png.
    - Add image relationship to the slide rels.
    Returns True if modified.
    """
    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return False

    # Find target txBox
    target_sp = None
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        cNvSpPr = sp.find(f".//{{{NS_P}}}cNvSpPr")
        if cNvSpPr is None or cNvSpPr.get("txBox") != "1":
            continue
        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is None:
            continue
        if any(p.find(f"{{{NS_A}}}pPr") is not None and
               p.find(f"{{{NS_A}}}pPr").find(f"{{{NS_A}}}buChar") is not None
               for p in txBody.findall(f"{{{NS_A}}}p")):
            target_sp = sp
            break

    if target_sp is None:
        return False

    # Add triangle image rel to slide rels
    rels_path = re.sub(r"ppt/slides/(slide\d+\.xml)$",
                       r"ppt/slides/_rels/\1.rels", slide_path)
    rels_bytes = file_map.get(rels_path, b"")
    rels_root_sl = etree.fromstring(rels_bytes)
    max_n = max(
        (int(m.group(1)) for r in rels_root_sl
         for m in [re.match(r"rId(\d+)", r.get("Id", ""))] if m),
        default=0
    )
    tri_rid = f"rId{max_n + 1}"
    img_rel = etree.SubElement(rels_root_sl, "Relationship")
    img_rel.set("Id", tri_rid)
    img_rel.set("Type", NS_R_IMG)
    img_rel.set("Target", "../media/triangle_bullet.png")
    file_map[rels_path] = etree.tostring(
        rels_root_sl, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Promote txBox → body ph, keep xfrm so position is unchanged
    cNvSpPr = target_sp.find(f".//{{{NS_P}}}cNvSpPr")
    cNvSpPr.attrib.pop("txBox", None)
    nvSpPr = target_sp.find(f"{{{NS_P}}}nvSpPr")
    nvPr   = nvSpPr.find(f"{{{NS_P}}}nvPr") if nvSpPr is not None else None
    if nvPr is not None and nvPr.find(f"{{{NS_P}}}ph") is None:
        ph = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
        ph.set("idx", "1")

    # Swap buChar/buFont → buBlip in every paragraph
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    BU_STRIP = {f"{{{NS_A}}}buChar", f"{{{NS_A}}}buFont",
                f"{{{NS_A}}}buClr",  f"{{{NS_A}}}buClrTx",
                f"{{{NS_A}}}buSzPct", f"{{{NS_A}}}buSzPts", f"{{{NS_A}}}buSzTx",
                f"{{{NS_A}}}buFontTx", f"{{{NS_A}}}buNone", f"{{{NS_A}}}buAutoNum"}
    txBody = target_sp.find(f"{{{NS_P}}}txBody")
    for p in txBody.findall(f"{{{NS_A}}}p"):
        pPr = p.find(f"{{{NS_A}}}pPr")
        if pPr is None:
            continue
        if pPr.find(f"{{{NS_A}}}buChar") is None:
            continue
        for tag in BU_STRIP:
            for el in pPr.findall(tag):
                pPr.remove(el)
        buBlip = etree.Element(f"{{{NS_A}}}buBlip")
        blip   = etree.SubElement(buBlip, f"{{{NS_A}}}blip")
        blip.set(f"{{{NS_R}}}embed", tri_rid)
        # buBlip must precede tabLst/defRPr/extLst per DrawingML schema order
        AFTER_TAGS = {f"{{{NS_A}}}tabLst", f"{{{NS_A}}}defRPr", f"{{{NS_A}}}extLst"}
        ins = next((i for i, c in enumerate(pPr) if c.tag in AFTER_TAGS), len(list(pPr)))
        pPr.insert(ins, buBlip)

    return True


def strip_footer_placeholders(slide_root, inject=True):
    """
    Remove dt/ftr shapes entirely.
    If inject=True: remove any existing sldNum, inject a fresh one with
    the master's exact position and styling.
    If inject=False: just remove all footer/sldNum shapes (used for cover slide).
    Returns count of shapes removed.
    """
    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return 0
    touched = 0

    for sp in list(spTree.findall(f"{{{NS_P}}}sp")):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is None:
            continue
        if ph.get("type", "") in ("dt", "ftr", "sldNum"):
            spTree.remove(sp)
            touched += 1

    if not inject:
        return touched

    # Inject a clean sldNum shape matching the master's styling exactly
    max_id = max(
        (int(el.get("id", 0)) for el in slide_root.iter(f"{{{NS_P}}}cNvPr")),
        default=100
    )
    sp = etree.SubElement(spTree, f"{{{NS_P}}}sp")
    nvSpPr  = etree.SubElement(sp, f"{{{NS_P}}}nvSpPr")
    cNvPr   = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvPr")
    cNvPr.set("id", str(max_id + 1)); cNvPr.set("name", "SlideNumber")
    cNvSpPr = etree.SubElement(nvSpPr, f"{{{NS_P}}}cNvSpPr")
    etree.SubElement(cNvSpPr, f"{{{NS_A}}}spLocks").set("noGrp", "1")
    nvPr    = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")
    ph_el   = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
    ph_el.set("type", "sldNum"); ph_el.set("idx", "14")
    spPr    = etree.SubElement(sp, f"{{{NS_P}}}spPr")
    xfrm    = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
    off     = etree.SubElement(xfrm, f"{{{NS_A}}}off")
    off.set("x", "5867400"); off.set("y", "6564276")
    ext     = etree.SubElement(xfrm, f"{{{NS_A}}}ext")
    ext.set("cx", "457200"); ext.set("cy", "128016")
    # txBody — copy master's bodyPr + lstStyle for correct font/color/size
    txBody  = etree.SubElement(sp, f"{{{NS_P}}}txBody")
    bodyPr  = etree.SubElement(txBody, f"{{{NS_A}}}bodyPr")
    bodyPr.set("vert", "horz"); bodyPr.set("wrap", "square"); bodyPr.set("anchor", "ctr")
    lstStyle = etree.SubElement(txBody, f"{{{NS_A}}}lstStyle")
    lvl1    = etree.SubElement(lstStyle, f"{{{NS_A}}}lvl1pPr")
    lvl1.set("algn", "l")
    defRPr  = etree.SubElement(lvl1, f"{{{NS_A}}}defRPr")
    defRPr.set("lang", "en-US"); defRPr.set("sz", "1000")
    solidFill = etree.SubElement(defRPr, f"{{{NS_A}}}solidFill")
    srgb    = etree.SubElement(solidFill, f"{{{NS_A}}}srgbClr")
    srgb.set("val", "767676")
    latin   = etree.SubElement(defRPr, f"{{{NS_A}}}latin")
    latin.set("typeface", "Arial"); latin.set("pitchFamily", "34"); latin.set("charset", "0")
    p_el    = etree.SubElement(txBody, f"{{{NS_A}}}p")
    fld     = etree.SubElement(p_el, f"{{{NS_A}}}fld")
    fld.set("id", "{" + str(uuid.uuid4()).upper() + "}"); fld.set("type", "slidenum")
    etree.SubElement(fld, f"{{{NS_A}}}t").text = "<#>"
    etree.SubElement(p_el, f"{{{NS_A}}}endParaRPr").set("lang", "en-US")

    return touched


# ── STEP 3: Vestige cleanup ───────────────────────────────────────────────────
_ARROW_PRSTS = {
    "leftArrow","rightArrow","upArrow","downArrow","leftRightArrow",
    "upDownArrow","bentArrow","uturnArrow","leftUpArrow","bentConnector3",
    "curvedRightArrow","curvedLeftArrow","stripedRightArrow","notchedRightArrow",
    "homePlate","rightArrowCallout","leftArrowCallout",
}

def _strip_all_decorative(spTree, keep=None):
    """Remove ALL non-placeholder sp, grpSp, and cxnSp shapes from spTree.
    Used on Cover/Ending slides — gives the FulcrumQ master full visual control.
    Placeholders and pic/graphicFrame elements are always preserved.
    `keep` is a single sp element that must not be removed (the promoted title)."""
    STRIP_TAGS = {
        f"{{{NS_P}}}sp",
        f"{{{NS_P}}}grpSp",
        f"{{{NS_P}}}cxnSp",
    }
    for child in list(spTree):
        if child.tag not in STRIP_TAGS:
            continue
        if child is keep:
            continue
        if child.find(f".//{{{NS_P}}}ph") is not None:
            continue
        spTree.remove(child)


def lift_text_into_placeholders(slide_root, layout_name):
    """
    For Cover and Divider slides: extract title (and subtitle for Cover) text
    into proper placeholders, then clear the canvas so the FulcrumQ master
    controls all visual styling.

    Cover strategy (aggressive — full canvas clear):
      1. Find best-font non-placeholder text → title placeholder
      2. Find second-best text → subTitle placeholder (if present)
      3. Strip ALL non-placeholder sp/grpSp/cxnSp (rectangles, decorative bands,
         groups) so leftover shapes can't interfere with z-order color detection.
      4. Strip all explicit rPr colors from the title shape so master colors apply.

    Divider/Ending strategy (conservative — only strip large filled blocks):
      Same steps 1-2, then _strip_large_filled_shapes only.

    Searches for text candidates inside grpSp children as well as top-level sp.
    Returns count of shapes promoted.
    """
    IS_COVER   = layout_name.startswith("Cover")
    IS_DIVIDER = "Divider" in layout_name or "Ending" in layout_name
    if not (IS_COVER or IS_DIVIDER):
        return 0

    spTree = slide_root.find(f".//{{{NS_P}}}spTree")
    if spTree is None:
        return 0

    SLIDE_W, SLIDE_H = 12_192_000, 6_858_000

    # ── 1. Check for existing title placeholder with content ──────────────────
    title_sp = None
    for sp in spTree.findall(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None and ph.get("type", "") in ("title", "ctrTitle"):
            title_sp = sp
            break

    # ── 2. Collect text candidates: top-level sp + sp inside grpSp ───────────
    # Body placeholders are included as candidates — some cover slides store
    # their title/subtitle text in body phs rather than title/ctrTitle phs.
    _SKIP_PH_TYPES = {"title", "ctrTitle", "subTitle", "sldNum", "dt", "ftr"}
    def _collect_candidates(parent):
        result = []
        for sp in parent.findall(f"{{{NS_P}}}sp"):
            ph = sp.find(f".//{{{NS_P}}}ph")
            if ph is not None and ph.get("type", "body") in _SKIP_PH_TYPES:
                continue
            text = _shape_text(sp)
            if text:
                result.append((_shape_font_size(sp), _shape_area(sp), sp))
        for grp in parent.findall(f"{{{NS_P}}}grpSp"):
            result.extend(_collect_candidates(grp))
        return result

    def _wipe_placeholder_formatting(sp):
        """Strip all explicit styling from an existing placeholder so the master's
        placeholder style owns font, size, weight, alignment, position, bullets, etc.

        Clears:
          - spPr xfrm  (let layout control position/size)
          - txBody lstStyle  (source override that sits above master in the cascade)
          - every rPr / endParaRPr / defRPr  (run-level overrides)
          - every pPr's attributes + children, leaving only <a:buNone>
        """
        # Remove xfrm so layout controls position
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is not None:
            for xfrm in spPr.findall(f"{{{NS_A}}}xfrm"):
                spPr.remove(xfrm)

        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is None:
            return

        # Remove lstStyle entirely — master placeholder owns this
        for lst in txBody.findall(f"{{{NS_A}}}lstStyle"):
            txBody.remove(lst)

        # Wipe all run-level formatting
        _KEEP_ATTRS = {"lang", "dirty", "smtClean"}
        for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
            for rPr in txBody.iter(tag):
                for child in list(rPr):
                    rPr.remove(child)
                for attr in list(rPr.attrib):
                    if attr not in _KEEP_ATTRS:
                        del rPr.attrib[attr]

        # Wipe pPr attributes (alignment, indent, etc.) and bullet formatting;
        # ensure buNone is present so no bullet is rendered
        for pPr in txBody.iter(f"{{{NS_A}}}pPr"):
            for attr in list(pPr.attrib):
                del pPr.attrib[attr]
            for child in list(pPr):
                pPr.remove(child)
            if pPr.find(f"{{{NS_A}}}buNone") is None:
                pPr.append(etree.Element(f"{{{NS_A}}}buNone"))

    if title_sp is not None and _shape_text(title_sp):
        # Title already has content — wipe source formatting then clear canvas
        _wipe_placeholder_formatting(title_sp)
        if IS_COVER:
            _strip_all_decorative(spTree, keep=title_sp)
        else:
            _strip_large_filled_shapes(spTree, SLIDE_W, SLIDE_H, keep=title_sp)
        return 0

    candidates = sorted(_collect_candidates(spTree),
                        key=lambda x: (x[0], x[1]), reverse=True)
    if not candidates:
        return 0

    _, _, best = candidates[0]
    remaining  = [sp for (_, _, sp) in candidates[1:]]
    promoted   = 0

    # ── 3. Convert best candidate → title placeholder ─────────────────────────
    def _promote(sp, ph_type, idx=None):
        nvSpPr = sp.find(f"{{{NS_P}}}nvSpPr")
        if nvSpPr is not None:
            nvPr = nvSpPr.find(f"{{{NS_P}}}nvPr")
            if nvPr is None:
                nvPr = etree.SubElement(nvSpPr, f"{{{NS_P}}}nvPr")
            for old in nvPr.findall(f"{{{NS_P}}}ph"):
                nvPr.remove(old)
            ph_el = etree.SubElement(nvPr, f"{{{NS_P}}}ph")
            ph_el.set("type", ph_type)
            if idx is not None:
                ph_el.set("idx", idx)
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is not None:
            for xfrm in spPr.findall(f"{{{NS_A}}}xfrm"):
                spPr.remove(xfrm)
            for fill_tag in (f"{{{NS_A}}}solidFill", f"{{{NS_A}}}gradFill",
                             f"{{{NS_A}}}pattFill",  f"{{{NS_A}}}blipFill"):
                for el in spPr.findall(fill_tag):
                    spPr.remove(el)
            if spPr.find(f"{{{NS_A}}}noFill") is None:
                spPr.insert(0, etree.Element(f"{{{NS_A}}}noFill"))
        BU_TAGS = {
            f"{{{NS_A}}}buClr", f"{{{NS_A}}}buClrTx", f"{{{NS_A}}}buChar",
            f"{{{NS_A}}}buAutoNum", f"{{{NS_A}}}buSzPct", f"{{{NS_A}}}buSzPts",
            f"{{{NS_A}}}buSzTx",   f"{{{NS_A}}}buFontTx", f"{{{NS_A}}}buFont",
        }
        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is not None:
            for pPr in txBody.iter(f"{{{NS_A}}}pPr"):
                for btag in BU_TAGS:
                    for el in pPr.findall(btag):
                        pPr.remove(el)
                if pPr.find(f"{{{NS_A}}}buNone") is None:
                    etree.SubElement(pPr, f"{{{NS_A}}}buNone")
            # Strip ALL explicit run/paragraph/list formatting for promoted
            # title/subTitle placeholders — master owns font, size, weight,
            # alignment, bullets.  Keep only lang/dirty/smtClean.
            if ph_type in ("title", "subTitle") or (ph_type == "body" and idx == "12"):
                # Remove lstStyle so master placeholder lstStyle wins
                for lst in txBody.findall(f"{{{NS_A}}}lstStyle"):
                    txBody.remove(lst)
                _KEEP_ATTRS = {"lang", "dirty", "smtClean"}
                for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
                    for rPr in txBody.iter(tag):
                        for child in list(rPr):
                            rPr.remove(child)
                        for attr in list(rPr.attrib):
                            if attr not in _KEEP_ATTRS:
                                del rPr.attrib[attr]
                for pPr in txBody.iter(f"{{{NS_A}}}pPr"):
                    for attr in list(pPr.attrib):
                        del pPr.attrib[attr]
                    for child in list(pPr):
                        pPr.remove(child)
                    if pPr.find(f"{{{NS_A}}}buNone") is None:
                        pPr.append(etree.Element(f"{{{NS_A}}}buNone"))

    _promote(best, "title")
    promoted += 1

    # ── 4. Cover: also extract subtitle from second candidate ─────────────────
    if IS_COVER and remaining:
        subtitle_sp = remaining[0]
        _promote(subtitle_sp, "body", idx="12")
        promoted += 1

    # ── 5. Strip decorative shapes ────────────────────────────────────────────
    if IS_COVER:
        # Full canvas clear — master owns all visual styling on covers
        _strip_all_decorative(spTree, keep=best)
    else:
        _strip_large_filled_shapes(spTree, SLIDE_W, SLIDE_H, keep=best)

    return promoted


def _strip_large_filled_shapes(spTree, slide_w, slide_h, keep=None):
    """Remove non-placeholder non-transparent shapes larger than 8% of the slide area.

    Uses the same broad fill detection as _has_colored_block_title: shapes are
    considered filled unless they have an explicit <a:noFill> as a direct child
    of spPr.  This catches shapes whose fill comes from p:style/fillRef or from
    theme inheritance (no solidFill/gradFill appears directly in spPr for those).
    """
    MIN_AREA = slide_w * slide_h * 0.08
    for sp in list(spTree.findall(f"{{{NS_P}}}sp")):
        if sp is keep:
            continue
        if sp.find(f".//{{{NS_P}}}ph") is not None:
            continue
        spPr = sp.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        # Skip explicitly transparent shapes
        if spPr.find(f"{{{NS_A}}}noFill") is not None:
            continue
        if _shape_area(sp) >= MIN_AREA:
            spTree.remove(sp)


def _is_vestige(sp, slide_w_emu, slide_h_emu):
    """
    Returns (True, reason) if shape is a vestige of the old master, else (False,'').
    Heuristics:
      1. Arrow auto-shape
      2. Zero or near-zero size (both cx and cy < 0.15")
      3. Full-slide background rect with no text
    """
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return False, ""

    # Check prstGeom
    pg   = spPr.find(f"{{{NS_A}}}prstGeom")
    prst = pg.get("prst","") if pg is not None else ""
    if prst in _ARROW_PRSTS:
        return True, f"arrow ({prst})"

    # Check size
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    ext  = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
    if ext is not None:
        cx = int(ext.get("cx", 0))
        cy = int(ext.get("cy", 0))
        if cx < int(0.15 * EMU) and cy < int(0.15 * EMU):
            texts = [t.text for t in sp.iter(f"{{{NS_A}}}t") if t.text]
            if not texts:
                return True, f"near-zero size ({cx}×{cy} EMU)"

        # Full-slide background rect with no text
        off  = xfrm.find(f"{{{NS_A}}}off") if xfrm is not None else None
        x    = int(off.get("x",0)) if off is not None else 0
        y    = int(off.get("y",0)) if off is not None else 0
        if (cx >= int(0.75 * slide_w_emu) and
                cy >= int(0.4 * slide_h_emu) and
                x <= int(0.1 * slide_w_emu)):
            texts = [t.text for t in sp.iter(f"{{{NS_A}}}t") if t.text]
            if not texts:
                return True, f"bg rect with no text ({cx//EMU:.1f}\"×{cy//EMU:.1f}\")"

    return False, ""


def clean_vestiges(slide_root, slide_w=12192000, slide_h=6858000):
    """Remove vestige shapes from a slide. Returns count removed."""
    csld   = slide_root.find(f"{{{NS_P}}}cSld")
    spTree = csld.find(f"{{{NS_P}}}spTree") if csld is not None else None
    if spTree is None:
        return 0

    removed = []
    for sp in list(spTree.findall(f"{{{NS_P}}}sp")):
        ok, reason = _is_vestige(sp, slide_w, slide_h)
        if ok:
            spTree.remove(sp)
            removed.append(reason)

    return removed


# ── STEP 4: Brand style pass ──────────────────────────────────────────────────
def _stamp(lat_el, typeface):
    lat_el.set("typeface", typeface)
    if typeface in FONT_FALLBACKS:
        pf, cs = FONT_FALLBACKS[typeface]
        lat_el.set("pitchFamily", str(pf))
        lat_el.set("charset", str(cs))
        lat_el.attrib.pop("panose", None)


def style_rpr(rPr):
    lat = rPr.find(f"{{{NS_A}}}latin")
    if lat is None:
        # Bold runs with no explicit font → make Aileron explicit for consistency
        if rPr.get("b") == "1":
            lat = etree.SubElement(rPr, f"{{{NS_A}}}latin")
            _stamp(lat, "Segoe UI")
            return True
        return False
    tf = lat.get("typeface","")
    if tf.startswith("+"):
        return False
    is_bold = rPr.get("b") == "1"
    if tf in BOLD_UPGRADE and is_bold:
        _stamp(lat, "Segoe UI")
        return True
    if tf in FONT_MAP:
        brand, force_b = FONT_MAP[tf]
        _stamp(lat, brand)
        if force_b is True:   rPr.set("b","1")
        elif force_b is False: rPr.attrib.pop("b",None)
        return True
    if tf in ("Segoe UI","Arial"):
        _stamp(lat, tf)
    return False


TITLE_PH_TYPES = {"title", "ctrTitle", "subTitle"}

# Vertical threshold for position-based title detection on non-placeholder shapes.
# Anything whose top edge sits above 1.5 in (content_top_in from style guide) is
# treated as a heading and gets Segoe UI Bold.  1 in = 914 400 EMU.
_TITLE_REGION_EMU = int(1.5 * 914_400)  # 1 371 600

_TITLE_SZ_CAP  = 3600   # 36pt — cap explicit title font sizes to prevent oversized headings
_SUBTITLE_SZ_CAP = 2400  # 24pt — cap subtitle sizes

def _style_rpr_as(rPr, is_title, is_subtitle=False):
    """Apply font + character spacing to one rPr element.

    is_subtitle — True for subTitle ph runs: Segoe UI, NOT bold, lighter spacing.
    is_title    — True for title/ctrTitle ph and title-region non-ph shapes.

    Title and subtitle placeholder runs also have their explicit sz capped so
    slides with oversized source fonts render consistently.
    """
    lat = rPr.find(f"{{{NS_A}}}latin")

    if is_subtitle:
        if lat is None:
            lat = etree.SubElement(rPr, f"{{{NS_A}}}latin")
        if not lat.get("typeface", "").startswith("+"):
            _stamp(lat, "Segoe UI")
            rPr.set("b", "0")
            rPr.set("spc", str(BODY_CHAR_SPC))
            # Cap explicit sz
            try:
                if int(rPr.get("sz", "0")) > _SUBTITLE_SZ_CAP:
                    rPr.set("sz", str(_SUBTITLE_SZ_CAP))
            except ValueError:
                pass
            return True
        return False

    if is_title:
        if lat is None:
            lat = etree.SubElement(rPr, f"{{{NS_A}}}latin")
        if not lat.get("typeface", "").startswith("+"):
            _stamp(lat, "Segoe UI")
            rPr.set("b", "1")
            rPr.set("spc", str(TITLE_CHAR_SPC))
            # Cap explicit sz so oversized source headings don't blow out
            try:
                if int(rPr.get("sz", "0")) > _TITLE_SZ_CAP:
                    rPr.set("sz", str(_TITLE_SZ_CAP))
            except ValueError:
                pass
            return True
        return False

    if lat is None or lat.get("typeface", "").startswith("+"):
        return False
    changed = style_rpr(rPr)
    if changed:
        rPr.set("spc", str(BODY_CHAR_SPC))
    return changed


PIC_RECOLOR = "765FFF"   # Pivot Purple — dark anchor for duotone on embedded graphics

# Images whose area exceeds this fraction of the slide are treated as screenshots
# (UI captures, spreadsheet grabs, etc.) and are NOT duotone-ified — they contain
# structured content that must remain legible through the overlaid redaction shapes.
_SCREENSHOT_AREA_FRAC = 0.20

# Internal attribute used to communicate screenshot identity across passes.
_FQ_SCREENSHOT_ATTR = "_fq_screenshot"


def _pic_bbox(pic):
    """Return (x, y, cx, cy) EMU tuple for a <p:pic>, or None."""
    spPr = pic.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return None
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if off is None or ext is None:
        return None
    try:
        return (int(off.get("x", 0)), int(off.get("y", 0)),
                int(ext.get("cx", 0)), int(ext.get("cy", 0)))
    except ValueError:
        return None


def recolor_pics(slide_root, slide_w=9144000, slide_h=5143500):
    """Apply duotone recolor to embedded graphics.

    Two tiers:
      · Small images (< 20% of slide area) — logos, icons, brand graphics.
        Gets purple/white duotone to align with FulcrumQ brand palette.
      · Large images (≥ 20% of slide area) — screenshots of dashboards, tools,
        spreadsheets, org charts.  These are left untouched so their content
        remains legible.  They are tagged with _fq_screenshot so downstream
        passes can detect overlay shapes sitting on top of them.

    Returns count of images duotone-ified (excludes screenshots).
    """
    slide_area = slide_w * slide_h
    count = 0
    for pic in slide_root.iter(f"{{{NS_P}}}pic"):
        blipFill = pic.find(f"{{{NS_P}}}blipFill")
        if blipFill is None:
            continue
        blip = blipFill.find(f"{{{NS_A}}}blip")
        if blip is None:
            continue

        # Classify by area: large = screenshot, small = logo/icon
        bbox = _pic_bbox(pic)
        if bbox is not None:
            pic_area = bbox[2] * bbox[3]
            if pic_area >= _SCREENSHOT_AREA_FRAC * slide_area:
                pic.set(_FQ_SCREENSHOT_ATTR, "1")
                continue   # preserve screenshot as-is

        # Remove SVG extension so PowerPoint uses the PNG fallback (SVGs ignore duotone)
        extLst = blip.find(f"{{{NS_A}}}extLst")
        if extLst is not None:
            blip.remove(extLst)
        # Remove any existing recolor effects
        for existing in blip.findall(f"{{{NS_A}}}duotone"):
            blip.remove(existing)
        # DrawingML duotone: dark pixels → brand purple, light pixels → white
        duotone = etree.SubElement(blip, f"{{{NS_A}}}duotone")
        dark_end = etree.SubElement(duotone, f"{{{NS_A}}}srgbClr")
        dark_end.set("val", PIC_RECOLOR)
        light_end = etree.SubElement(duotone, f"{{{NS_A}}}srgbClr")
        light_end.set("val", "FFFFFF")
        count += 1
    return count


def neutralize_screenshot_overlays(slide_root) -> int:
    """For every screenshot image on the slide, find the shapes that appear
    on top of it in z-order and reset their fill to the slide's background
    color (F2F2F2 for light layouts).

    These are redaction/annotation overlays — white or light rectangles placed
    by the author to hide sensitive values.  The normal color pass remaps their
    fills to brand colors which can break their camouflage against the slide
    background.  This pass restores that camouflage.

    Only shapes with a solidFill that is white or near-white (perceived
    luminance > 200) are treated; accent-coloured overlays (callout arrows,
    annotation labels) are left alone.

    Returns count of overlay shapes neutralized.
    """
    spTree = slide_root.find(f"{{{NS_P}}}cSld/{{{NS_P}}}spTree")
    if spTree is None:
        return 0

    # Walk spTree in z-order, collecting screenshot bboxes as we encounter them.
    # When we hit a shape later in z-order, check whether it overlaps any earlier screenshot.
    slide_bg = _slide_bg_hex(slide_root) or "F2F2F2"
    count = 0
    screenshot_bboxes = []   # bboxes of screenshots seen so far (lower z-order)

    for child in spTree:
        tag = child.tag.split("}")[-1]

        # Track screenshots as we pass them in z-order
        if tag == "pic" and child.get(_FQ_SCREENSHOT_ATTR):
            bb = _pic_bbox(child)
            if bb:
                screenshot_bboxes.append(bb)
            continue

        if tag != "sp" or not screenshot_bboxes:
            continue

        # Shape must have a near-white solidFill (luminance > 200)
        spPr = child.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            continue
        sf = spPr.find(f"{{{NS_A}}}solidFill")
        if sf is None:
            continue
        srgb = sf.find(f"{{{NS_A}}}srgbClr")
        if srgb is None:
            continue
        fill_hex = srgb.get("val", "").upper()
        try:
            r, g, b = int(fill_hex[0:2], 16), int(fill_hex[2:4], 16), int(fill_hex[4:6], 16)
            lum = 0.299 * r + 0.587 * g + 0.114 * b
        except Exception:
            continue
        if lum < 200:
            continue   # accent or dark fill — intentional, leave as-is

        # Check spatial overlap with any screenshot below this shape
        sp_bb = _get_shape_bbox(child)
        if sp_bb is None:
            continue
        sx, sy, scx, scy = sp_bb
        for px, py, pcx, pcy in screenshot_bboxes:
            if sx < px + pcx and sx + scx > px and sy < py + pcy and sy + scy > py:
                srgb.set("val", slide_bg)
                count += 1
                break

    # Clean up screenshot markers
    for pic in slide_root.iter(f"{{{NS_P}}}pic"):
        pic.attrib.pop(_FQ_SCREENSHOT_ATTR, None)

    return count


# schemeClr fills to convert: all standard OOXML scheme slots → vF brand colors.
# This prevents source-deck tint/shade transforms on scheme colors from producing
# non-brand colors like E5FFFF (light tint of an old accent4 teal).
# Note: tint/shade child elements are stripped along with the schemeClr so the
# bare brand color takes over — approved palette colors are used directly.
SCHEME_FILL_MAP = {
    # All standard OOXML scheme slots + presentation aliases (bg1=lt1, tx1=dk1, etc.)
    "dk1":     "1D1D1D",   "tx1":     "1D1D1D",   # Guiding Grey
    "dk2":     "281A42",   "tx2":     "281A42",   # Vector Dark Purple
    "lt1":     "F2F2F2",   "bg1":     "F2F2F2",   # → Lightest Grey (FFFFFF would vanish on white bg)
    "lt2":     "F2F2F2",   "bg2":     "F2F2F2",   # Lightest Grey
    "accent1": "765FFF",   # Pivot Purple
    "accent2": "281A42",   # Vector Dark Purple
    "accent3": "00C27A",   # Anchor Green
    "accent4": "60BDBC",   # Talent Teal
    "accent5": "FFB547",   # Momentum Amber
    "accent6": "FF2E88",   # Signal Magenta
    "hlink":   "765FFF",   # Pivot Purple
    "folHlink":"917FFF",   # Tint 1
}

def convert_scheme_fills(slide_root):
    """Replace ALL schemeClr color specs in solidFills and gradient stops with
    explicit srgbClr brand colors, so no color depends on theme resolution at
    render time. Tint/shade/lumMod child transforms are stripped with the schemeClr.

    Note: schemeClr inside p:style (lnRef/fillRef/effectRef/fontRef) are shape
    style inheritance refs — left untouched since they resolve via the patched theme.
    """
    count = 0
    STYLE_TAGS = {
        f"{{{NS_A}}}lnRef", f"{{{NS_A}}}fillRef",
        f"{{{NS_A}}}effectRef", f"{{{NS_A}}}fontRef",
    }

    # solidFill > schemeClr  (covers shape fills, line fills, text run fills, etc.)
    for solidFill in slide_root.iter(f"{{{NS_A}}}solidFill"):
        schemeClr = solidFill.find(f"{{{NS_A}}}schemeClr")
        if schemeClr is None:
            continue
        parent = solidFill.getparent()
        if parent is not None and parent.tag in STYLE_TAGS:
            continue   # shape-style ref — leave for theme resolution
        val   = schemeClr.get("val", "")
        brand = SCHEME_FILL_MAP.get(val, "1D1D1D")
        # For shape background fills (spPr context), honour any luminance-modifying
        # transforms so that lightly-tinted scheme slots stay light rather than
        # collapsing to the dark base brand color (e.g. dk2+tint(20%) = near-white card).
        transforms = list(schemeClr)
        is_shape_fill = (parent is not None and
                         parent.tag.split("}")[-1] == "spPr" and
                         bool(transforms))
        final = _apply_scheme_transforms(brand, transforms) if is_shape_fill else brand
        solidFill.remove(schemeClr)
        srgb = etree.SubElement(solidFill, f"{{{NS_A}}}srgbClr")
        srgb.set("val", final)
        count += 1

    # gradient stop > schemeClr  (gsLst > gs > schemeClr, no solidFill wrapper)
    for gs in slide_root.iter(f"{{{NS_A}}}gs"):
        schemeClr = gs.find(f"{{{NS_A}}}schemeClr")
        if schemeClr is None:
            continue
        val   = schemeClr.get("val", "")
        brand = SCHEME_FILL_MAP.get(val, "1D1D1D")
        gs.remove(schemeClr)
        srgb = etree.SubElement(gs, f"{{{NS_A}}}srgbClr")
        srgb.set("val", brand)
        count += 1

    # pattFill > fgClr/bgClr > schemeClr  (pattern fills — not wrapped in solidFill)
    for tag in (f"{{{NS_A}}}fgClr", f"{{{NS_A}}}bgClr"):
        for clr_el in slide_root.iter(tag):
            schemeClr = clr_el.find(f"{{{NS_A}}}schemeClr")
            if schemeClr is None:
                continue
            val   = schemeClr.get("val", "")
            brand = SCHEME_FILL_MAP.get(val, "1D1D1D")
            clr_el.remove(schemeClr)
            srgb = etree.SubElement(clr_el, f"{{{NS_A}}}srgbClr")
            srgb.set("val", brand)
            count += 1

    return count


# Brand accent colors in priority order for collision resolution within a shape.
# Full approved palette: primaries first, then tints, then neutrals.
BRAND_ACCENTS = [
    "765FFF",  # Pivot Purple
    "00C27A",  # Anchor Green
    "FF2E88",  # Signal Magenta
    "FFB547",  # Amber
    "60BDBC",  # Teal
    "917FFF",  # Pivot Purple Tint 1
    "AD9FFF",  # Pivot Purple Tint 2
    "C8BFFF",  # Pivot Purple Tint 3
    "E9E4FF",  # Shift Lavender
    "281A42",  # Vector Dark Purple
    "1D1D1D",  # Guiding Grey
]

# For shapes with mixed explicit+inherited runs (call-out: stat + body text).
CALLOUT_ACCENTS = [
    "765FFF",  # Pivot Purple   ← callout stat / emphasis
    "00C27A",  # Anchor Green
    "FF2E88",  # Signal Magenta
    "FFB547",  # Amber
    "60BDBC",  # Teal
    "917FFF",  # Pivot Purple Tint 1
    "AD9FFF",  # Pivot Purple Tint 2
    "281A42",  # Vector Dark Purple
    "1D1D1D",  # Guiding Grey
]

# For table cells: visually distinct colors so categorical rows stay readable.
TABLE_ACCENTS = [
    "00C27A",  # Anchor Green       ← first categorical color
    "FF2E88",  # Signal Magenta
    "FFB547",  # Amber
    "60BDBC",  # Teal
    "765FFF",  # Pivot Purple
    "917FFF",  # Pivot Purple Tint 1
    "281A42",  # Vector Dark Purple
    "1D1D1D",  # Guiding Grey
]

# Colors that are "neutral" in brand context — don't participate in accent assignment
NEUTRAL_BRAND = {"1D1D1D","3B3B3B","585858","767676","7A828D","B2BBCA","D0D7DF","E9E4FF","FFFFFF"}

def _remap_text_color(val, local_color_map, rag_preserve):
    """Per-shape text color remap with collision avoidance using BRAND_ACCENTS.

    Every non-RAG text color is snapped to a brand equivalent — same guarantee
    as shape fills.  Unknown colors (not in COLOR_REMAP) fall through to
    _nearest_brand_color so no non-brand color can survive unchanged."""
    v = val.upper()
    if v in rag_preserve:
        return None
    if v in local_color_map:
        return local_color_map[v]
    if v in COLOR_REMAP:
        candidate = COLOR_REMAP[v]
    elif _is_red_hue(v):
        candidate = RED_REMAP_TARGET
    else:
        # No explicit mapping — snap to nearest brand color perceptually
        try:
            candidate = _nearest_brand_color(v)
        except Exception:
            return None
    if candidate in NEUTRAL_BRAND:
        local_color_map[v] = candidate
        return candidate
    used = set(local_color_map.values()) - NEUTRAL_BRAND
    if candidate in used:
        for alt in BRAND_ACCENTS:
            if alt not in used and alt not in NEUTRAL_BRAND:
                candidate = alt
                break
    local_color_map[v] = candidate
    return candidate


def _remap_accent_sequence(val, color_map, rag_preserve, accent_list):
    """Remap text color by assigning accent colors from accent_list in encounter order.
    Neutral colors (dark greys, white) pass through COLOR_REMAP directly without
    consuming an accent slot.  Unknown colors fall through to _nearest_brand_color."""
    v = val.upper()
    if v in rag_preserve:
        return None
    if v in color_map:
        return color_map[v]
    if v in COLOR_REMAP:
        candidate = COLOR_REMAP[v]
    elif _is_red_hue(v):
        candidate = RED_REMAP_TARGET
    else:
        try:
            candidate = _nearest_brand_color(v)
        except Exception:
            return None
    if candidate in NEUTRAL_BRAND:
        color_map[v] = candidate
        return candidate
    # Non-neutral: assign next available slot from accent_list in order
    used = [c for c in color_map.values() if c not in NEUTRAL_BRAND]
    for accent in accent_list:
        if accent not in NEUTRAL_BRAND and accent not in used:
            color_map[v] = accent
            return accent
    color_map[v] = candidate
    return candidate


def _sentence_case_title(sp, is_subtitle: bool = False):
    """Apply CMOS title case (titles) or sentence case (subtitles) to a placeholder shape.

    - Titles  → _to_title_case  (Chicago Manual of Style)
    - Subtitles → _to_sentence_case (first-word cap only)

    All-caps source text is normalised first so acronyms inside mixed phrases
    are handled by the respective case function.

    Since case changes don't alter string length, runs are redistributed by
    character position.  Returns True if any text was changed."""
    txBody = sp.find(f"{{{NS_P}}}txBody")
    if txBody is None:
        return False

    case_fn = _to_sentence_case if is_subtitle else _to_title_case

    changed = False
    for p in txBody.findall(f"{{{NS_A}}}p"):
        runs = [r.find(f"{{{NS_A}}}t") for r in p.findall(f"{{{NS_A}}}r")]
        runs = [t for t in runs if t is not None]
        if not runs:
            continue
        full = "".join(t.text or "" for t in runs)
        if not full.strip():
            continue

        new_full = case_fn(full)

        if new_full == full:
            continue

        # Redistribute back to runs by character position (lengths unchanged)
        pos = 0
        for t in runs:
            old_len = len(t.text or "")
            t.text = new_full[pos:pos + old_len]
            pos += old_len
        changed = True

    return changed


def _style_soft_return_subtitle(sp) -> int:
    """In a title placeholder, detect soft returns (<a:br>) and apply subtitle
    styling to all runs AFTER the first break in each paragraph.

    Subtitle style: Segoe UI, no bold, mid-grey color (7A828D).
    This lets presenters hit Shift+Enter in a title placeholder to create a
    visually distinct sub-heading line without needing a separate placeholder.
    Returns the number of runs restyled."""
    SUBTITLE_COLOR = "7A828D"
    txBody = sp.find(f"{{{NS_P}}}txBody")
    if txBody is None:
        return 0
    count = 0
    for p in txBody.findall(f"{{{NS_A}}}p"):
        children = list(p)
        br_idx = None
        for ci, child in enumerate(children):
            if child.tag == f"{{{NS_A}}}br":
                br_idx = ci
                break
        if br_idx is None:
            continue
        # Style every <a:r> that appears after the break
        for child in children[br_idx + 1:]:
            if child.tag != f"{{{NS_A}}}r":
                continue
            rPr = child.find(f"{{{NS_A}}}rPr")
            if rPr is None:
                rPr = etree.Element(f"{{{NS_A}}}rPr")
                child.insert(0, rPr)
            # Remove bold
            rPr.attrib.pop("b", None)
            rPr.set("b", "0")
            # Set font to Segoe UI (non-bold variant)
            lat = rPr.find(f"{{{NS_A}}}latin")
            if lat is None:
                lat = etree.SubElement(rPr, f"{{{NS_A}}}latin")
            _stamp(lat, "Segoe UI")
            # Set subtitle color (schema-correct insertion)
            _set_rpr_color(rPr, SUBTITLE_COLOR)
            count += 1
    return count


def _is_text_run_color(srgb):
    """True if this srgbClr is inside a text run's rPr solidFill (not a shape fill)."""
    node = srgb.getparent()  # solidFill
    if node is None: return False
    node = node.getparent()  # rPr / endParaRPr / defRPr / tcPr / spPr …
    if node is None: return False
    return node.tag.split("}")[-1] in ("rPr", "endParaRPr", "defRPr")


# Dark hex applied to table headers and dark-filled shapes instead of near-black
HEADER_DARK = "281A42"


def _hex_is_dark(v):
    """True if a 6-char hex color has perceived luminance < 25% (i.e. visually dark)."""
    try:
        r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) < 64   # < ~25% of 255
    except Exception:
        return False


def _is_shape_bg_fill(srgb):
    """True if this srgbClr is a direct shape background fill: spPr > solidFill > srgbClr.
    Excludes: line fills (ln > solidFill), gradient stops, text runs, table cells."""
    solidFill = srgb.getparent()
    if solidFill is None or solidFill.tag.split("}")[-1] != "solidFill":
        return False
    parent = solidFill.getparent()
    if parent is None:
        return False
    return parent.tag.split("}")[-1] == "spPr"


def _is_table_cell_fill(srgb):
    """True if this srgbClr is a table-cell background fill: tcPr > solidFill > srgbClr."""
    solidFill = srgb.getparent()
    if solidFill is None or solidFill.tag.split("}")[-1] != "solidFill":
        return False
    parent = solidFill.getparent()
    if parent is None:
        return False
    return parent.tag.split("}")[-1] == "tcPr"


def _in_table_header_row(element):
    """True if element is inside the first <a:tr> of a table."""
    node = element
    first_tr = None
    tbl = None
    while node is not None:
        tag = node.tag.split("}")[-1]
        if tag == "tr" and first_tr is None:
            first_tr = node
        elif tag == "tbl":
            tbl = node
            break
        node = node.getparent()
    if tbl is None or first_tr is None:
        return False
    rows = tbl.findall(f"{{{NS_A}}}tr")
    return bool(rows) and rows[0] is first_tr


def _is_mixed_callout(sp):
    """True if the shape has BOTH explicit-color runs AND inherited-color runs.
    This pattern (e.g. colored stats + plain body text) needs a vibrant accent."""
    has_explicit = has_inherit = False
    for r in sp.iter(f"{{{NS_A}}}r"):
        rPr = r.find(f"{{{NS_A}}}rPr")
        if rPr is None:
            has_inherit = True
        else:
            sf = rPr.find(f"{{{NS_A}}}solidFill")
            if sf is not None and sf.find(f"{{{NS_A}}}srgbClr") is not None:
                has_explicit = True
            else:
                has_inherit = True
        if has_explicit and has_inherit:
            return True
    return False



# ── Contrast-aware text color enforcement ─────────────────────────────────────
# After shape fills + text remaps, ensure every run is readable against its bg.
# Accent / emphasis colors are never overridden.

_BG_DARK_THRESHOLD = 128   # perceived luminance 0–255; below this = "dark"

# These text colors represent intentional emphasis — never touched by contrast pass
_FONT_ACCENT_PRESERVE = frozenset({
    "765FFF", "917FFF", "AD9FFF", "C8BFFF",   # Pivot Purple family
    "00C27A",                                   # Anchor Green
    "FF2E88",                                   # Signal Magenta
    "FFB547",                                   # Momentum Amber
    "60BDBC",                                   # Talent Teal
})


def _bg_is_dark(v: str) -> bool:
    """True if perceived luminance < 50% (~128/255)."""
    try:
        r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) < _BG_DARK_THRESHOLD
    except Exception:
        return False


def _shape_fill_hex(sp) -> str | None:
    """Effective solid fill of a shape post-remap, or None if transparent.
    For gradients, returns the darkest stop (worst-case for text readability)."""
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return None
    solidFill = spPr.find(f"{{{NS_A}}}solidFill")
    if solidFill is not None:
        srgb = solidFill.find(f"{{{NS_A}}}srgbClr")
        if srgb is not None:
            return srgb.get("val", "").upper() or None
    gradFill = spPr.find(f"{{{NS_A}}}gradFill")
    if gradFill is not None:
        hexes = [
            s.get("val", "").upper()
            for gs in gradFill.iter(f"{{{NS_A}}}gs")
            for s in [gs.find(f"{{{NS_A}}}srgbClr")]
            if s is not None and s.get("val")
        ]
        if hexes:
            return min(hexes, key=lambda h: (
                0.299 * int(h[0:2], 16) + 0.587 * int(h[2:4], 16) + 0.114 * int(h[4:6], 16)
            ))
    return None   # noFill / blipFill / pattFill / unspecified → transparent


def _slide_bg_hex(slide_root) -> str:
    """Slide's explicit background solid fill hex, defaulting to FFFFFF."""
    bg = slide_root.find(f"{{{NS_P}}}cSld/{{{NS_P}}}bg")
    if bg is not None:
        bgPr = bg.find(f"{{{NS_P}}}bgPr")
        if bgPr is not None:
            sf = bgPr.find(f"{{{NS_A}}}solidFill")
            if sf is not None:
                srgb = sf.find(f"{{{NS_A}}}srgbClr")
                if srgb is not None:
                    return srgb.get("val", "FFFFFF").upper()
    return "FFFFFF"


def _set_rpr_color(rPr, hex_val: str):
    """Set (or replace) the solidFill > srgbClr color on an rPr/endParaRPr/defRPr element.

    CT_TextCharacterProperties schema order (OOXML §20.1.9.22):
      a:ln  →  fill (solidFill|gradFill|noFill|…)  →  a:effectLst  →  a:highlight
      →  a:uLn*  →  a:latin  →  a:ea  →  a:cs  →  a:sym  →  …

    We must INSERT solidFill before a:latin/a:ea/a:cs, not append to the end.
    """
    # Remove any existing fill element
    for fill_tag in (f"{{{NS_A}}}solidFill", f"{{{NS_A}}}gradFill",
                     f"{{{NS_A}}}noFill",    f"{{{NS_A}}}blipFill",
                     f"{{{NS_A}}}pattFill",  f"{{{NS_A}}}grpFill"):
        old = rPr.find(fill_tag)
        if old is not None:
            rPr.remove(old)

    sf = etree.Element(f"{{{NS_A}}}solidFill")
    sr = etree.SubElement(sf, f"{{{NS_A}}}srgbClr")
    sr.set("val", hex_val)

    # Insert at the correct schema position:
    # after a:ln (if present), but before everything else (latin, ea, cs, etc.)
    _BEFORE_FILL = {
        f"{{{NS_A}}}latin", f"{{{NS_A}}}ea", f"{{{NS_A}}}cs", f"{{{NS_A}}}sym",
        f"{{{NS_A}}}hlinkClick", f"{{{NS_A}}}hlinkMouseOver", f"{{{NS_A}}}rtl",
        f"{{{NS_A}}}extLst", f"{{{NS_A}}}uLn", f"{{{NS_A}}}uFill",
        f"{{{NS_A}}}uLnTx", f"{{{NS_A}}}uFillTx", f"{{{NS_A}}}highlight",
        f"{{{NS_A}}}effectLst", f"{{{NS_A}}}effectDag",
    }
    insert_idx = len(list(rPr))   # default: end
    for i, child in enumerate(rPr):
        if child.tag in _BEFORE_FILL:
            insert_idx = i
            break
    rPr.insert(insert_idx, sf)


def _get_shape_bbox(sp):
    """Return (x, y, cx, cy) EMU tuple for a shape, or None."""
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return None
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if off is None or ext is None:
        return None
    return (int(off.get("x", 0)), int(off.get("y", 0)),
            int(ext.get("cx", 0)), int(ext.get("cy", 0)))


def _overlapping_fill_below(sp, slide_root) -> str | None:
    """For a noFill shape, find the topmost (highest z-order) filled shape below it
    in the spTree that spatially overlaps.  Used to detect text overlaid on a colored
    background shape (e.g. a text box on a Pivot Purple rectangle).
    Returns the fill hex, or None if no overlapping filled shape is found."""
    bbox = _get_shape_bbox(sp)
    if bbox is None:
        return None
    spTree = slide_root.find(f"{{{NS_P}}}cSld/{{{NS_P}}}spTree")
    if spTree is None:
        return None
    ax, ay, acx, acy = bbox
    best = None
    for sibling in spTree:
        if sibling is sp:
            break   # stop at our own shape (only look at lower z-order)
        if sibling.tag.split("}")[-1] != "sp":
            continue
        fill = _shape_fill_hex(sibling)
        if fill is None:
            continue
        sbb = _get_shape_bbox(sibling)
        if sbb is None:
            continue
        bx, by, bcx, bcy = sbb
        if ax < bx + bcx and ax + acx > bx and ay < by + bcy and ay + acy > by:
            best = fill   # last overlapping fill wins (topmost below us)
    return best


def enforce_text_contrast(slide_root, rag_preserve: set) -> int:
    """Deterministic font-color pass: white text on dark fills, dark text on light fills.

    Shapes with explicit fills:
      · Dark fill  → FFFFFF (white text); inherited runs forced white too
      · Light fill → flip explicit light-on-light to 1D1D1D; leave inherited runs alone

    Shapes with noFill (transparent):
      · Z-order detection: if a dark filled shape overlaps below, force text white
      · Otherwise: leave text colors alone — white text may be intentional on a dark layout

    Accent + RAG colors in text are never overridden regardless of background.
    """
    n = 0

    def _enforce_runs(sp_or_tc, dark: bool):
        nonlocal n
        for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
            for rPr in sp_or_tc.iter(tag):
                sf = rPr.find(f"{{{NS_A}}}solidFill")
                if sf is not None:
                    srgb     = sf.find(f"{{{NS_A}}}srgbClr")
                    existing = srgb.get("val", "").upper() if srgb is not None else ""
                    if existing in rag_preserve or existing in _FONT_ACCENT_PRESERVE:
                        continue
                    want = "FFFFFF" if dark else "1D1D1D"
                    if (dark and _bg_is_dark(existing)) or \
                       (not dark and not _bg_is_dark(existing)):
                        _set_rpr_color(rPr, want)
                        n += 1
                elif dark:
                    # No explicit color → inherits from master (typically dark) → invisible
                    _set_rpr_color(rPr, "FFFFFF")
                    n += 1

    # ── Shape text ────────────────────────────────────────────────────────────
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        fill = _shape_fill_hex(sp)

        if fill is not None:
            # Explicit fill: full bidirectional contrast enforcement
            _enforce_runs(sp, _bg_is_dark(fill))
        else:
            # NoFill: only act when a dark shape is detected underneath via z-order.
            # Do NOT flip explicit text colors otherwise — text may be white-on-dark-layout.
            under = _overlapping_fill_below(sp, slide_root)
            if under is not None and _bg_is_dark(under):
                # Force white; but only flip dark→white (never light→dark here,
                # since we can't be sure the whole shape is over the dark area)
                for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
                    for rPr in sp.iter(tag):
                        sf = rPr.find(f"{{{NS_A}}}solidFill")
                        if sf is not None:
                            srgb     = sf.find(f"{{{NS_A}}}srgbClr")
                            existing = srgb.get("val", "").upper() if srgb is not None else ""
                            if existing in rag_preserve or existing in _FONT_ACCENT_PRESERVE:
                                continue
                            if _bg_is_dark(existing):   # dark text on detected dark bg
                                _set_rpr_color(rPr, "FFFFFF")
                                n += 1
                        else:
                            _set_rpr_color(rPr, "FFFFFF")
                            n += 1

    # ── Table cells (only when cell has explicit dark fill) ───────────────────
    slide_bg = _slide_bg_hex(slide_root)
    for tbl in slide_root.iter(f"{{{NS_A}}}tbl"):
        for tc in tbl.iter(f"{{{NS_A}}}tc"):
            if tc.get("_fq_circle"):
                continue
            tcPr      = tc.find(f"{{{NS_A}}}tcPr")
            cell_fill = None
            if tcPr is not None:
                sf = tcPr.find(f"{{{NS_A}}}solidFill")
                if sf is not None:
                    srgb = sf.find(f"{{{NS_A}}}srgbClr")
                    if srgb is not None:
                        cell_fill = srgb.get("val", "").upper()
            if cell_fill is None:
                continue   # no explicit cell fill — leave inherited colors alone
            _enforce_runs(tc, _bg_is_dark(cell_fill))

    return n


def style_slide(slide_root, slide_idx=None, layout_name=""):
    n_fonts = 0
    n_titles_cased = 0

    # On Cover / Divider / Ending layouts the master owns all placeholder
    # typography (font, size, weight, alignment).  We already wiped explicit
    # rPr from those placeholders in lift_text_into_placeholders; re-stamping
    # fonts here would undo that.  Skip the font pass for phs on these layouts.
    _MASTER_OWNS_PH = (
        layout_name.startswith("Cover") or
        layout_name.startswith("Divider") or
        layout_name.startswith("Ending")
    )

    # Placeholder shapes — titles get Segoe UI Bold + sentence case, body gets style_rpr.
    # Non-placeholder shapes (free text boxes, callouts, etc.) use vertical position:
    # top < 1.5 in → title region → Segoe UI Bold; otherwise → Arial via FONT_MAP.
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        ph = sp.find(f".//{{{NS_P}}}ph")
        if ph is not None:
            ph_type    = ph.get("type", "")
            is_subtitle = (ph_type == "subTitle" or
                           (ph_type == "body" and ph.get("idx") == "12"))
            is_title    = ph_type in TITLE_PH_TYPES and not is_subtitle
        else:
            is_subtitle = False
            xfrm = sp.find(f".//{{{NS_A}}}xfrm")
            off  = xfrm.find(f"{{{NS_A}}}off") if xfrm is not None else None
            y_emu = int(off.get("y", "9999999999")) if off is not None else 9999999999
            is_title = y_emu < _TITLE_REGION_EMU
        # Font stamp: skip for Cover/Divider/Ending phs — master owns typography.
        # Casing still applies everywhere.
        if not (ph is not None and _MASTER_OWNS_PH):
            for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
                for rPr in sp.iter(tag):
                    if _style_rpr_as(rPr, is_title, is_subtitle=is_subtitle): n_fonts += 1
        if is_title or is_subtitle:
            if _sentence_case_title(sp, is_subtitle=is_subtitle):
                n_titles_cased += 1
            if ph is not None and is_title:
                # Soft-return subtitle: runs after <a:br> get subtitle styling
                _style_soft_return_subtitle(sp)

    # Table cells — always Arial (header row bold is preserved, font is Arial)
    for tc in slide_root.iter(f"{{{NS_A}}}tc"):
        for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
            for rPr in tc.iter(tag):
                lat = rPr.find(f"{{{NS_A}}}latin")
                if lat is None or lat.get("typeface", "").startswith("+"):
                    continue
                _stamp(lat, "Arial")
                n_fonts += 1

    # Colors
    n_colors = 0

    # Mark circle cells (● bullets = RAG rating indicators — preserve their colors)
    CIRCLE_MARKER = "_fq_circle"
    for tc in slide_root.iter(f"{{{NS_A}}}tc"):
        if any("●" in (t.text or "") for t in tc.iter(f"{{{NS_A}}}t")):
            tc.set(CIRCLE_MARKER, "1")

    # RAG detection only fires for non-table shapes
    rag_preserve = detect_rag_colors(slide_root)

    # ── Shape fill colors ─────────────────────────────────────────────────────
    # Shape backgrounds (spPr > solidFill) use per-slide accent sequencing so
    # multiple differently-coloured shapes on the same slide get distinct brand
    # accents rather than all collapsing to the same dark/neutral tone.
    # Outlines, gradients, and other non-bg fills use direct COLOR_REMAP.

    for srgb in slide_root.iter(f"{{{NS_A}}}srgbClr"):
        if _is_text_run_color(srgb):
            continue  # handled per-shape below
        # Skip circle-marker table cells
        node = srgb.getparent()
        skip = False
        while node is not None:
            if node.tag == f"{{{NS_A}}}tc":
                if node.get(CIRCLE_MARKER):
                    skip = True
                break
            node = node.getparent()
        if skip:
            continue
        v = srgb.get("val", "").upper()
        if v in rag_preserve:
            continue
        mapped = _resolve_color(v, rag_preserve)
        if mapped is None:
            continue
        new_val = mapped
        if new_val is not None:
            srgb.set("val", new_val)
            n_colors += 1

    # ── Text run colors
    # Mixed callout shapes (some runs explicit-colored, some inherited) use
    # CALLOUT_ACCENTS (Vector Blue first) so the accent stays vibrant and
    # distinct from the inherited body text color.
    # Pure shapes use standard per-sp dedup with BRAND_ACCENTS.
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        local_map = {}
        accents = CALLOUT_ACCENTS if _is_mixed_callout(sp) else None
        for srgb in sp.iter(f"{{{NS_A}}}srgbClr"):
            if not _is_text_run_color(srgb):
                continue
            v = srgb.get("val", "").upper()
            if accents:
                new_val = _remap_accent_sequence(v, local_map, rag_preserve, accents)
            else:
                new_val = _remap_text_color(v, local_map, rag_preserve)
            if new_val is not None:
                srgb.set("val", new_val)
                n_colors += 1

    # Tables: per-TABLE shared map with TABLE_ACCENTS
    # (Vector Blue first, Anchor Green second — preserves categorical row distinctions)
    for tbl in slide_root.iter(f"{{{NS_A}}}tbl"):
        table_map = {}
        for tc in tbl.iter(f"{{{NS_A}}}tc"):
            if tc.get(CIRCLE_MARKER):
                continue
            for srgb in tc.iter(f"{{{NS_A}}}srgbClr"):
                if not _is_text_run_color(srgb):
                    continue
                v = srgb.get("val", "").upper()
                new_val = _remap_accent_sequence(v, table_map, rag_preserve, TABLE_ACCENTS)
                if new_val is not None:
                    srgb.set("val", new_val)
                    n_colors += 1

    # Contrast enforcement: white on dark fills, dark grey on light fills
    n_colors += enforce_text_contrast(slide_root, rag_preserve)

    # Remove markers
    for tc in slide_root.iter(f"{{{NS_A}}}tc"):
        tc.attrib.pop(CIRCLE_MARKER, None)

    return n_fonts, n_colors, n_titles_cased


# ── Text alignment pass ───────────────────────────────────────────────────────
# Rules:
#   Titles         → left-align (algn="l")
#   Body / txBox   → left-align (algn="l")
#   Table cells    → left-align, except header row (first row) which stays centered
#   Skip center/right that was set on non-placeholder free-form content (keep designer intent)
#
# In practice: force "l" on every <a:pPr> that has an explicit non-left algn,
# except inside table header rows.

_CEO_WORKS_RE = re.compile(
    r"CEO[\s.\-_]?WORKS?",
    re.IGNORECASE,
)


def remap_brand_names(slide_root) -> int:
    """Replace all variants of 'CEO Works' with 'FulcrumQ' in every text run.

    Catches: CEO Works, CEO.Works, CEOWORKS, ceo works, CEO-Works, CEOWork, etc.
    Preserves the surrounding run formatting — only the text node is changed.
    Returns the count of substitutions made."""
    count = 0
    for t_el in slide_root.iter(f"{{{NS_A}}}t"):
        if not t_el.text:
            continue
        new_text, n = _CEO_WORKS_RE.subn("FulcrumQ", t_el.text)
        if n:
            t_el.text = new_text
            count += n
    return count


def normalize_alignment(slide_root):
    """Left-align paragraph text in placeholder shapes only.
    Free-form text boxes and non-placeholder shapes keep their alignment —
    centered/right-aligned content there is intentional designer placement.
    Table body rows are left-aligned; header row (first row) is left alone."""
    count = 0

    # Placeholder shapes only
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        if sp.find(f".//{{{NS_P}}}ph") is None:
            continue
        for pPr in sp.iter(f"{{{NS_A}}}pPr"):
            if _in_table(pPr):
                continue
            if pPr.get("algn", "l") not in ("l", ""):
                pPr.set("algn", "l")
                count += 1

    # Table cells: left-align body rows, preserve header row alignment
    for tbl in slide_root.iter(f"{{{NS_A}}}tbl"):
        rows = tbl.findall(f"{{{NS_A}}}tr")
        header_ids = set()
        if rows:
            for tc in rows[0].iter(f"{{{NS_A}}}pPr"):
                header_ids.add(id(tc))
        for pPr in tbl.iter(f"{{{NS_A}}}pPr"):
            if id(pPr) in header_ids:
                continue
            if pPr.get("algn", "l") not in ("l", ""):
                pPr.set("algn", "l")
                count += 1

    return count


# ── Table header row styling ──────────────────────────────────────────────────
# Border element tags that must precede fills in CT_TableCellProperties
_TC_BORDER_TAGS = frozenset({
    f"{{{NS_A}}}lnL", f"{{{NS_A}}}lnR", f"{{{NS_A}}}lnT", f"{{{NS_A}}}lnB",
    f"{{{NS_A}}}lnTlToBr", f"{{{NS_A}}}lnBlToTr",
})

_TABLE_HEADER_FILL  = "281A42"   # Vector Dark Purple
_TABLE_HEADER_TEXT  = "FFFFFF"   # White

def _style_table_header_rows(slide_root) -> int:
    """Apply FulcrumQ header-row style to the first row of every table:
      · Cell fill: Vector Dark Purple (281A42)
      · Text:      white, bold
    Returns the number of text runs styled."""
    count = 0
    for tbl in slide_root.iter(f"{{{NS_A}}}tbl"):
        rows = tbl.findall(f"{{{NS_A}}}tr")
        if not rows:
            continue
        for tc in rows[0].findall(f"{{{NS_A}}}tc"):
            # ── Cell fill ─────────────────────────────────────────────────────
            tcPr = tc.find(f"{{{NS_A}}}tcPr")
            if tcPr is None:
                # Insert tcPr before txBody
                txBody_idx = next(
                    (i for i, ch in enumerate(tc)
                     if ch.tag == f"{{{NS_A}}}txBody"),
                    len(list(tc)),
                )
                tcPr = etree.Element(f"{{{NS_A}}}tcPr")
                tc.insert(txBody_idx, tcPr)
            # Remove any existing fill
            for fill_tag in (f"{{{NS_A}}}solidFill", f"{{{NS_A}}}gradFill",
                             f"{{{NS_A}}}noFill",    f"{{{NS_A}}}blipFill",
                             f"{{{NS_A}}}pattFill"):
                old = tcPr.find(fill_tag)
                if old is not None:
                    tcPr.remove(old)
            # Insert solidFill after border elements
            sf = etree.Element(f"{{{NS_A}}}solidFill")
            srgb_el = etree.SubElement(sf, f"{{{NS_A}}}srgbClr")
            srgb_el.set("val", _TABLE_HEADER_FILL)
            insert_idx = 0
            for ci, child in enumerate(tcPr):
                if child.tag in _TC_BORDER_TAGS:
                    insert_idx = ci + 1
            tcPr.insert(insert_idx, sf)

            # ── Text style: white + bold + Arial ≤ 12pt ──────────────────────
            for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}defRPr"):
                for rPr in tc.iter(tag):
                    rPr.set("b", "1")
                    _set_rpr_color(rPr, _TABLE_HEADER_TEXT)
                    # Font: Arial
                    lat = rPr.find(f"{{{NS_A}}}latin")
                    if lat is None:
                        lat = etree.SubElement(rPr, f"{{{NS_A}}}latin")
                    lat.set("typeface", "Arial")
                    lat.set("pitchFamily", "34")
                    lat.set("charset", "0")
                    # Size: cap at 12pt (1200 hundredths); shrink if larger
                    try:
                        current_sz = int(rPr.get("sz", "0"))
                    except ValueError:
                        current_sz = 0
                    if current_sz == 0 or current_sz > 1200:
                        rPr.set("sz", "1200")
                    count += 1

    return count


# ── Table body row banding ────────────────────────────────────────────────────

_BAND_COLORS = ("F2F2F2", "FFFFFF")   # even rows, odd rows


def _style_table_body_bands(slide_root) -> int:
    """Apply alternating F2F2F2 / white fills to table body rows (rows 1+).
    Cells that already carry an explicit solidFill are left untouched so
    intentional categorical fills (e.g. RAG dots, accent rows) survive.
    Returns the number of cells banded."""
    count = 0
    for tbl in slide_root.iter(f"{{{NS_A}}}tbl"):
        rows = tbl.findall(f"{{{NS_A}}}tr")
        if len(rows) < 2:
            continue
        for row_i, row in enumerate(rows[1:]):   # skip header
            fill_val = _BAND_COLORS[row_i % 2]
            for tc in row.findall(f"{{{NS_A}}}tc"):
                tcPr = tc.find(f"{{{NS_A}}}tcPr")
                # Skip if cell already has an explicit fill
                if tcPr is not None and tcPr.find(f"{{{NS_A}}}solidFill") is not None:
                    continue
                if tcPr is None:
                    txBody_idx = next(
                        (j for j, ch in enumerate(tc)
                         if ch.tag == f"{{{NS_A}}}txBody"),
                        len(list(tc)),
                    )
                    tcPr = etree.Element(f"{{{NS_A}}}tcPr")
                    tc.insert(txBody_idx, tcPr)
                # Insert after any border elements (same ordering as header rows)
                sf = etree.Element(f"{{{NS_A}}}solidFill")
                srgb_el = etree.SubElement(sf, f"{{{NS_A}}}srgbClr")
                srgb_el.set("val", fill_val)
                insert_idx = 0
                for ci, child in enumerate(tcPr):
                    if child.tag in _TC_BORDER_TAGS:
                        insert_idx = ci + 1
                tcPr.insert(insert_idx, sf)
                count += 1
    return count


# ── Compound group styling ────────────────────────────────────────────────────
# enforce_text_contrast() uses z-order in spTree, so shapes *inside* grpSp
# are invisible to its overlap detection.  This pass fixes that gap by
# treating each group as a self-contained unit: find the background fill from
# the rect member(s) and push contrast-correct text color to label members.

def _sp_text_content(sp) -> str:
    return "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()


def _sp_area_emu(sp) -> int:
    spPr = sp.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        return 0
    xfrm = spPr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return 0
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if ext is None:
        return 0
    try:
        return int(ext.get("cx", 0)) * int(ext.get("cy", 0))
    except ValueError:
        return 0


def _set_all_text_color(sp, hex_val: str) -> int:
    """Write hex_val to every text-run color element in sp. Returns run count."""
    n = 0
    for tag in (f"{{{NS_A}}}rPr", f"{{{NS_A}}}endParaRPr", f"{{{NS_A}}}defRPr"):
        for rPr in sp.iter(tag):
            _set_rpr_color(rPr, hex_val)
            n += 1
    return n


def style_compound_groups(slide_root, rag_preserve: set) -> int:
    """
    Style grouped shapes as compound semantic objects.

    Called *after* style_slide() so fills are already brand-snapped.

    Patterns handled
    ────────────────
    A  Self-contained card — sp has both solidFill AND text content.
       Apply white text for dark fills, 1D1D1D for light fills.

    B  Labeled card — GROUP contains separate bg rect (fill, no text) +
       label shapes (text, noFill or transparent).  Contrast-correct text
       color is derived from the largest bg rect's fill.

    C  Uniform strip — N ≥ 3 similarly-sized shapes in a row/column.
       Applies Pattern A to each self-contained child, Pattern B to
       bg+label pairs within the strip.

    Fills in rag_preserve are never touched.
    Returns count of text runs recolored.
    """
    count = 0

    for grpSp in slide_root.iter(f"{{{NS_P}}}grpSp"):
        direct_sp = [c for c in grpSp if c.tag == f"{{{NS_P}}}sp"]
        if not direct_sp:
            continue

        # Classify each direct child
        bg_shapes    = []   # (sp, fill_hex) — has fill, no text
        label_shapes = []   # sp — has text, no (or transparent) fill
        card_shapes  = []   # (sp, fill_hex) — has both fill and text

        for sp in direct_sp:
            fill = _shape_fill_hex(sp)   # handles solid + gradient
            text = _sp_text_content(sp)
            no_fill = (sp.find(f"{{{NS_P}}}spPr/{{{NS_A}}}noFill") is not None
                       or (sp.find(f"{{{NS_P}}}spPr/{{{NS_A}}}solidFill") is None
                           and sp.find(f"{{{NS_P}}}spPr/{{{NS_A}}}gradFill") is None))

            if fill and text:
                card_shapes.append((sp, fill))
            elif fill and not text:
                bg_shapes.append((sp, fill))
            elif text and no_fill:
                label_shapes.append(sp)

        # Pattern A — self-contained cards
        for sp, fill in card_shapes:
            if fill in rag_preserve:
                continue
            tc = "FFFFFF" if _bg_is_dark(fill) else "1D1D1D"
            count += _set_all_text_color(sp, tc)

        # Pattern B — bg rect + label pair(s)
        if bg_shapes and label_shapes:
            # Largest bg rect determines the group's background color
            bg_sp, bg_fill = max(bg_shapes, key=lambda t: _sp_area_emu(t[0]))
            if bg_fill and bg_fill not in rag_preserve:
                tc = "FFFFFF" if _bg_is_dark(bg_fill) else "1D1D1D"
                for lsp in label_shapes:
                    count += _set_all_text_color(lsp, tc)

    return count


# ── Text box margin pass ──────────────────────────────────────────────────────
TXBOX_INS_LR = 91440   # 0.10 inch in EMU  (lIns / rIns on <a:bodyPr>)
TXBOX_INS_TB = 45720   # 0.05 inch in EMU  (tIns / bIns on <a:bodyPr>)

def set_txbox_margins(slide_root):
    """Set internal text margins on all non-placeholder shapes with text content.
    Covers both strict txBox="1" shapes AND regular auto-shapes (txBox=0) with text.
    Skips placeholder shapes — their margins are controlled by the master/layout."""
    count = 0
    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        # Skip placeholders — margin is layout-controlled
        if sp.find(f".//{{{NS_P}}}ph") is not None:
            continue
        txBody = sp.find(f"{{{NS_P}}}txBody")
        if txBody is None:
            continue
        bodyPr = txBody.find(f"{{{NS_A}}}bodyPr")
        if bodyPr is None:
            bodyPr = etree.SubElement(txBody, f"{{{NS_A}}}bodyPr")
        # Skip shapes with explicitly-set margins — author intended that value (even 0)
        if bodyPr.get("lIns") is not None or bodyPr.get("rIns") is not None:
            continue
        bodyPr.set("lIns", str(TXBOX_INS_LR))
        bodyPr.set("rIns", str(TXBOX_INS_LR))
        bodyPr.set("tIns", str(TXBOX_INS_TB))
        bodyPr.set("bIns", str(TXBOX_INS_TB))
        count += 1
    return count


# ── Per-slide text color overrides ───────────────────────────────────────────
# Format: {slide_number (1-based): [(text_to_match, target_hex_color), ...]}
TEXT_COLOR_OVERRIDES = {
    6: [("VALUE", "60BDBC")],   # "VALUE" word → Teal
}

def apply_text_color_overrides(slide_root, slide_idx):
    """Override the color of specific runs by matching their text content."""
    rules = TEXT_COLOR_OVERRIDES.get(slide_idx)
    if not rules:
        return 0
    count = 0
    for r in slide_root.iter(f"{{{NS_A}}}r"):
        t = r.find(f"{{{NS_A}}}t")
        if t is None or not t.text:
            continue
        for match_text, target_hex in rules:
            if match_text in t.text:
                rPr = r.find(f"{{{NS_A}}}rPr")
                if rPr is None:
                    rPr = etree.Element(f"{{{NS_A}}}rPr")
                    r.insert(0, rPr)
                _set_rpr_color(rPr, target_hex)
                count += 1
    return count


# ── Helpers ───────────────────────────────────────────────────────────────────
def get_slide_size(prs_root):
    """Return (width_emu, height_emu) from presentation.xml, or 16:9 default."""
    sz = prs_root.find(f"{{{NS_P}}}sldSz")
    if sz is not None:
        return int(sz.get("cx", 12192000)), int(sz.get("cy", 6858000))
    return 12192000, 6858000


def get_source_layout_names(file_map):
    """Return {layout_file_path: layout_name} from the source master."""
    master = etree.fromstring(file_map["ppt/slideMasters/slideMaster1.xml"])
    m_rels = etree.fromstring(file_map["ppt/slideMasters/_rels/slideMaster1.xml.rels"])
    rid_map = {r.get("Id"): r.get("Target") for r in m_rels}
    idlst   = master.find(f"{{{NS_P}}}sldLayoutIdLst")
    names   = {}
    for el in idlst:
        rid = el.get(f"{{{NS_R}}}id")
        tgt = rid_map.get(rid,"").replace("../","ppt/")
        try:
            lx   = etree.fromstring(file_map[tgt])
            name = lx.find(f"{{{NS_P}}}cSld").get("name","")
            names[tgt] = name
        except Exception:
            pass
    return names


def get_slide_layout(file_map, slide_path):
    """Return the layout path referenced by a slide."""
    rels_path = re.sub(r"ppt/slides/(slide\d+\.xml)$",
                       r"ppt/slides/_rels/\1.rels", slide_path)
    rels_bytes = file_map.get(rels_path, b"")
    if not rels_bytes:
        return None
    rels_root = etree.fromstring(rels_bytes)
    for rel in rels_root:
        tgt = rel.get("Target","")
        if "slideLayout" in tgt:
            p = "ppt/slides/" + tgt
            return re.sub(r"ppt/slides/\.\./","ppt/", p)
    return None


def ordered_slides(file_map):
    """Return slide paths in presentation order."""
    prs      = etree.fromstring(file_map["ppt/presentation.xml"])
    prs_rels = etree.fromstring(file_map["ppt/_rels/presentation.xml.rels"])
    rid_map  = {r.get("Id"): r.get("Target") for r in prs_rels}
    slides   = []
    for el in prs.find(f"{{{NS_P}}}sldIdLst") or []:
        rid = el.get(f"{{{NS_R}}}id")
        tgt = rid_map.get(rid,"")
        if tgt and "slide" in tgt.lower():
            slides.append("ppt/" + tgt)
    return slides


# ── Main ─────────────────────────────────────────────────────────────────────
def convert(source_path: Path):
    output_path = source_path.parent / (source_path.stem + "_fq.pptx")

    primary_master = MASTER_X if MASTER_X.exists() else MASTER_PPTX

    print(f"\nSource : {source_path.name}")
    print(f"Master : {primary_master.name}")
    print(f"Output : {output_path.name}\n")

    with zipfile.ZipFile(str(source_path), "r") as zs:
        file_map = {n: zs.read(n) for n in zs.namelist()}

    with zipfile.ZipFile(str(primary_master), "r") as zm:
        vx_map = {n: zm.read(n) for n in zm.namelist()}

    prs_root = etree.fromstring(file_map["ppt/presentation.xml"])
    slide_w, slide_h = get_slide_size(prs_root)

    # Load triangle bullet PNG
    triangle_path = BASE_DIR / "logo package" / "PNG" / "triangle_default.png"
    triangle_bytes = triangle_path.read_bytes()

    # ── Step 1: inject X master (primary / default) ───────────────────────────
    print("[1] Swapping master…")
    old_layout_names = get_source_layout_names(file_map)
    x_name_map = swap_master(file_map, vx_map)

    _ensure_triangle_media(file_map, triangle_bytes)
    n_svg = remap_svg_colors(file_map)
    if n_svg:
        print(f"  SVG media recolored: {n_svg} file(s)")
    n_charts = remap_chart_colors(file_map)
    if n_charts:
        print(f"  Chart colors remapped: {n_charts} instance(s)")

    # ── Steps 2-4: process each slide ────────────────────────────────────────
    slides = ordered_slides(file_map)
    print(f"\n[2-4] Processing {len(slides)} slides…\n")

    total_vestiges = 0
    total_fonts    = 0
    total_colors   = 0

    for i, spath in enumerate(slides, 1):
        # Identify old layout name
        old_lpath = get_slide_layout(file_map, spath)
        old_name  = old_layout_names.get(old_lpath, "")

        # Step 2: remap layout — parse slide first to detect master choice
        slide_bytes = file_map.get(spath)
        if not slide_bytes:
            continue
        slide_root = etree.fromstring(slide_bytes)

        v7_layout, match = resolve_layout(
            old_name, x_name_map, slide_root,
            slide_idx=i, total_slides=len(slides),
        )
        v7_num = re.search(r"slideLayout(\d+)\.xml$", v7_layout).group(1)
        v7_layout_name = next(
            (k for k, v in x_name_map.items() if v == v7_layout), ""
        )
        remap_slide_layout(file_map, spath, v7_layout)

        print(f"  slide{i:2d}: '{old_name}' [{match}] → layout{v7_num} [{v7_layout_name}]")

        # Lift text from decorative shapes into placeholders (Cover/Divider slides)
        n_lifted = lift_text_into_placeholders(slide_root, v7_layout_name)
        if n_lifted:
            print(f"         lifted  : text from decorative shape → title placeholder")

        promoted = promote_slide_title(slide_root)
        if promoted:
            print(f"         promoted: body ph → title")

        # Subtitle detection: two sources —
        # (a) Shift+Enter in title → split_title_subtitle extracts it
        # (b) Existing subTitle placeholder already present in source slide
        # Either way: upgrade layout to the _Sub variant.
        subtitle_parts = split_title_subtitle(slide_root)
        has_existing_subtitle_ph = any(
            sp.find(f".//{{{NS_P}}}ph") is not None
            and (
                sp.find(f".//{{{NS_P}}}ph").get("type", "") == "subTitle"
                or (sp.find(f".//{{{NS_P}}}ph").get("type", "") == "body"
                    and sp.find(f".//{{{NS_P}}}ph").get("idx", "") == "12")
            )
            for sp in slide_root.iter(f"{{{NS_P}}}sp")
        )
        if subtitle_parts or has_existing_subtitle_ph:
            sub_layout_name = _SUB_LAYOUT_MAP.get(v7_layout_name)
            if sub_layout_name and sub_layout_name in x_name_map:
                v7_layout      = x_name_map[sub_layout_name]
                v7_layout_name = sub_layout_name
                remap_slide_layout(file_map, spath, v7_layout)
                reason = f"'{subtitle_parts[0]}' (split)" if subtitle_parts else "existing subTitle ph"
                print(f"         subtitle : {reason} | layout → {sub_layout_name}")
            elif subtitle_parts:
                print(f"         subtitle : '{subtitle_parts[0]}' → subTitle ph (no _Sub layout available)")

        n_snapped = snap_subtitle_to_layout(slide_root, file_map[v7_layout])
        if n_snapped:
            print(f"         snapped : {n_snapped} subtitle-vibe ph(s) → layout position")

        bullet_promoted = promote_bullet_txbox(slide_root, file_map, spath)
        if bullet_promoted:
            body_layout = x_name_map.get("Content_Light")
            if body_layout:
                body_num = re.search(r"slideLayout(\d+)\.xml$", body_layout).group(1)
                remap_slide_layout(file_map, spath, body_layout)
                print(f"         bullet txBox → body ph, layout → {body_num} (Title & Body)")

        n_footer = strip_footer_placeholders(slide_root, inject=(i > 1))
        if n_footer:
            print(f"         stripped: {n_footer} old footer/sldNum shape(s)")

        removed = clean_vestiges(slide_root, slide_w, slide_h)
        if removed:
            print(f"         removed: {removed}")
            total_vestiges += len(removed)

        convert_scheme_fills(slide_root)
        n_names = remap_brand_names(slide_root)
        if n_names:
            print(f"         renamed  : {n_names} 'CEO Works' → 'FulcrumQ'")
        n_fonts, n_colors, n_cased = style_slide(slide_root, slide_idx=i, layout_name=v7_layout_name)
        total_fonts  += n_fonts
        total_colors += n_colors
        parts = []
        if n_fonts:   parts.append(f"{n_fonts} font(s)")
        if n_colors:  parts.append(f"{n_colors} color(s)")
        if n_cased:   parts.append(f"{n_cased} title(s) sentence-cased")
        if parts:
            print(f"         styled : {', '.join(parts)}")

        n_pics = recolor_pics(slide_root, slide_w, slide_h)
        if n_pics:
            print(f"         recolored: {n_pics} picture(s) → brand palette")

        n_overlays = neutralize_screenshot_overlays(slide_root)
        if n_overlays:
            print(f"         redaction: {n_overlays} overlay shape(s) neutralized → slide bg")


        n_overrides = apply_text_color_overrides(slide_root, i)
        if n_overrides:
            print(f"         overrides: {n_overrides} text color(s) pinned")

        n_margins = set_txbox_margins(slide_root)
        if n_margins:
            print(f"         margins  : {n_margins} txBox(es) padded")

        n_tbl_hdrs = _style_table_header_rows(slide_root)
        if n_tbl_hdrs:
            print(f"         tbl-hdr  : {n_tbl_hdrs} table header cell(s) styled")

        n_tbl_bands = _style_table_body_bands(slide_root)
        if n_tbl_bands:
            print(f"         tbl-band : {n_tbl_bands} table body row(s) banded")

        n_grp = style_compound_groups(slide_root, detect_rag_colors(slide_root))
        if n_grp:
            print(f"         grp-style: {n_grp} group text run(s) contrasted")

        file_map[spath] = etree.tostring(
            slide_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── Save ──────────────────────────────────────────────────────────────────
    print(f"\n  Saving → {output_path.name} …")
    with zipfile.ZipFile(str(output_path), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_map.items():
            zout.writestr(name, data)

    print(f"\n{'='*55}")
    print("SUMMARY")
    print(f"  Slides         : {len(slides)}")
    print(f"  Vestiges removed: {total_vestiges} shape(s)")
    print(f"  Fonts remapped : {total_fonts} run(s)")
    print(f"  Colors remapped: {total_colors} instance(s)")
    print(f"  Output         : {output_path.name}")
    print(f"{'='*55}\n")

    # ── Slide classifier ──────────────────────────────────────────────────────
    from slide_classifier import (classify_deck, write_classifier_output,
                                   print_classifier_summary)
    classifier_results = classify_deck(file_map, slides, old_layout_names)
    print_classifier_summary(classifier_results)
    clf_path = write_classifier_output(classifier_results, output_path)
    print(f"  Classifier → {clf_path.name}\n")

    # ── QA structural validation ───────────────────────────────────────────────
    qa_report = run_qa_checks(file_map, slides,
                              classifier_results=classifier_results)
    _print_qa_report(qa_report)
    qa_path = output_path.with_suffix(".qa.json")
    import json
    qa_path.write_text(json.dumps(qa_report, indent=2))
    print(f"  QA report → {qa_path.name}\n")

    return output_path


# ── QA Validator (Phase 4 — structural checks) ────────────────────────────────

def _slide_layout_name_from_map(file_map, slide_path):
    """Resolve the human-readable layout name for a slide in file_map."""
    rels_path = re.sub(
        r"ppt/slides/(slide\d+\.xml)$",
        r"ppt/slides/_rels/\1.rels",
        slide_path,
    )
    rels_bytes = file_map.get(rels_path, b"")
    if not rels_bytes:
        return None, None
    rels_root = etree.fromstring(rels_bytes)
    layout_path = None
    for rel in rels_root:
        tgt = rel.get("Target", "")
        if "slideLayout" in tgt:
            # Target is relative to ppt/slides/ → resolve to ppt/slideLayouts/
            layout_path = "ppt/slideLayouts/" + tgt.split("/")[-1]
            break
    if not layout_path or layout_path not in file_map:
        return layout_path, None
    try:
        lroot = etree.fromstring(file_map[layout_path])
        cSld  = lroot.find(f".//{{{NS_P}}}cSld")
        name  = cSld.get("name", "") if cSld is not None else ""
    except Exception:
        name = ""
    return layout_path, name


def run_qa_checks(file_map, slides, classifier_results=None):
    """
    Run structural QA checks on the converted file_map.

    Checks
    ------
    cover_count         — exactly 1 slide using a Cover* layout, on slide 1
    ending_position     — Ending* layout (if present) must be the last slide
    adjacent_dividers   — no two consecutive slides both on a Divider* layout
    duplicate_fld_ids   — every <a:fld id="..."> must be unique across the deck
    layout_refs_valid   — every slide's layout rel target must exist in file_map
    classifier_review   — slides flagged for manual review by the classifier

    Returns a dict with:
      "checks":  list of {check, status, detail}  — one entry per check
      "slides":  list of {slide, layout, issues}   — one entry per slide
      "passed":  bool
    """
    import json as _json

    # Build a quick lookup: slide_num → classifier record
    clf_by_slide = {}
    if classifier_results:
        for r in classifier_results:
            clf_by_slide[r["slide"]] = r

    slide_entries = []   # [{slide: n, layout: name, layout_path: p, issues: []}]
    for i, spath in enumerate(slides, 1):
        layout_path, layout_name = _slide_layout_name_from_map(file_map, spath)
        slide_entries.append({
            "slide":       i,
            "path":        spath,
            "layout":      layout_name or "",
            "layout_path": layout_path or "",
            "issues":      [],
        })

    checks = []

    # ── 1. cover_count ────────────────────────────────────────────────────────
    cover_slides = [e for e in slide_entries if e["layout"].startswith("Cover")]
    if len(cover_slides) == 0:
        checks.append({"check": "cover_count", "status": "warn",
                        "detail": "No Cover layout found — slide 1 may be missing a cover"})
    elif len(cover_slides) > 1:
        nums = [e["slide"] for e in cover_slides]
        checks.append({"check": "cover_count", "status": "fail",
                        "detail": f"Multiple Cover slides: {nums}"})
        for e in cover_slides[1:]:
            e["issues"].append("extra cover slide")
    else:
        if cover_slides[0]["slide"] != 1:
            n = cover_slides[0]["slide"]
            checks.append({"check": "cover_count", "status": "fail",
                            "detail": f"Cover slide is slide {n}, not slide 1"})
            cover_slides[0]["issues"].append("cover not on slide 1")
        else:
            checks.append({"check": "cover_count", "status": "pass",
                            "detail": "1 cover slide on slide 1"})

    # ── 2. ending_position ────────────────────────────────────────────────────
    ending_slides = [e for e in slide_entries if e["layout"].startswith("Ending")]
    if not ending_slides:
        checks.append({"check": "ending_position", "status": "warn",
                        "detail": "No Ending layout found"})
    else:
        last_idx = slide_entries[-1]["slide"]
        bad = [e for e in ending_slides if e["slide"] != last_idx]
        if bad:
            nums = [e["slide"] for e in bad]
            checks.append({"check": "ending_position", "status": "fail",
                            "detail": f"Ending slide(s) not last: {nums}"})
            for e in bad:
                e["issues"].append("ending slide not last")
        else:
            checks.append({"check": "ending_position", "status": "pass",
                            "detail": f"Ending on slide {ending_slides[-1]['slide']} (last)"})

    # ── 3. adjacent_dividers ──────────────────────────────────────────────────
    adj_pairs = []
    for i in range(len(slide_entries) - 1):
        a, b = slide_entries[i], slide_entries[i + 1]
        if a["layout"].startswith("Divider") and b["layout"].startswith("Divider"):
            adj_pairs.append((a["slide"], b["slide"]))
    if adj_pairs:
        checks.append({"check": "adjacent_dividers", "status": "fail",
                        "detail": f"Adjacent dividers at slide pairs: {adj_pairs}"})
        for s1, s2 in adj_pairs:
            slide_entries[s1 - 1]["issues"].append("adjacent divider")
            slide_entries[s2 - 1]["issues"].append("adjacent divider")
    else:
        checks.append({"check": "adjacent_dividers", "status": "pass",
                        "detail": "No adjacent divider slides"})

    # ── 4. duplicate_fld_ids ──────────────────────────────────────────────────
    fld_id_seen = {}   # id_val → first slide num
    dupes = []
    for e in slide_entries:
        slide_bytes = file_map.get(e["path"], b"")
        if not slide_bytes:
            continue
        try:
            sroot = etree.fromstring(slide_bytes)
        except Exception:
            continue
        for fld in sroot.iter(f"{{{NS_A}}}fld"):
            fid = fld.get("id", "")
            if not fid:
                continue
            if fid in fld_id_seen:
                dupes.append((fid, fld_id_seen[fid], e["slide"]))
                e["issues"].append(f"duplicate fld id {fid}")
            else:
                fld_id_seen[fid] = e["slide"]
    if dupes:
        checks.append({"check": "duplicate_fld_ids", "status": "fail",
                        "detail": f"{len(dupes)} duplicate field ID(s): "
                                  + "; ".join(f"{d[0]} (slides {d[1]}&{d[2]})" for d in dupes[:5])})
    else:
        checks.append({"check": "duplicate_fld_ids", "status": "pass",
                        "detail": f"All {len(fld_id_seen)} field IDs unique"})

    # ── 5. layout_refs_valid ──────────────────────────────────────────────────
    missing = [e for e in slide_entries
               if e["layout_path"] and e["layout_path"] not in file_map]
    if missing:
        checks.append({"check": "layout_refs_valid", "status": "fail",
                        "detail": f"Missing layout targets: "
                                  + ", ".join(e["layout_path"] for e in missing)})
        for e in missing:
            e["issues"].append(f"layout ref missing: {e['layout_path']}")
    else:
        checks.append({"check": "layout_refs_valid", "status": "pass",
                        "detail": "All slide layout references resolve"})

    # ── 6. classifier_review ──────────────────────────────────────────────────
    if clf_by_slide:
        review_slides = [n for n, r in clf_by_slide.items() if r.get("review")]
        if review_slides:
            checks.append({"check": "classifier_review", "status": "warn",
                            "detail": f"{len(review_slides)} slide(s) need review: {review_slides}"})
            for n in review_slides:
                clf = clf_by_slide[n]
                slide_entries[n - 1]["issues"].append(
                    f"low classifier confidence ({clf['confidence']:.0%}): {clf['reasoning']}"
                )
        else:
            checks.append({"check": "classifier_review", "status": "pass",
                            "detail": "All slides classified above confidence threshold"})

    passed = all(c["status"] in ("pass", "warn") for c in checks)

    # Strip internal path keys before returning
    for e in slide_entries:
        e.pop("path", None)
        e.pop("layout_path", None)

    return {"checks": checks, "slides": slide_entries, "passed": passed}


def _print_qa_report(report):
    STATUS_ICON = {"pass": "✓", "warn": "~", "fail": "✗"}
    print(f"\n{'='*55}")
    print("QA STRUCTURAL CHECKS")
    for c in report["checks"]:
        icon = STATUS_ICON.get(c["status"], "?")
        print(f"  [{icon}] {c['check']:<22}  {c['detail']}")
    flagged = [e for e in report["slides"] if e["issues"]]
    if flagged:
        print(f"\n  Flagged slides:")
        for e in flagged:
            print(f"    slide {e['slide']:2d} [{e['layout']}]: {', '.join(e['issues'])}")
    overall = "PASS" if report["passed"] else "FAIL"
    print(f"\n  Overall: {overall}")
    print(f"{'='*55}")


if __name__ == "__main__":
    args = sys.argv[1:]

    if args and args[0] == "--palette":
        print_palette()
        sys.exit(0)

    if not args:
        args = ["Org_Efficiency.pptx"]

    sources = []
    for a in args:
        p = Path(a)
        if not p.is_absolute():
            p = BASE_DIR / p
        if not p.exists():
            print(f"[skip] not found: {p}")
            continue
        if p.is_dir():
            sources.extend(sorted(p.glob("*.pptx")))
        else:
            sources.append(p)

    if not sources:
        print("No .pptx files found.")
        sys.exit(1)

    print(f"Converting {len(sources)} file(s)…")
    failed = []
    for src in sources:
        try:
            convert(src)
        except Exception as e:
            print(f"\n[ERROR] {src.name}: {e}")
            failed.append(src.name)

    if len(sources) > 1:
        print(f"\nDone. {len(sources) - len(failed)}/{len(sources)} converted.")
        if failed:
            print("Failed:", ", ".join(failed))
