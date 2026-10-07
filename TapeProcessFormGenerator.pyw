"""
Tape Process Form Generator  –  IBAD-PF-009 Rev 00
----------------------------------------------------
Reads an Excel file (Monday.com / IBAD database export) and generates
a multi-page PDF by filling the Tape Process Form template for each tape row.

Pages per tape (Monday and RR124 formats alike) depend on the tape's Length
column — "Length, m" in Monday exports, "Length" in RR124 exports:
  - under 600 m        → 1 page
  - 600 m to 1099 m    → 2 pages
  - 1100 m and above   → 3 pages
The first page carries every extracted value including the coordinates; any
additional page repeats the same values but leaves the coordinate field as a
blank "(______-______)" for hand entry.

Fields filled per page:
  - Hastelloy thickness   → blank before "microns"       [Calibri Regular 11pt]
  - Name (#…(low-high))  → tape number + coordinates     [Calibri Regular 11pt]
  - max(Evaluation_start, Evaluation_end) → IBAD#3_EVAL  [Calibri Italic  10pt]
  - Reference Ic          → IBAD#3_ref                   [Calibri Italic  10pt]

Usage:
  1. Place this file in the same folder as IBAD-PF-009-Rev00_Tape_Process_Form.pdf
  2. Double-click TapeProcessFormGenerator.pyw
  3. Browse to your Excel file, confirm the output path, click Generate.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import io
import os
import re
import sys
import unicodedata
from collections import defaultdict

# ── Auto-install missing dependencies ─────────────────────────────────────────
try:
    import pandas as pd
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "pandas", "openpyxl", "pypdf", "reportlab",
                           "--quiet"])
    import pandas as pd
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

# ── PDF page geometry (A4 portrait) ───────────────────────────────────────────
PDF_W = 595.2
PDF_H = 841.92

# ── Field placement map ────────────────────────────────────────────────────────
# (x_from_left, y_baseline_reportlab, cover_width)
# y_baseline values taken directly from PDF content stream (bottom-origin)
# so text sits on exactly the same baseline as the surrounding template text
FIELDS = {
    # (x_from_left,  y_baseline_reportlab,  cover_width)
    # y values are exact text baselines from the PDF content stream (bottom-origin)
    "hastelloy":  (118.6, 746.1, 22.0),   # Buffer Tape line blank before "microns"
    "tape_num":   (228.8, 746.1, 45.0),   # tape identifier after "#"
    "coords":     (275.4, 746.1, 88.0),   # full "(low  -  high)" block
    "eval":       (191.1, 720.5, 35.0),   # value after IBAD#3_EVAL:
    "ref_ic":     (189.7, 707.2, 35.0),   # value after IBAD#3_ref:
    # NEW (RR124): XRD Tilt values. Baselines match the EVAL / ref lines.
    # The x positions below are ESTIMATES — calibrate against the real template
    # the same way eval/ref were. The template prints "______°"; the cover width
    # is wide enough to erase the underscores AND the printed °, because the
    # value is now drawn with its own degree sign attached (see DEGREE_SIGN),
    # which keeps the ° tight against the number instead of stranded at the end
    # of the blank. Erasing both is what stops a doubled "3.01° °".
    "xrd_min":    (300.0, 720.5, 42.0),   # value after "XRD Tilt_MIN:"  (same line as EVAL)
    "xrd_ave":    (300.0, 707.2, 42.0),   # value after "XRD Tilt_AVE:"  (same line as ref_ic)
}

# Degree sign appended to the XRD tilt values, e.g. "3.01°".
DEGREE_SIGN = "\u00B0"

# ── IBAD Deposition Date row placement ────────────────────────────────────────
# The deposition row is the first data row of the table. Its existing text sits
# on baseline y = 632.5 (bottom-origin). The template already prints the kanji
# "月" (month, at x=66.7) and "日" (day, at x=96.3) plus the slashes, so we only
# overlay the Calibri digits into the blanks — this matches the form exactly and
# needs no Japanese font.
DEP_BASELINE   = 632.5          # baseline of the deposition-date row text
DEP_YEAR_X     = 32.7           # left edge of the year digits ("202_")
DEP_YEAR_COVER = (31.5, 19.0)   # (x, width) white box covering the printed "202_"
DEP_MONTH_RX   = 65.8           # right edge to right-align the 2-digit month (before 月)
DEP_DAY_RX     = 95.3           # right edge to right-align the 2-digit day   (before 日)
DEP_FONT_SIZE  = 9              # matches the row's printed text size

# "Sent to PLD by:" value — centered in the Responsible Person column (col 4),
# on the same deposition row / baseline.
PLD_CENTER_X   = (315.5 + 420.2) / 2.0   # horizontal centre of the Responsible col
PLD_BASELINE   = 632.5
PLD_FONT_SIZE  = 9

# Input validation limits for the "Sent to PLD by:" field
PLD_MIN_CHARS = 4
PLD_MAX_CHARS = 10

# ── Pages per tape, by Length (metres) ────────────────────────────────────────
# under PAGE_LEN_2          → 1 page
# PAGE_LEN_2 .. PAGE_LEN_3  → 2 pages   (600 m itself gives 2)
# PAGE_LEN_3 and above      → 3 pages   (1100 m itself gives 3)
PAGE_LEN_2 = 600.0
PAGE_LEN_3 = 1100.0
MAX_PAGES  = 3

# Written into the coordinate field on continuation pages: the printed
# placeholder is erased and this blank line is drawn in its place.
COORDS_BLANK = "(______-______)"


def pages_for_length(length):
    """
    Number of form pages to generate for a tape of `length` metres.

    A missing / blank / non-numeric length falls back to a single page, which
    is the safe default: the operator gets one form rather than none.
    """
    if length is None:
        return 1
    try:
        L = float(length)
    except (TypeError, ValueError):
        return 1
    if L != L:                       # NaN
        return 1
    if L >= PAGE_LEN_3:
        return 3
    if L >= PAGE_LEN_2:
        return 2
    return 1

# Font specs matching the PDF template
# Buffer Tape line  → Calibri Regular 11pt
# IBAD#3 eval/ref   → Calibri Italic  10pt
FONT_REGULAR_SIZE = 11
FONT_ITALIC_SIZE  = 10

# Field → (font_key, size)
FIELD_FONTS = {
    "hastelloy": ("CalibrRegular", FONT_REGULAR_SIZE),
    "tape_num":  ("CalibrRegular", FONT_REGULAR_SIZE),
    "coords":    ("CalibrRegular", FONT_REGULAR_SIZE),
    "eval":      ("CalibrItalic",  FONT_ITALIC_SIZE),
    "ref_ic":    ("CalibrItalic",  FONT_ITALIC_SIZE),
    "xrd_min":   ("CalibrItalic",  FONT_ITALIC_SIZE),
    "xrd_ave":   ("CalibrItalic",  FONT_ITALIC_SIZE),
}


# ──────────────────────────────────────────────────────────────────────────────
#  FONT SETUP
# ──────────────────────────────────────────────────────────────────────────────

def _font_candidates():
    """Return (regular_path, italic_path) searching common OS locations."""
    win   = os.environ.get("SystemRoot", r"C:\Windows")
    paths = [
        # Windows
        (os.path.join(win, r"Fonts\calibri.ttf"),
         os.path.join(win, r"Fonts\calibrii.ttf")),
        # macOS (Office install)
        ("/Library/Fonts/Microsoft/Calibri.ttf",
         "/Library/Fonts/Microsoft/Calibri Italic.ttf"),
        # Linux Carlito (metric-identical free substitute)
        ("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
         "/usr/share/fonts/truetype/crosextra/Carlito-Italic.ttf"),
        # Ubuntu alternative path
        ("/usr/share/fonts/truetype/carlito/Carlito-Regular.ttf",
         "/usr/share/fonts/truetype/carlito/Carlito-Italic.ttf"),
    ]
    for reg, ital in paths:
        if os.path.isfile(reg) and os.path.isfile(ital):
            return reg, ital
    return None, None


def setup_fonts():
    """
    Register Calibri (or its substitute) with ReportLab.
    Falls back to Helvetica / Helvetica-Oblique if no TTF is found.
    Returns (regular_font_name, italic_font_name).
    """
    reg_path, ital_path = _font_candidates()
    if reg_path:
        try:
            pdfmetrics.registerFont(TTFont("CalibrRegular", reg_path))
            pdfmetrics.registerFont(TTFont("CalibrItalic",  ital_path))
            return "CalibrRegular", "CalibrItalic"
        except Exception:
            pass
    # Built-in fallback
    return "Helvetica", "Helvetica-Oblique"


# Register fonts once at module level
_REG_FONT, _ITAL_FONT = setup_fonts()

# Patch font keys to what was actually registered
FIELD_FONTS["hastelloy"] = (_REG_FONT,  FONT_REGULAR_SIZE)
FIELD_FONTS["tape_num"]  = (_REG_FONT,  FONT_REGULAR_SIZE)
FIELD_FONTS["coords"]    = (_REG_FONT,  FONT_REGULAR_SIZE)
FIELD_FONTS["eval"]      = (_ITAL_FONT, FONT_ITALIC_SIZE)
FIELD_FONTS["ref_ic"]    = (_ITAL_FONT, FONT_ITALIC_SIZE)
FIELD_FONTS["xrd_min"]   = (_ITAL_FONT, FONT_ITALIC_SIZE)
FIELD_FONTS["xrd_ave"]   = (_ITAL_FONT, FONT_ITALIC_SIZE)


# ──────────────────────────────────────────────────────────────────────────────
#  PARSING HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def parse_name(raw_name: str):
    """
    Extract tape identifier and sorted coordinate range from the Name column.

    Examples
    --------
    "#3244(1235-2165) NEW ULBRICH"  →  "3244",  1235, 2165
    "#KM3268(450-10)"               →  "KM3268",  10,  450  (sorted low→high)
    "#KM3264 (25-814)"              →  "KM3264",  25,  814
    Returns (None, None, None) when the pattern is absent.
    """
    m = re.search(r"#([A-Za-z0-9]+)\s*\(\s*(\d+)\s*[-\u2013]\s*(\d+)\s*\)",
                  str(raw_name))
    if not m:
        return None, None, None
    c1, c2 = int(m.group(2)), int(m.group(3))
    return m.group(1), min(c1, c2), max(c1, c2)


def fmt_num(val) -> str:
    """Convert Excel numeric cells to clean strings (drop trailing .0)."""
    try:
        f = float(val)
        if f != f:          # NaN check
            return ""
        return str(int(f)) if f == int(f) else str(round(f, 3))
    except (ValueError, TypeError):
        s = str(val).strip()
        return "" if s.lower() in ("nan", "none", "") else s


def pick_eval(eval_start_raw, eval_end_raw) -> str:
    """
    Return the highest of Evaluation_start and Evaluation_end.
    If only one is available, use that one.
    If both are missing, return empty string.
    """
    s = fmt_num(eval_start_raw)
    e = fmt_num(eval_end_raw)
    if s and e:
        try:
            return s if float(s) >= float(e) else e
        except ValueError:
            return s        # not numeric – fall back to start
    return s or e           # whichever is non-empty, or "" if both empty


def fmt_deposition_date(raw_value):
    """
    Convert the 'IBAD deposition date' cell into (year, month, day) strings.

    The cell is typically a date/timestamp (e.g. 2026-05-17) but may also be a
    plain YYYY-MM-DD string. Month and day are zero-padded to two digits to
    match the form layout, e.g. ("2026", "05", "17").

    Returns (None, None, None) when the cell is empty / not a valid date, so the
    form is left blank for that row.
    """
    if raw_value is None:
        return None, None, None

    # pandas Timestamp / datetime
    ts = pd.to_datetime(raw_value, errors="coerce")
    if pd.isna(ts):
        return None, None, None
    return f"{ts.year:04d}", f"{ts.month:02d}", f"{ts.day:02d}"


def find_header_row(excel_path: str) -> int:
    """Locate the real header row in Monday.com exports (scans first 10 rows)."""
    raw = pd.read_excel(excel_path, engine="openpyxl", header=None, nrows=10)
    for idx, row in raw.iterrows():
        if any(str(v).strip().lower() == "name" for v in row):
            return idx
    return 0


# ──────────────────────────────────────────────────────────────────────────────
#  FORMAT DETECTION + RR124 HELPERS
# ──────────────────────────────────────────────────────────────────────────────

# Signature columns used to recognise an RR124-format export (lower-cased).
_RR124_SIGNATURE = {
    "run_id", "tape_name", "tape_status", "cut_date",
    "sample_cuts_start_coordinate", "sample_cuts_end_coordinate",
}


def detect_format(excel_path: str) -> str:
    """
    Inspect the header row and return 'rr124' or 'monday'.

    Content-based: scans the first 15 rows for either the RR124 signature
    columns or the Monday columns. Raises ValueError if it matches neither.
    """
    raw = pd.read_excel(excel_path, engine="openpyxl", header=None, nrows=15)
    for _, row in raw.iterrows():
        vals = {str(v).strip().lower() for v in row}
        if _RR124_SIGNATURE.issubset(vals):
            return "rr124"
        if "name" in vals and (
            "evaluation_start" in vals or "evaluation_end" in vals
            or "reference ic" in vals
        ):
            return "monday"
    raise ValueError(
        "Unrecognized Excel format.\n\n"
        "The header row matches neither the RR124 columns "
        "(Run_ID, Tape_name, Tape_Status, Cut_Date, Sample_Cuts_*_Coordinate) "
        "nor the Monday columns (Name, Evaluation_start, Reference Ic)."
    )


def _find_rr124_header(excel_path: str) -> int:
    """Locate the RR124 header row (scans first 15 rows for the key columns)."""
    raw = pd.read_excel(excel_path, engine="openpyxl", header=None, nrows=15)
    for idx, row in raw.iterrows():
        vals = {str(v).strip().lower() for v in row}
        if {"run_id", "tape_name", "tape_status"}.issubset(vals):
            return idx
    return 0


def _col(df, name: str):
    """Case-insensitive, whitespace-stripped column lookup."""
    for c in df.columns:
        if str(c).strip().lower() == name.lower():
            return c
    return None


# Accepted headers for the tape-length column (matched like _col): Monday.com
# exports call it "Length, m", RR124 exports call it "Length".
_LENGTH_HEADERS = ("Length", "Length, m")


def _length_col(df):
    """The tape-length column under either export's header, or None."""
    for name in _LENGTH_HEADERS:
        c = _col(df, name)
        if c is not None:
            return c
    return None


# A length typed as text together with its unit, e.g. "855 m", "855m", "650,5 M".
# Group 1 is the number. "mm" (or any other unit) deliberately doesn't match.
_LENGTH_WITH_UNIT = re.compile(r"\s*([\d.,]+)\s*m\s*", re.IGNORECASE)


def parse_length(val):
    """
    Parse a length cell to metres, or None. Same as rr_to_float, except a text
    cell may also carry a trailing "m" unit: "855 m" → 855.0. Full-width
    characters typed with a Japanese IME ("８５５ｍ") are normalized first.
    """
    if isinstance(val, str):
        val = unicodedata.normalize("NFKC", val)
        m = _LENGTH_WITH_UNIT.fullmatch(val)
        if m:
            val = m.group(1)
    return rr_to_float(val)


def _length_and_pages(rec, len_col):
    """
    (length, n_pages) for one row. No length column, or a blank / unreadable
    length, gives (None, 1) — see parse_length and pages_for_length.
    """
    length = parse_length(rec[len_col]) if len_col is not None else None
    return length, pages_for_length(length)


def _status_norm(v) -> str:
    """Normalize Tape_Status: strip + case-fold (so 'cut Eval Sample' counts)."""
    return str(v).strip().casefold()


def _base_name(raw) -> str:
    """Characters before the parenthesis in a Tape_name, whitespace stripped."""
    return str(raw).split("(")[0].strip()


def rr_to_float(val):
    """
    Parse a possibly comma-decimal cell to float (Russian-locale export).
    Returns None for blanks / non-numeric. Only the XRD columns actually use
    comma decimals, but applying this everywhere is harmless.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return None if f != f else f          # NaN → None
    s = str(val).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def fmt_float(v, decimals: int = 2) -> str:
    """Format a float for the form: drop trailing .0, round to `decimals`."""
    if v is None:
        return ""
    if float(v).is_integer():
        return str(int(v))
    return f"{round(float(v), decimals):g}"


def rr124_cut_dates(excel_path: str):
    """Return the sorted distinct Cut_Date values (date objects) on PLD rows."""
    hdr = _find_rr124_header(excel_path)
    df = pd.read_excel(excel_path, engine="openpyxl", header=hdr)
    status_col = _col(df, "Tape_Status")
    cut_col    = _col(df, "Cut_Date")
    if status_col is None or cut_col is None:
        return []
    dates = set()
    for _, row in df.iterrows():
        if _status_norm(row[status_col]) == "sent to pld":
            cd = pd.to_datetime(row[cut_col], errors="coerce")
            if not pd.isna(cd):
                dates.add(cd.date())
    return sorted(dates)


def _cut_endpoints(rec, s_col, e_col):
    """The two coordinate endpoints of a cut-eval row as floats (or None)."""
    return rr_to_float(rec[s_col]), rr_to_float(rec[e_col])


def _dist_to(rec, coord, s_col, e_col) -> float:
    """Distance from `coord` to the nearest endpoint of a cut-eval row."""
    cs, ce = _cut_endpoints(rec, s_col, e_col)
    ds = abs(cs - coord) if cs is not None else float("inf")
    de = abs(ce - coord) if ce is not None else float("inf")
    return min(ds, de)


def _nearest(pool, coord, s_col, e_col, ic_col):
    """
    The cut-eval row nearest to `coord`. Ties broken by higher Ic, then by
    lower coordinate (purely for reproducibility).
    """
    if coord is None or not pool:
        return None

    def key(rec):
        d = _dist_to(rec, coord, s_col, e_col)
        ic = rr_to_float(rec[ic_col])
        ic = ic if ic is not None else 0.0
        cs, ce = _cut_endpoints(rec, s_col, e_col)
        low = min([x for x in (cs, ce) if x is not None], default=float("inf"))
        return (d, -ic, low)        # ascending: nearest, then highest Ic, then lowest coord

    return min(pool, key=key)


def _select_candidates(pool, pld_s, pld_e, s_col, e_col, ic_col):
    """Up to two distinct candidates: nearest to PLD start and to PLD end."""
    picks = []
    for coord in (pld_s, pld_e):
        c = _nearest(pool, coord, s_col, e_col, ic_col)
        if c is not None and not any(c is p for p in picks):
            picks.append(c)
    return picks


def _xrd_triplet(rec, xrd_cols):
    """Return the three XRD tilt floats if all present, else None (all-or-none)."""
    vals = [rr_to_float(rec[c]) for c in xrd_cols]
    if len(vals) >= 3 and all(v is not None for v in vals[:3]):
        return vals[:3]
    return None


def _derive_values(cands, pool, pld_s, pld_e, s_col, e_col, ic_col, ref_col, xrd_cols):
    """
    From the candidate list, derive (eval_val, ref_ic, xrd_min, xrd_ave) strings.

    Selection is NEAREST-first: candidates are ordered by distance to the PLD
    endpoints (ties broken by higher Ic, then lower coordinate). Each field is
    then taken from the nearest candidate that has it, falling back to the next
    nearest. Zero Ic / Reference_Ic is treated as no value (rendered blank).
    """
    if not cands:
        return "", "", "", ""

    def ic_of(rec):
        return rr_to_float(rec[ic_col])

    def cand_dist(rec):
        ds = _dist_to(rec, pld_s, s_col, e_col) if pld_s is not None else float("inf")
        de = _dist_to(rec, pld_e, s_col, e_col) if pld_e is not None else float("inf")
        return min(ds, de)

    def low_coord(rec):
        cs, ce = _cut_endpoints(rec, s_col, e_col)
        return min([x for x in (cs, ce) if x is not None], default=float("inf"))

    # Nearest first; distance tie → higher Ic; then lower coordinate (reproducible).
    ordered = sorted(
        cands,
        key=lambda r: (cand_dist(r),
                       -(ic_of(r) if ic_of(r) is not None else -1.0),
                       low_coord(r)),
    )

    # IC: nearest candidate with a positive Ic (0 / blank treated as no value).
    eval_val = ""
    for r in ordered:
        v = ic_of(r)
        if v is not None and v > 0:
            eval_val = fmt_float(v)
            break

    # XRD triplet: nearest candidate that carries one.
    xrd_min = xrd_ave = ""
    for r in ordered:
        trip = _xrd_triplet(r, xrd_cols)
        if trip:
            xrd_min = fmt_float(min(trip))
            xrd_ave = fmt_float(sum(trip) / len(trip))
            break

    # Reference_Ic: nearest candidate with a positive value; else nearest in pool.
    def ref_of(rec):
        v = rr_to_float(rec[ref_col])
        return v if (v is not None and v > 0) else None

    ref = None
    for r in ordered:
        if ref_of(r) is not None:
            ref = ref_of(r)
            break
    if ref is None:
        with_ref = [r for r in pool if ref_of(r) is not None]
        anchor = pld_s if pld_s is not None else pld_e
        if with_ref and anchor is not None:
            best = min(with_ref, key=lambda r: _dist_to(r, anchor, s_col, e_col))
            ref = ref_of(best)
        elif with_ref:
            ref = ref_of(with_ref[0])
    ref_ic = fmt_float(ref) if ref is not None else ""

    return eval_val, ref_ic, xrd_min, xrd_ave


# ──────────────────────────────────────────────────────────────────────────────
#  PDF OVERLAY BUILDER
# ──────────────────────────────────────────────────────────────────────────────

def _draw_field(c, field_key: str, text: str):
    """
    Erase the placeholder area, then write `text` with the correct font/size.
    x and rl_y are taken directly from the PDF content stream so the text
    lands on exactly the same baseline as the surrounding template characters.
    """
    x, rl_y, cover_w = FIELDS[field_key]          # rl_y is already bottom-origin
    font_name, font_size = FIELD_FONTS[field_key]

    # White erase rectangle (covers placeholder underscores)
    c.setFillColorRGB(1.0, 1.0, 1.0)
    c.rect(x - 1, rl_y - 2, cover_w + 2, font_size + 4, fill=1, stroke=0)

    # Draw text on the exact same baseline as the surrounding template text
    c.setFillColorRGB(0.0, 0.0, 0.0)
    c.setFont(font_name, font_size)
    c.drawString(x, rl_y, text)


def _draw_deposition_date(c, year, month, day):
    """
    Fill the IBAD Deposition Date cell.  The template already prints the kanji
    月 / 日 and the slashes, so we overlay only the Calibri digits:
      - the 4-digit year over the printed "202_" placeholder
      - the 2-digit month right-aligned just before 月
      - the 2-digit day  right-aligned just before 日
    """
    font = _REG_FONT
    # Year — cover the printed "202_" and write the full year
    cover_x, cover_w = DEP_YEAR_COVER
    c.setFillColorRGB(1.0, 1.0, 1.0)
    c.rect(cover_x, DEP_BASELINE - 2, cover_w, DEP_FONT_SIZE + 3, fill=1, stroke=0)
    c.setFillColorRGB(0.0, 0.0, 0.0)
    c.setFont(font, DEP_FONT_SIZE)
    c.drawString(DEP_YEAR_X, DEP_BASELINE, year)
    # Month — right-aligned before 月
    c.drawRightString(DEP_MONTH_RX, DEP_BASELINE, month)
    # Day — right-aligned before 日
    c.drawRightString(DEP_DAY_RX, DEP_BASELINE, day)


def _draw_pld_by(c, text):
    """Center the 'Sent to PLD by:' value in the Responsible Person column."""
    c.setFillColorRGB(0.0, 0.0, 0.0)
    c.setFont(_REG_FONT, PLD_FONT_SIZE)
    c.drawCentredString(PLD_CENTER_X, PLD_BASELINE, text)


def build_overlay(tape_num, low, high,
                  hastelloy, eval_val, ref_ic,
                  dep_year=None, dep_month=None, dep_day=None,
                  pld_by="", xrd_min="", xrd_ave="",
                  blank_coords=False) -> io.BytesIO:
    """
    Create a single-page ReportLab PDF containing only the filled values.
    Each value erases the template's printed underscore blank behind it.

    `blank_coords=True` is used for the continuation pages of a long tape: every
    other value is drawn exactly as on page 1, but the coordinate field is
    written as COORDS_BLANK — "(______-______)" — instead of the real range.
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(PDF_W, PDF_H))

    if hastelloy:
        _draw_field(c, "hastelloy", hastelloy)
    if tape_num:
        _draw_field(c, "tape_num", tape_num)
    if blank_coords:
        # Continuation page: erase the printed placeholder and lay down an empty
        # line in its place, so the template's own underscores cannot show
        # through the ones we draw.
        _draw_field(c, "coords", COORDS_BLANK)
    elif low is not None and high is not None:
        _draw_field(c, "coords", f"( {low}   -   {high} )")
    if eval_val:
        _draw_field(c, "eval", eval_val)
    if ref_ic:
        _draw_field(c, "ref_ic", ref_ic)

    # NEW (RR124): XRD Tilt values (left blank for Monday-format records).
    # The degree sign is drawn as part of the value so it sits directly after
    # the number; the template's own printed ° is erased underneath.
    if xrd_min:
        _draw_field(c, "xrd_min", f"{xrd_min}{DEGREE_SIGN}")
    if xrd_ave:
        _draw_field(c, "xrd_ave", f"{xrd_ave}{DEGREE_SIGN}")

    # IBAD Deposition Date (left blank if no valid date for this row)
    if dep_year and dep_month and dep_day:
        _draw_deposition_date(c, dep_year, dep_month, dep_day)

    # "Sent to PLD by:" — same value on every page (left blank if not entered)
    if pld_by:
        _draw_pld_by(c, pld_by)

    c.save()
    buf.seek(0)
    return buf


# ──────────────────────────────────────────────────────────────────────────────
#  CORE GENERATION
# ──────────────────────────────────────────────────────────────────────────────

def read_monday(excel_path: str, log_cb=None):
    """
    Monday.com / IBAD database export reader.
    Returns (records, skipped) where each record is a normalized per-tape dict.
    Behaviour is identical to the original program's reading logic, except that
    the page count now follows the "Length, m" column, as in the RR124 reader.
    """
    header_row = find_header_row(excel_path)
    df = pd.read_excel(excel_path, engine="openpyxl", header=header_row)

    name_col       = _col(df, "Name")
    hastelloy_col  = _col(df, "Hastelloy thickness")
    eval_start_col = _col(df, "Evaluation_start")
    eval_end_col   = _col(df, "Evaluation_end")
    ref_col        = _col(df, "Reference Ic")
    dep_date_col   = _col(df, "IBAD deposition date")
    len_col        = _length_col(df)              # optional – drives page count

    missing = [n for n, c_ in [
        ("Name", name_col),
        ("Hastelloy thickness", hastelloy_col),
        ("Reference Ic", ref_col),
    ] if c_ is None]
    if missing:
        raise ValueError(
            "The following column(s) were not found in the Excel file:\n\n"
            + "\n".join(f"  • {m}" for m in missing)
            + "\n\nPlease ensure the file contains these exact column names."
        )

    if eval_start_col is None and eval_end_col is None and log_cb:
        log_cb("⚠  Neither Evaluation_start nor Evaluation_end found — "
               "EVAL field will be blank.")
    if dep_date_col is None and log_cb:
        log_cb("⚠  'IBAD deposition date' column not found — "
               "deposition date will be blank.")
    if len_col is None and log_cb:
        log_cb("⚠  'Length, m' column not found — every tape will get 1 page.")

    records, skipped = [], []
    for i, (_, row) in enumerate(df.iterrows()):
        raw_name = str(row[name_col]) if name_col else ""

        if not re.search(r"#[A-Za-z0-9]+\s*\(", raw_name):
            if raw_name.strip() and raw_name.lower() != "nan":
                skipped.append(
                    f"Row {i + header_row + 2}: no tape pattern  –  \"{raw_name[:55]}\"")
            continue

        tape_num, low, high = parse_name(raw_name)
        if tape_num is None:
            skipped.append(
                f"Row {i + header_row + 2}: could not parse coords  –  \"{raw_name[:55]}\"")
            continue

        hastelloy = fmt_num(row[hastelloy_col]) if hastelloy_col else ""
        eval_val  = pick_eval(
            row[eval_start_col] if eval_start_col else None,
            row[eval_end_col]   if eval_end_col   else None,
        )
        ref_ic = fmt_num(row[ref_col]) if ref_col else ""

        dep_year, dep_month, dep_day = (None, None, None)
        if dep_date_col is not None:
            dep_year, dep_month, dep_day = fmt_deposition_date(row[dep_date_col])

        length, n_pages = _length_and_pages(row, len_col)

        dep_str = f"{dep_year}/{dep_month}/{dep_day}" if dep_year else "—"
        len_str = f"{fmt_float(length)} m" if length is not None else "—"
        records.append(dict(
            tape_num=tape_num, low=low, high=high, hastelloy=hastelloy,
            eval_val=eval_val, ref_ic=ref_ic,
            dep_year=dep_year, dep_month=dep_month, dep_day=dep_day,
            xrd_min="", xrd_ave="",
            length=length, pages=n_pages,
            log=(f"✔  #{tape_num} ({low}–{high})   "
                 f"Hastelloy={hastelloy} µm   "
                 f"EVAL={eval_val or '—'} A   "
                 f"REF={ref_ic or '—'} A   "
                 f"Date={dep_str}   "
                 f"Len={len_str} → {n_pages} pg"),
        ))

    return records, skipped


def read_rr124(excel_path: str, selected_date, log_cb=None):
    """
    RR124-format reader. One record per 'Sent to PLD' row whose Cut_Date matches
    `selected_date`. IC / XRD / Reference_Ic are pulled from the matching
    cut-eval samples per the finalized feature spec.
    Returns (records, skipped).
    """
    if not selected_date:
        raise ValueError("A Cut_Date must be selected before generating "
                         "from an RR124-format file.")

    hdr = _find_rr124_header(excel_path)
    df  = pd.read_excel(excel_path, engine="openpyxl", header=hdr)

    run_col    = _col(df, "Run_ID")
    name_col   = _col(df, "Tape_name")
    dep_col    = _col(df, "Date_of_IBAD_Deposition")
    thick_col  = _col(df, "Hastealloy_Thickness")   # note: misspelled in export
    status_col = _col(df, "Tape_Status")
    cut_col    = _col(df, "Cut_Date")
    s_col      = _col(df, "Sample_Cuts_Start_Coordinate")
    e_col      = _col(df, "Sample_Cuts_End_Coordinate")
    ic_col     = _col(df, "Ic")
    ref_col    = _col(df, "Reference_Ic")
    len_col    = _length_col(df)                    # optional – drives page count
    xrd_cols   = [c for c in df.columns
                  if str(c).strip().lower().startswith("2d_xrd_tilt_")]

    required = {
        "Run_ID": run_col, "Tape_name": name_col, "Tape_Status": status_col,
        "Cut_Date": cut_col, "Sample_Cuts_Start_Coordinate": s_col,
        "Sample_Cuts_End_Coordinate": e_col, "Ic": ic_col,
        "Reference_Ic": ref_col, "Date_of_IBAD_Deposition": dep_col,
        "Hastealloy_Thickness": thick_col,
    }
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(
            "The RR124 file is missing these column(s):\n\n"
            + "\n".join(f"  • {m}" for m in missing))
    if len(xrd_cols) < 3 and log_cb:
        log_cb(f"⚠  Found {len(xrd_cols)} XRD tilt column(s); expected 3 — "
               "XRD fields may stay blank.")
    if len_col is None and log_cb:
        log_cb("⚠  'Length' column not found — every tape will get 1 page.")

    rows = df.to_dict("records")

    # Candidate pools keyed on (Run_ID, base tape name) over cut-eval rows.
    pool = defaultdict(list)
    for r in rows:
        if _status_norm(r[status_col]) == "cut eval sample":
            pool[(r[run_col], _base_name(r[name_col]))].append(r)

    sel = pd.to_datetime(selected_date, errors="coerce")
    sel = sel.date() if not pd.isna(sel) else None

    records, skipped = [], []
    for r in rows:
        if _status_norm(r[status_col]) != "sent to pld":
            continue
        cd = pd.to_datetime(r[cut_col], errors="coerce")
        if pd.isna(cd) or sel is None or cd.date() != sel:
            continue

        base = _base_name(r[name_col])
        s = rr_to_float(r[s_col])
        e = rr_to_float(r[e_col])
        if s is not None and e is not None:
            low, high = int(round(min(s, e))), int(round(max(s, e)))
        elif s is not None:
            low = high = int(round(s))
        elif e is not None:
            low = high = int(round(e))
        else:
            low = high = None

        hastelloy = fmt_num(r[thick_col])
        dep_year, dep_month, dep_day = fmt_deposition_date(r[dep_col])

        key   = (r[run_col], base)
        cands = _select_candidates(pool[key], s, e, s_col, e_col, ic_col)
        eval_val, ref_ic, xrd_min, xrd_ave = _derive_values(
            cands, pool[key], s, e, s_col, e_col, ic_col, ref_col, xrd_cols)

        length, n_pages = _length_and_pages(r, len_col)

        dep_str = f"{dep_year}/{dep_month}/{dep_day}" if dep_year else "—"
        len_str = f"{fmt_float(length)} m" if length is not None else "—"
        records.append(dict(
            tape_num=base, low=low, high=high, hastelloy=hastelloy,
            eval_val=eval_val, ref_ic=ref_ic,
            dep_year=dep_year, dep_month=dep_month, dep_day=dep_day,
            xrd_min=xrd_min, xrd_ave=xrd_ave,
            length=length, pages=n_pages,
            log=(f"✔  {base} ({low}–{high})   "
                 f"Hastelloy={hastelloy} µm   "
                 f"EVAL={eval_val or '—'} A   "
                 f"REF={ref_ic or '—'} A   "
                 f"XRD_MIN={xrd_min or '—'}  XRD_AVE={xrd_ave or '—'}  "
                 f"Date={dep_str}   "
                 f"Len={len_str} → {n_pages} pg"),
        ))

    return records, skipped


def generate(excel_path: str, template_path: str, output_path: str,
             pld_by: str = "", selected_date=None, fmt: str = None,
             progress_cb=None, log_cb=None):
    """
    Detect the file format (unless `fmt` is supplied), dispatch to the right
    reader, then render onto the template the number of pages each record calls
    for (see pages_for_length). Page 1 carries the coordinates; any further page
    for the same tape repeats every other value with the coordinate field left
    blank.

    Returns (pages_generated, skipped_descriptions).
    """
    if fmt is None:
        fmt = detect_format(excel_path)

    if fmt == "rr124":
        records, skipped = read_rr124(excel_path, selected_date, log_cb=log_cb)
    elif fmt == "monday":
        records, skipped = read_monday(excel_path, log_cb=log_cb)
    else:
        raise ValueError(f"Unknown format: {fmt!r}")

    if not records:
        if fmt == "rr124":
            raise ValueError(
                "No 'Sent to PLD' rows match the selected Cut_Date.\n\n"
                "Pick a different date, or check the file.")
        raise ValueError(
            "No valid tape rows were found in the Excel file.\n\n"
            "Make sure the Name column contains entries like:\n"
            "  #3244(1235–2165)\n  #KM3268(450–10)")

    writer = PdfWriter()
    # Progress is counted in PAGES, not records: a single long tape can now
    # contribute up to MAX_PAGES pages.
    total = sum(max(1, int(rec.get("pages", 1))) for rec in records)
    ok    = 0

    for rec in records:
        n_pages = max(1, min(MAX_PAGES, int(rec.get("pages", 1))))

        if log_cb and rec.get("log"):
            log_cb(rec["log"])

        for p in range(n_pages):
            first = (p == 0)

            overlay_buf = build_overlay(
                rec["tape_num"],
                rec["low"]  if first else None,
                rec["high"] if first else None,
                rec["hastelloy"], rec["eval_val"], rec["ref_ic"],
                rec.get("dep_year"), rec.get("dep_month"), rec.get("dep_day"),
                pld_by, rec.get("xrd_min", ""), rec.get("xrd_ave", ""),
                blank_coords=not first,
            )
            # Re-read the template for every page: merge_page mutates the page
            # object, so a shared reader would stack overlays onto one page.
            template_reader = PdfReader(template_path)
            overlay_reader  = PdfReader(overlay_buf)
            page = template_reader.pages[0]
            page.merge_page(overlay_reader.pages[0])
            writer.add_page(page)
            ok += 1

            if progress_cb:
                progress_cb(ok, total)
            if log_cb and n_pages > 1 and not first:
                log_cb(f"   ↳ page {p + 1} of {n_pages} — coordinates blank")

    with open(output_path, "wb") as fh:
        writer.write(fh)

    if log_cb:
        log_cb(f"\n{len(records)} tape(s) → {ok} page(s).")

    return ok, skipped


# ──────────────────────────────────────────────────────────────────────────────
#  GUI
# ──────────────────────────────────────────────────────────────────────────────

ACCENT  = "#803C43"   # light maroon
BG      = "#F5F7FA"
BTN_BG  = "#803C43"   # light maroon
BTN_HVR = "#5E2C31"   # darker maroon for hover
BTN_FG  = "#FFFFFF"
GRAY    = "#90A4AE"
TEXT_DK = "#212121"
TEXT_SM = "#546E7A"
GREEN   = "#2E7D32"
RED     = "#C62828"


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Tape Process Form Generator")
        self.resizable(False, False)
        self.configure(bg=BG)

        self.excel_path    = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path   = tk.StringVar()
        self.pld_by        = tk.StringVar()
        self.cut_date      = tk.StringVar()

        self.detected_format = None     # 'rr124' | 'monday' | None
        self._cut_dates      = []       # list[date] available for the picker

        here    = os.path.dirname(os.path.abspath(__file__))
        default = os.path.join(here, "IBAD-PF-009-Rev00_Tape_Process_Form.pdf")
        if os.path.isfile(default):
            self.template_path.set(default)

        self._build_ui()
        self._center_window(510, 648)

    # ── Input validation ────────────────────────────────────────────────────
    def _validate_pld(self, proposed: str) -> bool:
        """Allow typing but cap length at PLD_MAX_CHARS (min is checked on Generate)."""
        return len(proposed) <= PLD_MAX_CHARS

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        hdr = tk.Frame(self, bg=ACCENT, height=56)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="IBAD Tape Process Form",
            font=("Segoe UI", 14, "bold"),
            fg="white", bg=ACCENT, anchor="center",
        ).pack(fill="x", pady=14)

        body = tk.Frame(self, bg=BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)

        self._file_row(body, "Excel Data File (.xlsx):",
                       self.excel_path, self._pick_excel, row=0)
        self._file_row(body, "Template PDF (IBAD-PF-009):",
                       self.template_path, self._pick_template, row=1)
        self._file_row(body, "Save Output PDF As:",
                       self.output_path, self._pick_output,
                       row=2, is_save=True)

        # 4th input row: "Sent to PLD by:" (typed value, not a file)
        tk.Label(
            body, text="Sent to PLD by:", font=("Segoe UI", 9),
            fg=TEXT_DK, bg=BG, anchor="w",
        ).grid(row=3, column=0, sticky="w", padx=(0, 8), pady=3)
        vcmd = (self.register(self._validate_pld), "%P")
        self.pld_entry = tk.Entry(
            body, textvariable=self.pld_by, width=34,
            font=("Segoe UI", 9), relief="solid", bd=1,
            validate="key", validatecommand=vcmd,
        )
        self.pld_entry.grid(row=3, column=1, sticky="ew", pady=3)
        tk.Label(
            body, text=f"{PLD_MIN_CHARS}–{PLD_MAX_CHARS} chars",
            font=("Segoe UI", 8), fg=GRAY, bg=BG, anchor="w",
        ).grid(row=3, column=2, sticky="w", padx=(6, 0), pady=3)

        # 5th input row: Cut_Date picker — shown only for RR124-format files.
        self.cut_date_label = tk.Label(
            body, text="Date sent to PLD:", font=("Segoe UI", 9),
            fg=TEXT_DK, bg=BG, anchor="w",
        )
        self.cut_date_label.grid(row=4, column=0, sticky="w", padx=(0, 8), pady=3)
        self.cut_date_combo = ttk.Combobox(
            body, textvariable=self.cut_date, width=32,
            font=("Segoe UI", 9), state="readonly",
        )
        self.cut_date_combo.grid(row=4, column=1, sticky="ew", pady=3)
        self.cut_date_hint = tk.Label(
            body, text="required", font=("Segoe UI", 8),
            fg=GRAY, bg=BG, anchor="w",
        )
        self.cut_date_hint.grid(row=4, column=2, sticky="w", padx=(6, 0), pady=3)
        # Hidden until an RR124 file is selected.
        self.cut_date_label.grid_remove()
        self.cut_date_combo.grid_remove()
        self.cut_date_hint.grid_remove()

        ttk.Separator(body, orient="horizontal").grid(
            row=6, column=0, columnspan=3, sticky="ew", pady=10)

        self.gen_btn = self._btn(body, "⚙   Generate PDF",
                                 self._start_generate, width=22)
        self.gen_btn.grid(row=7, column=0, columnspan=3, pady=(0, 6))

        self.progress = ttk.Progressbar(body, length=440, mode="determinate")
        self.progress.grid(row=8, column=0, columnspan=3,
                           sticky="ew", pady=(4, 2))

        self.status_var = tk.StringVar(value="Ready – select an Excel file to begin.")
        self._status_lbl = tk.Label(
            body, textvariable=self.status_var,
            font=("Segoe UI", 9), fg=TEXT_SM, bg=BG,
            wraplength=440, justify="left", anchor="w",
        )
        self._status_lbl.grid(row=9, column=0, columnspan=3,
                               sticky="w", pady=(2, 4))

        log_frame = tk.Frame(body, bg=BG)
        log_frame.grid(row=10, column=0, columnspan=3,
                       sticky="nsew", pady=(6, 0))
        body.grid_rowconfigure(10, weight=1)

        self.log_box = tk.Text(
            log_frame, height=9, width=60,
            font=("Consolas", 8), bg="#ECEFF1",
            relief="flat", state="disabled", wrap="word",
        )
        sb = ttk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        tk.Label(
            self,
            text="Faraday Factory Japan",
            font=("Segoe UI", 8), fg=GRAY, bg=BG,
        ).pack(pady=(4, 8))

    def _file_row(self, parent, label, var, cmd, row, is_save=False):
        tk.Label(
            parent, text=label, font=("Segoe UI", 9),
            fg=TEXT_DK, bg=BG, anchor="w",
        ).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        tk.Entry(
            parent, textvariable=var, width=34,
            font=("Segoe UI", 9), relief="solid", bd=1,
        ).grid(row=row, column=1, sticky="ew", pady=3)
        self._btn(parent, "Save As…" if is_save else "Browse…",
                  cmd, width=10
                  ).grid(row=row, column=2, padx=(6, 0), pady=3)

    def _btn(self, parent, text, cmd, width=None):
        kw = {"width": width} if width else {}
        b = tk.Button(
            parent, text=text, command=cmd,
            font=("Segoe UI", 9, "bold"),
            bg=BTN_BG, fg=BTN_FG,
            activebackground=BTN_HVR, activeforeground=BTN_FG,
            relief="flat", cursor="hand2", padx=10, pady=5, **kw,
        )
        b.bind("<Enter>", lambda _: b.configure(bg=BTN_HVR))
        b.bind("<Leave>", lambda _: b.configure(bg=BTN_BG))
        return b

    def _center_window(self, w, h):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── File pickers ──────────────────────────────────────────────────────────

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            if not self.output_path.get():
                base = os.path.splitext(path)[0]
                self.output_path.set(base + "_TapeProcessForms.pdf")
            self._log(f"Excel: {os.path.basename(path)}")
            self._detect_and_setup(path)

    def _detect_and_setup(self, path):
        """Detect the file format and show/hide the Cut_Date picker."""
        self.detected_format = None
        self._cut_dates = []
        self.cut_date.set("")
        try:
            self.detected_format = detect_format(path)
        except Exception as exc:
            self._log(f"⚠  Could not detect format: {exc}")

        if self.detected_format == "rr124":
            try:
                self._cut_dates = rr124_cut_dates(path)
            except Exception as exc:
                self._log(f"⚠  Could not read Cut_Date values: {exc}")
            self._log(f"Format: RR124  –  {len(self._cut_dates)} cut date(s) available.")
        elif self.detected_format == "monday":
            self._log("Format: Monday (existing reader).")

        self._update_cut_date_picker()

    def _update_cut_date_picker(self):
        """Reveal the picker (populated with available dates) only for RR124."""
        if self.detected_format == "rr124" and self._cut_dates:
            self.cut_date_combo["values"] = [d.isoformat() for d in self._cut_dates]
            self.cut_date_label.grid()
            self.cut_date_combo.grid()
            self.cut_date_hint.grid()
        else:
            self.cut_date_label.grid_remove()
            self.cut_date_combo.grid_remove()
            self.cut_date_hint.grid_remove()

    def _pick_template(self):
        path = filedialog.askopenfilename(
            title="Select Tape Process Form Template PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.template_path.set(path)
            self._log(f"Template: {os.path.basename(path)}")

    def _pick_output(self):
        init_file = (os.path.basename(self.output_path.get())
                     if self.output_path.get() else "TapeProcessForms.pdf")
        init_dir  = (os.path.dirname(self.output_path.get())
                     if self.output_path.get() else os.path.expanduser("~"))
        path = filedialog.asksaveasfilename(
            title="Save Output PDF As",
            defaultextension=".pdf",
            initialfile=init_file,
            initialdir=init_dir,
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self.output_path.set(path)

    # ── Log / status ──────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, msg: str, color: str = TEXT_SM):
        self.status_var.set(msg)
        self._status_lbl.configure(fg=color)

    # ── Generation ────────────────────────────────────────────────────────────

    def _start_generate(self):
        excel = self.excel_path.get().strip()
        tpl   = self.template_path.get().strip()
        out   = self.output_path.get().strip()
        pld   = self.pld_by.get().strip()
        sel_date = self.cut_date.get().strip()

        if not excel:
            messagebox.showwarning("Missing File",
                                   "Please select an Excel data file.")
            return
        if not tpl:
            messagebox.showwarning(
                "Missing Template",
                "Template PDF not found.\n\n"
                "Place IBAD-PF-009-Rev00_Tape_Process_Form.pdf in the same\n"
                "folder as this program, or use Browse… to locate it.")
            return
        if not out:
            messagebox.showwarning("Missing Path",
                                   "Please specify where to save the output PDF.")
            return
        # "Sent to PLD by:" — required, 4–10 characters
        if not (PLD_MIN_CHARS <= len(pld) <= PLD_MAX_CHARS):
            messagebox.showwarning(
                "Check \u201CSent to PLD by\u201D",
                f"The \u201CSent to PLD by:\u201D field must be between "
                f"{PLD_MIN_CHARS} and {PLD_MAX_CHARS} characters.\n\n"
                f"It currently has {len(pld)} character(s).")
            self.pld_entry.focus_set()
            return
        # RR124 files require a Cut_Date selection before generating
        if self.detected_format == "rr124" and not sel_date:
            messagebox.showwarning(
                "Select a Cut Date",
                "This is an RR124-format file, so a Cut Date is required.\n\n"
                "Choose a date from the dropdown before generating.")
            self.cut_date_combo.focus_set()
            return
        if not os.path.isfile(excel):
            messagebox.showerror("File Not Found", f"Excel file not found:\n{excel}")
            return
        if not os.path.isfile(tpl):
            messagebox.showerror("File Not Found", f"Template PDF not found:\n{tpl}")
            return

        self.gen_btn.configure(state="disabled", text="⏳  Generating…")
        self.progress["value"] = 0
        self._log(f"\n{'─' * 52}")
        self._log("Starting…")
        self._set_status("Generating PDF, please wait…")

        fmt = self.detected_format

        def _worker():
            try:
                def _prog(done, total):
                    self.progress["value"] = int(done / total * 100)
                    self._set_status(f"Processing page {done} of {total}…")

                pages, skipped = generate(excel, tpl, out,
                                          pld_by=pld,
                                          selected_date=sel_date or None,
                                          fmt=fmt,
                                          progress_cb=_prog,
                                          log_cb=self._log)
                self.progress["value"] = 100

                if skipped:
                    self._log(f"\n⚠  {len(skipped)} row(s) skipped:")
                    for s in skipped:
                        self._log(f"   {s}")

                self._log(f"\n✅  {pages} page(s) saved to:\n   {out}")
                self._set_status(
                    f"✅  Done!  {pages} page(s) →  {os.path.basename(out)}",
                    color=GREEN,
                )

                if messagebox.askyesno(
                    "PDF Generated",
                    f"✅  Success!\n\n"
                    f"{pages} page(s) saved to:\n{out}\n\n"
                    "Open the output folder?",
                    icon="info",
                ):
                    _open_folder(os.path.dirname(os.path.abspath(out)))

            except Exception as exc:
                self._log(f"\n❌  Error: {exc}")
                self._set_status(f"Error: {exc}", color=RED)
                messagebox.showerror("Generation Failed", str(exc))
            finally:
                self.gen_btn.configure(state="normal", text="⚙   Generate PDF")

        threading.Thread(target=_worker, daemon=True).start()


def _open_folder(path: str):
    import platform, subprocess
    sys_ = platform.system()
    if sys_ == "Windows":
        os.startfile(path)
    elif sys_ == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
