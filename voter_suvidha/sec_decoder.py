# -*- coding: utf-8 -*-
"""
sec_decoder.py — correct text extraction for Rajasthan SEC electoral roll PDFs.

WHY THIS EXISTS
---------------
These PDFs embed a legacy 8-bit Devanagari font whose glyphs live in the
Private Use Area (U+F021..U+F0FF). The PDF also carries a /ToUnicode CMap,
but that CMap is CORRUPT: several distinct glyph codes are all mapped to the
same Devanagari codepoint. Verified examples from BADNOR-Ward_No-001.pdf:

    codes 0x21 ' ', 0x27 "'", 0x37 '7', 0x41 'A', 0x53 'S'  ->  all claim  र

Because the mapping is many-to-one, text extraction is LOSSY. pdftotext and
PyMuPDF both produce the same garbage ("भलमररज" instead of "भीमराज"), and no
amount of post-processing or word-dictionary lookup can reliably invert it.
That is the root cause of wrong names on the generated slips.

THE FIX
-------
The raw byte codes in the page content stream are NOT lossy. We read them
directly out of the Form XObjects and decode with the real glyph table
(GLYPH_MAP below), then reorder the i-matra and reph, which this font stores
in visual order.

Proven by cross-checking raw bytes against the rendered page:
    '%!4'      -> न + ा + म        = नाम       (label "नाम")
    '43!%'     -> म + क + ा + न    = मकान      (label "मकान")
    '$/1!'     -> ि + प + त + ा    -> पिता     (label "पिता", needs reorder)
    '<0G#!'    -> स + ... + ख + ्य + ा         (label "संख्या")

USAGE
    python sec_decoder.py --pdf roll.pdf --verify      # sanity-check the table
    python sec_decoder.py --pdf roll.pdf --dump-codes  # list unmapped codes
"""

import re
import sys

# ---------------------------------------------------------------------------
# GLYPH TABLE
#
# Read off the embedded font by rendering codes 33..142 (see glyph_chart PNGs).
# Entries marked  # CHECK  are ones I could not read with certainty from the
# raster — open glyph_chart_1.png / glyph_chart_2.png, look up the code number,
# and correct the value. Everything unmarked was verified against real words.
# ---------------------------------------------------------------------------

GLYPH_MAP = {
    33: "ा",
    34: "ज",        # CHECK  (could be उ / a ja-variant)
    35: "य",
    36: "ि",
    37: "न",
    38: "व",
    39: "र्",       # CHECK  reph (hook above) — confirm it is reph, not a vowel
    40: "च",
    41: "आ",
    42: "ो",
    43: "ग",
    44: "ज",
    45: "स्",       # CHECK  half-sa
    46: "थ",
    47: "प",
    48: "़",        # CHECK  nukta vs. a low dot
    49: "त",
    50: "ु",        # CHECK  u-matra (hook below)
    51: "क",
    52: "म",
    53: "ल",
    54: "ी",
    55: "ि",        # CHECK  wide i-matra variant
    56: "ष",
    57: "द",
    58: "ठ",        # CHECK
    59: "ं",        # CHECK  anusvara vs. digit zero
    60: "स",
    61: "क्ष",
    62: "",         # CHECK  mark above — unidentified
    63: "त्र",
    64: "ब",
    65: "ड",
    66: "क्र",
    67: "ग्र",
    68: "ध",
    69: "भ",
    70: "ी",        # CHECK  wide i-matra variant
    71: "छ",        # CHECK
    72: "ए",
    73: "न्",       # CHECK  half-na
    74: "ह",
    75: "ण",
    76: "प्र",
    77: "अ",
    78: "ि",        # CHECK  wide i-matra variant
    79: "श",
    80: "",         # CHECK  mark above — unidentified
    81: "ओं",       # CHECK
    82: "म्",       # CHECK  half-ma
    83: "रू",
    84: "स्त्र",
    85: "त्",       # CHECK  half-ta
    86: "ु",        # CHECK  mark below
    87: "द्य",
    88: "",         # CHECK  mark — unidentified
    89: "िं",       # CHECK
    90: "घ",
    91: "क्",       # CHECK  half-ka
    92: "ष्ट",      # CHECK
    93: "ढ",
    94: "छ",
    95: "ट",
    96: "",         # CHECK
    97: "",         # CHECK
    98: "",         # CHECK
    99: "ख",
    100: "द्र",
    101: "फ",
    102: "द्द",
    103: "हु",
    104: "ऊ",
    105: "क़",
    106: "",        # CHECK
    107: "ज्ञ",
    108: "न्न",
    109: "ज़",
    110: "इ",
    111: "त्त",
    112: "ई",
    113: "फ़",
    114: "ौ",
    115: "ल्",      # CHECK
    116: "",        # dotted-circle placeholder glyph — emits nothing
    117: "़",       # CHECK
    118: "ठ",
    119: "द्",      # CHECK
    120: "ॉ",       # CHECK
    121: "स्त",
    122: "ड़",
    123: "",
    124: "उ",
    125: "र",
    126: "",        # CHECK
    127: "",        # CHECK
    128: "िं",      # CHECK  wide variant
    129: "ढ़",
    130: "फ्र",     # CHECK
    131: "रु",
    132: "श्र",
    133: "",
    134: "च्",      # CHECK
    135: "क्त",
    136: "ऐ",
    137: "श्व",
    138: "द्व",
    139: "ाँ",      # CHECK
    140: "ट्",      # CHECK
    141: "ऋ",
    142: "ब्र",
}

I_MATRAS = ("ि", "िं")
CONSONANTS = set("कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह")


def decode_bytes(raw):
    """Map legacy byte codes to Devanagari, then fix visual->logical order."""
    out = []
    for b in raw:
        if 32 <= b <= 142:
            out.append(GLYPH_MAP.get(b, " " if b == 32 else "\ufffd"))
        else:
            out.append(chr(b))
    s = "".join(out)
    return _reorder(s)


def _reorder(s):
    """This font stores the i-matra BEFORE its consonant (visual order) and
    the reph as a standalone mark. Move both into Unicode logical order."""
    # i-matra: <matra><consonant cluster> -> <consonant cluster><matra>
    for m in I_MATRAS:
        s = re.sub(
            r"(" + re.escape(m) + r")([" + "".join(CONSONANTS) + r"](?:्[" + "".join(CONSONANTS) + r"])*)",
            r"\2\1",
            s,
        )
    # reph: <consonant>र् -> र्<consonant>
    s = re.sub(r"([" + "".join(CONSONANTS) + r"])र्", r"र्\1", s)
    return re.sub(r"\s+", " ", s).strip()


def iter_raw_strings(pdf_path, page_no=None):
    """Yield (page_index, font_tag, raw_bytes) for every text-show op."""
    import pymupdf

    doc = pymupdf.open(pdf_path)
    pages = [page_no] if page_no is not None else range(len(doc))
    for pi in pages:
        page = doc[pi]
        streams = [doc.xref_stream(x[0]) for x in page.get_xobjects()]
        if not streams:
            streams = [page.read_contents()]
        for s in streams:
            font = None
            for t in re.finditer(rb"/(\w+)\s+[\d.]+\s+Tf|\((?:\\.|[^\\)])*\)\s*Tj", s):
                g = t.group(0)
                if g.endswith(b"Tf"):
                    font = re.match(rb"/(\w+)", g).group(1).decode()
                else:
                    raw = g[1 : g.rindex(b")")]
                    raw = re.sub(rb"\\([()\\])", rb"\1", raw)
                    yield pi, font, raw


# Labels that appear on every card — used as a self-test of the glyph table.
EXPECTED_LABELS = {
    "%!4": "नाम",
    "43!%": "मकान",
    "$/1!": "पिता",
    "3!": "का",
}


def verify(pdf_path):
    ok = fail = 0
    for probe, expect in EXPECTED_LABELS.items():
        got = decode_bytes(probe.encode("latin-1"))
        mark = "OK  " if got == expect else "FAIL"
        if got == expect:
            ok += 1
        else:
            fail += 1
        print(f"{mark} {probe!r:12} -> {got!r:20} expected {expect!r}")
    print(f"\n{ok} passed, {fail} failed")
    return fail == 0


def dump_codes(pdf_path):
    """Report which byte codes actually occur in this PDF and which of those
    are still unmapped — tells you exactly how much of the table matters."""
    seen = {}
    for _, font, raw in iter_raw_strings(pdf_path):
        for b in raw:
            seen[b] = seen.get(b, 0) + 1
    unmapped = [(c, n) for c, n in seen.items() if 33 <= c <= 142 and not GLYPH_MAP.get(c)]
    print(f"{len(seen)} distinct byte codes used in this PDF")
    print(f"{len(unmapped)} of them are unmapped/blank in GLYPH_MAP:\n")
    for c, n in sorted(unmapped, key=lambda t: -t[1]):
        print(f"  code {c:3}  {n:6} occurrences   <-- fill this in from the glyph chart")


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--verify", action="store_true")
    ap.add_argument("--dump-codes", action="store_true")
    ap.add_argument("--page", type=int)
    args = ap.parse_args()

    if args.verify:
        sys.exit(0 if verify(args.pdf) else 1)
    if args.dump_codes:
        dump_codes(args.pdf)
        sys.exit(0)

    for pi, font, raw in iter_raw_strings(args.pdf, args.page):
        txt = decode_bytes(raw)
        if txt.strip():
            print(f"p{pi} /{font} {txt}")
