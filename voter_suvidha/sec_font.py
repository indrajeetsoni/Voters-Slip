# -*- coding: utf-8 -*-
"""
sec_font.py — decode the legacy Devanagari font used in Rajasthan SEC
electoral roll PDFs into correct Unicode.

WHY: the PDF's own /ToUnicode CMap is corrupt (many glyph codes collapse onto
one codepoint), so pdftotext / PyMuPDF text extraction is LOSSY and cannot be
repaired afterwards. We read the raw glyph codes from the content stream and
decode them with the table below, which was derived by rendering the embedded
font and cross-checking against hundreds of known words in the roll.

THE SINGLE BIGGEST FIX: code 32 is NOT a space in this font — it is  र .
That one error alone destroyed most names ("बदनोर" -> "बदनो", "राजस्थान" ->
"ाजस्थान", "सत्यनारायण" -> "सत्यना ायण").
"""

import os
import re

# --- verified against real words in the roll -------------------------------
GLYPH = {
    32: "र",        # NOT a space.  [64,57,37,42,32] = बदनोर
    33: "ा",        # %!4 = नाम
    34: "ज्",       # {34}योित = ज्योति ;  ा{34}य = राज्य
    35: "य",
    36: "ि",        # visual order — reordered below
    37: "न",
    38: "व",
    39: "र्",       # reph.  िनवा{39}चन = निर्वाचन
    40: "च",
    41: "आ",        # [41,35,50] = आयु
    42: "ो",
    43: "ग",
    44: "ज",
    45: "स्",       # सद{45}य = सदस्य ;  ाज{45}थान = राजस्थान
    46: "थ",
    47: "प",
    48: "ं",        # <0G#! = संख्या ;  मा{48}ग{80}िंसह = मांगूसिंह
    49: "त",
    50: "ु",        # [41,35,50] = आयु ;  प{50}{127}पा = पुष्पा
    51: "क",
    52: "म",
    53: "ल",
    54: "ी",        # [84,54] = स्त्री
    55: "ि",        # wide variant.  स{55}(32)ता = सरिता
    56: "ष",
    57: "द",
    58: "ब्",       # {58}याव(32) = ब्यावर
    59: "\u0970",      # जि॰ प॰  (confirmed by chart + ToUnicode)
    60: "स",
    61: "क्ष",      # [61,62,63] = क्षेत्र
    62: "े",
    63: "त्र",
    64: "ब",
    65: "ड",
    66: "क्र",
    67: "ग्र",
    68: "ध",
    69: "भ",
    70: "ी",        # wide variant.  सोलंक{70} = सोलंकी
    71: "ख्",       # [60,48,71,35,33] = संख्या
    72: "ए",
    73: "न्",       # आसी{73}द = आसीन्द
    74: "ह",
    75: "ण",
    76: "प्र",
    77: "अ",
    78: "ि",        # narrow variant
    79: "श",
    80: "ू",        # बान{80} = बानू ;  बाब{80} = बाबू
    81: "ओं",
    82: "म्",       # मोह{82}मद = मोहम्मद
    83: "रू",       # [47,50,83,56] = पुरूष
    84: "स्त्र",
    85: "त्",       # स{85}य = सत्य ;  महा{85}मा = महात्मा
    86: "ं",        # गा{86}धी = गांधी
    87: "द्य",
    88: "ृ",        # प{88}ष्ट = पृष्ठ ;  त{88}तीय = तृतीय
    89: "िं",       # [89,53,43] = लिंग ;  [89,60,74] = सिंह
    90: "घ",
    91: "क्",       # न{91}शा = नक्शा
    92: "ष्ठ",      # पृ{92} = पृष्ठ
    93: "ढ",
    94: "छ",
    95: "ट",
    96: "",         # unresolved — reported by --audit, never guessed
    97: "ै",        # हुस{97}न = हुसैन ;  भ{97}रू = भैरू
    98: "क्ष्",     # ल{98}मण = लक्ष्मण
    99: "ख",
    100: "द्र",
    101: "फ",
    102: "द्द",
    103: "हु",
    104: "ऊ",
    105: "क़",
    106: "ेन्",     # गज{106}द्र = गजेन्द्र ; िवज{106}द्र = विजेन्द्र
    107: "ज्ञ",
    108: "न्न",
    109: "ज़",
    110: "इ",
    111: "त्त",
    112: "ई",
    113: "फ़",
    114: "ौ",
    115: "ल्",      # सु{115}तानखां = सुल्तानखां ; अ{115}लादीन = अल्लादीन
    116: "\u0966",  # abbreviation sign, as in "मो०"
    117: "",        # bare nukta; only appears after the ० abbreviation
    118: "ठ",
    119: "द",       # चां{119}पोल = चांदपोल
    120: "ां",      # ख{120} = खां
    121: "स्त",
    122: "ड़",
    123: "",
    124: "उ",
    125: "र",
    126: "श्",      # {126}याम = श्याम
    127: "ष्",      # व{97}{127}णव = वैष्णव ; प{50}{127}पा = पुष्पा
    128: "िं",
    129: "ढ़",
    130: "फ्र",
    131: "रु",
    132: "श्र",
    133: "",        # rare (1x) — unresolved
    134: "च्",      # क{134}छावा = कच्छावा
    135: "क्त",
    136: "ऐ",
    137: "श्व",
    138: "द्व",
    139: "ां",      # ग{139}धी = गांधी
    140: "प्",      # ाम{140}या(32)ी = रामप्यारी
    141: "ऋ",
    142: "ब्र",
}

# Glyphs missing from the reference subset, keyed by outline hash so a single
# entry works across every font subset and every ward's PDF.
# Add entries from the chart that `--audit` writes.
EXTRA_BY_HASH = {
    # read off the chart written by `sec_extractor.py <pdf> --audit`
    "d5b01ef4c3e1": "प्त",   # संक्षि+प्त = संक्षिप्त
    "091a4549869e": "ग्",    # ग्+यारसी = ग्यारसी
    "d3103facf566": "औ",     # औ+र = और
    "62fd37021226": "",      # tiny mark drawn after the ० abbreviation
    "0cbf1b91af09": "",      # blank glyph
}

UNRESOLVED = {c for c, v in GLYPH.items() if v == "" and c not in (116, 123, 133)}

CONS = "कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह"
_C = "[" + CONS + "]"
_CLUSTER = _C + r"(?:्" + _C + r")*"
_MATRA = r"[\u093e-\u094c\u0902\u0903\u093c]"

# Devanagari fonts of this family store the i-matra and the reph in VISUAL
# order (before the consonant they belong to). Unicode wants logical order.
_RE_IMATRA = re.compile(r"(िं|ि)(" + _CLUSTER + r")")
_RE_REPH = re.compile(r"(" + _CLUSTER + r")(" + _MATRA + r"*)र्")


def reorder(s: str) -> str:
    """Visual order -> Unicode logical order."""
    s = _RE_IMATRA.sub(r"\2\1", s)          # ि + प  ->  प + ि
    s = _RE_REPH.sub(r"र्\1\2", s)          # वा + र्  ->  र् + वा
    return s.strip()


def decode(raw: bytes) -> str:
    """Decode raw glyph codes using the reference table."""
    return reorder("".join(GLYPH.get(b, chr(b)) if b <= 142 else chr(b)
                           for b in raw))


# The reference subset's font program, needed to match other subsets by
# outline. Filled in by load_reference(); see sec_fontmap.
REFERENCE_FONT_BYTES = None


def load_reference(path=None):
    """Load the reference font subset.

    The path is resolved relative to THIS file, not the working directory.
    Getting that wrong is silent and fatal: without the reference font no
    font is recognised as Devanagari, every name comes out empty, and the
    extractor reports zero electors for every PDF.
    """
    global REFERENCE_FONT_BYTES
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "reference_font.ttf")
    with open(path, "rb") as f:            # deliberately not caught
        REFERENCE_FONT_BYTES = f.read()
    return REFERENCE_FONT_BYTES


def has_unresolved(raw: bytes) -> bool:
    return any(b in UNRESOLVED for b in raw)


# --- self-test -------------------------------------------------------------
TESTS = [
    (b"%!4", "नाम"),
    (b"3!", "का"),
    (b"43!%", "मकान"),
    (b"$/1!", "पिता"),
    (b"<0G#!", "संख्या"),
    (b")#2", "आयु"),
    (b"Y5+", "लिंग"),
    (b"Y<J", "सिंह"),
    (b"/2S8", "पुरूष"),
    (b"T6", "स्त्री"),
    (b"@9%* ", "बदनोर"),
    (b"=>?", "क्षेत्र"),
    (b"4*JR49", "मोहम्मद"),
    (b"$%&!'(%", "निर्वाचन"),
]

if __name__ == "__main__":
    ok = 0
    for raw, want in TESTS:
        got = decode(raw)
        flag = "OK  " if got == want else "FAIL"
        ok += got == want
        print(f"{flag} {raw!r:14} -> {got!r:14} want {want!r}")
    print(f"\n{ok}/{len(TESTS)} passed")
