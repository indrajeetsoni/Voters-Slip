# -*- coding: utf-8 -*-
"""
sec_fontmap.py — build a code->Unicode table for ANY Devanagari font subset
found in a Rajasthan SEC roll PDF, automatically.

WHY THIS IS NEEDED
------------------
A single roll PDF contains several subsets of the same legacy Devanagari
font, and each subset uses a DIFFERENT code->glyph assignment. The main
elector pages and the supplement (परिवर्धन / विलोपन) pages do not agree:
code 59 is ॰ on one and म on the other. So one hard-coded table is not
enough, and hand-writing a table per subset would not survive the next
ward's PDF.

HOW IT WORKS
------------
Every subset is cut from the same original font, so the same character has
byte-identical outlines everywhere. We therefore:

  1. take the reference subset, whose table we verified by hand
     (sec_font.GLYPH), and hash each of its glyph outlines
     ->  {outline_hash: "ा"}, {outline_hash: "े"}, ...
  2. for any other subset, hash its glyphs and look the hash up

The result is a correct table for a font nobody has ever seen, with no
manual work. Glyphs whose hash is unknown are reported, not guessed.
"""

import hashlib
import re

from fontTools.ttLib import TTFont

import sec_font

# When True, an unmapped glyph shows up as «code» instead of vanishing.
# Used by --audit so a missing character can be pinned to an exact code.
MARK_UNMAPPED = False
MARKER = "\u00ab%d\u00bb"

_CACHE = {}
_REF_BY_HASH = None


def _font_file(doc, font_xref):
    """Extract the embedded TrueType program for a font, or None."""
    obj = doc.xref_object(font_xref)
    m = re.search(r"/FontDescriptor (\d+) 0 R", obj)
    if not m:
        return None
    desc = doc.xref_object(int(m.group(1)))
    m = re.search(r"/FontFile2 (\d+) 0 R", desc)
    if not m:
        return None
    return doc.xref_stream(int(m.group(1)))


def _outline_hashes(ttf_bytes):
    """{code: hash of that code's glyph outline} for codes 32..255."""
    import io

    f = TTFont(io.BytesIO(ttf_bytes))
    cmap = {}
    for t in f["cmap"].tables:
        cmap.update(t.cmap)
    glyf = f["glyf"] if "glyf" in f else None
    if glyf is None:
        return {}

    out = {}
    for code in range(32, 256):
        name = cmap.get(code) or cmap.get(0xF000 + code)
        if not name or name not in glyf:
            continue
        g = glyf[name]
        h = hashlib.sha1()
        try:
            if g.numberOfContours == 0:
                h.update(b"empty")
            else:
                g.expand(glyf)
                h.update(bytes(g.flags))
                h.update(str(list(g.coordinates)).encode())
                h.update(str(list(g.endPtsOfContours)).encode())
        except Exception:
            continue
        out[code] = h.hexdigest()
    return out


def _reference():
    """{outline_hash: unicode} built from the hand-verified reference table."""
    global _REF_BY_HASH
    if _REF_BY_HASH is not None:
        return _REF_BY_HASH
    _REF_BY_HASH = {}
    ref = sec_font.REFERENCE_FONT_BYTES
    if not ref:
        return _REF_BY_HASH
    # a hash claimed by two codes with different meanings is ambiguous and
    # must not be used to map an unknown subset
    seen = {}
    for code, h in _outline_hashes(ref).items():
        u = sec_font.GLYPH.get(code)
        if not u:
            continue
        if h in seen and seen[h] != u:
            seen[h] = None
        else:
            seen.setdefault(h, u)
    _REF_BY_HASH = {h: u for h, u in seen.items() if u}
    return _REF_BY_HASH


def table_for(doc, font_xref):
    """Return (table, unknown_codes) for a font in this document."""
    key = (id(doc), font_xref)
    if key in _CACHE:
        return _CACHE[key]

    ttf = _font_file(doc, font_xref)
    if not ttf:
        _CACHE[key] = ({}, [])
        return _CACHE[key]

    # the reference subset itself: use the verified table as-is
    if sec_font.REFERENCE_FONT_BYTES and \
            hashlib.sha1(ttf).hexdigest() == \
            hashlib.sha1(sec_font.REFERENCE_FONT_BYTES).hexdigest():
        tbl = {c: u for c, u in sec_font.GLYPH.items() if u}
        # codes left blank in GLYPH may still be covered by an override
        extra = getattr(sec_font, "EXTRA_BY_HASH", {})
        if extra:
            unmapped = []
            for code, h in _outline_hashes(ttf).items():
                if code in tbl:
                    continue
                if h[:12] in extra:
                    tbl[code] = extra[h[:12]]
                else:
                    unmapped.append(code)
            _CACHE[key] = (tbl, unmapped)
            return _CACHE[key]
        _CACHE[key] = (tbl, [])
        return _CACHE[key]

    ref = _reference()
    extra = getattr(sec_font, "EXTRA_BY_HASH", {})
    hashes = _outline_hashes(ttf)
    table, unknown = {}, []
    for code, h in hashes.items():
        if h in ref:
            table[code] = ref[h]
        elif h[:12] in extra:
            table[code] = extra[h[:12]]
        else:
            unknown.append(code)

    # a font is only Devanagari if most of its glyphs resolved to Devanagari
    _CACHE[key] = (table, unknown)
    return _CACHE[key]


def hash_of(doc, font_xref, code):
    """Outline hash for one code in one font — the key used by EXTRA_BY_HASH."""
    ttf = _font_file(doc, font_xref)
    if not ttf:
        return None
    return _outline_hashes(ttf).get(code)


def render_unmapped(doc, wanted, out_png):
    """Draw the unmapped glyphs so a human can read them.

    `wanted` is [(hash, font_xref, code, sample_text), ...].
    """
    import io
    import pymupdf
    from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

    made = []
    for h, xref, code, sample in wanted:
        ttf = _font_file(doc, xref)
        if not ttf:
            continue
        f = TTFont(io.BytesIO(ttf))
        old = {}
        for t in f["cmap"].tables:
            old.update(t.cmap)
        new = dict(old)
        for k, v in old.items():
            if 0xF000 <= k <= 0xF0FF:
                new[k - 0xF000] = v
        sub = CmapSubtable.newSubtable(4)
        sub.platformID, sub.platEncID, sub.language = 3, 1, 0
        sub.cmap = new
        f["cmap"].tables = [sub]
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".ttf", delete=False)
        f.save(tmp.name)
        tmp.close()
        made.append((h, code, tmp.name, sample))

    if not made:
        return None

    out = pymupdf.open()
    per = 16
    for start in range(0, len(made), per):
        chunk = made[start:start + per]
        pg = out.new_page(width=760, height=60 + 48 * len(chunk))
        for row, (h, code, fontpath, sample) in enumerate(chunk):
            y = 45 + row * 48
            name = "G%d_%s" % (row, h[:6])
            pg.insert_font(fontname=name, fontfile=fontpath)
            pg.insert_text((30, y), h[:12], fontsize=10)
            pg.insert_text((120, y), "code %d" % code, fontsize=10)
            pg.insert_text((190, y), bytes([code]).decode("latin-1"),
                           fontname=name, fontsize=30)
            pg.insert_text((250, y), sample[:60], fontsize=9)
    out.save(out_png)
    return out_png


def is_devanagari(table):
    if len(table) < 40:
        return False
    joined = "".join(table.values())
    dev = sum(1 for ch in joined if 0x0900 <= ord(ch) <= 0x097F)
    return dev / max(1, len(joined)) > 0.8


def decode_with(table, raw: bytes) -> str:
    # codes 32..142 belong to the legacy encoding; if one is unmapped, drop
    # it rather than emitting a stray Latin letter into a voter's name
    out = []
    for b in raw:
        if b in table:
            out.append(table[b])
        elif 32 <= b <= 142:
            out.append(MARKER % b if MARK_UNMAPPED else "")
        else:
            out.append(chr(b))
    return sec_font.reorder("".join(out))
