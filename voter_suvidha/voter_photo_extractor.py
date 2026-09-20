# -*- coding: utf-8 -*-
"""
voter_photo_extractor.py — pulls each voter's photo out of a "WithPhoto"
SEC electoral roll PDF and maps it to that voter's EPIC number (or, for
supplement / newly-added entries that carry no EPIC yet, to their serial
number) so it can be stitched onto the generated voter slip.

WHY THIS WORKS WITHOUT THE GLYPH TABLE
---------------------------------------
Names and relative-names on these rolls come from a corrupt custom font and
need sec_font.py's glyph table to read correctly (see extractor.py). EPIC
numbers and serial numbers, however, are set in a normal Latin/ASCII font
and come back correct straight from PyMuPDF's ordinary text layer — so this
module never touches the glyph table at all, and stays independent of it.

MATCHING STRATEGY
------------------
Each voter card prints its serial number and EPIC on one line, with the
photo sitting slightly below and to the right of that line, inside the
same card. For every embedded image on a page:
  1. Collect every text word on the page that looks like an EPIC (an
     uppercase alnum token, e.g. "NQJ1154392" or "RJ/21/159/160087") or a
     plain serial number.
  2. Keep only the ones that sit to the left of the image and roughly on
     the same row (small y-gap).
  3. Pick whichever candidate is closest horizontally. EPIC candidates are
     preferred over bare-serial ones because EPIC is unique and stable;
     bare serial numbers are only used as a fallback for freshly-added
     voters on a "परिवर्धन सूची" supplement page, who have no EPIC printed.

Photos smaller than a passport-photo box / bigger than the one full-page
ward-map illustration are filtered out by size, so the map never gets
mistaken for a voter photo.
"""

import os
import re

import pymupdf

_EPIC_RE = re.compile(r'^[A-Z]{2,4}[A-Z0-9/]{4,20}$')
_SERIAL_RE = re.compile(r'^\d{1,4}$')

# A voter photo on these rolls is roughly 42.5 x 56.9 pt. Give that some
# headroom for other layouts while still rejecting the one big ward-map
# illustration (which runs to hundreds of points) that ships in every PDF.
_MIN_W, _MAX_W = 15, 100
_MIN_H, _MAX_H = 20, 120
_ROW_Y_TOLERANCE = 20   # pt: how far above/below the image top a label can sit
_MAX_LEFT_GAP = 160     # pt: how far left of the image a label can sit


def _closest_labels(words, ix0, iy0):
    """Return (epic_text_or_None, serial_text_or_None) best matching this image."""
    best_epic, best_epic_gap = None, 1e9
    best_serial, best_serial_gap = None, 1e9
    for w in words:
        wx0, wy0, wx1, wy1, text = w[0], w[1], w[2], w[3], w[4]
        if wx1 > ix0 + 5 or abs(wy0 - iy0) > _ROW_Y_TOLERANCE:
            continue
        gap = ix0 - wx1
        if gap < 0 or gap > _MAX_LEFT_GAP:
            continue
        if _EPIC_RE.match(text):
            if gap < best_epic_gap:
                best_epic_gap, best_epic = gap, text
        elif _SERIAL_RE.match(text):
            if gap < best_serial_gap:
                best_serial_gap, best_serial = gap, text
    return best_epic, best_serial


def _safe_filename(label, ext):
    safe = re.sub(r'[^A-Za-z0-9_-]+', '_', label).strip('_') or "photo"
    return f"{safe}.{ext}"


def extract_voter_photos(pdf_path, out_dir):
    """Pull every voter photo out of pdf_path and save it under out_dir.

    Returns (by_epic, by_serial): two dicts mapping EPIC / serial-number
    strings to the saved photo's absolute file path. by_epic should always
    be checked first since EPIC is unique across the whole roll; by_serial
    is only a fallback for voters with no EPIC printed (fresh additions on
    a supplement page), and serial numbers are only unique within one
    uploaded PDF.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    os.makedirs(out_dir, exist_ok=True)
    by_epic, by_serial = {}, {}

    doc = pymupdf.open(pdf_path)
    try:
        for page in doc:
            words = page.get_text("words")
            for im in page.get_image_info(xrefs=True):
                x0, y0, x1, y1 = im["bbox"]
                w, h = x1 - x0, y1 - y0
                if not (_MIN_W <= w <= _MAX_W and _MIN_H <= h <= _MAX_H):
                    continue  # not a voter-photo-sized image (e.g. the ward map)

                epic, serial = _closest_labels(words, x0, y0)
                if not epic and not serial:
                    continue

                try:
                    img_info = doc.extract_image(im["xref"])
                except Exception:
                    continue

                ext = img_info.get("ext", "png")
                label = epic or f"serial_{serial}"
                out_path = os.path.join(out_dir, _safe_filename(label, ext))
                try:
                    with open(out_path, "wb") as f:
                        f.write(img_info["image"])
                except Exception:
                    continue

                if epic:
                    by_epic[epic] = out_path
                if serial:
                    # Don't let an unrelated later card steal a serial number
                    # that was already resolved via a proper EPIC match.
                    by_serial.setdefault(serial, out_path)
    finally:
        doc.close()

    return by_epic, by_serial


if __name__ == "__main__":
    import argparse
    import json

    ap = argparse.ArgumentParser(description="Extract voter photos from a WithPhoto roll PDF")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    by_epic, by_serial = extract_voter_photos(args.pdf, args.out_dir)
    print(json.dumps({
        "photosByEpic": len(by_epic),
        "photosBySerial": len(by_serial),
    }, ensure_ascii=False, indent=2))
