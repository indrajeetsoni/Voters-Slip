# -*- coding: utf-8 -*-
"""
sec_extractor.py — accurate elector extraction for Rajasthan SEC roll PDFs.

Reads the raw glyph codes straight out of the page content stream, together
with their x/y positions, and rebuilds each voter card. The PDF's text layer
is never used, because its /ToUnicode map is corrupt (74 of 111 glyph codes
are wrong — every matra is dropped). See sec_font.py.

    python sec_extractor.py roll.pdf --json out.json --check
"""

import json
import os
import re
import sys

import pymupdf

import sec_font
import sec_fontmap

sec_font.load_reference()   # resolves next to sec_font.py on any OS

# ---------------------------------------------------------------------------
# content-stream tokenizer with text positioning
# ---------------------------------------------------------------------------

_N = rb"-?[\d.]+"
_TOK = re.compile(
    rb"/(?P<font>\w+)\s+[-\d.]+\s+Tf"
    rb"|(?P<tm>" + _N + rb"(?:\s+" + _N + rb"){5})\s+Tm"
    rb"|(?P<td>" + _N + rb"\s+" + _N + rb")\s+(?P<tdop>TD|Td)"
    rb"|(?P<star>T\*)"
    rb"|(?P<tl>" + _N + rb")\s+TL"
    rb"|(?P<bt>BT)"
    rb"|\((?P<s>(?:\\.|[^\\)])*)\)\s*Tj"
)


def _widths(doc, font_xref):
    """{code: width/1000} from the font's /Widths array, for x advance."""
    obj = doc.xref_object(font_xref)
    fc = re.search(r"/FirstChar (\d+)", obj)
    wr = re.search(r"/Widths (\d+) 0 R", obj)
    if not (fc and wr):
        return {}
    arr = doc.xref_object(int(wr.group(1)))
    nums = [float(n) for n in re.findall(r"-?[\d.]+", arr)]
    first = int(fc.group(1))
    return {first + i: n / 1000.0 for i, n in enumerate(nums)}


def page_items(doc, pno):
    """Every text-show op on a page as {x, y, font, raw, i} in stream order.

    x is advanced by the real glyph widths. This matters: names are drawn as
    several fragments (लाद + ू + राम) and without correct advance the matra
    sorts to the wrong end of the word.
    """
    page = doc[pno]
    xo = page.get_xobjects()
    if xo:
        streams = [doc.xref_stream(x[0]) for x in xo]
        maps = []
        for x in xo:
            o = doc.xref_object(x[0])
            maps.append(
                dict(re.findall(r"/(\w+)\s+(\d+) 0 R", o.split("/Font")[1]))
                if "/Font" in o else {}
            )
    else:
        streams, maps = [page.read_contents()], [{}]

    out, i = [], 0
    wcache = {}
    for stream, fonts in zip(streams, maps):
        font = None
        size = 10.0
        tx = ty = lx = ly = 0.0
        leading = 0.0
        for m in _TOK.finditer(stream):
            g = m.groupdict()
            if g["font"]:
                font = g["font"].decode()
                size = float(re.search(rb"([-\d.]+)\s+Tf", m.group(0)).group(1))
            elif g["tm"]:
                v = [float(n) for n in g["tm"].split()]
                tx = lx = v[4]
                ty = ly = v[5]
            elif g["td"]:
                dx, dy = [float(n) for n in g["td"].split()]
                lx += dx
                ly += dy
                tx, ty = lx, ly
                if g["tdop"] == b"TD":
                    leading = -dy
            elif g["star"]:
                ly -= leading
                tx, ty = lx, ly
            elif g["tl"]:
                leading = float(g["tl"])
            elif g["bt"] is not None:
                tx = ty = lx = ly = 0.0
            elif g["s"] is not None:
                raw = re.sub(rb"\\([()\\])", rb"\1", g["s"])
                out.append(dict(x=tx, y=ty, font=font, raw=raw, i=i, fonts=fonts))
                i += 1
                if font in fonts:
                    key = fonts[font]
                    if key not in wcache:
                        wcache[key] = _widths(doc, int(key))
                    w = wcache[key]
                    tx += sum(w.get(b, 0.5) for b in raw) * size
    return out


def _bold_fonts(doc, items):
    """The serial number of each card is drawn in the bold font."""
    tags = set()
    for it in items:
        f, fonts = it["font"], it["fonts"]
        if f in tags or f not in fonts:
            continue
        if "Bold" in doc.xref_object(int(fonts[f])):
            tags.add(f)
    return tags


def _dev_fonts(doc, items):
    """{font tag: glyph table} for the Devanagari subsets used on this page.

    Each subset has its own code assignment, so the table is rebuilt per
    font by matching glyph outlines against the reference (sec_fontmap).
    """
    tabs = {}
    for it in items:
        f, fonts = it["font"], it["fonts"]
        if f in tabs or f not in fonts:
            continue
        table, _ = sec_fontmap.table_for(doc, int(fonts[f]))
        if sec_fontmap.is_devanagari(table):
            tabs[f] = table
    return tabs


def text_of(it, devtags):
    table = devtags.get(it["font"]) if isinstance(devtags, dict) else None
    if table:
        return sec_fontmap.decode_with(table, it["raw"])
    if isinstance(devtags, (set, frozenset)) and it["font"] in devtags:
        return sec_font.decode(it["raw"])
    return it["raw"].decode("latin-1")


# ---------------------------------------------------------------------------
# card assembly
# ---------------------------------------------------------------------------

# serial anchors are drawn in the bold font as " 12 " or "E 49 "
_SERIAL = re.compile(r"^\s*(?P<mark>[ESR])?\s*(?P<sn>\d{1,4})\s*$")

# drawn inside every card box; must not end up inside a name
_PHOTO_TEXT = {"Photo is", "Available", "Photo", "is"}

CARD_W = 165.0          # a card is ~165pt wide, 3 per row
CARD_H = 62.0           # and ~62pt tall


def rows_of(items, devtags):
    """Group items into visual rows, then order each row left-to-right.

    Within a row, x must be rounded before falling back to stream order:
    a name like लाद + ू + राम is drawn as three ops at almost the same x,
    and only stream order puts the matra in the right place.
    """
    clusters = []
    for it in sorted(items, key=lambda t: -t["y"]):
        for c in clusters:
            if abs(c[0]["y"] - it["y"]) <= 2.5:
                c.append(it)
                break
        else:
            clusters.append([it])
    out = {}
    for c in clusters:
        c.sort(key=lambda t: (round(t["x"]), t["i"]))
        out[round(c[0]["y"], 1)] = c
    return out


def _value_after_label(row, devtags, labels):
    """Row text is 'label : value'; return the part after the last colon."""
    txt = "".join(text_of(it, devtags) for it in row)
    for lab in labels:
        if lab in txt:
            tail = txt.split(":", 1)[1] if ":" in txt else ""
            return re.sub(r"\s+", " ", tail).strip()
    return ""


def cards_on_page(doc, pno):
    items = page_items(doc, pno)
    devtags = _dev_fonts(doc, items)
    page_text = "".join(text_of(it, devtags) for it in items)

    bold = _bold_fonts(doc, items)
    bolds = [it for it in items if it["font"] in bold]
    marks = [it for it in bolds
             if it["raw"].decode("latin-1").strip() in ("E", "S", "R")]

    anchors = []
    for it in bolds:
        m = _SERIAL.match(it["raw"].decode("latin-1"))
        if not m or int(m.group("sn")) < 1:
            continue
        mark = m.group("mark")
        if not mark:
            for mk in marks:
                if abs(mk["y"] - it["y"]) < 3 and -14 < mk["x"] - it["x"] < 0:
                    mark = mk["raw"].decode("latin-1").strip()
                    break
        anchors.append((int(m.group("sn")), mark, it))

    out = []
    for sn, mark, anc in anchors:
        x0, y0 = anc["x"] - 14, anc["y"]
        region = [
            it for it in items
            if x0 <= it["x"] < x0 + CARD_W and y0 - CARD_H <= it["y"] <= y0 + 2
            and it["raw"].decode("latin-1").strip() not in _PHOTO_TEXT
        ]
        if len(region) < 6:
            continue
        rows = rows_of(region, devtags)
        card = dict(SerialNo=sn, Mark=mark, VoterName="", RelativeName="",
                    RelativeType="पिता", HouseNo="", Age="", Gender="", EPIC="")
        for k in sorted(rows, reverse=True):
            row = rows[k]
            txt = "".join(text_of(it, devtags) for it in row)
            flat = re.sub(r"\s+", " ", txt).strip()
            if flat.startswith("नाम"):
                card["VoterName"] = _value_after_label(row, devtags, ["नाम"])
            elif flat.startswith("पिता") or flat.startswith("पति") or flat.startswith("अन्य"):
                card["RelativeType"] = ("पति" if flat.startswith("पति")
                                        else "अन्य" if flat.startswith("अन्य") else "पिता")
                card["RelativeName"] = _value_after_label(
                    row, devtags, ["पिता", "पति", "अन्य"])
            elif flat.startswith("मकान"):
                card["HouseNo"] = _value_after_label(row, devtags, ["मकान"])
            elif flat.startswith("आयु"):
                m = re.search(r"(\d{1,3})", flat)
                card["Age"] = m.group(1) if m else ""
                card["Gender"] = ("स्त्री" if "स्त्री" in flat
                                  else "तृतीय लिंग" if "तृतीय" in flat
                                  else "पुरूष" if "पुरूष" in flat else "")
            else:
                e = re.search(r"\b((?:RJ/[\d/]+)|(?:[A-Z]{3}\d{7}))\b", flat)
                if e and not card["EPIC"]:
                    card["EPIC"] = e.group(1)
        if not card["VoterName"]:
            continue          # summary-table / page-number digits, not a card
        out.append(card)
    return out, page_text


# ---------------------------------------------------------------------------
# whole-roll extraction
# ---------------------------------------------------------------------------

def audit(pdf_path):
    """Report every glyph code this PDF uses that the table cannot decode.

    Run this whenever a name looks wrong or has a character missing:
        python sec_extractor.py roll.pdf --audit
    Send the output and the codes can be added to sec_font.GLYPH.
    """
    from collections import Counter

    sec_fontmap.MARK_UNMAPPED = True
    doc = pymupdf.open(pdf_path)
    hits = Counter()
    samples, where = {}, {}
    for pno in range(len(doc)):
        items = page_items(doc, pno)
        devtags = _dev_fonts(doc, items)
        for it in items:
            table = devtags.get(it["font"])
            if not table:
                continue
            missing = [b for b in it["raw"] if 32 <= b <= 142 and b not in table]
            if not missing:
                continue
            txt = text_of(it, devtags)
            xref = int(it["fonts"][it["font"]])
            for b in set(missing):
                h = sec_fontmap.hash_of(doc, xref, b)
                if not h:
                    continue
                hits[h] += 1
                samples.setdefault(h, []).append(txt)
                where.setdefault(h, (xref, b))

    if not hits:
        print("No unmapped glyphs. Every character in this PDF decodes.")
        return hits

    print(f"{len(hits)} unmapped glyph(s) in {pdf_path.split('/')[-1]}:\n")
    wanted = []
    for h, n in hits.most_common():
        xref, code = where[h]
        ex = " , ".join(dict.fromkeys(samples[h]))[:100]
        print(f"  {h[:12]}  used {n:5} times   e.g. {ex}")
        wanted.append((h, xref, code, ex))

    png = sec_fontmap.render_unmapped(doc, wanted, "unmapped_glyphs.png")
    if png:
        print(f"\nWrote {png} — open it. Each row shows the missing character")
        print("next to its id. Send me that image, or add the characters yourself")
        print("to EXTRA_BY_HASH in sec_font.py as  \"<id>\": \"<char>\",")
    return hits


def _score(card):
    """How complete a card is — used when a serial appears more than once."""
    return sum(bool(card[k]) for k in
               ("VoterName", "RelativeName", "HouseNo", "Age", "Gender", "EPIC"))


def extract(pdf_path):
    doc = pymupdf.open(pdf_path)

    # --- header facts from page 1
    it0 = page_items(doc, 0)
    dev0 = _dev_fonts(doc, it0)
    rows0 = rows_of(it0, dev0)
    ward = gram = booth = ""
    for k in sorted(rows0, reverse=True):
        line = re.sub(r"\s+", " ",
                      "".join(text_of(i, dev0) for i in rows0[k])).strip()
        m = re.search(r"क्रमांक\s*:\s*(\d+)", line)
        if m:
            ward = m.group(1)
        m = re.search(r"ग्रामपंचायत\s*:\s*(.+?)(?:वार्ड|क्रमांक|$)", line)
        if m and m.group(1).strip():
            gram = m.group(1).strip()
        if "मतदान" in line and "पता" in line:
            booth = line.split(":", 1)[-1].strip()
    if not booth:
        for k in sorted(rows0, reverse=True):
            line = re.sub(r"\s+", " ",
                          "".join(text_of(i, dev0) for i in rows0[k])).strip()
            if re.match(r"^\d+\s*-", line):
                booth = line
                break

    # --- cards, page by page
    active, deleted, added = {}, {}, {}
    for pno in range(len(doc)):
        cards, ptext = cards_on_page(doc, pno)
        is_add_page = "परिवर्धन" in ptext
        for c in cards:
            sn = c["SerialNo"]
            if c["Mark"]:
                deleted[sn] = c
            else:
                bucket = added if is_add_page else active
                old = bucket.get(sn)
                if old is None or _score(c) > _score(old):
                    bucket[sn] = c

    for sn in deleted:
        active.pop(sn, None)
        added.pop(sn, None)
    active.update(added)

    voters = []
    for sn in sorted(active):
        c = active[sn]
        voters.append({
            "Ward": ward or "1",
            "Part": "1",
            "Booth": booth,
            "SerialNo": sn,
            "VoterName": c["VoterName"],
            "RelativeType": c["RelativeType"],
            "RelativeName": c["RelativeName"],
            "HouseNo": c["HouseNo"] or "-",
            "Age": c["Age"],
            "Gender": c["Gender"] or "पुरूष",
            "EPIC": c["EPIC"],
            "GramPanchayat": gram,
        })

    return {
        "fileName": pdf_path.split("/")[-1],
        "ward": ward or "1",
        "part": "1",
        "booth": booth,
        "gramPanchayat": gram,
        "totalSerials": max(active) if active else 0,
        "deletedCount": len(deleted),
        "activeCount": len(voters),
        "voters": voters,
    }


# ---------------------------------------------------------------------------
# self-check against the roll's own summary table
# ---------------------------------------------------------------------------

def check(data):
    v = data["voters"]
    male = sum(1 for x in v if x["Gender"] == "पुरूष")
    female = sum(1 for x in v if x["Gender"] == "स्त्री")
    other = len(v) - male - female
    print(f"ward            : {data['ward']}")
    print(f"gram panchayat  : {data['gramPanchayat']}")
    print(f"booth           : {data['booth']}")
    print(f"active electors : {len(v)}")
    print(f"  male          : {male}")
    print(f"  female        : {female}")
    print(f"  other         : {other}")
    print(f"deleted         : {data['deletedCount']}")
    blank = [x for x in v if not x["VoterName"] or not x["Age"]]
    print(f"incomplete rows : {len(blank)}")
    for x in blank[:10]:
        print("   ", x["SerialNo"], repr(x["VoterName"]), x["Age"])
    return len(v), male, female


if __name__ == "__main__":
    ap = sys.argv[1:]
    pdf = ap[0]
    if "--audit" in ap:
        audit(pdf)
        sys.exit(0)
    data = extract(pdf)
    if "--audit" in ap:
        audit(pdf)
        sys.exit(0)
    if "--check" in ap:
        check(data)
    if "--json" in ap:
        out = ap[ap.index("--json") + 1]
        with open(out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        print("wrote", out)
    if "--show" in ap:
        n = int(ap[ap.index("--show") + 1])
        for x in data["voters"][:n]:
            print(f"{x['SerialNo']:4} {x['VoterName']:28} "
                  f"{x['RelativeType']}: {x['RelativeName']:24} "
                  f"मकान {x['HouseNo']:5} आयु {x['Age']:3} {x['Gender']:7} {x['EPIC']}")
