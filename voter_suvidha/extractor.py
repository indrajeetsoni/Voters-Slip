# -*- coding: utf-8 -*-
"""
extractor.py — Voter Suvidha elector extractor (rebuilt).

WHAT CHANGED AND WHY
--------------------
These SEC Rajasthan roll PDFs ship a corrupt /ToUnicode map: 74 of its 111
glyph codes are wrong and every matra is dropped, so `page.get_text()` can
never return correct names. The old version tried to patch the damage with a
dictionary of known words (WORD_MAP / words_dict.json), which is why only a
handful of names printed correctly.

This version ignores the text layer completely. It reads the raw glyph codes
straight from the page content stream, decodes them with the real font table,
and rebuilds each card from its x/y positions. See sec_font.py, sec_fontmap.py
and sec_extractor.py.

Verified against BADNOR-Ward_No-001.pdf, matching the roll's own summary page:
    405 active electors, 204 male, 201 female, 9 deleted.

The public API is unchanged, so server.py needs no edits:
    extract_pdf_elector_data(pdf_path)
    export_voters_to_excel(voters, out_path)
    read_voters_from_excel(excel_path)
"""

import json
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import sec_extractor
import sec_font


def _self_check():
    """Fail loudly at import if the glyph table is not working.

    A broken setup used to show up as "no electors found" on every PDF,
    which looks like a bad upload but is really a missing file.
    """
    for raw, want in (b"%!4", "\u0928\u093e\u092e"), (b"43!%", "\u092e\u0915\u093e\u0928"):
        got = sec_font.decode(raw)
        if got != want:
            raise RuntimeError(
                "Devanagari glyph table is not working (expected %r, got %r). "
                "Check that sec_font.py, sec_fontmap.py, sec_extractor.py and "
                "reference_font.ttf are all present in the voter_suvidha folder."
                % (want, got))


_self_check()


def _part_from_filename(pdf_path):
    """Roll PDFs are one file per भाग; the number is only in the filename."""
    name = os.path.basename(pdf_path)
    m = re.search(r"(?i)part\s*(?:no)?[-_\s]*0*(\d+)", name)
    if m:
        return m.group(1)
    m = re.search(r"(?:भाग)\s*0*(\d+)", name)
    return m.group(1) if m else "1"


def extract_pdf_elector_data(pdf_path):
    """Extract ward, booth and all active electors from a roll PDF."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    data = sec_extractor.extract(pdf_path)

    part = _part_from_filename(pdf_path)
    data["part"] = part
    for v in data["voters"]:
        v["Part"] = part

    # fall back to the filename for the ward if the header scan came up empty
    if not data.get("ward") or data["ward"] == "1":
        m = re.search(r"(?i)ward\s*(?:no)?[-_\s]*0*(\d+)", os.path.basename(pdf_path))
        if m:
            data["ward"] = m.group(1)
            for v in data["voters"]:
                v["Ward"] = m.group(1)

    if data["activeCount"] == 0:
        raise RuntimeError(
            "इस PDF से कोई मतदाता नहीं निकला / No electors extracted from "
            f"{os.path.basename(pdf_path)}.\n"
            "संभावित कारण / likely causes:\n"
            "  1. reference_font.ttf voter_suvidha फोल्डर में नहीं है\n"
            "  2. यह मूल निर्वाचक नामावली नहीं, पहले से बनी वोटर पर्ची PDF है\n"
            "  3. इस PDF का फ़ॉन्ट subset नया है — फाइल भेजें, तालिका जोड़ दी जाएगी\n"
            "जाँच के लिए चलाएं: python sec_extractor.py \"<pdf>\" --check")

    print(f"Extracted ward {data['ward']} part {part}: "
          f"{data['activeCount']} active, {data['deletedCount']} deleted")
    return data


def export_voters_to_excel(voters, out_excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "मतदाता सूची (Master Voter List)"

    # Exact 12 columns requested by user
    headers = [
        "Sr No", "वार्ड संख्या", "भाग संख्या", "मतदान केंद्र की संख्या व पता",
        "क्रम संख्या", "निर्वाचक का नाम", "पिता/पति का नाम",
        "मकान संख्या", "आयु", "लिंग", "EPIC No", "ग्राम पंचायत"
    ]
    ws.append(headers)

    # Soft blue/lavender accent fill (#D9E1F2) as shown in template
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF")
    )

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    ws.row_dimensions[1].height = 26

    for idx, v in enumerate(voters, 1):
        row = [
            idx,
            v.get("Ward", "1"),
            v.get("Part", "1"),
            v.get("Booth", ""),
            v.get("SerialNo", ""),
            v.get("VoterName", ""),
            v.get("RelativeName", ""),
            v.get("HouseNo", "-"),
            v.get("Age", ""),
            v.get("Gender", ""),
            v.get("EPIC", ""),
            v.get("GramPanchayat", "")
        ]
        ws.append(row)

    # Style data rows
    for row_idx in range(2, len(voters) + 2):
        ws.row_dimensions[row_idx].height = 20
        for col_idx in range(1, 13):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            # Left align booth, voter name, relative name; Center align others
            if col_idx in [4, 6, 7]:
                cell.alignment = left_align
            else:
                cell.alignment = center_align

    # Set column widths matching 12-column layout
    col_widths = {
        1: 8,    # Sr No
        2: 12,   # वार्ड संख्या
        3: 12,   # भाग संख्या
        4: 36,   # मतदान केंद्र की संख्या व पता
        5: 12,   # क्रम संख्या
        6: 22,   # निर्वाचक का नाम
        7: 22,   # पिता/पति का नाम
        8: 12,   # मकान संख्या
        9: 8,    # आयु
        10: 10,  # लिंग
        11: 18,  # EPIC No
        12: 18   # ग्राम पंचायत
    }
    for col_idx, w in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w

    os.makedirs(os.path.dirname(os.path.abspath(out_excel_path)), exist_ok=True)
    wb.save(out_excel_path)
    wb.close()
    return out_excel_path

def read_voters_from_excel(excel_path):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    # Detect header mapping from row 1
    header_row = [cell.value for cell in ws[1]]
    col_map = {}
    for idx, h in enumerate(header_row):
        if not h:
            continue
        h_clean = re.sub(r'[\s_]+', '', str(h)).lower()
        if any(k in h_clean for k in ["srno", "sr", "क्रमांक"]) and "क्रमसंख्या" not in h_clean:
            col_map["sr_no"] = idx
        elif any(k in h_clean for k in ["वार्डसंख्या", "वार्ड"]):
            col_map["ward"] = idx
        elif any(k in h_clean for k in ["भागसंख्या", "भाग"]):
            col_map["part"] = idx
        elif any(k in h_clean for k in ["मतदानकेंद्र", "मतदान", "booth"]):
            col_map["booth"] = idx
        elif any(k in h_clean for k in ["क्रमसंख्या", "serialno", "serial"]):
            col_map["serial_no"] = idx
        elif any(k in h_clean for k in ["निर्वाचककानाम", "मतदाताकानाम", "निर्वाचक", "मतदाता", "votername", "name"]):
            col_map["voter_name"] = idx
        elif any(k in h_clean for k in ["पिता/पतिकाराम", "पिता/पतिकानाम", "पिताकानाम", "पतिकाराम", "रिश्तेदार", "relativename", "father"]):
            col_map["rel_name"] = idx
        elif any(k in h_clean for k in ["मकानसंख्या", "मकान", "houseno"]):
            col_map["house_no"] = idx
        elif any(k in h_clean for k in ["आयु", "उम्र", "age"]):
            col_map["age"] = idx
        elif any(k in h_clean for k in ["लिंग", "gender", "sex"]):
            col_map["gender"] = idx
        elif any(k in h_clean for k in ["epicno", "epic", "पहचानपत्र"]):
            col_map["epic"] = idx
        elif any(k in h_clean for k in ["ग्रामपंचायत", "panchayat", "पंचायत"]):
            col_map["gram_panchayat"] = idx

    voters = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        def get_val(key, default_col):
            col_idx = col_map.get(key, default_col)
            if col_idx is not None and col_idx < len(row) and row[col_idx] is not None:
                val = str(row[col_idx]).strip()
                if val.endswith(".0"):
                    val = val[:-2]
                return val
            return ""

        serial = get_val("serial_no", 4)
        v_name = get_val("voter_name", 5)
        if not v_name and not serial:
            continue

        rel_name = get_val("rel_name", 6)
        gender = get_val("gender", 9)
        # Check if relation is husband based on gender
        rel_type = "पति" if gender in ["स्त्री", "महिला", "F", "Female"] else "पिता"

        v = {
            "Ward": get_val("ward", 1) or "1",
            "Part": get_val("part", 2) or "1",
            "Booth": get_val("booth", 3),
            "SerialNo": int(serial) if serial.isdigit() else serial,
            "VoterName": v_name,
            "RelativeType": rel_type,
            "RelativeName": rel_name,
            "HouseNo": get_val("house_no", 7) or "-",
            "Age": get_val("age", 8),
            "Gender": gender or "पुरुष",
            "EPIC": get_val("epic", 10),
            "GramPanchayat": get_val("gram_panchayat", 11)
        }
        voters.append(v)

    wb.close()
    return voters


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Extract electors from a roll PDF")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out-json")
    ap.add_argument("--out-excel")
    args = ap.parse_args()

    data = extract_pdf_elector_data(args.pdf)

    if args.out_excel:
        export_voters_to_excel(data["voters"], args.out_excel)
        print(f"Exported {len(data['voters'])} voters to Excel: {args.out_excel}")

    if args.out_json:
        with open(args.out_json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Saved JSON: {args.out_json}")
    else:
        print(json.dumps({
            "ward": data["ward"],
            "part": data["part"],
            "booth": data["booth"],
            "gramPanchayat": data["gramPanchayat"],
            "totalSerials": data["totalSerials"],
            "deletedCount": data["deletedCount"],
            "activeCount": data["activeCount"],
            "votersCount": len(data["voters"]),
        }, ensure_ascii=False, indent=2))
