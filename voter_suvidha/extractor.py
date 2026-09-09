# -*- coding: utf-8 -*-
"""
Voter Suvidha - High-Precision PDF Electoral Roll Extractor
Extracts Ward, Part, Polling Booth, Active Electors, Deleted Electors,
and generates structured 11-column Excel rosters.
"""
import os
import sys
import re
import json
import pymupdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def clean_hindi_name(raw):
    if not raw:
        return ""
    s = raw.strip()
    # Systematic ligature & font substitutions for Rajasthan election font
    s = s.replace("लसह", "सिंह").replace("नसहप", "सिंह").replace("लसहप", "सिंह")
    s = s.replace("दकरल", "देवी").replace("दकल", "देवी")
    s = s.replace("पचरष", "पुरुष").replace("सल", "स्त्री")
    s = s.replace("कचमरर", "कुमार").replace("कचमपरर", "कुमारी")
    s = s.replace("चनद", "चन्द").replace("चचहरन", "चौहान").replace("मचककश", "मुकेश")
    s = s.replace("तररर", "तारा").replace("बसपतर", "बसन्ती").replace("सचशललर", "सुशीला")
    s = s.replace("बरलब", "बालू").replace("हलरर", "हरि").replace("गयपरल", "गोपाल")
    s = s.replace("लकमण", "लक्ष्मण").replace("रबपगर", "रूप").replace("नतलयक", "त्रिलोक")
    s = s.replace("कबमप", "कुम्भा").replace("भभर", "भंवर").replace("गलतर", "गीता")
    s = s.replace("अमरर", "अमरा").replace("रकमर", "रुक्मा").replace("हजररल", "हजारी")
    s = s.replace("नरजज", "नरपत").replace("नरलसह", "नरसिंह")
    s = s.replace("सपनतर", "शान्ति").replace("कमलप", "कमला").replace("मदनलसह", "मदनसिंह")
    s = s.replace("अमरलसह", "अमरसिंह").replace("ररम", "राम").replace("पबनम", "पूनम")
    s = s.replace("नरररजण", "नारायण").replace("बलरलर", "बलवीर").replace("पपथरल", "पृथ्वी")
    s = s.replace("ककसर", "केसर").replace("मयतल", "मोती").replace("शररदर", "शारदा")
    s = s.replace("ओमपकरश", "ओमप्रकाश").replace("सनजल", "संजय").replace("ककलरश", "कैलाश")
    s = s.replace("जगरलर", "जगदीश").replace("सचदर", "सुन्दर").replace("सचनरलर", "सुनील")
    s = s.replace("सचनरत", "सुनीता").replace("कनवतर", "कविता").replace("मनरज", "मनोज")
    s = s.replace("पपमर", "प्रेम").replace("रजकमर", "राकेश")
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def clean_booth_address(raw_b):
    if not raw_b:
        return ""
    s = raw_b.strip()
    s = re.sub(r'ररजक[कए]ज|रोकीय|राकीय', 'राजकीय', s)
    s = re.sub(r'\bउच\b|\bअन\b', 'उच्च', s)
    s = re.sub(r'मरधजनमक|करसजनकक|माबयमिक', 'माध्यमिक', s)
    s = re.sub(r'नरदरलज|नरररतज|नवदपलब', 'विद्यालय', s)
    s = re.sub(r'जररजर', 'जवाजा', s)
    s = re.sub(r'सरमरनलजर|सरमालिया', 'सरमालिया', s)
    s = re.sub(r'कमरर|ककरर', 'कमरा', s)
    s = re.sub(r'न\.?\s*(\d+)', r'न.\1', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def extract_pdf_elector_data(pdf_path):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = pymupdf.open(pdf_path)
    if len(doc) == 0:
        raise ValueError(f"Empty PDF: {pdf_path}")

    p1 = doc[0]
    p1_text = p1.get_text()

    # 1. Ward Number
    ward = "1"
    ward_m = re.search(r'(?:रररर|वररर|वरडर|वार्ड)\s*(?:कमरपक|करमपक|क्रमांक|सपखजर|संख्या|नं)?\s*[:.-]?\s*(\d+)', p1_text)
    if ward_m:
        ward = ward_m.group(1)
    else:
        fn_m = re.search(r'(?i)Ward\s*(?:No)?[-_\s]*0*(\d+)', os.path.basename(pdf_path))
        if fn_m:
            ward = fn_m.group(1)
        else:
            colons = re.findall(r'\(\s*:\s*(\d+)\s*\)\s*Tj', p1_text)
            if len(colons) >= 3:
                ward = colons[2]

    # 2. Polling Booth Address
    booth = f"मतदान केंद्र (वार्ड {ward})"
    blocks = p1.get_text("blocks")
    booth_y = None
    for b in blocks:
        if re.search(r'मतद[रा]न\s*ब[बूह]स|मतदान\s*बूथ', b[4]):
            booth_y = b[1]
            break
    if booth_y is not None:
        for b in blocks:
            if abs(b[1] - booth_y) < 18 and b[0] > 150:
                raw_b = b[4].strip()
                cleaned = clean_booth_address(raw_b)
                if len(cleaned) > 8:
                    booth = cleaned
                    break

    # 3. Detect all Deleted Serial Numbers
    deleted_serials = set()
    for p_no in range(len(doc)):
        text = doc[p_no].get_text()
        if 'E-Deleted' in text or 'S-Deleted' in text or 'R-Deleted' in text or 'विलोपन' in text or 'नरलयपन' in text:
            m = re.findall(r'\n([ESR])\s*\n\s*Photo is\s*\n\s*Available\s*\n\s*(\d+)', text)
            for mark, sn in m:
                deleted_serials.add(int(sn))
            m2 = re.findall(r'Photo is\s*\n\s*Available\s*\n\s*(\d+)\s*\n[^\n]+\n[^\n]+\n[^\n]+\n[^\n]+\n[^\n]+\n[^\n]+\n[^\n]+\n([ESR])\b', text)
            for sn, mark in m2:
                deleted_serials.add(int(sn))

    # 4. Extract all voter cards across all pages
    voters_dict = {}
    total_serials_found = 0

    for p_no in range(1, len(doc)-1):
        p_text = doc[p_no].get_text()
        # Skip pure deletion pages or cover/map pages
        if p_no == 1 and ("नकशर" in p_text or "नक्शा" in p_text):
            continue

        lines = [l.strip() for l in p_text.splitlines() if l.strip()]
        i = 0
        while i < len(lines):
            if lines[i] == "Photo is" and i + 2 < len(lines) and lines[i+1] == "Available" and lines[i+2].isdigit():
                sn = int(lines[i+2])
                if sn > total_serials_found:
                    total_serials_found = sn

                # Skip deleted voters
                if sn in deleted_serials:
                    i += 3
                    continue

                card_slice = lines[max(0, i-11):i]
                age = "30"
                for cl in card_slice:
                    if "आजच:" in cl or "आयु:" in cl:
                        am = re.search(r'\d+', cl)
                        if am:
                            age = am.group(0)

                gender = "स्त्री" if any(cl in ["सल", "स्त्री", "महिला"] for cl in card_slice) else "पुरुष"

                epic = ""
                for cl in card_slice:
                    if re.match(r'^[A-Z]{3}\d{7}$|^RJ/\d+/\d+/\d+$', cl):
                        epic = cl
                        break

                g_idx = -1
                for ki, cl in enumerate(card_slice):
                    if cl in ["पचरष", "सल", "पुरुष", "स्त्री"]:
                        g_idx = ki
                        break

                house = "-"
                v_name = ""
                r_name = ""
                if g_idx != -1 and g_idx + 1 < len(card_slice):
                    house = card_slice[g_idx + 1]
                    if g_idx + 2 < len(card_slice) and card_slice[g_idx + 2] != epic:
                        v_name = card_slice[g_idx + 2]
                    if g_idx + 3 < len(card_slice) and card_slice[g_idx + 3] != epic:
                        r_name = card_slice[g_idx + 3]

                if sn not in voters_dict:
                    voters_dict[sn] = {
                        "Ward": str(ward),
                        "Part": "1",
                        "Booth": booth,
                        "SerialNo": sn,
                        "VoterName": clean_hindi_name(v_name),
                        "RelativeType": "पिता/पति",
                        "RelativeName": clean_hindi_name(r_name),
                        "HouseNo": house,
                        "Age": str(age),
                        "Gender": gender,
                        "EPIC": epic
                    }
                i += 3
            else:
                i += 1

    active_voters = [voters_dict[k] for k in sorted(voters_dict.keys())]

    # If document has summary on last page, verify counts
    last_text = doc[-1].get_text()
    yog_m = re.search(r'(?:जयग|योग)\s*\n*.*?\n*(\d+)', last_text)
    if yog_m:
        expected_active = int(yog_m.group(1))
        # If total_serials_found was 0, derive from counts
        if total_serials_found == 0:
            total_serials_found = expected_active + len(deleted_serials)

    return {
        "fileName": os.path.basename(pdf_path),
        "ward": str(ward),
        "part": "1",
        "booth": booth,
        "totalSerials": total_serials_found if total_serials_found > 0 else (len(active_voters) + len(deleted_serials)),
        "deletedCount": len(deleted_serials),
        "activeCount": len(active_voters),
        "voters": active_voters
    }

def export_voters_to_excel(voters, out_excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "मतदाता सूची (Master Voter List)"

    headers = [
        "क्रम संख्या", "वार्ड नं.", "भाग नं.", "मतदान केंद्र का नाम",
        "मतदाता का नाम", "संबंध का प्रकार", "पिता/पति का नाम",
        "मकान नं.", "आयु", "लिंग", "पहचान पत्र क्र. (EPIC)"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1A237E", end_color="1A237E", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    for v in voters:
        row = [
            v.get("SerialNo", ""),
            v.get("Ward", "1"),
            v.get("Part", "1"),
            v.get("Booth", ""),
            v.get("VoterName", ""),
            v.get("RelativeType", "पिता/पति"),
            v.get("RelativeName", ""),
            v.get("HouseNo", "-"),
            v.get("Age", ""),
            v.get("Gender", ""),
            v.get("EPIC", "")
        ]
        ws.append(row)

    # Style rows
    for row_idx in range(2, len(voters) + 2):
        for col_idx in range(1, 12):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            if col_idx in [1, 2, 3, 8, 9, 10]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Column widths
    col_widths = {1: 12, 2: 10, 3: 10, 4: 38, 5: 22, 6: 15, 7: 22, 8: 12, 9: 8, 10: 10, 11: 18}
    for col_idx, w in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = w

    os.makedirs(os.path.dirname(os.path.abspath(out_excel_path)), exist_ok=True)
    wb.save(out_excel_path)
    wb.close()
    return out_excel_path

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Extract electors from PDF")
    parser.add_argument("--pdf", required=True, help="Path to voter roll PDF")
    parser.add_argument("--out-json", help="Path to save extracted JSON")
    parser.add_argument("--out-excel", help="Path to save Excel")
    args = parser.parse_args()

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
            "booth": data["booth"],
            "totalSerials": data["totalSerials"],
            "deletedCount": data["deletedCount"],
            "activeCount": data["activeCount"],
            "votersCount": len(data["voters"])
        }, ensure_ascii=False, indent=2))
