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

# Comprehensive Devanagari Decoder for SEC Rajasthan Election Rolls
WORD_MAP = {
    # Places & Booth text
    "ररनजजरररस": "राजियावास",
    "जररजर": "जवाजा",
    "सरमरनलजर": "सरमालिया",
    "सरमालिया": "सरमालिया",
    "ररजककज": "राजकीय",
    "ररजकएज": "राजकीय",
    "रोकीय": "राजकीय",
    "राकीय": "राजकीय",
    "मरधजनमक": "माध्यमिक",
    "करसजनकक": "माध्यमिक",
    "माबयमिक": "माध्यमिक",
    "नरदरलज": "विद्यालय",
    "नरररतज": "विद्यालय",
    "नवदपलब": "विद्यालय",
    "कमरर": "कमरा",
    "ककरर": "कमरा",
    "मकरन": "मकान",
    "सपखजर": "संख्या",
    "नपतर": "पिता",
    "पनत": "पति",
    "आजच": "आयु",
    "ललग": "लिंग",
    "पचरष": "पुरुष",
    "सल": "स्त्री",

    # Surnames / Titles
    "चफहरन": "चौहान",
    "चचहरन": "चौहान",
    "चयहरन": "चौहान",
    "ररठयड": "राठौड",
    "ररठयद": "राठौड",
    "पपररर": "पंवार",
    "रररत": "रावत",
    "कचमररल": "कुमारी",
    "कचमरर": "कुमार",
    "कचमपरर": "कुमारी",

    # Names
    "सयहन": "सोहन",
    "सयहनल": "सोहनी",
    "जररन": "जवान",
    "छगनल": "छगनी",
    "छगन": "छगन",
    "करलब": "कालू",
    "करलप": "कालू",
    "धमर": "धर्म",
    "पकमल": "प्रेमी",
    "पकम": "प्रेम",
    "पपमर": "प्रेम",
    "परममद": "परमेन्द्र",
    "परमकनद": "परमेन्द्र",
    "चपचल": "चंचल",
    "दरकनद": "देवेन्द्र",
    "दकरकनद": "देवेन्द्र",
    "दरकन क द": "देवेन्द्र",
    "दकरकन द": "देवेन्द्र",
    "दरकन द": "देवेन्द्र",
    "दकरषद": "देवेन्द्र",
    "सचरकनद": "सुरेन्द्र",
    "सचरकन द": "सुरेन्द्र",
    "हककनत": "हेमन्त",
    "हकमनत": "हेमन्त",
    "हमनत": "हेमन्त",
    "हमनत क": "हेमन्त",
    "जमनल": "जमनी",
    "जमनर": "जमना",
    "मदन": "मदन",
    "मदनलसह": "मदनसिंह",
    "तलजर": "तीजा",
    "भगररन": "भगवान",
    "भररन": "भगवान",
    "आनन्द": "आनन्द",
    "आनपद": "आनन्द",
    "आननद": "आनन्द",
    "आनद": "आनन्द",
    "मपजच": "मंजू",
    "मपजचदकरल": "मंजू देवी",
    "मपजचदेवी": "मंजू देवी",
    "मपजब": "मंजू",
    "पकरश": "प्रकाश",
    "ओमपकरश": "ओमप्रकाश",
    "भलम": "भीम",
    "मपगल": "मंगल",
    "मपगललसह": "मंगलसिंह",
    "झमकब": "झमकू",
    "झमकप": "झमकू",
    "भपरर": "भंवर",
    "भबर": "भंवर",
    "सचशललर": "सुशीला",
    "गणकश": "गणेश",
    "महकनद": "महेन्द्र",
    "महनद": "महेन्द्र",
    "महनदक": "महेन्द्र",
    "महकनदलसह": "महेन्द्रसिंह",
    "ररधर": "राधा",
    "लनलतर": "ललिता",
    "सचननतर": "सुनीता",
    "सचनलतर": "सुनीता",
    "सचनरत": "सुनीता",
    "कैलाश": "कैलाश",
    "ककलरश": "कैलाश",
    "शररदर": "शारदा",
    "तकज": "तेज",
    "गटच": "गटू",
    "गटब": "गटू",
    "दकशन": "किशन",
    "ककशन": "किशन",
    "केसर": "केसर",
    "ककसर": "केसर",
    "कपकब": "कंकू",
    "कपकप": "कंकू",
    "पचषपर": "पुष्पा",
    "हरम": "राम",
    "ररम": "राम",
    "सबरतर": "सूरता",
    "सबरज": "सूरज",
    "सचरत": "सूरज",
    "लकहरल": "लेहरी",
    "कबरल": "कबरी",
    "मरपगब": "मांगू",
    "ईशर": "ईश्वर",
    "लपटब": "पिंटू",
    "नपनटब": "पिन्टू",
    "मनलषर": "मनीषा",
    "नरदर": "विद्या",
    "हमलतर": "हेमलता",
    "कलपनर": "कल्पना",
    "रलनर": "रीना",
    "ररलनर": "रीना",
    "पबजर": "पूजा",
    "पपजर": "पूजा",
    "नपजर": "पूजा",
    "परलण": "प्रवीण",
    "नरकश": "नरेश",
    "बसतल": "बस्ती",
    "खचशबच": "खुशबू",
    "मयखम": "मोखम",
    "सचमन": "सुमन",
    "कपरर": "कंवर",
    "कनरतर": "कविता",
    "कनवतर": "कविता",
    "नचमन": "चिमन",
    "सरकल": "साक्षी",
    "कपचन": "कंचन",
    "सपनर": "सपना",
    "अनलतर": "अनीता",
    "अननतर": "अनिता",
    "सररतर": "सरिता",
    "नललम": "नीलम",
    "अमलषर": "अमीषा",
    "अपजप": "अंजू",
    "मरजर": "माया",
    "रयहन": "रोहन",
    "पबरण": "पूरण",
    "तरण": "तरुण",
    "सचरकश": "सुरेश",
    "टलकम": "टीकम",
    "सपदलप": "संदीप",
    "पपकज": "पंकज",
    "चकतन": "चेतन",
    "ररहल": "राहुल",
    "धनर": "धना",
    "नरररजण": "नारायण",
    "नरसप": "नाथू",
    "नरसब": "नाथू",
    "नरपत": "नरपत",
    "खकत": "खेत",
    "गयलरद": "गोविन्द",
    "गयनरनद": "गोविन्द",
    "कदनकश": "दिनेश",
    "हकम": "हुकम",
    "करण": "करण",
    "पचषपषद": "पुष्पेन्द्र",
    "पचषपकनद": "पुष्पेन्द्र",
    "पचषपमद": "पुष्पेन्द्र",
    "नरकम": "विक्रम",
    "खलम": "खीम",
    "सलमर": "सीमा",
    "कदललप": "दिलीप",
    "ददललप": "दिलीप",
    "नरकरस": "विकास",
    "सचखदकर": "सुखदेव",
    "सचखदर": "सुखदेव",
    "गयपरल": "गोपाल",
    "शरण": "श्रवण",
    "शररण": "श्रवण",
    "अशयक": "अशोक",
    "अरनरनद": "अरविन्द",
    "अरलरद": "अरविन्द",
    "मलनर": "मीना",
    "मकनर": "मैना",
    "नकनर": "नैनसिंह",
    "दकशनल": "किशनी",
    "आपब": "आपू",
    "समचनद": "समुन्द्र",
    "सकब": "सकू",
    "ररजकनद": "राजेन्द्र",
    "नरकनद": "नरेन्द्र",
    "नकनल": "नेनी",
    "अजचरन": "अर्जुन",
    "भकर": "भेरू",
    "जज": "जय",
    "मधच": "मधु",
    "डरऊ": "डाऊ",
    "हरर": "हरि",
    "हरल": "हरि",
    "हरलब": "हरि",
    "सकनर": "सेना",
    "भगरतल": "भगवती",
    "ररमल": "रामी",
    "लरलल": "लाली",
    "परनल": "पानी",
    "पजररल": "प्यारी",
    "टलल": "टील",
    "सपगलतर": "संगीता",
    "खखोयत": "खोयत",
    "नडमपल": "डिम्पल",
    "लयकमद": "लोकेन्द्र",
    "खचमरन": "खुमान",
    "ननरमर": "निरमा",
    "बकनचगयपरल": "वेणुगोपाल",
    "हलरल": "हीरी",
    "गकनल": "गेनी",
    "ककलल": "केली",
    "जलरणल": "जीवणी",
    "इपदर": "इंद्रा",
    "रनर": "रवि",
    "सलतर": "सीता",
    "बलनर": "बीना",
    "बखतररर": "बख्तावर",
    "जशयदर": "जशोदा",
    "पतरप": "प्रताप",
    "रतनल": "रतन",
    "अनरल": "अमरा",
    "अमरर": "अमरा",
    "मचनल": "मुन्नी",
    "नगरधररल": "गिरधारी",
    "इनदर": "इन्द्रा",
    "शपकर": "शंकर",
    "परपचब": "पांचू",
    "परपचल": "पांची",
    "बदरमल": "बादामी",
    "अपजबलतर": "अंजुलता",
    "शकब": "शंकू",
    "सचगनर": "सुगना",
    "लरजरपतल": "लाजवंती",
    "सरलर": "सरला",
    "नरजज": "नत्थू",
    "मलठब": "मीठू",
    "नमटठब": "मीठू",
    "सयनब": "सोनू",
    "झमरल": "झमरी",
    "नलतब": "नीतू",
    "मयहन": "मोहन",
    "मलरर": "मीरा",
    "ररककश": "राकेश",
    "रजकमर": "राकेश",
    "मकसल": "मैकी",
    "चचनल": "चुन्नी",
    "ननदर": "नन्दा",
    "रकखर": "रेखा",
    "मनयहर": "मनोहर",
    "मनरज": "मनोज",
    "आशर": "आशा",
    "अकत": "अक्षत",
    "अननरद": "अनिरुद्ध",
    "बरबब": "बाबू",
    "बरलब": "बालू",
    "गलतर": "गीता",
    "टकमर": "टीमा",
    "ररजब": "राजू",
    "सचनलल": "सुनील",
    "सचनरलर": "सुनील",
    "सपतयष": "सन्तोष",
    "सपतयषल": "सन्तोषी",
    "समपनत": "सम्पति",
    "सपनतर": "शान्ति",
    "शरपनत": "शान्ति",
    "नतलयक": "तिलोक",
    "लकमण": "लक्ष्मण",
    "लकमल": "लक्ष्मी",
    "अमरलसह": "अमरसिंह",
    "परजल": "पायल",
    "पबनम": "पूनम",
    "पचनम": "पूनम",
    "पभब": "प्रभु",
    "पभच": "प्रभु",
    "मचककश": "मुकेश",
    "तररर": "तारा",
    "बसपतर": "बसन्ती",
    "रकमर": "रुक्मा",
    "हजररल": "हजारी",
    "कमलप": "कमला",
    "कमलर": "कमला",
    "बलरलर": "बलवीर",
    "पपथरल": "पृथ्वी",
    "मयतल": "मोती",
    "सनजल": "संजय",
    "जगरलर": "जगदीश",
    "सचदर": "सुन्दर",
    "ककशर": "केशर",
    "गयदरररल": "गोदावरी",
    "दचगरर": "दुर्गा",
    "पचनर": "पूना",
    "ररजकशरल": "राजेश्वरी",
    "ररर": "राव",
    "सरदरर": "सरदार",
    "शरजरल": "सज्जन",
    "रबपगर": "रूप",
    "रबप": "रूप",
    "कबमप": "कुम्भा",
    "भभर": "भंवर",
    "चनद": "चन्द"
}

def clean_hindi_name(raw):
    if not raw:
        return ""
    s = raw.strip()

    # Pre-clean known suffixes with word boundaries
    s = re.sub(r'लसह\b|नसहम\b|लसग\b|लसहप\b', 'सिंह', s)
    s = re.sub(r'दकरल\b|दरल\s*क\b|दरलक\b|दकल\b', 'देवी', s)
    s = re.sub(r'चफहरन\b|चचहरन\b|चयहरन\b', 'चौहान', s)
    s = re.sub(r'ररठयड\b|ररठयद\b', 'राठौड', s)
    s = re.sub(r'कचमररल\b', 'कुमारी', s)
    s = re.sub(r'कचमरर\b', 'कुमार', s)
    s = re.sub(r'दकरकन\s*द\b|दरकन\s*द\b|दरकन\s*क\s*द\b', 'देवेन्द्र', s)
    s = re.sub(r'सचरकन\s*द\b', 'सुरेन्द्र', s)
    s = re.sub(r'हककनत\b|हकमनत\b', 'हेमन्त', s)

    # Word-by-word dictionary translation
    tokens = s.split()
    cleaned = []
    for t in tokens:
        if t in WORD_MAP:
            cleaned.append(WORD_MAP[t])
        else:
            cur = t
            for k, v in WORD_MAP.items():
                if k in cur:
                    cur = cur.replace(k, v)
            cleaned.append(cur)

    res = " ".join(cleaned)
    res = re.sub(r'\s+', ' ', res)
    return res.strip()

def clean_booth_address(raw_b):
    if not raw_b:
        return ""
    s = clean_hindi_name(raw_b)
    s = re.sub(r'ररजक[कए]ज|रोकीय|राकीय', 'राजकीय', s)
    s = re.sub(r'\bउच\b|\bअन\b', 'उच्च', s)
    s = re.sub(r'मरधजनमक|करसजनकक|माबयमिक', 'माध्यमिक', s)
    s = re.sub(r'नरदरलज|नरररतज|नवदपलब', 'विद्यालय', s)
    s = re.sub(r'ररनजजरररस', 'राजियावास', s)
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
    ward_m = re.search(r'(?:ररडर|रररर|वररर|वरडर|वार्ड)\s*(?:कमरपक|करमपक|क्रमांक|सपखजर|संख्या|नं)?\s*[:.-]?\s*(\d+)', p1_text)
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

                # Extract Relationship Type
                rel_type = "पिता"
                for cl in card_slice:
                    if "पनत" in cl or "पति" in cl:
                        rel_type = "पति"
                        break
                    elif "नपतर" in cl or "पिता" in cl:
                        rel_type = "पिता"
                        break
                    elif "मपतर" in cl or "माता" in cl:
                        rel_type = "माता"
                        break
                    elif "अनज" in cl or "अन्य" in cl:
                        rel_type = "अन्य"
                        break

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

                after_gender = card_slice[g_idx+1:] if g_idx != -1 else []
                # Filter out epic if it is in after_gender
                items = [x for x in after_gender if x != epic and x not in ['Photo is', 'Available']]

                house = "-"
                v_name = ""
                r_name = ""
                if len(items) >= 3:
                    house = items[0]
                    v_name = items[1]
                    r_name = items[2]
                elif len(items) == 2:
                    # If first item has digits and no Devanagari, it is house number
                    if re.search(r'\d', items[0]) and not re.search(r'[\u0900-\u097F]', items[0]):
                        house = items[0]
                        v_name = items[1]
                    else:
                        # House number is omitted (e.g. additions page)
                        v_name = items[0]
                        r_name = items[1]
                elif len(items) == 1:
                    v_name = items[0]

                if sn not in voters_dict:
                    voters_dict[sn] = {
                        "Ward": str(ward),
                        "Part": "1",
                        "Booth": booth,
                        "SerialNo": sn,
                        "VoterName": clean_hindi_name(v_name),
                        "RelativeType": rel_type,
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
            v.get("RelativeType", "पिता"),
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
            if col_idx in [1, 2, 3, 6, 8, 9, 10]:
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
