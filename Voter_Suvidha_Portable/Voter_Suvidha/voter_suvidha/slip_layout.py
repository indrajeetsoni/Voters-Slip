# -*- coding: utf-8 -*-
"""
slip_layout.py — voter slip design for Voter Suvidha.

Drop-in replacement for render_slip_html() and get_full_page_css() in
server.py. Keeping the layout in its own file means server.py does not have
to be edited again the next time the design changes.

CHANGES IN THIS VERSION
  * "भाग : [ ]" removed from every slip
  * "क्रम संख्या" now sits on the same row as "वार्ड नं", to its right
  * new "ग्राम पंचायत" line directly under the "वोटर सुविधा स्लिप" title,
    filled from the uploaded PDF (identical for every slip of one village)
  * spacing tightened throughout so more fits without shrinking the text
  * a vertical dashed cut line before the candidate box, carrying the
    message "मतदान देने जाने से पूर्व यहाँ से काटें" printed sideways

font_scale keys are read with .get() and sensible defaults, so this works
with any of the 4 / 6 / 8 / 10 / 12 slips-per-page presets.
"""

CUT_MESSAGE = "मतदान से पूर्व यहाँ से काटें"


def render_slip_html(v, candidate, party, appeal, candidate_photo,
                     party_symbol, font_scale, candidate_post="सरपंच"):
    photo_html = (
        f'<img class="cand-photo" src="{candidate_photo}" alt="Candidate">'
        if candidate_photo else '<div class="cand-avatar">&#128100;</div>'
    )
    symbol_html = (
        f'<div class="cand-symbol-frame">'
        f'<img class="cand-symbol-img" src="{party_symbol}" alt="चुनाव चिन्ह">'
        f'</div>'
        if party_symbol else ''
    )

    rel_type = v.get("RelativeType", "")
    rel_label = (f"{rel_type} का नाम"
                 if rel_type in ("पिता", "पति", "माता") else "पिता/पति का नाम")

    gp = (v.get("GramPanchayat") or "").strip()
    ps = (v.get("PanchayatSamiti") or "").strip()

    gp_html = (f'<div class="gp-line"><span class="meta-lbl">ग्राम पंचायत :</span> <strong>{gp}</strong></div>'
               if gp else '')
    ps_html = (f'<div class="ps-line"><span class="meta-lbl">पंचायत समिति :</span> <strong>{ps}</strong></div>'
               if ps else '')

    return f"""
        <div class="slip">
          <div class="slip-left">
            <div class="slip-header-block">
              <div class="slip-title">वोटर सुविधा स्लिप</div>
              {gp_html}
              {ps_html}
              <div class="meta-row">
                <span class="badge-item"><span class="meta-lbl">वार्ड नं :</span> <strong>{v.get('Ward', '1')}</strong></span>
                <span class="badge-item serial-badge"><span class="meta-lbl">क्रम संख्या :</span> <strong>{v.get('SerialNo', '')}</strong></span>
              </div>
            </div>
            <div class="voter-body">
              <div class="voter-line voter-name">{v.get('VoterName', '')}</div>
              <div class="voter-line"><span class="meta-lbl">{rel_label} :</span> <strong>{v.get('RelativeName', '')}</strong></div>
              <div class="voter-line"><span class="meta-lbl">मकान नं. :</span> <strong>{v.get('HouseNo', '')}</strong> | <span class="meta-lbl">उम्र :</span> <strong>{v.get('Age', '')}</strong> | <span class="meta-lbl">लिंग :</span> <strong>{v.get('Gender', '')}</strong></div>
              <div class="voter-line"><span class="meta-lbl">EPIC :</span> <strong>{v.get('EPIC', '')}</strong></div>
            </div>
            <div class="booth-line"><span class="booth-label">मतदान केंद्र :</span> <span class="booth-val"><strong>{v.get('Booth', '')}</strong></span></div>
          </div>
          <div class="cut-strip"><span class="cut-text">{CUT_MESSAGE}</span></div>
          <div class="slip-right-box">
            <div class="cand-post">{candidate_post} पद हेतु</div>
            <div class="cand-photo-frame">{photo_html}</div>
            <div class="cand-name">{candidate}</div>
            {symbol_html}
            <div class="cand-party">({party})</div>
            <div class="cand-appeal">{appeal}</div>
          </div>
        </div>
    """


def get_full_page_css(grid_css, font_scale):
    f = font_scale.get
    return f"""
    @page {{ size: A4 portrait; margin: 0; }}
    * {{ box-sizing: border-box; margin: 0; padding: 0;
         -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    body {{ font-family: "Nirmala UI", "Mangal", "Segoe UI", Arial, sans-serif;
            background: #fff; color: #111; }}

    .a4-page {{ width: 210mm; height: 297mm; padding: 3.5mm;
                page-break-after: always; overflow: hidden; }}
    .slips-grid {{ display: grid; {grid_css} width: 100%; height: 100%; }}

    .slip {{
      border: 1.5px solid #111;
      border-radius: 4px;
      padding: 3px 4px;
      display: flex;
      flex-direction: row;
      align-items: stretch;
      gap: 3px;
      background: #fff;
      overflow: hidden;
    }}

    /* ---------- left: voter details ---------- */
    .slip-left {{
      flex: 1 1 auto;
      min-width: 0;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      overflow: hidden;
    }}
    .slip-header-block {{
      border-bottom: 1.2px dashed #444;
      padding-bottom: 2px;
      margin-bottom: 2px;
    }}
    .slip-title {{
      text-align: center;
      font-weight: 900;
      font-size: {f('title', '15px')};
      letter-spacing: 0.3px;
      color: #000;
      line-height: 1.15;
    }}
    .gp-line, .ps-line {{
      text-align: center;
      font-size: {f('panchayat', f('meta', '12px'))};
      color: #000;
      line-height: 1.2;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      margin-top: 1px;
    }}
    .gp-line .meta-lbl, .ps-line .meta-lbl {{
      font-weight: 700;
      color: #222;
    }}
    .gp-line strong, .ps-line strong {{
      font-weight: 900;
      color: #000;
    }}
    .meta-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 4px;
      font-size: {f('meta', '12.5px')};
      margin-top: 2px;
    }}
    .meta-lbl {{
      font-weight: 700;
      color: #222;
    }}
    .badge-item {{ color: #111; white-space: nowrap; }}
    .badge-item strong {{ color: #000; font-weight: 900; }}
    .serial-badge {{
      background: #f0f1f4;
      border: 1.2px solid #111;
      border-radius: 3px;
      padding: 0.5px 5px;
      font-weight: 900;
      font-size: {f('serial', f('meta', '13px'))};
    }}

    .voter-body {{
      flex: 1 1 auto;
      display: flex;
      flex-direction: column;
      justify-content: space-evenly;
      gap: 2.5px;
      padding: 2px 0;
      min-width: 0;
    }}
    .voter-line {{
      font-size: {f('detail', '12.7px')};
      color: #111;
      line-height: 1.25;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .voter-name {{
      font-size: {f('name', '19px')};
      font-weight: 900;
      color: #000;
      line-height: 1.15;
      letter-spacing: 0.2px;
      margin-bottom: 1px;
    }}
    .voter-line strong {{ color: #000; font-weight: 900; }}

    .booth-line {{
      font-size: {f('booth', '12px')};
      font-weight: 800;
      line-height: 1.18;
      border-top: 1.5px solid #111;
      padding-top: 2px;
      color: #000;
      word-break: break-word;
    }}
    .booth-label {{ font-weight: 900; color: #000; }}
    .booth-val {{ font-weight: 800; color: #000; }}

    /* ---------- the cut line ---------- */
    .cut-strip {{
      position: relative;
      flex: 0 0 14px;
      width: 14px;
      border-left: 1.2px dashed #444;
      overflow: hidden;
    }}
    .cut-text {{
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%) rotate(-90deg);
      transform-origin: center center;
      font-size: {f('cut', '7.5px')};
      color: #444;
      font-weight: 700;
      white-space: nowrap;
      line-height: 1;
    }}

    /* ---------- right: candidate ---------- */
    .slip-right-box {{
      flex: 0 0 33%;
      max-width: 33%;
      border: 1.5px solid #1a237e;
      border-radius: 4px;
      background: #fbfbfd;
      padding: 2.5px 2px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: space-evenly;
      text-align: center;
      overflow: hidden;
    }}
    .cand-post {{
      font-size: {f('c_post', '13px')};
      font-weight: 900;
      color: #b71c1c;
      line-height: 1.12;
      width: 100%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .cand-photo-frame {{
      width: {f('img_w', '64px')};
      height: {f('img_h', '70px')};
      border-radius: 4px;
      border: 1.2px solid #7986cb;
      background: #e8eaf6;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
      flex-shrink: 0;
    }}
    .cand-photo {{ width: 100%; height: 100%; object-fit: cover; display: block; }}
    .cand-avatar {{ font-size: {f('avatar', '48px')}; line-height: 1; }}
    .cand-name {{
      font-size: {f('c_name', '16.5px')};
      font-weight: 900;
      color: #0d47a1;
      line-height: 1.12;
    }}
    .cand-symbol-frame {{
      width: {f('sym_w', '50px')};
      height: {f('sym_h', '50px')};
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }}
    .cand-symbol-img {{ width: 100%; height: 100%; object-fit: contain; }}
    .cand-party {{
      font-size: {f('c_party', '11.8px')};
      font-weight: 800;
      color: #2e7d32;
      line-height: 1.12;
    }}
    .cand-appeal {{
      font-size: {f('c_app', '10.8px')};
      font-weight: 800;
      color: #b71c1c;
      line-height: 1.12;
    }}
    """
