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
    gp_html = (f'<div class="gp-line">ग्राम पंचायत : <strong>{gp}</strong></div>'
               if gp else '')

    return f"""
        <div class="slip">
          <div class="slip-left">
            <div class="slip-header-block">
              <div class="slip-title">वोटर सुविधा स्लिप</div>
              {gp_html}
              <div class="meta-row">
                <span class="badge-item">वार्ड नं : <strong>{v.get('Ward', '1')}</strong></span>
                <span class="badge-item serial-badge">क्रम संख्या : <strong>{v.get('SerialNo', '')}</strong></span>
              </div>
            </div>
            <div class="voter-body">
              <div class="voter-line voter-name">{v.get('VoterName', '')}</div>
              <div class="voter-line">{rel_label} : <strong>{v.get('RelativeName', '')}</strong></div>
              <div class="voter-line">मकान नं. : <strong>{v.get('HouseNo', '')}</strong> &nbsp;|&nbsp; उम्र : <strong>{v.get('Age', '')}</strong> &nbsp;|&nbsp; लिंग : <strong>{v.get('Gender', '')}</strong></div>
              <div class="voter-line">EPIC : <strong>{v.get('EPIC', '')}</strong></div>
            </div>
            <div class="booth-line"><span class="booth-label">मतदान केंद्र :</span> {v.get('Booth', '')}</div>
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

    .a4-page {{ width: 210mm; height: 297mm; padding: 4mm;
                page-break-after: always; overflow: hidden; }}
    .slips-grid {{ display: grid; {grid_css} width: 100%; height: 100%; }}

    .slip {{
      border: 1.4px solid #111;
      border-radius: 4px;
      padding: 2px 3px;
      display: flex;
      flex-direction: row;
      align-items: stretch;
      gap: 2px;
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
      border-bottom: 1px dashed #555;
      padding-bottom: 1px;
      margin-bottom: 1px;
    }}
    .slip-title {{
      text-align: center;
      font-weight: 900;
      font-size: {f('title', '14px')};
      letter-spacing: 0.3px;
      color: #000;
      line-height: 1.1;
    }}
    .gp-line {{
      text-align: center;
      font-size: {f('meta', '11px')};
      font-weight: 700;
      color: #000;
      line-height: 1.15;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .meta-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 4px;
      font-size: {f('meta', '11px')};
      margin-top: 1px;
    }}
    .badge-item {{ color: #111; white-space: nowrap; }}
    .badge-item strong {{ color: #000; }}
    .serial-badge {{
      background: #f2f3f5;
      border: 1px solid #111;
      border-radius: 3px;
      padding: 0 4px;
      font-weight: 800;
    }}

    .voter-body {{
      flex: 1 1 auto;
      display: flex;
      flex-direction: column;
      justify-content: space-around;
      gap: 1px;
      padding: 2px 0;
      min-width: 0;
    }}
    .voter-line {{
      font-size: {f('detail', '12px')};
      color: #111;
      line-height: 1.18;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .voter-name {{
      font-size: {f('name', '16px')};
      font-weight: 900;
      color: #000;
      line-height: 1.15;
      margin-bottom: 0.5px;
    }}
    .voter-line strong {{ color: #000; }}

    .booth-line {{
      font-size: {f('booth', '10px')};
      line-height: 1.15;
      border-top: 1px dashed #666;
      padding-top: 1px;
      color: #000;
    }}
    .booth-label {{ font-weight: 900; }}

    /* ---------- the cut line ---------- */
    /* rotate(), not writing-mode: Devanagari stacks badly in vertical mode
       and rotate is supported identically by Chrome and Edge */
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
      font-size: {f('cut', '6.5px')};
      color: #555;
      white-space: nowrap;
      line-height: 1;
    }}

    /* ---------- right: candidate ---------- */
    .slip-right-box {{
      flex: 0 0 33%;
      max-width: 33%;
      border: 1.4px solid #1a237e;
      border-radius: 4px;
      background: #fbfbfd;
      padding: 1px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: space-evenly;
      text-align: center;
      overflow: hidden;
    }}
    .cand-post {{
      font-size: {f('c_post', '10px')};
      font-weight: 900;
      color: #b71c1c;
      line-height: 1.1;
      width: 100%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .cand-photo-frame {{
      width: {f('img_w', '52px')};
      height: {f('img_h', '58px')};
      border-radius: 3px;
      border: 1px solid #7986cb;
      background: #e8eaf6;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
      flex-shrink: 0;
    }}
    .cand-photo {{ width: 100%; height: 100%; object-fit: cover; display: block; }}
    .cand-avatar {{ font-size: {f('avatar', '38px')}; line-height: 1; }}
    .cand-name {{
      font-size: {f('c_name', '12px')};
      font-weight: 900;
      color: #0d47a1;
      line-height: 1.1;
    }}
    .cand-symbol-frame {{
      width: {f('sym_w', '46px')};
      height: {f('sym_h', '46px')};
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }}
    .cand-symbol-img {{ width: 100%; height: 100%; object-fit: contain; }}
    .cand-party {{
      font-size: {f('c_party', '9px')};
      font-weight: 800;
      color: #2e7d32;
      line-height: 1.1;
    }}
    .cand-appeal {{
      font-size: {f('c_app', '8px')};
      font-weight: 800;
      color: #b71c1c;
      line-height: 1.1;
    }}
    """
