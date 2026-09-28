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
  * the voter's own photo (from the WithPhoto roll PDF, matched by EPIC /
    serial number — see voter_photo_extractor.py) sits beside the voter's
    name and father's/husband's name, sized like the candidate photo box.
    Only one PDF is ever uploaded (server.py runs the photo extractor on
    it automatically); if that PDF has no photos, voter_photo is empty for
    every voter and the name block below falls back to the exact original
    single-column layout — no placeholder icon, no empty box, no shift.

font_scale keys are read with .get() and sensible defaults, so this works
with any of the 4 / 6 / 8 / 10 / 12 slips-per-page presets.
"""

CUT_MESSAGE = "मतदान से पूर्व यहाँ से काटें"


def _has_image(data_uri: str) -> bool:
    """True only for a value that actually carries image bytes.

    Guards against a broken-image icon showing up when no photo/symbol was
    uploaded. The literal empty string "" is the normal "nothing uploaded"
    case, but a data URI can also arrive as just the header with nothing
    after the comma — e.g. "data:image/png;base64," — which is non-empty,
    passes a plain `if value:` check, yet is not a decodable image. Both
    the live preview (which embeds this string directly as <img src=...>)
    and the server's temp-file cache need this same check, since the
    preview never goes through the server's file-caching step at all.
    """
    if not data_uri or not isinstance(data_uri, str):
        return False
    s = data_uri.strip()
    if s in ("", "undefined", "null", "false", "0"):
        return False
    if s.startswith("data:"):
        _, _, payload = s.partition(",")
        return bool(payload.strip())
    return True


def render_slip_html(v, candidate, party, appeal, candidate_photo,
                     party_symbol, font_scale, candidate_post="सरपंच",
                     voter_photo=""):
    photo_html = (
        f'<img class="cand-photo" src="{candidate_photo}" alt="Candidate">'
        if _has_image(candidate_photo) else '<div class="cand-avatar">&#128100;</div>'
    )
    symbol_html = (
        f'<div class="cand-symbol-frame">'
        f'<img class="cand-symbol-img" src="{party_symbol}" alt="चुनाव चिन्ह">'
        f'</div>'
        if _has_image(party_symbol) else ''
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

    name_line = f'<div class="voter-line voter-name">{v.get("VoterName", "")}</div>'
    rel_line = f'<div class="voter-line"><span class="meta-lbl">{rel_label} :</span> <strong>{v.get("RelativeName", "")}</strong></div>'

    if _has_image(voter_photo):
        # Uploaded PDF had a matching photo for this voter — split the row
        # so the photo sits beside the name / father's name, like the
        # candidate photo does on the right side of the slip.
        name_block_html = f"""
              <div class="voter-top-row">
                <div class="voter-top-text">
                  {name_line}
                  {rel_line}
                </div>
                <div class="voter-photo-frame"><img class="voter-photo" src="{voter_photo}" alt="Voter"></div>
              </div>"""
    else:
        # No photo for this voter (or the uploaded PDF has no photos at
        # all) — render exactly as before the photo feature existed. No
        # placeholder icon and no empty box eating into the layout.
        name_block_html = f"""
              {name_line}
              {rel_line}"""

    return f"""
        <div class="slip">
          <div class="slip-left">
            <div class="slip-header-block">
              {gp_html}
              {ps_html}
              <div class="meta-row">
                <span class="badge-item"><span class="meta-lbl">वार्ड नं :</span> <strong>{v.get('Ward', '1')}</strong></span>
                <span class="badge-item serial-badge"><span class="meta-lbl">क्रम संख्या :</span> <strong>{v.get('SerialNo', '')}</strong></span>
              </div>
            </div>
            <div class="voter-body">
              {name_block_html}
              <div class="voter-line"><span class="meta-lbl">मकान नं. :</span> <strong>{v.get('HouseNo', '')}</strong> | <span class="meta-lbl">उम्र :</span> <strong>{v.get('Age', '')}</strong> | <span class="meta-lbl">लिंग :</span> <strong>{v.get('Gender', '')}</strong></div>
              <div class="voter-line"><span class="meta-lbl">EPIC :</span> <strong>{v.get('EPIC', '')}</strong></div>
            </div>
            <div class="booth-line"><span class="booth-label">मतदान केंद्र :</span> <span class="booth-val"><strong>{v.get('Booth', '')}</strong></span></div>
          </div>
          <div class="cut-strip"><span class="cut-text">{CUT_MESSAGE}</span></div>
          <div class="slip-right-box">
            <div class="cand-stack">
              <div class="cand-post">{candidate_post} पद हेतु</div>
              <div class="cand-photo-frame">{photo_html}</div>
              <div class="cand-name">{candidate}</div>
              {symbol_html}
              <div class="cand-party">({party})</div>
              <div class="cand-appeal">{appeal}</div>
            </div>
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
      padding-bottom: 1px;
      margin-bottom: 1px;
    }}
    .gp-line, .ps-line {{
      text-align: center;
      font-size: {f('panchayat', f('meta', '12px'))};
      color: #000;
      line-height: 1.08;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      margin-top: 0.5px;
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
      margin-top: 1px;
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
      gap: 1.5px;
      padding: 1px 0;
      min-width: 0;
      /* Without this, a flex column child never shrinks below its own
         content height, which silently stole room from .booth-line
         whenever a long booth address needed 2-3 lines. */
      min-height: 0;
    }}
    /* name + father's/husband's name sit to the left; the voter's own
       photo (from the WithPhoto roll PDF, matched by EPIC / serial number
       — see voter_photo_extractor.py) sits to their right, sized like the
       candidate photo box so it reads as a proper ID photo, not an icon */
    .voter-top-row {{
      display: flex;
      flex-direction: row;
      align-items: flex-start;
      justify-content: space-between;
      gap: 5px;
      min-width: 0;
    }}
    .voter-top-text {{
      flex: 1 1 auto;
      min-width: 0;
      display: flex;
      flex-direction: column;
      gap: 1px;
      overflow: hidden;
    }}
    .voter-photo-frame {{
      flex: 0 0 auto;
      width: {f('vphoto_w', '54px')};
      height: {f('vphoto_h', '62px')};
      border-radius: 4px;
      border: 1.2px solid #7986cb;
      background: #e8eaf6;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
    }}
    .voter-photo {{ width: 100%; height: 100%; object-fit: cover; display: block; }}
    .voter-line {{
      font-size: {f('detail', '12.7px')};
      color: #111;
      line-height: 1.15;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .voter-name {{
      font-size: {f('name', '19px')};
      font-weight: 900;
      color: #000;
      line-height: 1.05;
      letter-spacing: 0.2px;
      margin-bottom: 0px;
    }}
    .voter-line strong {{ color: #000; font-weight: 900; }}

    .booth-line {{
      font-size: {f('booth', '12px')};
      font-weight: 900;
      line-height: 1.1;
      border-top: 1.5px solid #111;
      padding-top: 1px;
      color: #000;
      word-break: break-word;
      /* Always show the full booth address from the roll, however many
         lines it takes — never truncated, never clipped. */
      flex-shrink: 0;
      min-height: 0;
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
      align-items: center;
      justify-content: center;
      text-align: center;
      overflow: hidden;
    }}
    /* Fixed gap, centred as one block. Whether the party symbol is
       present or absent, the spacing between every other label stays
       identical — only the block's total height changes, and any
       leftover height sits above/below the block, never inside it. */
    .cand-stack {{
      width: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: {f('c_gap', '3px')};
    }}
    .cand-post {{
      font-size: {f('c_post', '13px')};
      font-weight: 900;
      color: #b71c1c;
      line-height: 1.1;
      width: 100%;
      white-space: normal;
      word-break: break-word;
      overflow: visible;
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
