import http.server
import socketserver
import socket
import json
import os
import sys
import subprocess
import tempfile
import base64
from urllib.parse import quote
import openpyxl

PORT = 5000
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)
DOWNLOADS_DIR = os.path.join(SCRIPT_DIR, "downloads")
WEB_DIR = os.path.join(SCRIPT_DIR, "web")
UPLOADS_DIR = os.path.join(SCRIPT_DIR, "uploads")

os.makedirs(DOWNLOADS_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)

try:
    from voter_suvidha.extractor import extract_pdf_elector_data, export_voters_to_excel
except ImportError:
    from extractor import extract_pdf_elector_data, export_voters_to_excel

# Load master voter data from verified Excel
excel_path = os.path.join(WORKSPACE_DIR, "beawar_ward_001_part_001.xlsx")
if not os.path.exists(excel_path):
    excel_path = os.path.join(SCRIPT_DIR, "template.xlsx")
wb = openpyxl.load_workbook(excel_path, read_only=True)
ws = wb.active

voters_list = []
rows = list(ws.iter_rows(values_only=True))
header = rows[0]
for r in rows[1:]:
    if r[0] is not None:
        voters_list.append({
            "SerialNo": r[4] if len(r) > 4 and r[4] is not None else r[0],
            "Ward": str(r[1]) if len(r) > 1 and r[1] is not None else "1",
            "Part": str(r[2]) if len(r) > 2 and r[2] is not None else "1",
            "Booth": str(r[3] or "") if len(r) > 3 else "1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)",
            "VoterName": str(r[5] or "") if len(r) > 5 else "",
            "RelativeType": "पिता/पति",
            "RelativeName": str(r[6] or "") if len(r) > 6 else "",
            "HouseNo": str(r[7] or "") if len(r) > 7 else "",
            "Age": str(r[8] or "") if len(r) > 8 else "",
            "Gender": str(r[9] or "") if len(r) > 9 else "",
            "EPIC": str(r[10] or "") if len(r) > 10 else ""
        })
wb.close()

active_voters = voters_list
deleted_voters = []
print(f"Loaded {len(voters_list)} active voters from beawar Excel.")

global_session = {
    "totalSerials": len(voters_list),
    "activeVoters": active_voters,
    "deletedVoters": deleted_voters,
    "ward": "001",
    "parts": [
        {
            "part": 1,
            "booth": "1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)",
            "totalVoters": len(voters_list),
            "activeVoters": len(active_voters)
        }
    ]
}

def get_grid_and_font(slips_per_page):
    grid_css = {
        4: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(2, 1fr); gap: 4mm 5mm;",
        6: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(3, 1fr); gap: 3.5mm 4.5mm;",
        8: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(4, 1fr); gap: 2.8mm 4mm;",
        10: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(5, 1fr); gap: 2.2mm 3.5mm;",
        12: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(6, 1fr); gap: 1.8mm 3mm;"
    }.get(slips_per_page, "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(4, 1fr); gap: 2.8mm 4mm;")

    font_scale = {
        4: {
            "title": "16px", "meta": "13px", "name": "18px", "detail": "13.8px", "booth": "12.5px",
            "c_post": "14px", "c_name": "17.5px", "c_party": "13.5px", "c_app": "11.5px",
            "img_w": "95px", "img_h": "100px", "sym_w": "82px", "sym_h": "82px", "avatar": "60px"
        },
        6: {
            "title": "15px", "meta": "12px", "name": "16.5px", "detail": "12.5px", "booth": "11px",
            "c_post": "12.5px", "c_name": "15px", "c_party": "12px", "c_app": "10px",
            "img_w": "82px", "img_h": "86px", "sym_w": "68px", "sym_h": "68px", "avatar": "52px"
        },
        8: {
            "title": "13.5px", "meta": "10.5px", "name": "14.5px", "detail": "11px", "booth": "9.8px",
            "c_post": "11.5px", "c_name": "13px", "c_party": "10.5px", "c_app": "8.8px",
            "img_w": "68px", "img_h": "72px", "sym_w": "56px", "sym_h": "56px", "avatar": "42px"
        },
        10: {
            "title": "12px", "meta": "9.5px", "name": "13px", "detail": "9.8px", "booth": "8.8px",
            "c_post": "10px", "c_name": "11.5px", "c_party": "9.5px", "c_app": "7.8px",
            "img_w": "56px", "img_h": "58px", "sym_w": "44px", "sym_h": "44px", "avatar": "34px"
        },
        12: {
            "title": "10.5px", "meta": "8.8px", "name": "11.8px", "detail": "8.8px", "booth": "8px",
            "c_post": "9px", "c_name": "10.5px", "c_party": "8.5px", "c_app": "7.2px",
            "img_w": "48px", "img_h": "50px", "sym_w": "38px", "sym_h": "38px", "avatar": "30px"
        }
    }.get(slips_per_page, {
        "title": "13.5px", "meta": "10.5px", "name": "14.5px", "detail": "11px", "booth": "9.8px",
        "c_post": "11.5px", "c_name": "13px", "c_party": "10.5px", "c_app": "8.8px",
        "img_w": "68px", "img_h": "72px", "sym_w": "56px", "sym_h": "56px", "avatar": "42px"
    })

    return grid_css, font_scale

def render_slip_html(v, candidate, party, appeal, candidate_photo, party_symbol, font_scale, candidate_post="सरपंच"):
    photo_html = f'<img class="cand-photo" src="{candidate_photo}" alt="Candidate">' if candidate_photo else '<div class="cand-avatar">👤</div>'
    symbol_html = f'<div class="cand-symbol-frame"><img class="cand-symbol-img" src="{party_symbol}" alt="चुनाव चिन्ह"></div>' if party_symbol else ''

    return f"""
        <div class="slip">
          <div class="slip-left">
            <div class="slip-header-block">
              <div class="slip-title">वोटर सुविधा स्लिप</div>
              <div class="meta-row">
                <span class="badge-item">वार्ड नं : <strong>[ {v.get('Ward', '1')} ]</strong></span>
                <span class="badge-item">भाग : <strong>[ {v.get('Part', '1')} ]</strong></span>
              </div>
              <div class="meta-row-serial">
                <span class="badge-item serial-badge">क्रम संख्या : <strong>[ {v.get('SerialNo', '')} ]</strong></span>
              </div>
            </div>
            <div class="voter-body">
              <div class="voter-line voter-name">मतदाता का नाम : <strong>{v.get('VoterName', '')}</strong></div>
              <div class="voter-line">{v.get('RelativeType', 'पिता')} का नाम : <span>{v.get('RelativeName', '')}</span></div>
              <div class="voter-line">मकान नं. : <strong>{v.get('HouseNo', '')}</strong> &nbsp;|&nbsp; उम्र : <strong>{v.get('Age', '')}</strong> &nbsp;|&nbsp; लिंग : <strong>{v.get('Gender', '')}</strong></div>
              <div class="voter-line">पहचान पत्र क्र. (EPIC) : <strong>{v.get('EPIC', '')}</strong></div>
            </div>
            <div class="booth-line">
              <span class="booth-label">मतदान केंद्र :</span> {v.get('Booth', '')}
            </div>
          </div>
          <div class="slip-right-box">
            <div class="cand-post"><strong>{candidate_post} पद हेतु</strong></div>
            <div class="cand-photo-frame">
              {photo_html}
            </div>
            <div class="cand-name">{candidate}</div>
            {symbol_html}
            <div class="cand-party">({party})</div>
            <div class="cand-appeal">{appeal}</div>
          </div>
        </div>
    """

def get_full_page_css(grid_css, font_scale):
    return f"""
    @page {{ size: A4 portrait; margin: 0; }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    body {{ font-family: "Nirmala UI", "Mangal", "Segoe UI", Arial, sans-serif; background: #fff; color: #111; }}
    .a4-page {{ width: 210mm; height: 297mm; padding: 5mm 5mm; page-break-after: always; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }}
    .slips-grid {{ display: grid; {grid_css} width: 100%; height: 100%; }}
    .slip {{
      border: 1.5px solid #111;
      border-radius: 5px;
      padding: 4px 5px;
      display: flex;
      flex-direction: row;
      justify-content: space-between;
      align-items: stretch;
      gap: 6px;
      background: #fff;
      overflow: hidden;
    }}
    .slip-left {{
      width: 63%;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      overflow: hidden;
    }}
    .slip-header-block {{
      border-bottom: 1.2px dashed #444;
      padding-bottom: 2px;
      margin-bottom: 1.5px;
    }}
    .slip-title {{
      text-align: center;
      font-weight: 900;
      font-size: {font_scale['title']};
      letter-spacing: 0.6px;
      color: #000;
      margin-bottom: 1.5px;
    }}
    .meta-row {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: {font_scale['meta']};
      margin-bottom: 1.5px;
    }}
    .meta-row-serial {{
      display: flex;
      justify-content: flex-start;
      align-items: center;
    }}
    .badge-item {{
      color: #111;
    }}
    .badge-item strong {{
      color: #000;
    }}
    .serial-badge {{
      background: #f4f5f7;
      border: 1.3px solid #111;
      border-radius: 3px;
      padding: 0.5px 5px;
      font-size: {font_scale['meta']};
      font-weight: 800;
      display: inline-block;
    }}
    .voter-body {{
      flex: 1;
      display: flex;
      flex-direction: column;
      justify-content: space-evenly;
      padding: 1px 0;
    }}
    .voter-line {{
      font-size: {font_scale['detail']};
      color: #111;
      line-height: 1.25;
    }}
    .voter-name {{
      font-size: {font_scale['name']};
      font-weight: 900;
      color: #000;
    }}
    .voter-line strong {{
      color: #000;
    }}
    .booth-line {{
      font-size: {font_scale['booth']};
      line-height: 1.22;
      border-top: 1.2px dashed #666;
      padding-top: 2px;
      margin-top: 1px;
      color: #000;
      background: #fafafa;
      border-radius: 2px;
      padding-left: 2px;
    }}
    .booth-label {{
      font-weight: 900;
      color: #000;
    }}
    .slip-right-box {{
      width: 36%;
      min-width: 36%;
      max-width: 37%;
      border: 1.8px solid #1a237e;
      border-radius: 6px;
      background: #fbfbfd;
      padding: 3px 2px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: space-around;
      text-align: center;
      box-sizing: border-box;
    }}
    .cand-post {{
      font-size: {font_scale['c_post']};
      font-weight: 900;
      color: #b71c1c;
      line-height: 1.15;
      margin-bottom: 2px;
      text-align: center;
      width: 100%;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      letter-spacing: 0.2px;
    }}
    .cand-photo-frame {{
      width: {font_scale['img_w']};
      height: {font_scale['img_h']};
      border-radius: 4px;
      border: 1.2px solid #7986cb;
      background: #e8eaf6;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
      flex-shrink: 0;
    }}
    .cand-photo {{
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
    }}
    .cand-avatar {{
      font-size: {font_scale['avatar']};
      line-height: 1;
    }}
    .cand-name {{
      font-size: {font_scale['c_name']};
      font-weight: 900;
      color: #0d47a1;
      line-height: 1.15;
      margin: 1px 0;
    }}
    .cand-symbol-frame {{
      width: {font_scale['sym_w']};
      height: {font_scale['sym_h']};
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 1px 0;
      flex-shrink: 0;
    }}
    .cand-symbol-img {{
      width: 100%;
      height: 100%;
      object-fit: contain;
    }}
    .cand-party {{
      font-size: {font_scale['c_party']};
      font-weight: 800;
      color: #2e7d32;
      line-height: 1.15;
    }}
    .cand-appeal {{
      font-size: {font_scale['c_app']};
      font-weight: 800;
      color: #b71c1c;
      line-height: 1.15;
      margin-top: 1px;
    }}
    """

def find_browser():
    candidates = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    import shutil
    for name in ["chrome", "msedge", "google-chrome", "chromium"]:
        p = shutil.which(name)
        if p:
            return p
    return None

def save_base64_image_to_temp_file(data_uri, prefix):
    if not data_uri or not isinstance(data_uri, str):
        return ""
    if data_uri.startswith("data:image/"):
        try:
            comma_idx = data_uri.find(",")
            if comma_idx != -1:
                header = data_uri[:comma_idx]
                raw_b64 = data_uri[comma_idx + 1:]
                ext = ".jpg"
                if "png" in header:
                    ext = ".png"
                elif "webp" in header:
                    ext = ".webp"
                elif "svg" in header:
                    ext = ".svg"
                temp_path = os.path.join(tempfile.gettempdir(), f"{prefix}{ext}")
                with open(temp_path, "wb") as f:
                    f.write(base64.b64decode(raw_b64))
                norm_path = os.path.abspath(temp_path).replace("\\", "/")
                return f"file:///{quote(norm_path, safe=':/')}"
        except Exception as e:
            print(f"Warning: Failed to cache temp image: {e}")
            return data_uri
    return data_uri

class VoterSuvidhaHandler(http.server.SimpleHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def do_GET(self):
        clean_path = self.path.split('?', 1)[0].split('#', 1)[0]
        if clean_path.startswith('/downloads/'):
            filename = os.path.basename(clean_path)
            target_path = None
            for d in [DOWNLOADS_DIR, WORKSPACE_DIR, os.path.join(WEB_DIR, "downloads")]:
                candidate = os.path.join(d, filename)
                if os.path.exists(candidate) and os.path.isfile(candidate):
                    target_path = candidate
                    break
            if target_path:
                self.serve_file_with_range(target_path)
                return
            else:
                self.send_error(404, f"File {filename} not found")
                return
        super().do_GET()

    def serve_file_with_range(self, file_path):
        file_size = os.path.getsize(file_path)
        filename = os.path.basename(file_path)

        if filename.endswith(".pdf"):
            content_type = "application/pdf"
        elif filename.endswith(".xlsx"):
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif filename.endswith(".png"):
            content_type = "image/png"
        elif filename.endswith(".jpg") or filename.endswith(".jpeg"):
            content_type = "image/jpeg"
        else:
            content_type = "application/octet-stream"

        range_header = self.headers.get("Range")
        start = 0
        end = file_size - 1
        status_code = 200

        if range_header and range_header.startswith("bytes="):
            try:
                ranges = range_header.replace("bytes=", "").split("-")
                if ranges[0]:
                    start = int(ranges[0])
                if len(ranges) > 1 and ranges[1]:
                    end = int(ranges[1])
                if start <= end and end < file_size:
                    status_code = 206
            except Exception:
                start = 0
                end = file_size - 1
                status_code = 200

        length = end - start + 1
        self.send_response(status_code)
        self.send_header("Content-Type", content_type)
        self.send_header("Accept-Ranges", "bytes")
        self.send_header("Content-Length", str(length))
        if status_code == 206:
            self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
        
        self.send_header("Content-Disposition", f'inline; filename="{filename}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Connection", "close")
        self.close_connection = True
        self.end_headers()

        with open(file_path, "rb") as f:
            f.seek(start)
            remaining = length
            chunk_size = 131072
            while remaining > 0:
                to_read = min(chunk_size, remaining)
                buf = f.read(to_read)
                if not buf:
                    break
                self.wfile.write(buf)
                remaining -= len(buf)
        try:
            self.wfile.flush()
        except Exception:
            pass

    def finish(self):
        try:
            self.wfile.flush()
        except Exception:
            pass
        try:
            self.connection.shutdown(socket.SHUT_WR)
        except Exception:
            pass
        super().finish()

    def translate_path(self, path):
        clean_path = path.split('?', 1)[0].split('#', 1)[0]
        if clean_path.startswith('/downloads/'):
            filename = os.path.basename(clean_path)
            for d in [DOWNLOADS_DIR, WORKSPACE_DIR, os.path.join(WEB_DIR, "downloads")]:
                candidate = os.path.join(d, filename)
                if os.path.exists(candidate):
                    return candidate
        return super().translate_path(path)

    def send_json_response(self, obj, status_code=200):
        body = json.dumps(obj).encode('utf-8')
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)
        try:
            self.wfile.flush()
        except Exception:
            pass

    def send_html_response(self, html_str, status_code=200):
        body = html_str.encode('utf-8')
        self.send_response(status_code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)
        try:
            self.wfile.flush()
        except Exception:
            pass

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        post_data = self.rfile.read(content_length)

        if self.path in ["/api/upload", "/api/process-pdfs"]:
            self.handle_upload(post_data)
        elif self.path == "/api/preview":
            payload = json.loads(post_data.decode("utf-8")) if post_data else {}
            self.handle_preview(payload)
        elif self.path == "/api/generate-excel":
            payload = json.loads(post_data.decode("utf-8")) if post_data else {}
            self.handle_generate_excel(payload)
        elif self.path == "/api/generate-pdf":
            payload = json.loads(post_data.decode("utf-8")) if post_data else {}
            self.handle_generate_pdf(payload)
        else:
            self.send_error(404, "Unknown API endpoint")

    def handle_upload(self, post_data):
        import base64
        # Completely clear existing session memory / cache on every upload
        global_session.clear()
        global_session.update({
            "totalSerials": 0,
            "activeVoters": [],
            "deletedVoters": [],
            "ward": "",
            "parts": []
        })

        # Clear previous uploaded temp PDF files from UPLOADS_DIR
        try:
            for old_f in os.listdir(UPLOADS_DIR):
                if old_f.startswith("uploaded_") or old_f.endswith(".tmp"):
                    old_p = os.path.join(UPLOADS_DIR, old_f)
                    if os.path.isfile(old_p):
                        os.remove(old_p)
        except Exception as e:
            pass

        parts_list = []
        all_active_voters = []
        detected_ward = ""
        total_serials_sum = 0
        total_deleted_sum = 0

        try:
            payload = json.loads(post_data.decode("utf-8")) if post_data else {}
            files = payload.get("files", [])
            for idx, f in enumerate(files, 1):
                fn = f.get("filename", f"uploaded_part_{idx}.pdf")
                b64_data = f.get("data", "")
                if b64_data:
                    target_path = os.path.join(UPLOADS_DIR, fn)
                    with open(target_path, "wb") as upf:
                        upf.write(base64.b64decode(b64_data))
                    print(f"Saved uploaded PDF: {fn}")

                    extracted = extract_pdf_elector_data(target_path)
                    part_num = extracted.get("part", str(idx))
                    booth_name = extracted.get("booth", "")
                    part_serials = extracted.get("totalSerials", 0)
                    part_del = extracted.get("deletedCount", 0)
                    part_act = extracted.get("activeCount", 0)
                    part_voters = extracted.get("voters", [])

                    if not detected_ward:
                        detected_ward = extracted.get("ward", "1")

                    parts_list.append({
                        "part": part_num,
                        "booth": booth_name,
                        "totalSerials": part_serials,
                        "deletedCount": part_del,
                        "activeCount": part_act
                    })

                    total_serials_sum += part_serials
                    total_deleted_sum += part_del
                    all_active_voters.extend(part_voters)

        except Exception as e:
            print(f"Upload processing error: {e}")
            self.send_json_response({"success": False, "error": str(e)}, status_code=500)
            return

        global_session["ward"] = detected_ward if detected_ward else "1"
        global_session["totalSerials"] = total_serials_sum
        global_session["activeVoters"] = all_active_voters
        global_session["parts"] = parts_list
        global_session["deletedVoters"] = list(range(total_deleted_sum))

        resp_obj = {
            "success": True,
            "ward": global_session["ward"],
            "totalSerials": global_session["totalSerials"],
            "totalDeleted": total_deleted_sum,
            "totalActive": len(all_active_voters),
            "activeVoters": len(all_active_voters),
            "deletedVoters": total_deleted_sum,
            "parts": parts_list,
            "message": f"मतदाता सूची (वार्ड {global_session['ward']}) सफलतापूर्वक विश्लेषित!"
        }
        self.send_json_response(resp_obj)

    def handle_preview(self, payload):
        candidate_post = payload.get("candidatePost", "सरपंच")
        candidate = payload.get("candidateName", "मनोज बाबेल")
        party = payload.get("partyName", "भारतीय जनता पार्टी (BJP)")
        appeal = payload.get("bottomMessage", "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!")
        slips_per_page = int(payload.get("slipsPerPage", 8))
        candidate_photo = payload.get("candidatePhoto", "")
        party_symbol = payload.get("partySymbol", "")

        # Apply booth overrides if provided
        if payload.get("parts"):
            part_map = {str(p.get("part")): p.get("booth") for p in payload["parts"] if p.get("booth")}
            for v in global_session["activeVoters"]:
                vp = str(v.get("Part", "1"))
                if vp in part_map:
                    v["Booth"] = part_map[vp]

        voters = global_session["activeVoters"][:slips_per_page]
        grid_css, font_scale = get_grid_and_font(slips_per_page)
        
        slips_html = ""
        for v in voters:
            slips_html += render_slip_html(v, candidate, party, appeal, candidate_photo, party_symbol, font_scale, candidate_post)

        css_text = get_full_page_css(grid_css, font_scale)
        html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <title>वोटर सुविधा स्लिप</title>
  <style>
{css_text}
  </style>
</head>
<body>
  <div class="a4-page">
    <div class="slips-grid">
      {slips_html}
    </div>
  </div>
</body>
</html>
"""
        self.send_html_response(html)

    def handle_generate_excel(self, payload):
        # Apply booth overrides if provided
        if payload.get("parts"):
            part_map = {str(p.get("part")): p.get("booth") for p in payload["parts"] if p.get("booth")}
            for v in global_session["activeVoters"]:
                vp = str(v.get("Part", "1"))
                if vp in part_map:
                    v["Booth"] = part_map[vp]

        ward_num = global_session.get("ward", "1")
        out_filename = f"voter_list_ward_{ward_num}.xlsx"
        dst_excel = os.path.join(DOWNLOADS_DIR, out_filename)
        ws_excel = os.path.join(WORKSPACE_DIR, out_filename)
        web_dl_dir = os.path.join(WEB_DIR, "downloads")
        os.makedirs(web_dl_dir, exist_ok=True)
        web_excel = os.path.join(web_dl_dir, out_filename)

        try:
            export_voters_to_excel(global_session["activeVoters"], dst_excel)
            import shutil
            shutil.copy(dst_excel, ws_excel)
            shutil.copy(dst_excel, web_excel)
        except Exception as e:
            print(f"Excel export error: {e}")
            self.send_json_response({"success": False, "error": str(e)}, status_code=500)
            return

        resp_obj = {
            "success": True,
            "totalVoters": len(global_session["activeVoters"]),
            "filename": out_filename,
            "downloadUrl": f"/downloads/{out_filename}"
        }
        self.send_json_response(resp_obj)

    def handle_generate_pdf(self, payload):
        candidate_post = payload.get("candidatePost", "सरपंच")
        candidate = payload.get("candidateName", "मनोज बाबेल")
        party = payload.get("partyName", "भारतीय जनता पार्टी (BJP)")
        appeal = payload.get("bottomMessage", "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!")
        slips_per_page = int(payload.get("slipsPerPage", 8))
        candidate_photo = payload.get("candidatePhoto", "")
        party_symbol = payload.get("partySymbol", "")

        # Apply booth overrides if provided
        if payload.get("parts"):
            part_map = {str(p.get("part")): p.get("booth") for p in payload["parts"] if p.get("booth")}
            for v in global_session["activeVoters"]:
                vp = str(v.get("Part", "1"))
                if vp in part_map:
                    v["Booth"] = part_map[vp]

        ward_num = global_session.get("ward", "1")
        out_filename = f"voter_slips_ward_{ward_num}.pdf"
        dst_pdf = os.path.join(DOWNLOADS_DIR, out_filename)

        self.generate_custom_pdf(
            global_session["activeVoters"], candidate, party, appeal,
            slips_per_page, candidate_photo, party_symbol, dst_pdf, candidate_post
        )

        import shutil
        ws_pdf = os.path.join(WORKSPACE_DIR, out_filename)
        web_dl_dir = os.path.join(WEB_DIR, "downloads")
        os.makedirs(web_dl_dir, exist_ok=True)
        shutil.copy(dst_pdf, ws_pdf)
        shutil.copy(dst_pdf, os.path.join(web_dl_dir, out_filename))

        import math
        total_pages = math.ceil(len(global_session["activeVoters"]) / slips_per_page)

        resp_obj = {
            "success": True,
            "totalVoters": len(global_session["activeVoters"]),
            "totalPages": total_pages,
            "filename": out_filename,
            "downloadUrl": f"/downloads/{out_filename}"
        }
        self.send_json_response(resp_obj)

    def generate_custom_pdf(self, voters, candidate, party, appeal, slips_per_page, candidate_photo, party_symbol, out_pdf_path, candidate_post="सरपंच"):
        temp_html = os.path.join(tempfile.gettempdir(), "voter_slips_render.html")
        chunk_size = slips_per_page
        chunks = [voters[i:i + chunk_size] for i in range(0, len(voters), chunk_size)]

        grid_css, font_scale = get_grid_and_font(slips_per_page)
        css_text = get_full_page_css(grid_css, font_scale)

        # Cache base64 images once to temp file to save memory
        cand_photo_url = save_base64_image_to_temp_file(candidate_photo, "voter_cand_photo")
        party_sym_url = save_base64_image_to_temp_file(party_symbol, "voter_party_symbol")

        html_body = ""
        for chunk in chunks:
            slips_html = ""
            for v in chunk:
                slips_html += render_slip_html(v, candidate, party, appeal, cand_photo_url, party_sym_url, font_scale, candidate_post)

            if len(chunk) < chunk_size:
                for _ in range(chunk_size - len(chunk)):
                    slips_html += '<div class="slip" style="visibility:hidden;"></div>'

            html_body += f"""
      <div class="a4-page">
        <div class="slips-grid">
          {slips_html}
        </div>
      </div>
            """

        full_doc = f"""<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <title>वोटर सुविधा स्लिप</title>
  <style>
{css_text}
  </style>
</head>
<body>
{html_body}
</body>
</html>
"""
        with open(temp_html, "w", encoding="utf-8") as f:
            f.write(full_doc)

        browser_exe = find_browser()
        if not browser_exe:
            raise RuntimeError("Neither Google Chrome nor Microsoft Edge was found for offline PDF generation.")

        cmd = [
            browser_exe,
            "--headless",
            "--disable-gpu",
            "--allow-file-access-from-files",
            "--no-pdf-header-footer",
            f"--print-to-pdf={out_pdf_path}",
            temp_html
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0 and not os.path.exists(out_pdf_path):
            raise RuntimeError(f"Browser PDF printing failed: {res.stderr}")
        print(f"Generated PDF: {out_pdf_path}")

class ThreadedHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True

if __name__ == "__main__":
    import webbrowser
    with ThreadedHTTPServer(("127.0.0.1", PORT), VoterSuvidhaHandler) as httpd:
        print(f"Voter Suvidha server running on http://127.0.0.1:{PORT}")
        sys.stdout.flush()
        try:
            webbrowser.open(f"http://127.0.0.1:{PORT}")
        except Exception:
            pass
        httpd.serve_forever()
