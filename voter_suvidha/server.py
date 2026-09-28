import http.server
import socketserver
import socket
import json
import os
import re
import sys
import subprocess
import tempfile
import base64
from urllib.parse import quote, unquote
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
    from voter_suvidha.extractor import extract_pdf_elector_data, export_voters_to_excel, read_voters_from_excel
except ImportError:
    from extractor import extract_pdf_elector_data, export_voters_to_excel, read_voters_from_excel

try:
    from voter_suvidha.voter_photo_extractor import extract_voter_photos
except ImportError:
    from voter_photo_extractor import extract_voter_photos

PHOTOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads", "voter_photos")

def _sanitize_filename_part(text):
    """Strip characters Windows forbids in filenames and tidy whitespace."""
    text = (text or "").strip()
    text = re.sub(r'[\\/:*?"<>|]+', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def build_voter_list_basename(ward_num, gram_panchayat, panchayat_samiti):
    """पंचायत समिति-ग्राम पंचायत(गांव)-वार्ड नं के फॉर्मेट में फाइल का नाम
    बनाता है, जैसे 'जैतारण-फालका-9.xlsx' / 'जैतारण-फालका-9.pdf'। जो हिस्सा
    PDF से नहीं मिला उसकी जगह एक सुरक्षित default डाला जाता है ताकि नाम
    कभी खाली न रहे।"""
    ps = _sanitize_filename_part(panchayat_samiti) or "पंचायत-समिति"
    gp = _sanitize_filename_part(gram_panchayat) or "गांव"
    ward = _sanitize_filename_part(str(ward_num)) or "1"
    return f"{ps}-{gp}-{ward}"

# Clean empty global session - only populated when user uploads a PDF
global_session = {
    "totalSerials": 0,
    "activeVoters": [],
    "deletedVoters": [],
    "ward": "",
    "parts": [],
    "voterPhotosByEpic": {},
    "voterPhotosBySerial": {}
}

def get_grid_and_font(slips_per_page):
    grid_css = {
        4: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(2, 1fr); gap: 4mm 5mm;",
        6: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(3, 1fr); gap: 3.5mm 4.5mm;",
        8: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(4, 1fr); gap: 2.5mm 3.5mm;",
        10: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(5, 1fr); gap: 2mm 3mm;",
        12: "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(6, 1fr); gap: 1.6mm 2.5mm;"
    }.get(slips_per_page, "grid-template-columns: 1fr 1fr; grid-template-rows: repeat(4, 1fr); gap: 2.5mm 3.5mm;")

    font_scale = {
        4: {
            "title": "18px", "panchayat": "14.5px", "meta": "14.5px", "serial": "15.5px",
            "name": "22px", "detail": "15.5px", "booth": "14.5px",
            "c_post": "15.5px", "c_name": "20px", "c_party": "15px", "c_app": "13px",
            "img_w": "95px", "img_h": "102px", "sym_w": "80px", "sym_h": "80px", "avatar": "60px", "cut": "9px",
            "vphoto_w": "88px", "vphoto_h": "100px"
        },
        6: {
            "title": "16.5px", "panchayat": "13.5px", "meta": "13.5px", "serial": "14.5px",
            "name": "20px", "detail": "14px", "booth": "13.2px",
            "c_post": "14px", "c_name": "18px", "c_party": "13.5px", "c_app": "11.5px",
            "img_w": "80px", "img_h": "86px", "sym_w": "68px", "sym_h": "68px", "avatar": "52px", "cut": "8px",
            "vphoto_w": "74px", "vphoto_h": "84px"
        },
        8: {
            "title": "15px", "panchayat": "12px", "meta": "12.5px", "serial": "13.5px",
            "name": "18.5px", "detail": "12.7px", "booth": "12px",
            "c_post": "13px", "c_name": "16px", "c_party": "11.8px", "c_app": "10.5px",
            "img_w": "62px", "img_h": "68px", "sym_w": "48px", "sym_h": "48px", "avatar": "46px", "cut": "7.5px",
            "vphoto_w": "58px", "vphoto_h": "66px"
        },
        10: {
            "title": "13px", "panchayat": "11px", "meta": "11px", "serial": "12px",
            "name": "15.5px", "detail": "11.5px", "booth": "11.5px",
            "c_post": "11px", "c_name": "13.5px", "c_party": "10.5px", "c_app": "9px",
            "img_w": "50px", "img_h": "54px", "sym_w": "42px", "sym_h": "42px", "avatar": "34px", "cut": "6.5px",
            "vphoto_w": "46px", "vphoto_h": "52px"
        },
        12: {
            "title": "11.5px", "panchayat": "9.8px", "meta": "10px", "serial": "11px",
            "name": "13.5px", "detail": "10.2px", "booth": "7.0px",
            "c_post": "10px", "c_name": "12px", "c_party": "9.5px", "c_app": "8px",
            "img_w": "44px", "img_h": "48px", "sym_w": "36px", "sym_h": "36px", "avatar": "30px", "cut": "6px",
            "vphoto_w": "40px", "vphoto_h": "46px"
        }
    }.get(slips_per_page, {
        "title": "15px", "panchayat": "12px", "meta": "12.5px", "serial": "13.5px",
        "name": "18px", "detail": "13px", "booth": "12.2px",
        "c_post": "12.5px", "c_name": "15px", "c_party": "11px", "c_app": "10px",
        "img_w": "58px", "img_h": "64px", "sym_w": "48px", "sym_h": "48px", "avatar": "42px", "cut": "7.2px",
        "vphoto_w": "54px", "vphoto_h": "62px"
    })

    return grid_css, font_scale

from slip_layout import render_slip_html, get_full_page_css  # new slip design (see slip_layout.py)

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
                raw_b64 = data_uri[comma_idx + 1:].strip()
                # A data URI can carry the "data:image/png;base64," HEADER
                # with nothing after the comma (an empty payload) — e.g. if
                # the browser starts building the string before a file is
                # actually chosen, or a selection gets cleared. That string
                # is non-empty and passes every check above, so without this
                # guard we decode an empty payload (no error), write a
                # ZERO-BYTE file, and hand back a file:// URL that looks
                # perfectly valid — Chrome then renders its own broken-image
                # icon for it. Treat an empty payload as "no image" instead.
                if not raw_b64:
                    return ""
                image_bytes = base64.b64decode(raw_b64)
                if not image_bytes:
                    return ""
                ext = ".jpg"
                if "png" in header:
                    ext = ".png"
                elif "webp" in header:
                    ext = ".webp"
                elif "svg" in header:
                    ext = ".svg"
                temp_path = os.path.join(tempfile.gettempdir(), f"{prefix}{ext}")
                with open(temp_path, "wb") as f:
                    f.write(image_bytes)
                norm_path = os.path.abspath(temp_path).replace("\\", "/")
                return f"file:///{quote(norm_path, safe=':/')}"
        except Exception as e:
            print(f"Warning: Failed to cache temp image: {e}")
            return data_uri
    return data_uri

def path_to_file_url(path):
    """Turn an on-disk photo path (already extracted from the WithPhoto
    PDF) into the same file:/// form the browser-print step expects for
    every other image on the slip."""
    if not path or not os.path.exists(path):
        return ""
    norm_path = os.path.abspath(path).replace("\\", "/")
    return f"file:///{quote(norm_path, safe=':/')}"

def resolve_voter_photo_url(v, by_epic, by_serial):
    """EPIC is unique across the whole roll, so it is tried first; the
    serial-number map is only a fallback for supplement entries that carry
    no EPIC yet (see voter_photo_extractor.py)."""
    if not by_epic and not by_serial:
        return ""
    epic = str(v.get("EPIC", "") or "").strip()
    if epic and epic in by_epic:
        return path_to_file_url(by_epic[epic])
    serial = str(v.get("SerialNo", "") or "").strip()
    if serial and serial in by_serial:
        return path_to_file_url(by_serial[serial])
    return ""

class VoterSuvidhaHandler(http.server.SimpleHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def do_GET(self):
        clean_path = self.path.split('?', 1)[0].split('#', 1)[0]
        if clean_path.startswith('/downloads/'):
            # self.path is percent-encoded (spaces -> %20, Devanagari -> %E0..),
            # so it must be decoded before it can match the real filename on disk.
            filename = os.path.basename(unquote(clean_path))
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

        # HTTP headers must be latin-1 — Devanagari filenames (जैतारण-फालका-9.pdf)
        # can't go in a plain filename="..." param, so we send an ASCII fallback
        # plus the real UTF-8 name via the RFC 6266 filename*= form.
        ascii_fallback = filename.encode('ascii', 'ignore').decode('ascii').strip() or "voter-suvidha-download"
        self.send_header(
            "Content-Disposition",
            f"inline; filename=\"{ascii_fallback}\"; filename*=UTF-8''{quote(filename)}"
        )
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
            filename = os.path.basename(unquote(clean_path))
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

        if self.path in ["/api/clear-session", "/api/reset"]:
            self.handle_clear_session()
        elif self.path in ["/api/upload", "/api/process-pdfs"]:
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

    def handle_clear_session(self):
        import gc
        global_session.clear()
        global_session.update({
            "totalSerials": 0,
            "activeVoters": [],
            "deletedVoters": [],
            "ward": "",
            "parts": [],
            "gramPanchayat": "",
            "panchayatSamiti": "",
            "voterPhotosByEpic": {},
            "voterPhotosBySerial": {}
        })
        try:
            for old_f in os.listdir(UPLOADS_DIR):
                old_p = os.path.join(UPLOADS_DIR, old_f)
                if os.path.isfile(old_p):
                    os.remove(old_p)
        except Exception:
            pass
        try:
            import shutil
            if os.path.isdir(PHOTOS_DIR):
                shutil.rmtree(PHOTOS_DIR, ignore_errors=True)
        except Exception:
            pass
        gc.collect()
        self.send_json_response({"success": True, "message": "सत्र मेमोरी और कैशे पूरी तरह साफ़ कर दिया गया है।"})

    def handle_upload(self, post_data):
        import base64
        # Completely clear existing session memory / cache on every upload
        global_session.clear()
        global_session.update({
            "totalSerials": 0,
            "activeVoters": [],
            "deletedVoters": [],
            "ward": "",
            "parts": [],
            "gramPanchayat": "",
            "panchayatSamiti": "",
            "voterPhotosByEpic": {},
            "voterPhotosBySerial": {}
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
        detected_gp = ""
        detected_ps = ""
        total_serials_sum = 0
        total_deleted_sum = 0
        photos_by_epic = {}
        photos_by_serial = {}

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
                    if not detected_gp:
                        detected_gp = extracted.get("gramPanchayat", "")
                    if not detected_ps:
                        detected_ps = extracted.get("panchayatSamiti", "")

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

                    # यही अपलोड की गई PDF अगर फोटो वाली (WithPhoto) निर्वाचक
                    # नामावली है तो हर मतदाता की फोटो EPIC / क्रम संख्या से
                    # मिलाकर निकाल ली जाती है (voter_photo_extractor.py)।
                    # सामान्य (बिना फोटो वाली) PDF पर यह सिर्फ खाली dict
                    # लौटाता है, कोई dummy/placeholder फोटो नहीं जोड़ी जाती।
                    try:
                        epic_map, serial_map = extract_voter_photos(target_path, PHOTOS_DIR)
                        if epic_map or serial_map:
                            photos_by_epic.update(epic_map)
                            for s, p in serial_map.items():
                                photos_by_serial.setdefault(s, p)
                            print(f"Extracted {len(epic_map)} voter photos (by EPIC) from {fn}")
                    except Exception as photo_ex:
                        print(f"Voter-photo extraction skipped for {fn}: {photo_ex}")

        except Exception as e:
            print(f"Upload processing error: {e}")
            self.send_json_response({"success": False, "error": str(e)}, status_code=500)
            return

        global_session["voterPhotosByEpic"] = photos_by_epic
        global_session["voterPhotosBySerial"] = photos_by_serial

        global_session["ward"] = detected_ward if detected_ward else "1"
        global_session["gramPanchayat"] = detected_gp
        global_session["panchayatSamiti"] = detected_ps
        global_session["totalSerials"] = total_serials_sum
        global_session["activeVoters"] = all_active_voters
        global_session["parts"] = parts_list
        global_session["deletedVoters"] = list(range(total_deleted_sum))

        # फाइल नाम फॉर्मेट: पंचायत समिति नाम-गांव (ग्राम पंचायत) नाम-वार्ड नं
        ward_num = global_session["ward"]
        output_basename = build_voter_list_basename(ward_num, detected_gp, detected_ps)
        global_session["outputBasename"] = output_basename

        # Auto-generate 12-column master Excel file immediately upon upload
        out_excel_name = f"{output_basename}.xlsx"
        dst_excel = os.path.join(DOWNLOADS_DIR, out_excel_name)
        ws_excel = os.path.join(WORKSPACE_DIR, out_excel_name)
        web_dl_dir = os.path.join(WEB_DIR, "downloads")
        os.makedirs(web_dl_dir, exist_ok=True)
        web_excel = os.path.join(web_dl_dir, out_excel_name)

        excel_url = ""
        try:
            export_voters_to_excel(all_active_voters, dst_excel)
            import shutil
            shutil.copy(dst_excel, ws_excel)
            shutil.copy(dst_excel, web_excel)
            global_session["excelPath"] = dst_excel
            excel_url = f"/downloads/{out_excel_name}"
            print(f"Master 12-column Excel auto-generated: {dst_excel}")
        except Exception as ex:
            print(f"Initial Excel export error: {ex}")

        # Also persist metadata (ward, GP, PS, booth)
        meta_info = {
            "ward": global_session["ward"],
            "gramPanchayat": detected_gp,
            "panchayatSamiti": detected_ps,
            "totalSerials": total_serials_sum,
            "totalDeleted": total_deleted_sum,
            "totalActive": len(all_active_voters),
            "parts": parts_list
        }
        try:
            with open(os.path.join(DOWNLOADS_DIR, f"metadata_ward_{ward_num}.json"), "w", encoding="utf-8") as mf:
                json.dump(meta_info, mf, ensure_ascii=False, indent=2)
            with open(os.path.join(WORKSPACE_DIR, f"metadata_ward_{ward_num}.json"), "w", encoding="utf-8") as mf:
                json.dump(meta_info, mf, ensure_ascii=False, indent=2)
        except Exception:
            pass

        photos_matched = sum(1 for v in all_active_voters if resolve_voter_photo_url(v, photos_by_epic, photos_by_serial))

        resp_obj = {
            "success": True,
            "ward": global_session["ward"],
            "gramPanchayat": detected_gp,
            "panchayatSamiti": detected_ps,
            "totalSerials": global_session["totalSerials"],
            "totalDeleted": total_deleted_sum,
            "totalActive": len(all_active_voters),
            "activeVoters": len(all_active_voters),
            "deletedVoters": total_deleted_sum,
            "parts": parts_list,
            "excelUrl": excel_url,
            "excelFilename": out_excel_name,
            "photosMatched": photos_matched,
            "message": f"मतदाता सूची (वार्ड {global_session['ward']}) सफलतापूर्वक विश्लेषित एवं 12-कॉलम एक्सेल तैयार!" + (
                f" इस PDF में {photos_matched} मतदाताओं की फोटो भी मिल गई — पर्ची में अपने आप लग जाएगी।" if photos_matched else ""
            )
        }
        self.send_json_response(resp_obj)

    def _output_basename(self):
        """वर्तमान session के लिए 'पंचायत समिति-गांव-वार्ड नं' फॉर्मेट वाला
        बेस फाइल-नाम (extension के बिना), Excel और PDF दोनों के लिए एक जैसा।"""
        basename = global_session.get("outputBasename")
        if basename:
            return basename
        basename = build_voter_list_basename(
            global_session.get("ward", "1"),
            global_session.get("gramPanchayat", ""),
            global_session.get("panchayatSamiti", ""))
        global_session["outputBasename"] = basename
        return basename

    def get_voters_from_active_excel_or_session(self):
        ward_num = global_session.get("ward", "1")
        basename = global_session.get("outputBasename") or build_voter_list_basename(
            ward_num, global_session.get("gramPanchayat", ""), global_session.get("panchayatSamiti", ""))
        excel_candidates = [
            global_session.get("excelPath"),
            os.path.join(DOWNLOADS_DIR, f"{basename}.xlsx"),
            os.path.join(WORKSPACE_DIR, f"{basename}.xlsx"),
            os.path.join(WEB_DIR, "downloads", f"{basename}.xlsx"),
            # पुराने नामकरण से बनी फाइलों के साथ भी काम करता रहे (backward compatibility)
            os.path.join(DOWNLOADS_DIR, f"voter_list_ward_{ward_num}.xlsx"),
            os.path.join(WORKSPACE_DIR, f"voter_list_ward_{ward_num}.xlsx"),
            os.path.join(WEB_DIR, "downloads", f"voter_list_ward_{ward_num}.xlsx")
        ]
        active_excel = next((p for p in excel_candidates if p and os.path.exists(p)), None)

        voters = []
        if active_excel:
            try:
                print(f"Loading voter records directly from Excel: {active_excel}")
                voters = read_voters_from_excel(active_excel)
                if voters:
                    print(f"Successfully loaded {len(voters)} voters from Excel.")
            except Exception as e:
                print(f"Warning: Failed to read from Excel ({e}), falling back to session memory")

        if not voters:
            voters = global_session.get("activeVoters", [])

        # Recover GP / PS from metadata if not in session
        if not global_session.get("gramPanchayat") or not global_session.get("panchayatSamiti"):
            meta_candidates = [
                os.path.join(DOWNLOADS_DIR, f"metadata_ward_{ward_num}.json"),
                os.path.join(WORKSPACE_DIR, f"metadata_ward_{ward_num}.json")
            ]
            for mp in meta_candidates:
                if os.path.exists(mp):
                    try:
                        with open(mp, "r", encoding="utf-8") as mf:
                            mdata = json.load(mf)
                            if not global_session.get("gramPanchayat"):
                                global_session["gramPanchayat"] = mdata.get("gramPanchayat", "")
                            if not global_session.get("panchayatSamiti"):
                                global_session["panchayatSamiti"] = mdata.get("panchayatSamiti", "")
                        break
                    except Exception:
                        pass

        gp = global_session.get("gramPanchayat", "")
        ps = global_session.get("panchayatSamiti", "")
        for v in voters:
            if not v.get("GramPanchayat") and gp:
                v["GramPanchayat"] = gp
            if not v.get("PanchayatSamiti") and ps:
                v["PanchayatSamiti"] = ps

        return voters

    def handle_preview(self, payload):
        voters_source = self.get_voters_from_active_excel_or_session()
        if not voters_source:
            self.send_html_response("<p style='color:red;font-family:sans-serif;padding:20px;'>कोई मतदाता सूची डेटा लोड नहीं है। कृपया पहले एक नया PDF अपलोड करें।</p>", status_code=400)
            return

        candidate_post = payload.get("candidatePost", "सरपंच")
        candidate = payload.get("candidateName", "मनोज बाबेल")
        party = payload.get("partyName", "भारतीय जनता पार्टी (BJP)")
        appeal = payload.get("bottomMessage", "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!")
        slips_per_page = int(payload.get("slipsPerPage", 8))
        candidate_photo = payload.get("candidatePhoto", "")
        party_symbol = payload.get("partySymbol", "")

        voters = voters_source[:slips_per_page]
        grid_css, font_scale = get_grid_and_font(slips_per_page)

        photos_by_epic = global_session.get("voterPhotosByEpic", {})
        photos_by_serial = global_session.get("voterPhotosBySerial", {})

        slips_html = ""
        for v in voters:
            v_photo_url = resolve_voter_photo_url(v, photos_by_epic, photos_by_serial)
            slips_html += render_slip_html(v, candidate, party, appeal, candidate_photo, party_symbol, font_scale, candidate_post, voter_photo=v_photo_url)

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
        if not global_session.get("activeVoters"):
            self.send_json_response({"success": False, "error": "कोई मतदाता सूची डेटा लोड नहीं है। कृपया पहले एक नया PDF अपलोड करें।"}, status_code=400)
            return

        ward_num = global_session.get("ward", "1")
        out_filename = f"{self._output_basename()}.xlsx"
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
            global_session["excelPath"] = dst_excel
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
        voters_source = self.get_voters_from_active_excel_or_session()
        if not voters_source:
            self.send_json_response({"success": False, "error": "कोई मतदाता सूची डेटा लोड नहीं है। कृपया पहले एक नया PDF अपलोड करें।"}, status_code=400)
            return

        candidate_post = payload.get("candidatePost", "सरपंच")
        candidate = payload.get("candidateName", "मनोज बाबेल")
        party = payload.get("partyName", "भारतीय जनता पार्टी (BJP)")
        appeal = payload.get("bottomMessage", "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!")
        slips_per_page = int(payload.get("slipsPerPage", 8))
        candidate_photo = payload.get("candidatePhoto", "")
        party_symbol = payload.get("partySymbol", "")

        ward_num = global_session.get("ward", "1")
        out_filename = f"{self._output_basename()}.pdf"
        dst_pdf = os.path.join(DOWNLOADS_DIR, out_filename)

        print(f"Generating PDF slips directly from Excel records ({len(voters_source)} voters)...")
        self.generate_custom_pdf(
            voters_source, candidate, party, appeal,
            slips_per_page, candidate_photo, party_symbol, dst_pdf, candidate_post
        )

        import shutil
        ws_pdf = os.path.join(WORKSPACE_DIR, out_filename)
        web_dl_dir = os.path.join(WEB_DIR, "downloads")
        os.makedirs(web_dl_dir, exist_ok=True)
        shutil.copy(dst_pdf, ws_pdf)
        shutil.copy(dst_pdf, os.path.join(web_dl_dir, out_filename))

        import math
        total_pages = math.ceil(len(voters_source) / slips_per_page)

        resp_obj = {
            "success": True,
            "totalVoters": len(voters_source),
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

        photos_by_epic = global_session.get("voterPhotosByEpic", {})
        photos_by_serial = global_session.get("voterPhotosBySerial", {})

        html_body = ""
        for chunk in chunks:
            slips_html = ""
            for v in chunk:
                v_photo_url = resolve_voter_photo_url(v, photos_by_epic, photos_by_serial)
                slips_html += render_slip_html(v, candidate, party, appeal, cand_photo_url, party_sym_url, font_scale, candidate_post, voter_photo=v_photo_url)

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
    server_port = PORT
    httpd = None
    for p in range(PORT, PORT + 11):
        try:
            httpd = ThreadedHTTPServer(("127.0.0.1", p), VoterSuvidhaHandler)
            server_port = p
            break
        except (OSError, PermissionError):
            continue

    if not httpd:
        raise RuntimeError(f"Could not bind server to any port between {PORT} and {PORT + 10}.")

    print(f"Voter Suvidha server running on http://127.0.0.1:{server_port}")
    sys.stdout.flush()
    try:
        webbrowser.open(f"http://127.0.0.1:{server_port}")
    except Exception:
        pass
    httpd.serve_forever()
