"""
Auntie Anne's Onboarding PDF Filler API
----------------------------------------
POST /fill                    → fill PDFs, upload to Drive, log to Sheet, return ZIP
GET  /health                  → 200 OK
GET  /submissions             → employee list from Sheet  (X-API-Key required)
PATCH /submissions/<id>/i9    → fill I-9 Section 2, replace Drive file, mark Sheet complete
POST /welcome                 → send welcome email with handbook to new hire (X-API-Key required)
"""

import os
import io
import base64
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date, datetime

import requests as http_requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject

app = Flask(__name__)
CORS(app)

# ─── Paths ───────────────────────────────────────────────────────────────────

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
W4_PATH    = os.path.join(BASE_DIR, "forms", "w4.pdf")
DE_W4_PATH = os.path.join(BASE_DIR, "forms", "de_w4.pdf")
I9_PATH    = os.path.join(BASE_DIR, "forms", "i9.pdf")

# ─── Environment ─────────────────────────────────────────────────────────────

GAS_WEBHOOK_URL      = os.environ.get("GAS_WEBHOOK_URL", "")
ADMIN_API_KEY        = os.environ.get("ADMIN_API_KEY", "")
RESEND_API_KEY       = os.environ.get("RESEND_API_KEY", "")
HANDBOOK_PATH        = os.path.join(BASE_DIR, "handbook.pdf")
GAS_SECRET           = os.environ.get("GAS_SECRET", "")
DRIVE_FOLDER_ID      = os.environ.get("DRIVE_FOLDER_ID", "")

# ── Demo store (separate GAS deployment + Drive folder) ──────────────────────
DEMO_API_KEY         = os.environ.get("DEMO_API_KEY", "demo")
DEMO_GAS_URL         = os.environ.get("DEMO_GAS_URL", "")
DEMO_DRIVE_FOLDER_ID = os.environ.get("DEMO_DRIVE_FOLDER_ID", "")


def check_api_key(req):
    key = req.headers.get("X-API-Key", "")
    if not ADMIN_API_KEY:
        return True
    if key == ADMIN_API_KEY:
        return True
    if DEMO_API_KEY and key == DEMO_API_KEY:
        return True
    return False


def resolve_store(req):
    """Return (gas_url, drive_folder_id) for the request's API key."""
    key = req.headers.get("X-API-Key", "")
    if DEMO_API_KEY and key == DEMO_API_KEY:
        return (DEMO_GAS_URL or GAS_WEBHOOK_URL), (DEMO_DRIVE_FOLDER_ID or DRIVE_FOLDER_ID)
    return GAS_WEBHOOK_URL, DRIVE_FOLDER_ID


# ─── Employer config cache (backed by GAS PropertiesService) ─────────────────
import time as _time
_cfg_cache: dict = {}  # keyed by gas_url

def get_employer_config(gas_url=None) -> dict:
    """Return shared employer config from GAS PropertiesService (10-min TTL cache)."""
    url = gas_url or GAS_WEBHOOK_URL
    entry = _cfg_cache.get(url, {"data": {}, "ts": 0.0})
    if _time.time() - entry["ts"] > 600:
        result = gas_post({"action": "getConfig"}, timeout=10, gas_url=url)
        if result and "businessName" in result:
            _cfg_cache[url] = {"data": result, "ts": _time.time()}
    return _cfg_cache.get(url, {"data": {}})["data"]


# ─── Google Apps Script helpers ──────────────────────────────────────────────

def gas_post(payload, timeout=30, gas_url=None):
    """POST to GAS webhook, follow redirect as GET, return parsed JSON or None."""
    url = gas_url or GAS_WEBHOOK_URL
    if not url:
        return None
    try:
        payload["secret"] = GAS_SECRET
        headers = {"Content-Type": "application/json"}
        res = http_requests.post(
            url, json=payload,
            headers=headers, timeout=timeout,
            allow_redirects=False,
        )
        if res.status_code in (301, 302, 303, 307, 308):
            redirect_url = res.headers.get("Location")
            res = http_requests.get(redirect_url, timeout=timeout)
        res.raise_for_status()
        return res.json()
    except Exception as exc:
        print(f"[GAS] error: {exc}")
        return None


def create_employee_folder(folder_name, gas_url=None, drive_folder_id=None):
    """Create a subfolder in the main Drive folder. Returns (folderId, url) or (None, None)."""
    folder_id = drive_folder_id or DRIVE_FOLDER_ID
    if not folder_id:
        return None, None
    result = gas_post({
        "action":        "createFolder",
        "parentFolderId": folder_id,
        "folderName":    folder_name,
    }, timeout=30, gas_url=gas_url)
    print(f"[Drive] createFolder '{folder_name}' result: {result}")
    if result and "folderId" in result:
        return result["folderId"], result.get("url", "")
    return None, None


def upload_file_to_drive(filename, file_bytes, mimetype, folder_id=None, gas_url=None):
    """Upload file_bytes to Drive via GAS. Returns (fileId, url) or (None, None)."""
    target = folder_id or DRIVE_FOLDER_ID
    if not target:
        return None, None
    result = gas_post({
        "action":   "uploadFile",
        "folderId": target,
        "filename": filename,
        "mimeType": mimetype,
        "fileData": base64.b64encode(file_bytes).decode("utf-8"),
    }, timeout=60, gas_url=gas_url)
    if result and "fileId" in result:
        return result["fileId"], result.get("url", "")
    return None, None


def download_file_from_drive(file_id):
    """Download a Drive file via GAS. Returns bytes or None."""
    result = gas_post({"action": "getFile", "fileId": file_id}, timeout=60)
    if result and "fileData" in result:
        return base64.b64decode(result["fileData"])
    return None


def replace_drive_file(file_id, filename, file_bytes):
    """Replace a Drive file via GAS. Returns new fileId or None."""
    result = gas_post({
        "action":   "replaceFile",
        "fileId":   file_id,
        "filename": filename,
        "fileData": base64.b64encode(file_bytes).decode("utf-8"),
    }, timeout=60)
    if result and "fileId" in result:
        return result["fileId"], result.get("url", "")
    return None, None


def log_to_sheet(data, zip_drive_url="", i9_file_id="", gas_url=None):
    ssn    = data.get("ssn", "")
    digits = "".join(c for c in ssn if c.isdigit())
    masked = f"***-**-{digits[-4:]}" if len(digits) >= 4 else "***"
    result = gas_post({
        "action":         "log",
        "submittedAt":    datetime.utcnow().isoformat(),
        "firstName":      data.get("firstName",      ""),
        "lastName":       data.get("lastName",        ""),
        "email":          data.get("email",           ""),
        "phone":          data.get("phone",           ""),
        "ssn":            masked,
        "dob":            data.get("dob",             ""),
        "address1":       data.get("address1",        ""),
        "city":           data.get("city",            ""),
        "state":          data.get("state",           ""),
        "zip":            data.get("zip",             ""),
        "zipDriveUrl":    zip_drive_url,
        "i9FileId":       i9_file_id,
        "startDate":      data.get("startDate",       ""),
        "ecName":         data.get("ecName",          ""),
        "ecRelationship": data.get("ecRelationship",  ""),
        "ecPhone":        data.get("ecPhone",         ""),
    }, gas_url=gas_url)
    return result.get("rowId") if result else None


def _is_date(val):
    """Return True if val looks like an ISO date/datetime string."""
    try:
        from datetime import datetime
        datetime.fromisoformat(str(val).replace("Z", "+00:00"))
        return True
    except Exception:
        return False


# ─── PDF helpers ─────────────────────────────────────────────────────────────

def fmt_date(val):
    if not val:
        return ""
    try:
        return date.fromisoformat(str(val)[:10]).strftime("%m/%d/%Y")
    except Exception:
        return str(val)


def fmt_ssn(val):
    if not val:
        return ""
    digits = "".join(c for c in val if c.isdigit())
    if len(digits) == 9:
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
    return val


def child_credit_amount(n):
    try:
        amt = int(n) * 2000
        return str(amt) if amt > 0 else ""
    except Exception:
        return ""


def other_dependent_amount(n):
    try:
        amt = int(n) * 500
        return str(amt) if amt > 0 else ""
    except Exception:
        return ""


def fill_pdf_to_bytes(input_path_or_stream, text_fields, checkbox_fields=None):
    if hasattr(input_path_or_stream, "read"):
        reader = PdfReader(input_path_or_stream)
    else:
        reader = PdfReader(input_path_or_stream)
    writer = PdfWriter()
    writer.append(reader)

    # ── Strip XFA so our AcroForm changes are what every viewer renders ──────
    # W4 is a hybrid XFA+AcroForm PDF; leaving the XFA in causes Adobe/etc. to
    # ignore the AcroForm /V + /AS values we set below.
    acroform_ref = writer._root_object.get("/AcroForm")
    if acroform_ref:
        acroform_obj = acroform_ref.get_object()
        if "/XFA" in acroform_obj:
            del acroform_obj["/XFA"]
        # Tell viewers to regenerate appearance streams from our field values
        acroform_obj.update({NameObject("/NeedAppearances"): BooleanObject(True)})

    for page in writer.pages:
        writer.update_page_form_field_values(page, text_fields, auto_regenerate=False)

    if checkbox_fields:
        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    annot_obj = annot_ref.get_object()
                    t = annot_obj.get("/T")

                    if t:
                        # ── Standard checkbox / push-button with /T ──────────
                        t_str = str(t)
                        for field_name, value in checkbox_fields.items():
                            if t_str == field_name or field_name.endswith("." + t_str):
                                annot_obj.update({
                                    NameObject("/V"):  NameObject(value),
                                    NameObject("/AS"): NameObject(value),
                                })
                    else:
                        # ── Radio-button kid: no /T, but has /Parent ─────────
                        # e.g. DE W4 '3status' group whose kids have T=None
                        parent_ref = annot_obj.get("/Parent")
                        if not parent_ref:
                            continue
                        parent_obj = parent_ref.get_object()
                        parent_t = parent_obj.get("/T")
                        if not parent_t:
                            continue
                        parent_t_str = str(parent_t)
                        for field_name, desired_value in checkbox_fields.items():
                            if parent_t_str == field_name or field_name.endswith("." + parent_t_str):
                                # Find which on-state this particular kid represents
                                ap = annot_obj.get("/AP")
                                if not ap:
                                    continue
                                ap_obj = ap.get_object()
                                n = ap_obj.get("/N")
                                if not n:
                                    continue
                                n_obj = n.get_object()
                                if not hasattr(n_obj, "keys"):
                                    continue
                                on_states = [k for k in n_obj.keys() if k != "/Off"]
                                if not on_states:
                                    continue
                                kid_on_state = on_states[0]  # e.g. "/Single" or "/Married"
                                if kid_on_state == desired_value:
                                    annot_obj.update({NameObject("/AS"): NameObject(desired_value)})
                                    parent_obj.update({NameObject("/V"): NameObject(desired_value)})
                                else:
                                    annot_obj.update({NameObject("/AS"): NameObject("/Off")})

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


# ─── Signature image + date text overlay ────────────────────────────────────

def overlay_text(pdf_bytes, text, page_num, x, y, font_size=10):
    """Draw a string at PDF coordinates (x, y) on page_num. Origin = bottom-left."""
    try:
        from reportlab.pdfgen import canvas as rl_canvas
        orig_reader = PdfReader(io.BytesIO(pdf_bytes))
        writer = PdfWriter()
        writer.append(PdfReader(io.BytesIO(pdf_bytes)))
        page = orig_reader.pages[page_num]
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        overlay_buf = io.BytesIO()
        c = rl_canvas.Canvas(overlay_buf, pagesize=(pw, ph))
        c.setFont("Helvetica", font_size)
        c.drawString(x, y, text)
        c.save()
        overlay_buf.seek(0)
        writer.pages[page_num].merge_page(PdfReader(overlay_buf).pages[0])
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return out.read()
    except Exception as exc:
        print(f"[overlay_text] error: {exc}")
        return pdf_bytes


def overlay_signature_image(pdf_bytes, sig_b64, placements):
    """Overlay a drawn signature PNG onto specific pages/coordinates of a PDF.

    placements: list of (page_number, x, y, width, height)
                coordinates in PDF points, origin = bottom-left of page.
    Returns new PDF bytes, or original bytes on any error.
    """
    if not sig_b64:
        return pdf_bytes
    try:
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib.utils import ImageReader
        from PIL import Image

        # Decode the base64 data-URL or raw base64
        raw = sig_b64.split(",", 1)[1] if "," in sig_b64 else sig_b64
        img_bytes = base64.b64decode(raw)
        pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGBA")

        # Build a reader from the original PDF to learn page sizes
        orig_reader = PdfReader(io.BytesIO(pdf_bytes))

        # Create one overlay page per placement
        writer = PdfWriter()
        writer.append(PdfReader(io.BytesIO(pdf_bytes)))

        for page_num, x, y, w, h in placements:
            page = orig_reader.pages[page_num]
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)

            overlay_buf = io.BytesIO()
            c = rl_canvas.Canvas(overlay_buf, pagesize=(pw, ph))
            c.drawImage(ImageReader(pil_img), x, y, width=w, height=h, mask="auto")
            c.save()
            overlay_buf.seek(0)

            overlay_page = PdfReader(overlay_buf).pages[0]
            writer.pages[page_num].merge_page(overlay_page)

        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return out.read()
    except Exception as exc:
        print(f"[overlay_signature_image] error: {exc}")
        return pdf_bytes


# ─── Form fillers ─────────────────────────────────────────────────────────────

def fill_w4(data):
    first_mid = data.get("firstName", "")
    if data.get("middleName"):
        first_mid += " " + data["middleName"][0]
    city_state_zip = f"{data.get('city','')} {data.get('state','')} {data.get('zip','')}"

    text_fields = {
        "topmostSubform[0].Page1[0].Step1a[0].f1_01[0]": first_mid,
        "topmostSubform[0].Page1[0].Step1a[0].f1_02[0]": data.get("lastName", ""),
        "topmostSubform[0].Page1[0].f1_05[0]":            fmt_ssn(data.get("ssn", "")),
        "topmostSubform[0].Page1[0].Step1a[0].f1_03[0]": data.get("address1", ""),
        "topmostSubform[0].Page1[0].Step1a[0].f1_04[0]": city_state_zip,
        "topmostSubform[0].Page1[0].Step3_ReadOrder[0].f1_06[0]": child_credit_amount(data.get("childDependents", 0)),
        "topmostSubform[0].Page1[0].Step3_ReadOrder[0].f1_07[0]": other_dependent_amount(data.get("otherDependents", 0)),
        "topmostSubform[0].Page1[0].f1_08[0]": data.get("otherIncome",           ""),
        "topmostSubform[0].Page1[0].f1_09[0]": data.get("deductions",             ""),
        "topmostSubform[0].Page1[0].f1_10[0]": data.get("additionalWithholding",  ""),
        # f1_11 is NOT the signature date — it sits in the Step 4(c) visual area.
        # The W4 "Sign Here" area has no fillable date AcroForm field; signature
        # is overlaid as an image by overlay_signature_image() below.
        "topmostSubform[0].Page1[0].f1_12[0]": data.get("employerName", "Auntie Anne's"),
        "topmostSubform[0].Page1[0].f1_13[0]": fmt_date(data.get("startDate")),
        "topmostSubform[0].Page1[0].f1_14[0]": data.get("employerEIN", ""),
    }

    fs = data.get("filingStatus", "single")
    checkbox_fields = {
        "topmostSubform[0].Page1[0].c1_1[0]": "/1" if fs == "single"  else "/Off",
        "topmostSubform[0].Page1[0].c1_1[1]": "/2" if fs == "married" else "/Off",
        "topmostSubform[0].Page1[0].c1_1[2]": "/3" if fs == "hoh"     else "/Off",
        "topmostSubform[0].Page1[0].c1_2[0]": "/1" if data.get("multipleJobs") == "yes" else "/Off",
        "topmostSubform[0].Page1[0].c1_3[0]": "/1" if data.get("exempt")        == "yes" else "/Off",
    }
    return fill_pdf_to_bytes(W4_PATH, text_fields, checkbox_fields)


def fill_de_w4(data):
    first_initial = data.get("firstName", "")
    if data.get("middleName"):
        first_initial += " " + data["middleName"][0]

    text_fields = {
        "1Firstnameinitial":       first_initial,
        "1Lastname":               data.get("lastName", ""),
        "2Taxpayerid":             fmt_ssn(data.get("ssn", "")),
        "2Homeaddress":            data.get("address1", ""),
        "3Cityortown":             data.get("city", ""),
        "3State":                  data.get("state", ""),
        "3Zipcode":                data.get("zip", ""),
        "4Totalnumberdependents":  str(data.get("deAllowances", "0")),
        "5Additionalamount":       data.get("deAdditional", ""),
        "6Employersname":          data.get("employerName", "Auntie Anne's"),
        "7Firstdayofemployment":   fmt_date(data.get("startDate")),
        "8Taxpayeridein":          data.get("employerEIN", ""),
    }
    de_fs = data.get("deFilingStatus", "single")
    checkbox_fields = {"3status": "/Single" if de_fs == "single" else "/Married"}
    return fill_pdf_to_bytes(DE_W4_PATH, text_fields, checkbox_fields)


def fill_i9_section1(data):
    """Fill employee Section 1 of the I-9."""
    today = fmt_date(data.get("signatureDate") or date.today().isoformat())
    text_fields = {
        "Last Name (Family Name)":                  data.get("lastName", ""),
        "First Name Given Name":                    data.get("firstName", ""),
        "Employee Middle Initial (if any)":         data.get("middleName", "")[:1] if data.get("middleName") else "",
        "Employee Other Last Names Used (if any)":  data.get("otherNames", "N/A"),
        "Address Street Number and Name":           data.get("address1", ""),
        "Apt Number (if any)":                      data.get("address2", ""),
        "City or Town":                             data.get("city", ""),
        "ZIP Code":                                 data.get("zip", ""),
        "Date of Birth mmddyyyy":                   fmt_date(data.get("dob")),
        "US Social Security Number":                fmt_ssn(data.get("ssn", "")),
        "Telephone Number":                         data.get("phone", ""),
        "Employees E-mail Address":                 data.get("email", ""),
        "Today's Date mmddyyy":                     today,
        "State":                                    data.get("state", ""),
        # Signature of Employee is overlaid as a drawn image by overlay_signature_image()
        # — do NOT fill the text field here or it will show typed text under the image.
        "USCIS ANumber":                            data.get("uscisNumber", ""),
    }
    cit     = data.get("citizenship", "citizen")
    cit_map = {"citizen": "CB_1", "noncitizen": "CB_2", "lpr": "CB_3", "authorized": "CB_4"}
    # CB_1–CB_4 on-state is /On (not /Yes — confirmed by inspecting AP/N keys)
    checkbox_fields = {
        field_id: "/On" if cit == key else "/Off"
        for key, field_id in cit_map.items()
    }
    return fill_pdf_to_bytes(I9_PATH, text_fields, checkbox_fields)


def fill_i9_section2(i9_bytes, section2_data, start_date=""):
    """Fill employer Section 2 onto an existing I-9 PDF (already has Section 1).

    docType="listA"  → one List A document (e.g. U.S. Passport).
    docType="listBC" → List B + List C documents (e.g. Driver's License + SSN card).

    The List A Document Title field on page 0 has no AcroForm /T name, so it is
    written via overlay_text() after fill_pdf_to_bytes returns.
    Supplement B (page 3) fields are intentionally left blank — reverification only.
    """
    today     = fmt_date(datetime.utcnow().date().isoformat())
    first_day = fmt_date(start_date) if start_date else today
    emp_name  = section2_data.get("empName", "")
    emp_addr  = section2_data.get("employerAddress", "")
    emp_org   = section2_data.get("employerOrg", "Auntie Anne's")
    doc_type  = section2_data.get("docType", "listA")

    text_fields = {
        # ── Employer certification ──────────────────────────────────────────
        "FirstDayEmployed mmddyyyy": first_day,
        "Last Name First Name and Title of Employer or Authorized Representative": emp_name,
        # "Signature of Employer or AR" left blank — drawn image overlaid by caller
        "S2 Todays Date mmddyyyy":       today,
        "Employers Business or Org Name":    emp_org,
        "Employers Business or Org Address": emp_addr,
    }

    if doc_type == "listA":
        # ── List A: single document (e.g. U.S. Passport) ───────────────────
        # Document Title has no field name on page 0 — overlaid as text below.
        text_fields.update({
            "Issuing Authority 1":        section2_data.get("issuer",    ""),
            "Document Number 0 (if any)": section2_data.get("docNumber", ""),
            "Expiration Date if any":     fmt_date(section2_data.get("expDate", "")) or section2_data.get("expDate", ""),
        })
    else:
        # ── List B + List C: two documents ─────────────────────────────────
        text_fields.update({
            "List B Document 1 Title":    section2_data.get("listBTitle",   ""),
            "List B Issuing Authority 1": section2_data.get("listBIssuer",  ""),
            "List B Document Number 1":   section2_data.get("listBNumber",  ""),
            "List B Expiration Date 1":   fmt_date(section2_data.get("listBExpDate", "")) or section2_data.get("listBExpDate", ""),
            "List C Document Title 1":    section2_data.get("listCTitle",   ""),
            "List C Issuing Authority 1": section2_data.get("listCIssuer",  ""),
            "List C Document Number 1":   section2_data.get("listCNumber",  ""),
            "List C Expiration Date 1":   fmt_date(section2_data.get("listCExpDate", "")) or section2_data.get("listCExpDate", ""),
        })

    result = fill_pdf_to_bytes(io.BytesIO(i9_bytes), text_fields)

    # List A doc title: unnamed field at page 0, x=127–263, y≈342 — use text overlay.
    if doc_type == "listA" and section2_data.get("docTitle"):
        result = overlay_text(result, section2_data["docTitle"], 0, 128, 345, font_size=10)

    return result


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200


@app.route("/fill", methods=["POST"])
def fill():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data received"}), 400

    try:
        first = data.get("firstName", "employee").strip()
        last  = data.get("lastName",  "").strip()
        name  = f"{first}_{last}".replace(" ", "_")

        # Resolve which GAS/Drive to use based on API key
        store_gas_url, store_folder_id = resolve_store(request)

        # Inject employer info from server-side config store (falls back to defaults)
        cfg = get_employer_config(gas_url=store_gas_url)
        data.setdefault("employerName", cfg.get("businessName", "Your Store"))
        data.setdefault("employerEIN",  cfg.get("ein", ""))

        w4_bytes    = fill_w4(data)
        de_w4_bytes = fill_de_w4(data)
        i9_bytes    = fill_i9_section1(data)

        # Overlay drawn signature image if provided
        sig_b64 = data.get("signatureImage", "")
        today_str = fmt_date(date.today().isoformat())
        if sig_b64:
            # W4: Step 5 signature line at y≈86; blank signing space y=88–108.
            # Date column starts at x≈408. (Measured by rendering PDF at 150 DPI.)
            w4_bytes = overlay_signature_image(w4_bytes, sig_b64, [(0, 100, 88, 290, 18)])
            w4_bytes = overlay_text(w4_bytes, today_str, 0, 415, 91)
            # DE W4: signature/date line at y≈432; blank space y=433–453.
            # Date column starts at x≈401.
            de_w4_bytes = overlay_signature_image(de_w4_bytes, sig_b64, [(0, 79, 432, 295, 20)])
            de_w4_bytes = overlay_text(de_w4_bytes, today_str, 0, 401, 436)
            # I-9: AcroForm "Signature of Employee" field rect=[42.1, 420.8, 365.3, 433.7].
            # Today's Date is already filled via AcroForm — no text overlay needed.
            i9_bytes = overlay_signature_image(i9_bytes, sig_b64, [(0, 42, 421, 315, 22)])

        zip_filename = f"{name}_onboarding_docs.zip"
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{name}_W4_Federal.pdf",  w4_bytes)
            zf.writestr(f"{name}_W4_Delaware.pdf", de_w4_bytes)
            zf.writestr(f"{name}_I9.pdf",          i9_bytes)
        zip_bytes = zip_buf.getvalue()

        # Create per-employee subfolder in Drive
        emp_folder_id, emp_folder_url = create_employee_folder(
            f"{first} {last}", gas_url=store_gas_url, drive_folder_id=store_folder_id
        )
        target_folder = emp_folder_id or store_folder_id

        # Upload individual PDFs into the employee's folder
        i9_file_id, _ = upload_file_to_drive(
            f"{name}_I9.pdf", i9_bytes, "application/pdf", target_folder, gas_url=store_gas_url
        )
        upload_file_to_drive(
            f"{name}_W4_Federal.pdf", w4_bytes, "application/pdf", target_folder, gas_url=store_gas_url
        )
        upload_file_to_drive(
            f"{name}_W4_Delaware.pdf", de_w4_bytes, "application/pdf", target_folder, gas_url=store_gas_url
        )

        # Log to Sheet — driveUrl points to the employee's folder
        log_to_sheet(data, zip_drive_url=emp_folder_url or "", i9_file_id=i9_file_id or "", gas_url=store_gas_url)

        return send_file(
            io.BytesIO(zip_bytes),
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_filename,
        )

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/submissions", methods=["GET"])
def get_submissions():
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    store_gas_url, _ = resolve_store(request)
    result = gas_post({"action": "getAll"}, gas_url=store_gas_url)
    if result is None:
        return jsonify([])
    if "error" in result:
        return jsonify({"error": result["error"]}), 500
    return jsonify(result.get("employees", []))


@app.route("/submissions/<int:row_id>/i9", methods=["PATCH"])
def complete_i9(row_id):
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data"}), 400

    store_gas_url, _ = resolve_store(request)
    # ── Round trip 1: get row data (start date + I-9 file ID) ────────────────
    row_result = gas_post({"action": "getRow", "rowId": row_id}, timeout=60, gas_url=store_gas_url)
    if row_result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in row_result:
        return jsonify({"error": row_result["error"]}), 500

    row = row_result.get("row", [])
    # The id returned by /submissions is already the 1-based sheet row number
    # (GAS getAll uses `id = i + 1` which accounts for any header row), so we
    # can use row_id directly as the actual row without any offset detection.
    actual_row_id = row_id

    i9_file_id = str(row[19]).strip() if len(row) > 19 else ""
    start_date = row[20] if len(row) > 20 else ""

    if not i9_file_id:
        return jsonify({"error": "I-9 file not found for this employee"}), 404

    # ── Round trip 2: download current I-9 PDF from Drive ────────────────────
    file_result = gas_post({"action": "getFile", "fileId": i9_file_id}, timeout=60, gas_url=store_gas_url)
    if file_result is None or "error" in (file_result or {}):
        return jsonify({"error": "Failed to download I-9 from Drive"}), 500

    i9_bytes   = base64.b64decode(file_result["fileData"])
    # Inject employer org name from server-side config if not provided
    if not data.get("employerOrg"):
        cfg = get_employer_config()
        data["employerOrg"] = cfg.get("businessName", "Auntie Anne's")
    updated_i9 = fill_i9_section2(i9_bytes, data, str(start_date))

    # Overlay employer's drawn signature on Section 2 (page 0 only).
    # "Signature of Employer or AR" field rect=[294.3, 79.6, 485.3, 99.2].
    # Page 3 Supplement B fields are left blank (reverification/rehires).
    sig_b64 = data.get("sigImage", "")
    if sig_b64:
        updated_i9 = overlay_signature_image(updated_i9, sig_b64, [
            (0, 294, 79, 188, 22),
        ])

    # ── Round trip 3: replace I-9 file in Drive ───────────────────────────────
    replace_result = gas_post({
        "action":   "replaceFile",
        "fileId":   i9_file_id,
        "filename": f"I9_COMPLETED_row{actual_row_id}.pdf",
        "fileData": base64.b64encode(updated_i9).decode("utf-8"),
    }, timeout=60, gas_url=store_gas_url)
    if replace_result is None or "error" in (replace_result or {}):
        return jsonify({"error": "Failed to replace I-9 file in Drive"}), 500

    new_file_id = replace_result.get("fileId", i9_file_id)

    # ── Round trip 4: mark I-9 complete in Sheet ──────────────────────────────
    complete_result = gas_post({
        "action":      "completeI9",
        "rowId":       actual_row_id,
        "docTitle":    data.get("docTitle",  ""),
        "docNumber":   data.get("docNumber", ""),
        "issuer":      data.get("issuer",    ""),
        "expDate":     data.get("expDate",   ""),
        "empName":     data.get("empName",   ""),
        "newI9FileId": new_file_id,
    }, timeout=60, gas_url=store_gas_url)
    if complete_result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in complete_result:
        return jsonify({"error": complete_result["error"]}), 500
    return jsonify({"status": "ok"})


@app.route("/submissions/<int:row_id>/status", methods=["PATCH"])
def update_status(row_id):
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    data = request.get_json()
    if not data or "overallStatus" not in data:
        return jsonify({"error": "Missing overallStatus"}), 400
    store_gas_url, _ = resolve_store(request)
    # Detect header-row offset same as complete_i9
    row_result = gas_post({"action": "getRow", "rowId": row_id}, gas_url=store_gas_url)
    actual_row_id = row_id
    if row_result and "row" in row_result:
        row = row_result["row"]
        if row and row[0] and not _is_date(str(row[0])):
            actual_row_id = row_id + 1
    result = gas_post({
        "action":        "updateStatus",
        "rowId":         actual_row_id,
        "overallStatus": data["overallStatus"],
    }, gas_url=store_gas_url)
    if result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in result:
        return jsonify({"error": result["error"]}), 500
    return jsonify({"status": "ok"})


@app.route("/submissions/<int:row_id>/working-papers", methods=["POST"])
def upload_working_papers(row_id):
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    data = request.get_json()
    if not data or not data.get("fileData"):
        return jsonify({"error": "Missing fileData"}), 400
    store_gas_url, _ = resolve_store(request)
    # Detect header-row offset
    actual_row_id = row_id
    row_result = gas_post({"action": "getRow", "rowId": row_id}, gas_url=store_gas_url)
    if row_result and "row" in row_result:
        row = row_result["row"]
        if row and row[0] and not _is_date(str(row[0])):
            actual_row_id = row_id + 1
    result = gas_post({
        "action":   "uploadWorkingPapers",
        "rowId":    actual_row_id,
        "folderId": data.get("folderId", ""),
        "fileData": data.get("fileData", ""),
        "mimeType": data.get("mimeType", "image/jpeg"),
        "filename": data.get("filename", "working_papers"),
    }, gas_url=store_gas_url)
    if result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in result:
        return jsonify({"error": result["error"]}), 500
    return jsonify({"status": "ok", "fileId": result.get("fileId", "")})


@app.route("/debug", methods=["GET"])
def debug():
    return jsonify({
        "gas_configured":    bool(GAS_WEBHOOK_URL),
        "drive_configured":  bool(DRIVE_FOLDER_ID),
        "admin_key_set":     bool(ADMIN_API_KEY),
        "gas_secret_set":    bool(GAS_SECRET),
    })


@app.route("/config", methods=["GET"])
def get_config():
    """Return shared employer config from GAS PropertiesService."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    store_gas_url, _ = resolve_store(request)
    result = gas_post({"action": "getConfig"}, gas_url=store_gas_url)
    return jsonify(result or {})


@app.route("/config", methods=["PATCH"])
def update_config():
    """Save shared employer config to GAS PropertiesService and bust cache."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data"}), 400
    allowed = {"businessName", "managerName", "managerPhone", "ein", "address", "email", "state"}
    payload = {k: v for k, v in data.items() if k in allowed}
    if not payload:
        return jsonify({"error": "No valid config fields provided"}), 400
    store_gas_url, _ = resolve_store(request)
    result = gas_post({"action": "setConfig", **payload}, gas_url=store_gas_url)
    if result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    # Bust cache for this store
    if store_gas_url in _cfg_cache:
        _cfg_cache[store_gas_url]["ts"] = 0.0
    return jsonify({"status": "ok"})


WELCOME_EMAIL_TEMPLATE = """\
Hi {firstName},

We are all really excited to welcome you to our team at Auntie Anne's Christiana Mall! We believe that you will be a wonderful addition to our team! Your starting pay rate will be {payRate}. We are paid bi-weekly on Fridays via Direct Deposit. With your start date being the week of {startWeek} you will receive your first paycheck on {firstPaycheck} and then every other Friday after that.

Joining our team at Auntie Anne's is a 5-step process, so let's get started!

STEP 1 - REVIEW THE EMPLOYEE HANDBOOK
Attached to this email is the employee handbook. When you have a few minutes, take some time to read through it. You are expected to read the entire document, but please pay special attention to Section 6, "Employee Conduct," and Section 7, "Timekeeping and Payroll," so you are familiar with these policies. If you have any questions about what you read in the handbook, please reach out to me ({senderName} - {senderPhone}).

STEP 2 - FILL OUT YOUR NEW EMPLOYEE PAPERWORK ONLINE
If you haven't already, you should soon be receiving a link to fill out all of the new hire paperwork online. This will be done on your mobile device or any other computer. You will be asked if you have read the employee handbook, so make sure you complete Step 1 before completing Step 2. If you have any questions while filling out new employee paperwork, please reach out to myself ({senderName}).

STEP 3 - SET UP YOUR ACCOUNT ON ADP IN THE FIRST WEEK
In the next week or so you will receive an email from ADP, our payroll processor. Please follow the instructions in the email to setup your ADP account so that you can confirm your payroll settings and view paystubs. You won't get this email until we set you up in ADP, so please give us 7-10 days to get you registered. In the meantime, you can move on to Step 4.

STEP 4 - GET READY FOR YOUR FIRST DAY
Here is what you'll need:

  • Have all of your new hire paperwork completed before you come in.
  • Bring in the forms of identification so we can fill out your I-9 form.
  • If you are under the age of 18, bring in your working papers.
  • Come prepared to work on the floor, with jeans or tan/black khaki pants or shorts and comfortable closed toe shoes. Long hair must be pulled back.
  • We will provide you with an Auntie Anne's t-shirt and a hat or visor.
  • Have your log in code. I will get you set up on our register system and send you a text with the code you'll need to clock in/out and log in on the registers.

STEP 5 - DOWNLOAD OUR SCHEDULING APP
Once you have completed your initial 4-day training you will get login information for Humanity, our scheduling service. You can log in here: https://auntieannes105112.humanity.com/app/dashboard/
In the meantime, you can download the free Humanity app on any smartphone or tablet — just search for "Humanity - employee scheduling" in your app store.

We are looking forward to working with you and seeing you achieve great things! Thanks for joining the Auntie Anne's Christiana Mall family!

Best regards,
{senderName}
Store Manager, Auntie Anne's Christiana Mall
{senderPhone}
"""


@app.route("/welcome", methods=["POST"])
def send_welcome():
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data"}), 400

    required = ["firstName", "email", "payRate", "firstPaycheck", "senderName", "senderEmail"]
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({"error": f"Missing fields: {', '.join(missing)}"}), 400

    if not RESEND_API_KEY:
        return jsonify({"error": "RESEND_API_KEY not configured in Railway"}), 500

    first_name    = data["firstName"]
    sender_name   = data["senderName"]
    sender_phone  = data.get("senderPhone", "")
    sender_email  = data["senderEmail"]

    body = WELCOME_EMAIL_TEMPLATE.format(
        firstName=first_name,
        payRate=data["payRate"],
        startWeek=data.get("startWeek", ""),
        firstPaycheck=data["firstPaycheck"],
        senderName=sender_name,
        senderPhone=sender_phone,
    )

    # Build Resend payload
    to_list = [data["email"]]
    payload = {
        "from":     f"Auntie Anne's Christiana Mall <tyler@gomiddleman.com>",
        "to":       to_list,
        "reply_to": sender_email,
        "subject":  f"Welcome to the Team, {first_name}! 🥨",
        "text":     body,
    }
    cc = data.get("cc", "").strip()
    if cc:
        payload["cc"] = [a.strip() for a in cc.split(",") if a.strip()]

    # Attach handbook from Drive via GAS if available
    handbook_attachment = None
    handbook_file_id = "1MkAHFaMDj3Ejkjid73nvNSpfxcq3eo2s"
    gas_result = gas_post({"action": "getFile", "fileId": handbook_file_id}, timeout=20)
    if gas_result and gas_result.get("fileData"):
        handbook_attachment = gas_result["fileData"]
        payload["attachments"] = [{
            "filename": "Auntie_Annes_Employee_Handbook.pdf",
            "content":  handbook_attachment,
        }]

    try:
        res = http_requests.post(
            "https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {RESEND_API_KEY}", "Content-Type": "application/json"},
            json=payload,
            timeout=20,
        )
        if res.status_code not in (200, 201):
            return jsonify({"error": res.json().get("message", f"Resend error {res.status_code}")}), 500
        return jsonify({"status": "ok"})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
