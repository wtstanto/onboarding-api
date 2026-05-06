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
import json
import base64
import zipfile
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

GAS_WEBHOOK_URL  = os.environ.get("GAS_WEBHOOK_URL", "")
ADMIN_API_KEY    = os.environ.get("ADMIN_API_KEY", "")
GMAIL_USER       = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASS   = os.environ.get("GMAIL_APP_PASSWORD", "")
HANDBOOK_PATH    = os.path.join(BASE_DIR, "handbook.pdf")
GAS_SECRET       = os.environ.get("GAS_SECRET", "")
DRIVE_FOLDER_ID  = os.environ.get("DRIVE_FOLDER_ID", "")
DEMO_GAS_URL      = os.environ.get("DEMO_GAS_URL", "")
DEMO_DRIVE_FOLDER = os.environ.get("DEMO_DRIVE_FOLDER_ID", "")
DEMO_ADMIN_API_KEY = os.environ.get("DEMO_ADMIN_API_KEY", "")

TWILIO_ACCOUNT_SID     = os.environ.get("TWILIO_ACCOUNT_SID", "")
TWILIO_AUTH_TOKEN      = os.environ.get("TWILIO_AUTH_TOKEN", "")
TWILIO_FROM_NUMBER     = os.environ.get("TWILIO_FROM_NUMBER", "")
TWILIO_MESSAGING_SID   = os.environ.get("TWILIO_MESSAGING_SERVICE_SID", "")
RESEND_API_KEY         = os.environ.get("RESEND_API_KEY", "")
RESEND_FROM_EMAIL      = os.environ.get("RESEND_FROM_EMAIL", "onboarding@de112.com")


def check_api_key(req):
    key = req.headers.get("X-API-Key", "")
    if DEMO_ADMIN_API_KEY and key == DEMO_ADMIN_API_KEY:
        return True
    if not ADMIN_API_KEY:
        return True
    return key == ADMIN_API_KEY


def is_demo_request(req):
    """True when the request was authenticated with the demo API key."""
    key = req.headers.get("X-API-Key", "")
    return bool(DEMO_ADMIN_API_KEY and key == DEMO_ADMIN_API_KEY)


def demo_gas_url(req):
    return DEMO_GAS_URL if is_demo_request(req) else None


def demo_folder(req):
    return DEMO_DRIVE_FOLDER if is_demo_request(req) else None


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


def create_employee_folder(folder_name, gas_url=None, folder_id=None):
    """Create a subfolder in the main Drive folder. Returns (folderId, url) or (None, None)."""
    parent = folder_id or DRIVE_FOLDER_ID
    if not parent:
        return None, None
    result = gas_post({
        "action":        "createFolder",
        "parentFolderId": parent,
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
        "action":           "log",
        "submittedAt":      datetime.utcnow().isoformat(),
        "firstName":        data.get("firstName",      ""),
        "lastName":         data.get("lastName",        ""),
        "email":            data.get("email",           ""),
        "phone":            data.get("phone",           ""),
        "ssn":              masked,
        "dob":              data.get("dob",             ""),
        "address1":         data.get("address1",        ""),
        "city":             data.get("city",            ""),
        "state":            data.get("state",           ""),
        "zip":              data.get("zip",             ""),
        "zipDriveUrl":      zip_drive_url,
        "i9FileId":         i9_file_id,
        "startDate":        data.get("startDate",       ""),
        "ecName":           data.get("ecName",          ""),
        "ecRelationship":   data.get("ecRelationship",  ""),
        "ecPhone":          data.get("ecPhone",         ""),
        "gender":           data.get("gender",          ""),
        "tshirtSize":       data.get("tshirtSize",      ""),
        "i9s1docs":         json.dumps(data.get("i9s1docs") or {}),
        "testEntry":        bool(data.get("testEntry")),
        "bankName":         data.get("bankName",        ""),
        "routingNumber":    data.get("routingNumber",   ""),
        "accountNumber":    data.get("accountNumber",   ""),
        "accountType":      data.get("accountType",     ""),
    }, timeout=90, gas_url=gas_url)
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
    emp_org   = "Auntie Anne's"
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

    first = data.get("firstName", "employee").strip()
    last  = data.get("lastName",  "").strip()
    name  = f"{first}_{last}".replace(" ", "_")

    # All PDF generation, signature overlay, Drive upload, and Sheet logging
    # happen in a background thread. The employee sees success instantly.
    import threading
    def _background(data, first, last, name):
        try:
            today_str = fmt_date(date.today().isoformat())
            sig_b64   = data.get("signatureImage", "")

            w4_bytes    = fill_w4(data)
            de_w4_bytes = fill_de_w4(data)
            i9_bytes    = fill_i9_section1(data)

            if sig_b64:
                w4_bytes    = overlay_signature_image(w4_bytes,    sig_b64, [(0, 100, 88,  290, 18)])
                w4_bytes    = overlay_text(w4_bytes,    today_str, 0, 415, 91)
                de_w4_bytes = overlay_signature_image(de_w4_bytes, sig_b64, [(0,  79, 432, 295, 20)])
                de_w4_bytes = overlay_text(de_w4_bytes, today_str, 0, 401, 436)
                i9_bytes    = overlay_signature_image(i9_bytes,    sig_b64, [(0,  42, 421, 315, 22)])

            is_demo  = data.get("mode") == "demo"
            g_url    = DEMO_GAS_URL      if is_demo else None
            g_folder = DEMO_DRIVE_FOLDER if is_demo else None

            emp_folder_id, emp_folder_url = create_employee_folder(
                f"{first} {last}", gas_url=g_url, folder_id=g_folder
            )
            target_folder = emp_folder_id or (g_folder or DRIVE_FOLDER_ID)

            i9_file_id, _ = upload_file_to_drive(
                f"{name}_I9.pdf", i9_bytes, "application/pdf", target_folder, gas_url=g_url
            )
            upload_file_to_drive(
                f"{name}_W4_Federal.pdf", w4_bytes, "application/pdf", target_folder, gas_url=g_url
            )
            upload_file_to_drive(
                f"{name}_W4_Delaware.pdf", de_w4_bytes, "application/pdf", target_folder, gas_url=g_url
            )
            log_to_sheet(data, zip_drive_url=emp_folder_url or "", i9_file_id=i9_file_id or "", gas_url=g_url)
            print(f"[/fill] background complete for {name}")
        except Exception as exc:
            print(f"[/fill] background error for {name}: {exc}")

    threading.Thread(target=_background, args=(data, first, last, name), daemon=False).start()

    return jsonify({"status": "ok"})


@app.route("/submissions", methods=["GET"])
def get_submissions():
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    result = gas_post({"action": "getAll"}, gas_url=demo_gas_url(request))
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

    g_url = demo_gas_url(request)
    # ── Round trip 1: get row data (start date + I-9 file ID) ────────────────
    row_result = gas_post({"action": "getRow", "rowId": row_id}, timeout=60, gas_url=g_url)
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
    file_result = gas_post({"action": "getFile", "fileId": i9_file_id}, timeout=60, gas_url=g_url)
    if file_result is None or "error" in (file_result or {}):
        return jsonify({"error": "Failed to download I-9 from Drive"}), 500

    i9_bytes   = base64.b64decode(file_result["fileData"])
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
    }, timeout=60, gas_url=g_url)
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
    }, timeout=60, gas_url=g_url)
    if complete_result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in complete_result:
        return jsonify({"error": complete_result["error"]}), 500
    return jsonify({"status": "ok"})


# (Legacy /submissions/<id>/status endpoint removed — replaced by lifecycle
#  status endpoint at the bottom of this file. The old one wrote to the
#  overallStatus column which is no longer read by the admin.)


@app.route("/debug", methods=["GET"])
def debug():
    return jsonify({
        "gas_configured":    bool(GAS_WEBHOOK_URL),
        "drive_configured":  bool(DRIVE_FOLDER_ID),
        "admin_key_set":     bool(ADMIN_API_KEY),
        "gas_secret_set":    bool(GAS_SECRET),
    })


WELCOME_EMAIL_TEMPLATE = """\
{firstName},

We are all really excited to welcome you to our team at Auntie Anne's Christiana Mall! We believe that you will be a wonderful addition to our team! Your starting pay rate will be {payRate}. We are paid bi-weekly on Fridays via Direct Deposit. With your start date being the week of {startWeek} you will receive your first paycheck on {firstPaycheck} and then every other Friday after that.

Joining our team at Auntie Anne's is a 4-step process, so let's get started!


STEP 1 - FILL OUT YOUR NEW EMPLOYEE PAPERWORK ONLINE

Please click the link below to complete our onboarding process, which includes collecting your information and reviewing our employee handbook:

{onboardLink}

The handbook covers everything you need to know before your first day. Please pay special attention to Section 6, "Employee Conduct," and Section 7, "Timekeeping and Payroll."

If you have any questions, please reach out to me ({senderName} at {senderPhone}).


STEP 2 - SET UP YOUR ACCOUNT ON ADP IN THE FIRST WEEK

Within a week from completing your paperwork you will receive an email from ADP, our payroll processor. Please follow the instructions in the email to setup your ADP account so that you can confirm your payroll settings and view paystubs.


STEP 3 - GET READY FOR YOUR FIRST DAY

Here is what you'll need for your first day:

  - Have all of your new hire paperwork completed (Step 2 of this email)
  - Bring in the forms of identification you specified on your I-9 form so we can verify.
  - If you are under the age of 18, bring in your working papers.
  - Come prepared to work on the floor, with jeans or tan/black khaki pants or shorts and also comfortable closed toe shoes. Long hair must be pulled back. (We will provide you with an Auntie Anne's t-shirt and a hat or visor.)


STEP 4 - DOWNLOAD OUR SCHEDULING APP

Once you have completed your initial 4-day training you will get login information for Humanity, our scheduling service. Once you have that info you can login here:

https://auntieannes105112.humanity.com/app/dashboard/

In the meantime, you can download the free Humanity app on any smartphone or tablet. Just search for "Humanity - employee scheduling" in your app store. You will be able to log in once we give you your login info after your 4-day training.

We are looking forward to working with you and seeing you achieve great things! Thanks for joining the Auntie Anne's Christiana Mall family!
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

    first_name    = data["firstName"]
    last_name     = data.get("lastName", "")
    to_email      = data["email"]
    hire_phone    = data.get("hirePhone", "").strip()
    pay_rate      = data["payRate"]
    first_paycheck = data["firstPaycheck"]
    start_week    = data.get("startWeek", "")
    onboard_link  = data.get("onboardLink", "")
    sender_name   = data["senderName"]
    sender_phone  = data.get("senderPhone", "")
    sender_email  = data["senderEmail"]

    # Format pay rate: "15" → "$15.00/hr", "15.50" → "$15.50/hr"
    try:
        pay_rate_fmt = f"${float(pay_rate):.2f}/hr"
    except Exception:
        pay_rate_fmt = pay_rate

    body = WELCOME_EMAIL_TEMPLATE.format(
        firstName=first_name,
        payRate=pay_rate_fmt,
        startWeek=start_week,
        firstPaycheck=first_paycheck,
        senderName=sender_name,
        senderPhone=sender_phone,
        onboardLink=onboard_link or "https://previews.gomiddleman.com/de112c/",
    )

    subject = f"Welcome to the Auntie Anne's Christiana Mall Team, {first_name}!"

    # Send via Resend HTTPS API — no SMTP ports, no Google OAuth issues
    if not RESEND_API_KEY:
        return jsonify({"error": "Email not configured — set RESEND_API_KEY in Railway"}), 500

    try:
        resend_resp = http_requests.post(
            "https://api.resend.com/emails",
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "from":     f"{sender_name} <{RESEND_FROM_EMAIL}>",
                "to":       [to_email],
                "reply_to": sender_email,
                "subject":  subject,
                "text":     body,
            },
            timeout=10,
        )
        if resend_resp.status_code not in (200, 201):
            err = resend_resp.json().get("message", resend_resp.text)
            return jsonify({"error": f"Failed to send email: {err}"}), 500
    except Exception as exc:
        return jsonify({"error": f"Failed to send email: {exc}"}), 500

    # Optional SMS via Twilio
    sms_status = None
    if hire_phone and TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN:
        # Normalize to E.164 (assume US if no country code)
        digits = "".join(c for c in hire_phone if c.isdigit())
        if len(digits) == 10:
            digits = "1" + digits
        to_number = "+" + digits
        link_line = f" Start your paperwork here: {onboard_link}" if onboard_link else ""
        sms_body = (
            f"Hey {first_name}, welcome to Auntie Anne's! An email was just sent to {to_email} "
            f"with important onboarding details — if you don't see it, check your spam folder.{link_line} "
            f"Reply STOP to opt out."
        )
        sms_data = {"To": to_number, "Body": sms_body}
        if TWILIO_MESSAGING_SID:
            sms_data["MessagingServiceSid"] = TWILIO_MESSAGING_SID
        else:
            sms_data["From"] = TWILIO_FROM_NUMBER
        try:
            http_requests.post(
                f"https://api.twilio.com/2010-04-01/Accounts/{TWILIO_ACCOUNT_SID}/Messages.json",
                auth=(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN),
                data=sms_data,
                timeout=10,
            )
            sms_status = "sent"
        except Exception:
            sms_status = "failed"

    return jsonify({"status": "ok", "sms": sms_status})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
"""
APPEND to /mnt/onboarding-api/app.py — do not replace any existing code.

Adds 5 routes and 4 CSV/cheat-sheet builders for the /de112b downstream-export
workflow. Uses existing helpers: gas_post, check_api_key, upload_file_to_drive,
fmt_date, ADMIN_API_KEY, DRIVE_FOLDER_ID.
"""

# ─── Downstream-system exports (for /de112b admin) ───────────────────────────

import csv as _csv
import io as _io
import html as _html

# Sheet indexes for the new columns (0-based, matching the row[] arrays from
# GAS getRow). Column letters in comments. If the sheet ever shifts these must
# update here AND in google-apps-script.js — keep in sync.
_COL_TSHIRT_SIZE      = 44  # AS
_COL_GENDER           = 25  # Z
_COL_PAY_RATE         = 26  # AA
_COL_POSITION         = 27  # AB
_COL_LOCATION         = 28  # AC
_COL_DEPT_CODE        = 29  # AD
_COL_HIRE_DATE        = 30  # AE
_COL_HUMANITY_STAMP   = 31  # AF
_COL_QU_STAMP         = 32  # AG
_COL_ZIGNAL_STAMP     = 33  # AH
_COL_ADP_SHEET_STAMP  = 34  # AI
_COL_ADP_CSV_STAMP    = 35  # AJ

# Map system name → stamp column index + sheet column number (1-based, for GAS)
_EXPORT_SYSTEMS = {
    "humanity": {"stamp_idx": _COL_HUMANITY_STAMP,  "stamp_col": 32, "label": "Humanity"},
    "qu":       {"stamp_idx": _COL_QU_STAMP,        "stamp_col": 33, "label": "Qu POS"},
    "zignal":   {"stamp_idx": _COL_ZIGNAL_STAMP,    "stamp_col": 34, "label": "Zygnal"},
    "adp":      {"stamp_idx": _COL_ADP_CSV_STAMP,   "stamp_col": 36, "label": "ADP Run"},
    "adp_sheet":{"stamp_idx": _COL_ADP_SHEET_STAMP, "stamp_col": 35, "label": "ADP cheat sheet"},
}


def _row_to_employee(row):
    """Normalize a sheet row (list of cell values from GAS getRow) into a dict
    with every field the CSV/cheat-sheet builders need. Missing cells become ''.

    Two date formats are exposed:
    - `dob`, `startDate`, `hireDate`: human MM/DD/YYYY (used in cheat sheet)
    - `dobIso`, `hireDateIso`: ISO YYYY-MM-DD (used in CSVs for machine imports)
    """
    def g(idx):
        try:
            return row[idx] if row[idx] is not None else ""
        except IndexError:
            return ""

    def iso_date(val):
        """Force YYYY-MM-DD for CSV outputs. Handles ISO, GAS-format, and empty."""
        if not val:
            return ""
        try:
            return date.fromisoformat(str(val)[:10]).isoformat()
        except Exception:
            try:
                # Fallback for GAS format: "Mon Apr 20 2026 00:00:00 GMT-0400"
                from datetime import datetime as _dt
                return _dt.strptime(str(val)[:15], "%a %b %d %Y").date().isoformat()
            except Exception:
                return str(val)

    raw_dob        = g(6)
    raw_start      = g(20)
    raw_hire       = g(_COL_HIRE_DATE) or raw_start

    return {
        "submittedAt":    g(0),
        "firstName":      g(1),
        "lastName":       g(2),
        "email":          g(3),
        "phone":          g(4),
        "ssn_masked":     g(5),   # sheet stores masked SSN; real SSN is NOT here
        "dob":            fmt_date(raw_dob),
        "dobIso":         iso_date(raw_dob),
        "address1":       g(7),
        "city":           g(8),
        "state":          g(9),
        "zip":            str(g(10)),
        "i9Status":       g(11),
        "driveUrl":       g(12),
        "startDate":      fmt_date(raw_start),
        "ecName":         g(21),
        "ecRelationship": g(22),
        "ecPhone":        g(23),
        "overallStatus":  g(24),
        "gender":         g(_COL_GENDER),
        "tshirtSize":     g(_COL_TSHIRT_SIZE),
        "payRate":        g(_COL_PAY_RATE),
        "position":       g(_COL_POSITION),
        "location":       g(_COL_LOCATION),
        "deptCode":       g(_COL_DEPT_CODE),
        "hireDate":       fmt_date(raw_hire),
        "hireDateIso":    iso_date(raw_hire),
    }


def _employment_complete(emp):
    """Pay rate, position, location are the three hard requirements before any
    export is allowed. Dept code and hire date we can infer if blank."""
    return bool(str(emp.get("payRate", "")).strip()
                and str(emp.get("position", "")).strip()
                and str(emp.get("location", "")).strip())


def _get_row_and_drive_folder(row_id):
    """Fetch the employee's sheet row AND their individual Drive folder ID.

    The Drive folder ID is needed to save the exported CSV alongside their
    PDFs. We get it by asking GAS for the I-9 file's parent folder. If there
    is no I-9 file on record (shouldn't happen post-submission), folder_id is
    None and the caller should skip the Drive upload.
    """
    result = gas_post({"action": "getRow", "rowId": row_id}, timeout=30)
    if not result or "error" in result:
        return None, None
    row = result.get("row", [])
    emp = _row_to_employee(row)

    folder_id = None
    i9_file_id = str(row[19]).strip() if len(row) > 19 else ""
    if i9_file_id:
        parent_result = gas_post({
            "action": "getFileParent", "fileId": i9_file_id
        }, timeout=15)
        if parent_result and "folderId" in parent_result:
            folder_id = parent_result["folderId"]
    return emp, folder_id


def _stamp_export(row_id, system):
    """Write the current timestamp into the right stamp column."""
    sys_info = _EXPORT_SYSTEMS.get(system)
    if not sys_info:
        return
    gas_post({
        "action": "stampExport",
        "rowId": row_id,
        "col": sys_info["stamp_col"],
        "timestamp": datetime.utcnow().isoformat(),
    }, timeout=15)


def _csv_response(filename, rows):
    """Build a CSV file as a Flask response. rows is a list of lists; first
    row is the header."""
    buf = _io.StringIO()
    writer = _csv.writer(buf, quoting=_csv.QUOTE_MINIMAL)
    for r in rows:
        writer.writerow(r)
    buf.seek(0)
    return send_file(
        _io.BytesIO(buf.getvalue().encode("utf-8")),
        as_attachment=True,
        download_name=filename,
        mimetype="text/csv",
    )


# ─── CSV builders — one per system ───────────────────────────────────────────

def build_humanity_csv(emp):
    """Humanity's import accepts First Name, Last Name, Email, Phone, Location,
    Position, Wage, Start Date as column headers. Custom fields (any extra
    header) get auto-created. Format confirmed per helpcenter.humanity.com."""
    return [
        ["First Name", "Last Name", "Email", "Phone",
         "Location", "Position", "Wage", "Start Date"],
        [
            emp["firstName"], emp["lastName"], emp["email"], emp["phone"],
            emp["location"], emp["position"], emp["payRate"], emp["hireDateIso"],
        ],
    ]


def build_qu_csv(emp):
    """Qu POS template not yet confirmed — generic staff columns. Ask Bryan for
    a real export sample and update this mapping."""
    return [
        ["First Name", "Last Name", "Email", "Phone",
         "Role", "Location", "Start Date"],
        [
            emp["firstName"], emp["lastName"], emp["email"], emp["phone"],
            emp["position"], emp["location"], emp["hireDateIso"],
        ],
    ]


def build_zignal_csv(emp):
    """Zygnal template not yet confirmed — generic staff columns including a
    blank Qu PIN column Ashley fills in manually once Qu generates it."""
    return [
        ["First Name", "Last Name", "Role", "Location",
         "Qu PIN", "Start Date", "Email", "Phone"],
        [
            emp["firstName"], emp["lastName"], emp["position"], emp["location"],
            "",  # Qu PIN unknown until after Qu onboarding
            emp["hireDateIso"], emp["email"], emp["phone"],
        ],
    ]


def build_adp_csv(emp):
    """Generic ADP-compatible column set. RUN has no self-serve new-hire CSV
    import, but we ship one anyway in case Bryan connects a Marketplace
    integration (Deputy, HR Cloud, etc.) that accepts this shape."""
    return [
        ["First Name", "Middle Name", "Last Name", "SSN", "Date of Birth",
         "Gender", "Address Line 1", "City", "State", "ZIP",
         "Phone", "Email", "Hire Date", "Pay Rate", "Pay Frequency",
         "Department Code", "Position", "Location"],
        [
            emp["firstName"], "", emp["lastName"],
            emp["ssn_masked"],   # sheet only has masked SSN; real SSN lives in the PDFs
            emp["dobIso"], emp["gender"],
            emp["address1"], emp["city"], emp["state"], emp["zip"],
            emp["phone"], emp["email"],
            emp["hireDateIso"], emp["payRate"],
            "Biweekly",  # hardcoded default — confirm with Bryan
            emp["deptCode"], emp["position"], emp["location"],
        ],
    ]


def build_adp_cheatsheet_html(emp):
    """Printable one-page HTML formatted to mirror ADP Run's New Hire Wizard
    step order. Ashley keeps this open on a second monitor while she tabs
    through Run's wizard."""
    esc = _html.escape
    ssn_note = "(Full SSN on printed W-4 in Drive folder — this sheet shows masked only)"
    rows = [
        ("Personal", [
            ("Legal First Name", emp["firstName"]),
            ("Legal Last Name",  emp["lastName"]),
            ("SSN",              f"{emp['ssn_masked']} <small style='color:#8792a2'>{ssn_note}</small>"),
            ("Date of Birth",    emp["dob"]),
            ("Gender",           emp["gender"] or "—"),
        ]),
        ("Contact", [
            ("Address",          f"{emp['address1']}, {emp['city']} {emp['state']} {emp['zip']}"),
            ("Phone",            emp["phone"]),
            ("Email",            emp["email"]),
        ]),
        ("Employment", [
            ("Hire Date",        emp["hireDate"]),
            ("Pay Rate",         f"${emp['payRate']}/hr" if emp["payRate"] else "—"),
            ("Pay Frequency",    "Biweekly (confirm)"),
            ("Position",         emp["position"]),
            ("Location",         emp["location"]),
            ("Department Code",  emp["deptCode"] or "—"),
        ]),
        ("Emergency Contact", [
            ("Name",             emp["ecName"]),
            ("Relationship",     emp["ecRelationship"]),
            ("Phone",            emp["ecPhone"]),
        ]),
    ]
    sections_html = ""
    for section_title, fields in rows:
        field_html = "".join(
            f"<tr><td class='lbl'>{esc(label)}</td>"
            f"<td class='val'>{val if '<small' in str(val) else esc(str(val))}</td></tr>"
            for label, val in fields
        )
        sections_html += (
            f"<section><h2>{esc(section_title)}</h2>"
            f"<table>{field_html}</table></section>"
        )

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>ADP Run cheat sheet — {esc(emp['firstName'])} {esc(emp['lastName'])}</title>
<style>
  @media print {{ @page {{ margin: 0.5in; }} .no-print {{ display: none; }} }}
  body {{ font-family: -apple-system, 'Segoe UI', Roboto, sans-serif;
         font-size: 13px; color: #1a1f36; max-width: 720px;
         margin: 24px auto; padding: 0 24px; }}
  h1 {{ font-size: 18px; margin-bottom: 4px; }}
  .sub {{ color: #8792a2; font-size: 12px; margin-bottom: 20px; }}
  section {{ margin-bottom: 18px; page-break-inside: avoid; }}
  section h2 {{ font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em;
               color: #8792a2; margin-bottom: 6px; border-bottom: 1px solid #e3e8ee;
               padding-bottom: 4px; }}
  table {{ width: 100%; border-collapse: collapse; }}
  td {{ padding: 5px 0; vertical-align: top; }}
  td.lbl {{ color: #4f566b; width: 35%; }}
  td.val {{ font-weight: 500; }}
  .no-print {{ background: #f6f9fc; padding: 12px; border-radius: 6px;
              margin-bottom: 16px; font-size: 12px; color: #4f566b; }}
  .no-print button {{ background: #000f9f; color: white; border: 0; padding: 8px 14px;
                     border-radius: 6px; font-size: 12px; cursor: pointer; margin-left: 8px; }}
</style>
</head><body>
<div class="no-print">
  Open ADP Run on your main screen. Work through the New Hire Wizard and copy
  each field below into the matching wizard field.
  <button onclick="window.print()">Print</button>
</div>
<h1>{esc(emp['firstName'])} {esc(emp['lastName'])}</h1>
<div class="sub">ADP Run new-hire cheat sheet · generated {datetime.utcnow().strftime('%b %d, %Y')}</div>
{sections_html}
</body></html>"""


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/submissions/<int:row_id>/employment", methods=["PATCH"])
def update_employment(row_id):
    """Ashley saves pay rate / position / location / dept code / hire date
    from the admin sidebar. Writes to columns AA–AE via GAS."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    data = request.get_json() or {}
    result = gas_post({
        "action":   "updateEmployment",
        "rowId":    row_id,
        "payRate":  data.get("payRate", ""),
        "position": data.get("position", ""),
        "location": data.get("location", ""),
        "deptCode": data.get("deptCode", ""),
        "hireDate": data.get("hireDate", ""),
    }, timeout=20, gas_url=demo_gas_url(request))
    if not result or "error" in result:
        return jsonify({"error": (result or {}).get("error", "GAS error")}), 500
    return jsonify({"status": "ok"})


@app.route("/submissions/<int:row_id>/export/<system>", methods=["GET"])
def export_for_system(row_id, system):
    """Download a CSV for one of {humanity, qu, zignal, adp} OR the ADP
    cheat-sheet as an HTML page (system='adp_sheet' and format=html).

    Uses ?key=<ADMIN_API_KEY> query param for auth because this is a direct
    file download triggered by window.open() — can't set custom headers.
    """
    # Query-param auth for file downloads (no custom header possible)
    provided = request.args.get("key", "") or request.headers.get("X-API-Key", "")
    valid_keys = {k for k in [ADMIN_API_KEY, DEMO_ADMIN_API_KEY] if k}
    if valid_keys and provided not in valid_keys:
        return jsonify({"error": "Unauthorized"}), 401

    if system not in _EXPORT_SYSTEMS:
        return jsonify({"error": f"Unknown system: {system}"}), 400

    emp, folder_id = _get_row_and_drive_folder(row_id)
    if not emp:
        return jsonify({"error": "Employee not found"}), 404

    if not _employment_complete(emp):
        return jsonify({
            "error": "Employment details incomplete — set pay rate, position, "
                     "and location before exporting"
        }), 400

    name_slug = f"{emp['firstName']}_{emp['lastName']}".replace(" ", "_")

    # Special case: ADP cheat sheet is HTML, not CSV
    if system == "adp_sheet":
        html_content = build_adp_cheatsheet_html(emp)
        # Save to Drive as HTML for audit trail
        if folder_id:
            upload_file_to_drive(
                f"ADP_cheatsheet_{name_slug}.html",
                html_content.encode("utf-8"),
                "text/html",
                folder_id=folder_id,
            )
        _stamp_export(row_id, "adp_sheet")
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}

    # CSV path
    builders = {
        "humanity": build_humanity_csv,
        "qu":       build_qu_csv,
        "zignal":   build_zignal_csv,
        "adp":      build_adp_csv,
    }
    rows = builders[system](emp)
    filename = f"{system}_{name_slug}.csv"

    # Also save a copy to the employee's Drive folder (audit trail)
    if folder_id:
        csv_buf = _io.StringIO()
        _csv.writer(csv_buf, quoting=_csv.QUOTE_MINIMAL).writerows(rows)
        upload_file_to_drive(
            filename, csv_buf.getvalue().encode("utf-8"),
            "text/csv", folder_id=folder_id,
        )

    _stamp_export(row_id, system)
    return _csv_response(filename, rows)


# ── /de112b: working papers (under-18 employment certificate) ─────────────
# Two endpoints:
#   PATCH /submissions/<id>/working-papers       — toggle 'given' or 'returned' boolean
#   POST  /submissions/<id>/working-papers/upload — upload photo of completed papers
# Both delegate to GAS for sheet writes / Drive uploads.

@app.route("/submissions/<int:row_id>/working-papers", methods=["PATCH"])
def update_working_papers(row_id):
    """Toggle a working-papers boolean (given or returned)."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    body = request.get_json(silent=True) or {}
    field = body.get("field")
    if field not in ("given", "returned"):
        return jsonify({"error": "field must be 'given' or 'returned'"}), 400
    result = gas_post({
        "action": "setWorkingPapers",
        "rowId":  row_id,
        "field":  field,
        "value":  body.get("value", True),
    }, timeout=90, gas_url=demo_gas_url(request))
    if result is None:
        return jsonify({"error": "GAS error"}), 502
    return jsonify(result)


@app.route("/submissions/<int:row_id>/working-papers/upload", methods=["POST"])
def upload_working_papers(row_id):
    """Upload a photo of the completed working papers to the employee's Drive folder.

    Expects JSON body with:
      - fileData: base64-encoded image
      - mimeType: e.g. 'image/jpeg' (optional, defaults to image/jpeg)
      - filename: optional display name
    """
    if request.headers.get("X-API-Key") != ADMIN_API_KEY:
        return jsonify({"error": "Unauthorized"}), 401
    body = request.get_json(silent=True) or {}
    if not body.get("fileData"):
        return jsonify({"error": "fileData (base64) required"}), 400
    payload = {
        "secret":   GAS_SECRET,
        "action":   "uploadWorkingPapers",
        "rowId":    row_id,
        "fileData": body["fileData"],
        "mimeType": body.get("mimeType", "image/jpeg"),
        "filename": body.get("filename") or f"working-papers-{row_id}.jpg",
    }
    try:
        r = http_requests.post(GAS_WEBHOOK_URL, json=payload, timeout=90)
        r.raise_for_status()
        return jsonify(r.json())
    except Exception as e:
        return jsonify({"error": f"GAS error: {e}"}), 502


# ── /de112b: employee lifecycle status ────────────────────────────────────
# Employees live in one of three states tracked by the Status column (AN):
#   onboarding → active → inactive
# Transitions:
#   onboarding → active    when Ashley clicks "Mark as fully onboarded"
#   active     → inactive  when Ashley clicks "Mark as inactive"
# Reactivation (inactive → active/onboarding) isn't wired up yet — will come later.

_VALID_STATUSES = ("onboarding", "active", "inactive")


@app.route("/submissions/<int:row_id>/status", methods=["PATCH"])
def update_status(row_id):
    """Update an employee's lifecycle status (onboarding/active/inactive)."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    body = request.get_json(silent=True) or {}
    status = body.get("status")
    if status not in _VALID_STATUSES:
        return jsonify({"error": f"status must be one of {_VALID_STATUSES}"}), 400
    result = gas_post({
        "action": "setStatus",
        "rowId":  row_id,
        "status": status,
        "reason": body.get("reason", ""),
    }, timeout=90, gas_url=demo_gas_url(request))
    if result is None:
        return jsonify({"error": "GAS error"}), 502
    return jsonify(result)


@app.route("/submissions/<int:row_id>/folder-url", methods=["GET"])
def get_folder_url(row_id):
    """Return the Drive folder URL for an employee.

    GAS checks the cached driveFolderId column (AR) first; falls back to
    deriving from the I-9 file's parent. This means imported/legacy employees
    that point at a shared folder (no I-9 file in our system) still work.
    """
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    result = gas_post({
        "action": "getEmployeeFolderUrl",
        "rowId":  row_id,
    }, timeout=90, gas_url=demo_gas_url(request))
    if result is None:
        return jsonify({"error": "GAS error"}), 502
    if result.get("error"):
        return jsonify({"url": None, "error": result["error"]}), 404
    url = result.get("url")
    if not url:
        return jsonify({"url": None, "error": "No Drive folder found"}), 404
    return jsonify({"url": url, "folderId": result.get("folderId"), "source": result.get("source")})


@app.route("/submissions/<int:row_id>", methods=["DELETE"])
def delete_submission(row_id):
    """Delete a test/demo submission. GAS only allows deletion of rows flagged testEntry=TRUE."""
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401
    result = gas_post({
        "action": "deleteRow",
        "rowId":  row_id,
    }, timeout=30, gas_url=demo_gas_url(request))
    if result is None:
        return jsonify({"error": "GAS error"}), 502
    if result.get("error"):
        return jsonify({"error": result["error"]}), 400
    return jsonify({"status": "ok"})
