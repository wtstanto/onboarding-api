"""
Auntie Anne's Onboarding PDF Filler API
----------------------------------------
POST /fill                    → fill PDFs, upload to Drive, log to Sheet, return ZIP
GET  /health                  → 200 OK
GET  /submissions             → employee list from Sheet  (X-API-Key required)
PATCH /submissions/<id>/i9    → fill I-9 Section 2, replace Drive file, mark Sheet complete
"""

import os
import io
import base64
import zipfile
from datetime import date, datetime

import requests as http_requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject

app = Flask(__name__)
CORS(app)

# ─── Paths ───────────────────────────────────────────────────────────────────

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
W4_PATH    = os.path.join(BASE_DIR, "forms", "w4.pdf")
DE_W4_PATH = os.path.join(BASE_DIR, "forms", "de_w4.pdf")
I9_PATH    = os.path.join(BASE_DIR, "forms", "i9.pdf")

# ─── Environment ─────────────────────────────────────────────────────────────

GAS_WEBHOOK_URL = os.environ.get("GAS_WEBHOOK_URL", "")
ADMIN_API_KEY   = os.environ.get("ADMIN_API_KEY", "")
GAS_SECRET      = os.environ.get("GAS_SECRET", "")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")


def check_api_key(req):
    if not ADMIN_API_KEY:
        return True
    return req.headers.get("X-API-Key", "") == ADMIN_API_KEY


# ─── Google Apps Script helpers ──────────────────────────────────────────────

def gas_post(payload, timeout=30):
    """POST to GAS webhook, follow redirect as GET, return parsed JSON or None."""
    if not GAS_WEBHOOK_URL:
        return None
    try:
        payload["secret"] = GAS_SECRET
        headers = {"Content-Type": "application/json"}
        res = http_requests.post(
            GAS_WEBHOOK_URL, json=payload,
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


def upload_file_to_drive(filename, file_bytes, mimetype):
    """Upload file_bytes to Drive via GAS. Returns (fileId, url) or (None, None)."""
    if not DRIVE_FOLDER_ID:
        print(f"[Drive] skipping upload — DRIVE_FOLDER_ID not set")
        return None, None
    print(f"[Drive] uploading {filename} ({len(file_bytes)} bytes) to folder {DRIVE_FOLDER_ID}")
    result = gas_post({
        "action":   "uploadFile",
        "folderId": DRIVE_FOLDER_ID,
        "filename": filename,
        "mimeType": mimetype,
        "fileData": base64.b64encode(file_bytes).decode("utf-8"),
    }, timeout=60)
    print(f"[Drive] upload result for {filename}: {result}")
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


def log_to_sheet(data, zip_drive_url="", i9_file_id=""):
    ssn    = data.get("ssn", "")
    digits = "".join(c for c in ssn if c.isdigit())
    masked = f"***-**-{digits[-4:]}" if len(digits) >= 4 else "***"
    result = gas_post({
        "action":      "log",
        "submittedAt": datetime.utcnow().isoformat(),
        "firstName":   data.get("firstName",  ""),
        "lastName":    data.get("lastName",   ""),
        "email":       data.get("email",      ""),
        "phone":       data.get("phone",      ""),
        "ssn":         masked,
        "dob":         data.get("dob",        ""),
        "address1":    data.get("address1",   ""),
        "city":        data.get("city",       ""),
        "state":       data.get("state",      ""),
        "zip":         data.get("zip",        ""),
        "zipDriveUrl": zip_drive_url,
        "i9FileId":    i9_file_id,
        "startDate":   data.get("startDate",  ""),
    })
    return result.get("rowId") if result else None


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

    for page in writer.pages:
        writer.update_page_form_field_values(page, text_fields, auto_regenerate=False)

    if checkbox_fields:
        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    annot_obj = annot_ref.get_object()
                    t = annot_obj.get("/T")
                    if not t:
                        continue
                    t_str = str(t)
                    for field_name, value in checkbox_fields.items():
                        if t_str == field_name or field_name.endswith("." + t_str):
                            annot_obj.update({
                                NameObject("/V"):  NameObject(value),
                                NameObject("/AS"): NameObject(value),
                            })

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


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
        "topmostSubform[0].Page1[0].f1_10[0]": data.get("additionalWithholding", ""),
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
        "USCIS ANumber":                            data.get("uscisNumber", ""),
    }
    cit     = data.get("citizenship", "citizen")
    cit_map = {"citizen": "CB_1", "noncitizen": "CB_2", "lpr": "CB_3", "authorized": "CB_4"}
    checkbox_fields = {
        field_id: "/Yes" if cit == key else "/Off"
        for key, field_id in cit_map.items()
    }
    return fill_pdf_to_bytes(I9_PATH, text_fields, checkbox_fields)


def fill_i9_section2(i9_bytes, section2_data, start_date=""):
    """Fill employer Section 2 onto an existing I-9 PDF (already has Section 1)."""
    today      = fmt_date(datetime.utcnow().date().isoformat())
    first_day  = fmt_date(start_date) if start_date else today
    doc_title  = section2_data.get("docTitle",  "")
    doc_number = section2_data.get("docNumber", "")
    issuer     = section2_data.get("issuer",    "")
    exp_date   = section2_data.get("expDate",   "")
    emp_name   = section2_data.get("empName",   "")
    emp_addr   = section2_data.get("employerAddress", "")
    emp_org    = "Auntie Anne's"

    # Fill all numbered copies of the Section 2 document fields (copies 0, 1, 2)
    text_fields = {
        # Document info — covers List A copies (0/1/2 = employer/employee/retention copy)
        "Document Title 0":   doc_title,
        "Document Title 1":   doc_title,
        "Document Title 2":   doc_title,
        "Document Number 0":  doc_number,
        "Document Number 1":  doc_number,
        "Document Number 2":  doc_number,
        "Expiration Date 0":  exp_date,
        "Expiration Date 1":  exp_date,
        "Expiration Date 2":  exp_date,
        # List A issuing authority
        "List A":             issuer,
        # List B / C fallback (in case it's B+C instead of A)
        "List B Document 1 Title":    doc_title,
        "List B Issuing Authority 1": issuer,
        "List B Document Number 1":   doc_number,
        "List B Expiration Date 1":   exp_date,
        # Employer info
        "FirstDayEmployed mmddyyyy": first_day,
        "Last Name First Name and Title of Employer or Authorized Representative": emp_name,
        "Name of Emp or Auth Rep 0": emp_name,
        "Name of Emp or Auth Rep 1": emp_name,
        "Name of Emp or Auth Rep 2": emp_name,
        # Typed name as signature (actual digital signatures need additional tooling)
        "Signature of Emp Rep 0":    emp_name,
        "Signature of Emp Rep 1":    emp_name,
        "Signature of Emp Rep 2":    emp_name,
        "Signature of Employer or AR": emp_name,
        # Section 2 date
        "S2 Todays Date mmddyyyy": today,
        "Todays Date 0":           today,
        "Todays Date 1":           today,
        "Todays Date 2":           today,
        # Employer org
        "Employers Business or Org Name":    emp_org,
        "Employers Business or Org Address": emp_addr,
    }
    return fill_pdf_to_bytes(io.BytesIO(i9_bytes), text_fields)


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

        w4_bytes    = fill_w4(data)
        de_w4_bytes = fill_de_w4(data)
        i9_bytes    = fill_i9_section1(data)

        zip_filename = f"{name}_onboarding_docs.zip"
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{name}_W4_Federal.pdf",  w4_bytes)
            zf.writestr(f"{name}_W4_Delaware.pdf", de_w4_bytes)
            zf.writestr(f"{name}_I9.pdf",          i9_bytes)
        zip_bytes = zip_buf.getvalue()

        # Upload I-9 PDF to Drive (needed later for Section 2 fill)
        i9_file_id, _ = upload_file_to_drive(
            f"{name}_I9.pdf", i9_bytes, "application/pdf"
        )

        # Upload full ZIP to Drive
        _, zip_drive_url = upload_file_to_drive(
            zip_filename, zip_bytes, "application/zip"
        )

        # Log to Sheet
        log_to_sheet(data, zip_drive_url=zip_drive_url or "", i9_file_id=i9_file_id or "")

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
    result = gas_post({"action": "getAll"})
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

    new_i9_file_id = ""

    # Get the employee's row to find their I-9 Drive file ID and start date
    row_result = gas_post({"action": "getRow", "rowId": row_id})
    if row_result and "row" in row_result:
        row        = row_result["row"]
        i9_file_id = row[19] if len(row) > 19 else ""  # column T
        start_date = row[20] if len(row) > 20 else ""  # column U

        if i9_file_id:
            # Download the existing I-9 (has Section 1 already filled)
            i9_bytes = download_file_from_drive(i9_file_id)
            if i9_bytes:
                # Fill Section 2 on top of it
                updated_i9 = fill_i9_section2(i9_bytes, data, str(start_date))
                # Replace the Drive file with the completed version
                new_file_id, _ = replace_drive_file(
                    i9_file_id,
                    f"I9_COMPLETED_row{row_id}.pdf",
                    updated_i9,
                )
                new_i9_file_id = new_file_id or ""

    # Update the Sheet row (mark complete, store I-9 completion data)
    result = gas_post({
        "action":      "completeI9",
        "rowId":       row_id,
        "docTitle":    data.get("docTitle",  ""),
        "docNumber":   data.get("docNumber", ""),
        "issuer":      data.get("issuer",    ""),
        "expDate":     data.get("expDate",   ""),
        "empName":     data.get("empName",   ""),
        "newI9FileId": new_i9_file_id,
    })

    if result is None:
        return jsonify({"error": "GAS webhook not configured"}), 500
    if "error" in result:
        return jsonify({"error": result["error"]}), 500
    return jsonify({"status": "ok"})


@app.route("/debug", methods=["GET"])
def debug():
    return jsonify({
        "gas_configured":    bool(GAS_WEBHOOK_URL),
        "drive_configured":  bool(DRIVE_FOLDER_ID),
        "admin_key_set":     bool(ADMIN_API_KEY),
        "gas_secret_set":    bool(GAS_SECRET),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
