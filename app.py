"""
Auntie Anne's Onboarding PDF Filler API
----------------------------------------
POST /fill                    → fill PDFs, upload ZIP to Drive, log row to Sheet, return ZIP
GET  /health                  → 200 OK (Railway health check)
GET  /submissions             → JSON list of employees from Google Sheet  (requires X-API-Key)
PATCH /submissions/<id>/i9    → mark I-9 complete in Google Sheet         (requires X-API-Key)
"""

import os
import io
import json
import zipfile
from datetime import date, datetime

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

GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS", "")
DRIVE_FOLDER_ID   = os.environ.get("DRIVE_FOLDER_ID", "")
SHEET_ID          = os.environ.get("SHEET_ID", "")
ADMIN_API_KEY     = os.environ.get("ADMIN_API_KEY", "")

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Lazy-initialised Google clients (reused across requests)
_drive_client  = None
_sheets_client = None


def get_google_clients():
    global _drive_client, _sheets_client
    if _drive_client and _sheets_client:
        return _drive_client, _sheets_client
    if not GOOGLE_CREDS_JSON:
        return None, None
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds = service_account.Credentials.from_service_account_info(
            json.loads(GOOGLE_CREDS_JSON), scopes=SCOPES
        )
        _drive_client  = build("drive",  "v3", credentials=creds)
        _sheets_client = build("sheets", "v4", credentials=creds)
        return _drive_client, _sheets_client
    except Exception as exc:
        print(f"[Google] client init error: {exc}")
        return None, None


def upload_to_drive(drive, zip_bytes, filename):
    """Upload zip_bytes to the configured Drive folder; return the web-view link."""
    if not drive or not DRIVE_FOLDER_ID:
        return None
    try:
        from googleapiclient.http import MediaIoBaseUpload
        file_meta = {"name": filename, "parents": [DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(io.BytesIO(zip_bytes), mimetype="application/zip")
        f = drive.files().create(
            body=file_meta, media_body=media, fields="id,webViewLink"
        ).execute()
        return f.get("webViewLink", "")
    except Exception as exc:
        print(f"[Drive] upload error: {exc}")
        return None


def log_to_sheet(sheets, data, drive_url):
    """Append one row to Sheet1 and return the sheet row number (1-indexed)."""
    if not sheets or not SHEET_ID:
        return None
    try:
        ssn    = data.get("ssn", "")
        digits = "".join(c for c in ssn if c.isdigit())
        masked = f"***-**-{digits[-4:]}" if len(digits) >= 4 else "***"

        row = [
            datetime.utcnow().isoformat(),  # A  submittedAt
            data.get("firstName",  ""),     # B
            data.get("lastName",   ""),     # C
            data.get("email",      ""),     # D
            data.get("phone",      ""),     # E
            masked,                         # F  ssn (masked)
            data.get("dob",        ""),     # G
            data.get("address1",   ""),     # H
            data.get("city",       ""),     # I
            data.get("state",      ""),     # J
            data.get("zip",        ""),     # K
            "pending",                      # L  i9Status
            drive_url or "",                # M  driveUrl
            "", "", "", "", "", "",         # N-S  i9 fields (filled later)
        ]

        result = sheets.spreadsheets().values().append(
            spreadsheetId=SHEET_ID,
            range="Sheet1!A:S",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()

        # Parse the actual row number from the returned range (e.g. "Sheet1!A3:S3")
        updated = result.get("updates", {}).get("updatedRange", "")
        try:
            return int(updated.split("!")[1].split(":")[0][1:])
        except Exception:
            return None
    except Exception as exc:
        print(f"[Sheets] log error: {exc}")
        return None


def check_api_key(req):
    """Return True if the request carries a valid admin API key (or none is configured)."""
    if not ADMIN_API_KEY:
        return True
    return req.headers.get("X-API-Key", "") == ADMIN_API_KEY


# ─── PDF helpers ─────────────────────────────────────────────────────────────

def fmt_date(val):
    if not val:
        return ""
    try:
        return date.fromisoformat(val[:10]).strftime("%m/%d/%Y")
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


def fill_pdf_to_bytes(input_path, text_fields, checkbox_fields=None):
    reader = PdfReader(input_path)
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
    checkbox_fields = {
        "3status": "/Single" if de_fs == "single" else "/Married",
    }

    return fill_pdf_to_bytes(DE_W4_PATH, text_fields, checkbox_fields)


def fill_i9(data):
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
        i9_bytes    = fill_i9(data)

        zip_filename = f"{name}_onboarding_docs.zip"
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{name}_W4_Federal.pdf",  w4_bytes)
            zf.writestr(f"{name}_W4_Delaware.pdf", de_w4_bytes)
            zf.writestr(f"{name}_I9.pdf",          i9_bytes)
        zip_bytes = zip_buf.getvalue()

        # Upload to Drive and log to Sheet — errors here don't fail the request
        drive, sheets = get_google_clients()
        drive_url = upload_to_drive(drive, zip_bytes, zip_filename)
        log_to_sheet(sheets, data, drive_url)

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

    _, sheets = get_google_clients()
    if not sheets or not SHEET_ID:
        return jsonify([])

    try:
        result = sheets.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range="Sheet1!A:S",
        ).execute()

        employees = []
        for i, row in enumerate(result.get("values", [])):
            while len(row) < 19:
                row.append("")

            i9_complete = (row[11] or "pending") == "complete"
            employees.append({
                "id":          i + 1,          # 1-indexed sheet row number
                "submittedAt": row[0],
                "firstName":   row[1],
                "lastName":    row[2],
                "email":       row[3],
                "phone":       row[4],
                "ssn":         row[5],
                "dob":         row[6],
                "address1":    row[7],
                "city":        row[8],
                "state":       row[9],
                "zip":         row[10],
                "i9Status":    row[11] or "pending",
                "driveUrl":    row[12],
                # Nest i9 detail the way the admin UI expects it
                "i9s2": {
                    "docTitle":      row[13],
                    "docNumber":     row[14],
                    "issuer":        row[15],
                    "expDate":       row[16],
                    "verifiedDate":  row[17],
                    "empName":       row[18],
                } if i9_complete else None,
            })

        return jsonify(employees)

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/submissions/<int:row_id>/i9", methods=["PATCH"])
def complete_i9(row_id):
    if not check_api_key(request):
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data"}), 400

    _, sheets = get_google_clients()
    if not sheets or not SHEET_ID:
        return jsonify({"error": "Google Sheets not configured"}), 500

    try:
        # Columns L–S (i9Status, driveUrl, i9Doc, i9DocNumber, i9Issuer, i9ExpDate, i9VerifiedDate, i9VerifiedBy)
        values = [
            "complete",
            data.get("driveUrl", ""),           # preserve the drive link
            data.get("docTitle",  ""),
            data.get("docNumber", ""),
            data.get("issuer",    ""),
            data.get("expDate",   ""),
            datetime.utcnow().isoformat(),
            data.get("empName",   ""),
        ]

        sheets.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"Sheet1!L{row_id}:S{row_id}",
            valueInputOption="RAW",
            body={"values": [values]},
        ).execute()

        return jsonify({"status": "ok"})

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
