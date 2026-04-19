"""
Auntie Anne's Onboarding PDF Filler API
----------------------------------------
POST /fill  →  accepts JSON employee data, returns ZIP of filled PDFs
GET  /health → returns 200 OK (for Railway health checks)
"""

import os
import io
import zipfile
from datetime import date
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject

app = Flask(__name__)
CORS(app)

# Paths to blank PDF templates (relative to this file)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
W4_PATH    = os.path.join(BASE_DIR, "forms", "w4.pdf")
DE_W4_PATH = os.path.join(BASE_DIR, "forms", "de_w4.pdf")
I9_PATH    = os.path.join(BASE_DIR, "forms", "i9.pdf")


# ─── Helpers ────────────────────────────────────────────────────────────────

def fmt_date(val):
    if not val:
        return ""
    try:
        d = date.fromisoformat(val[:10])
        return d.strftime("%m/%d/%Y")
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
    """Fill a PDF and return it as bytes."""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.append(reader)

    # Fill text fields across all pages
    for page in writer.pages:
        writer.update_page_form_field_values(
            page, text_fields, auto_regenerate=False
        )

    # Fill checkboxes/radio buttons via direct annotation update
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
                                NameObject("/V"): NameObject(value),
                                NameObject("/AS"): NameObject(value),
                            })

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


# ─── Form Fillers ────────────────────────────────────────────────────────────

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
        "topmostSubform[0].Page1[0].f1_10[0]":            data.get("additionalWithholding", ""),
        "topmostSubform[0].Page1[0].f1_12[0]":            data.get("employerName", "Auntie Anne's"),
        "topmostSubform[0].Page1[0].f1_13[0]":            fmt_date(data.get("startDate")),
        "topmostSubform[0].Page1[0].f1_14[0]":            data.get("employerEIN", ""),
    }

    fs = data.get("filingStatus", "single")
    checkbox_fields = {
        "topmostSubform[0].Page1[0].c1_1[0]": "/1"  if fs == "single"  else "/Off",
        "topmostSubform[0].Page1[0].c1_1[1]": "/2"  if fs == "married" else "/Off",
        "topmostSubform[0].Page1[0].c1_1[2]": "/3"  if fs == "hoh"     else "/Off",
        "topmostSubform[0].Page1[0].c1_2[0]": "/1"  if data.get("multipleJobs") == "yes" else "/Off",
        "topmostSubform[0].Page1[0].c1_3[0]": "/1"  if data.get("exempt") == "yes"        else "/Off",
    }

    return fill_pdf_to_bytes(W4_PATH, text_fields, checkbox_fields)


def fill_de_w4(data):
    first_initial = data.get("firstName", "")
    if data.get("middleName"):
        first_initial += " " + data["middleName"][0]

    text_fields = {
        "1Firstnameinitial":    first_initial,
        "1Lastname":            data.get("lastName", ""),
        "2Taxpayerid":          fmt_ssn(data.get("ssn", "")),
        "2Homeaddress":         data.get("address1", ""),
        "3Cityortown":          data.get("city", ""),
        "3State":               data.get("state", ""),
        "3Zipcode":             data.get("zip", ""),
        "4Totalnumberdependents": str(data.get("deAllowances", "0")),
        "5Additionalamount":    data.get("deAdditional", ""),
        "6Employersname":       data.get("employerName", "Auntie Anne's"),
        "7Firstdayofemployment": fmt_date(data.get("startDate")),
        "8Taxpayeridein":       data.get("employerEIN", ""),
    }

    de_fs = data.get("deFilingStatus", "single")
    checkbox_fields = {
        "3status": "/Single" if de_fs == "single" else "/Married",
    }

    return fill_pdf_to_bytes(DE_W4_PATH, text_fields, checkbox_fields)


def fill_i9(data):
    today = fmt_date(data.get("signatureDate") or date.today().isoformat())

    text_fields = {
        "Last Name (Family Name)":             data.get("lastName", ""),
        "First Name Given Name":               data.get("firstName", ""),
        "Employee Middle Initial (if any)":    data.get("middleName", "")[:1] if data.get("middleName") else "",
        "Employee Other Last Names Used (if any)": data.get("otherNames", "N/A"),
        "Address Street Number and Name":      data.get("address1", ""),
        "Apt Number (if any)":                 data.get("address2", ""),
        "City or Town":                        data.get("city", ""),
        "ZIP Code":                            data.get("zip", ""),
        "Date of Birth mmddyyyy":              fmt_date(data.get("dob")),
        "US Social Security Number":           fmt_ssn(data.get("ssn", "")),
        "Telephone Number":                    data.get("phone", ""),
        "Employees E-mail Address":            data.get("email", ""),
        "Today's Date mmddyyy":                today,
        "USCIS ANumber":                       data.get("uscisNumber", ""),
    }

    cit = data.get("citizenship", "citizen")
    cit_map = {"citizen": "CB_1", "noncitizen": "CB_2", "lpr": "CB_3", "authorized": "CB_4"}
    checkbox_fields = {
        field_id: "/Yes" if cit == key else "/Off"
        for key, field_id in cit_map.items()
    }

    return fill_pdf_to_bytes(I9_PATH, text_fields, checkbox_fields)


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200


@app.route("/fill", methods=["POST"])
def fill():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No JSON data received"}), 400

    try:
        first  = data.get("firstName", "employee").strip()
        last   = data.get("lastName", "").strip()
        name   = f"{first}_{last}".replace(" ", "_")

        w4_bytes    = fill_w4(data)
        de_w4_bytes = fill_de_w4(data)
        i9_bytes    = fill_i9(data)

        # Package into a ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{name}_W4_Federal.pdf",  w4_bytes)
            zf.writestr(f"{name}_W4_Delaware.pdf", de_w4_bytes)
            zf.writestr(f"{name}_I9.pdf",          i9_bytes)
        zip_buf.seek(0)

        zip_filename = f"{name}_onboarding_docs.zip"
        return send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_filename
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
