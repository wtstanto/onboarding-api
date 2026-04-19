"""
fill_forms.py
Fills W-4, DE W-4, and I-9 from onboarding form JSON submission.

Usage:
  python fill_forms.py <submission.json>

Outputs:
  w4_filled.pdf
  de_w4_filled.pdf
  i9_filled.pdf

Expected submission JSON keys (matching onboarding form):
  firstName, middleName, lastName, ssn, dob
  address1, address2, city, state, zip
  phone, email
  filingStatus         (single | married | hoh)
  multipleJobs         (yes | no)
  childDependents      (number)
  otherDependents      (number)
  additionalWithholding
  exempt               (yes | no)
  deFilingStatus       (single | married)
  deAllowances         (number)
  deAdditional
  otherNames           (I-9 other last names)
  citizenship          (citizen | noncitizen | lpr | authorized)
  uscisNumber          (if lpr)
  startDate            (YYYY-MM-DD, first day of employment)
  employerName         (stub for now)
  employerAddress      (stub)
  employerEIN          (stub)
  signatureDate        (YYYY-MM-DD)
"""

import sys
import json
import copy
from datetime import date
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, ArrayObject


# ─── Helpers ────────────────────────────────────────────────────────────────

def fmt_date(val, fmt="mmddyyyy"):
    """Convert YYYY-MM-DD to MM/DD/YYYY or MMDDYYYY."""
    if not val:
        return ""
    try:
        d = date.fromisoformat(val)
        if fmt == "mmddyyyy":
            return d.strftime("%m/%d/%Y")
        elif fmt == "compact":
            return d.strftime("%m%d%Y")
        return str(val)
    except Exception:
        return str(val)

def fmt_ssn(val):
    """Ensure SSN is formatted XXX-XX-XXXX."""
    if not val:
        return ""
    digits = "".join(c for c in val if c.isdigit())
    if len(digits) == 9:
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
    return val

def child_credit_amount(n):
    """$2000 per child under 17."""
    try:
        return str(int(n) * 2000) if int(n) > 0 else ""
    except Exception:
        return ""

def other_dependent_amount(n):
    """$500 per other dependent."""
    try:
        return str(int(n) * 500) if int(n) > 0 else ""
    except Exception:
        return ""

def fill_pdf(input_path, output_path, field_values: dict, checkbox_fields: dict = None):
    """
    Fill a fillable PDF.
    field_values: {field_id: text_value}
    checkbox_fields: {field_id: True/False or checked_value/unchecked_value}
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    writer.append(reader)

    # Fill text fields
    writer.update_page_form_field_values(
        writer.pages[0], field_values, auto_regenerate=False
    )

    # Handle multi-page text fields
    all_fields = {}
    for page in writer.pages:
        writer.update_page_form_field_values(
            page, field_values, auto_regenerate=False
        )

    # Handle checkboxes and radio buttons via direct annotation update
    if checkbox_fields:
        for page in writer.pages:
            if "/Annots" in page:
                for annot_ref in page["/Annots"]:
                    annot_obj = annot_ref.get_object()
                    t = annot_obj.get("/T")
                    if not t:
                        continue
                    t_str = str(t)
                    # Match on short name (e.g. "c1_1[0]") OR full path
                    for field_name, value in checkbox_fields.items():
                        if t_str == field_name or field_name.endswith("." + t_str):
                            annot_obj.update({
                                NameObject("/V"): NameObject(value),
                                NameObject("/AS"): NameObject(value),
                            })

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"  ✓ Written: {output_path}")


# ─── W-4 Fill ───────────────────────────────────────────────────────────────

def fill_w4(data, output_path):
    print("Filling Federal W-4...")

    first_mid = data.get("firstName", "")
    if data.get("middleName"):
        first_mid += " " + data["middleName"][0]  # just initial

    city_state_zip = f"{data.get('city','')} {data.get('state','')} {data.get('zip','')}"

    text_fields = {
        # Row 1: first name | last name | SSN  (f1_01, f1_02 left cols; f1_05 right col)
        "topmostSubform[0].Page1[0].Step1a[0].f1_01[0]": first_mid,
        "topmostSubform[0].Page1[0].Step1a[0].f1_02[0]": data.get("lastName", ""),
        "topmostSubform[0].Page1[0].f1_05[0]": fmt_ssn(data.get("ssn", "")),
        # Row 2: Address
        "topmostSubform[0].Page1[0].Step1a[0].f1_03[0]": data.get("address1", ""),
        # Row 3: City, state, ZIP
        "topmostSubform[0].Page1[0].Step1a[0].f1_04[0]": city_state_zip,
        # Step 3 - dependents
        "topmostSubform[0].Page1[0].Step3_ReadOrder[0].f1_06[0]": child_credit_amount(data.get("childDependents", 0)),
        "topmostSubform[0].Page1[0].Step3_ReadOrder[0].f1_07[0]": other_dependent_amount(data.get("otherDependents", 0)),
        # Step 4
        "topmostSubform[0].Page1[0].f1_10[0]": data.get("additionalWithholding", ""),
        # Employer fields (stubs)
        "topmostSubform[0].Page1[0].f1_12[0]": data.get("employerName", "Auntie Anne's"),
        "topmostSubform[0].Page1[0].f1_13[0]": fmt_date(data.get("startDate"), "mmddyyyy"),
        "topmostSubform[0].Page1[0].f1_14[0]": data.get("employerEIN", ""),
    }

    # Filing status: visually confirmed c1_1[0]=single(y625), c1_1[1]=married(y613), c1_1[2]=hoh(y602)
    fs = data.get("filingStatus", "single")
    checkbox_fields = {
        "topmostSubform[0].Page1[0].c1_1[0]": "/1" if fs == "single" else "/Off",
        "topmostSubform[0].Page1[0].c1_1[1]": "/2" if fs == "married" else "/Off",
        "topmostSubform[0].Page1[0].c1_1[2]": "/3" if fs == "hoh" else "/Off",
        # Multiple jobs checkbox
        "topmostSubform[0].Page1[0].c1_2[0]": "/1" if data.get("multipleJobs") == "yes" else "/Off",
        # Exempt checkbox
        "topmostSubform[0].Page1[0].c1_3[0]": "/1" if data.get("exempt") == "yes" else "/Off",
    }

    fill_pdf("/home/claude/w4.pdf", output_path, text_fields, checkbox_fields)


# ─── DE W-4 Fill ────────────────────────────────────────────────────────────

def fill_de_w4(data, output_path):
    print("Filling Delaware W-4...")

    first_initial = data.get("firstName", "")
    if data.get("middleName"):
        first_initial += " " + data["middleName"][0]

    text_fields = {
        "1Firstnameinitial": first_initial,
        "1Lastname": data.get("lastName", ""),
        "2Taxpayerid": fmt_ssn(data.get("ssn", "")),
        "2Homeaddress": data.get("address1", ""),
        "3Cityortown": data.get("city", ""),
        "3State": data.get("state", ""),
        "3Zipcode": data.get("zip", ""),
        "4Totalnumberdependents": str(data.get("deAllowances", "0")),
        "5Additionalamount": data.get("deAdditional", ""),
        # Employer stubs
        "6Employersname": data.get("employerName", "Auntie Anne's"),
        "7Firstdayofemployment": fmt_date(data.get("startDate"), "mmddyyyy"),
        "8Taxpayeridein": data.get("employerEIN", ""),
    }

    # DE filing status radio: /Single or /Married
    de_fs = data.get("deFilingStatus", "single")
    checkbox_fields = {
        "3status": "/Single" if de_fs == "single" else "/Married",
    }

    fill_pdf("/home/claude/de_w4.pdf", output_path, text_fields, checkbox_fields)


# ─── I-9 Fill (Section 1 only) ──────────────────────────────────────────────

def fill_i9(data, output_path):
    print("Filling I-9 (Section 1)...")

    today = fmt_date(data.get("signatureDate") or date.today().isoformat(), "mmddyyyy")

    text_fields = {
        # Section 1 - Employee info
        "Last Name (Family Name)": data.get("lastName", ""),
        "First Name Given Name": data.get("firstName", ""),
        "Employee Middle Initial (if any)": data.get("middleName", "")[:1] if data.get("middleName") else "",
        "Employee Other Last Names Used (if any)": data.get("otherNames", "N/A"),
        "Address Street Number and Name": data.get("address1", ""),
        "Apt Number (if any)": data.get("address2", ""),
        "City or Town": data.get("city", ""),
        "ZIP Code": data.get("zip", ""),
        "Date of Birth mmddyyyy": fmt_date(data.get("dob"), "mmddyyyy"),
        "US Social Security Number": fmt_ssn(data.get("ssn", "")),
        "Telephone Number": data.get("phone", ""),
        "Employees E-mail Address": data.get("email", ""),
        "Today's Date mmddyyy": today,
        # USCIS number if LPR
        "USCIS ANumber": data.get("uscisNumber", ""),
    }

    # Citizenship checkboxes: CB_1=citizen, CB_2=noncitizen national, CB_3=LPR, CB_4=authorized
    cit = data.get("citizenship", "citizen")
    citizenship_map = {
        "citizen": "CB_1",
        "noncitizen": "CB_2",
        "lpr": "CB_3",
        "authorized": "CB_4",
    }
    # These are radio groups - set the selected one to /Yes, others to /Off
    checkbox_fields = {}
    for key, field_id in citizenship_map.items():
        checkbox_fields[field_id] = "/Yes" if cit == key else "/Off"

    fill_pdf("/home/claude/i9.pdf", output_path, text_fields, checkbox_fields)


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        # Use sample data for testing
        data = {
            "firstName": "Jane",
            "middleName": "Marie",
            "lastName": "Smith",
            "ssn": "123-45-6789",
            "dob": "1995-06-15",
            "address1": "42 Pretzel Lane",
            "address2": "Apt 3B",
            "city": "Wilmington",
            "state": "DE",
            "zip": "19801",
            "phone": "(302) 555-0100",
            "email": "jane.smith@email.com",
            "filingStatus": "single",
            "multipleJobs": "no",
            "childDependents": 0,
            "otherDependents": 0,
            "additionalWithholding": "",
            "exempt": "no",
            "deFilingStatus": "single",
            "deAllowances": 1,
            "deAdditional": "",
            "otherNames": "N/A",
            "citizenship": "citizen",
            "startDate": "2025-05-01",
            "signatureDate": date.today().isoformat(),
            "employerName": "Auntie Anne's",
            "employerAddress": "123 Mall Rd, Wilmington DE 19801",
            "employerEIN": "XX-XXXXXXX",
        }
        print("No submission file provided — using sample data.")
    else:
        with open(sys.argv[1]) as f:
            data = json.load(f)

    fill_w4(data, "/home/claude/w4_filled.pdf")
    fill_de_w4(data, "/home/claude/de_w4_filled.pdf")
    fill_i9(data, "/home/claude/i9_filled.pdf")

    print("\n✓ All forms filled. Output files:")
    print("  w4_filled.pdf")
    print("  de_w4_filled.pdf")
    print("  i9_filled.pdf")


if __name__ == "__main__":
    main()
