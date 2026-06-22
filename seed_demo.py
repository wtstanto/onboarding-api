#!/usr/bin/env python3
"""Seed 5 demo employees into the onboarding demo sheet/folder."""
import urllib.request, urllib.error, json, time, sys

API      = "https://onboarding-api-production-d42e.up.railway.app"
DEMO_KEY = "ak_demo_onboarding_middleman_2026"

EMPLOYEES = [
    {
        "mode": "demo", "firstName": "Kayla", "lastName": "Simmons",
        "dob": "2009-02-14", "ssn": "111-22-3344",
        "address1": "428 Valley Road", "city": "Wilmington", "state": "DE", "zip": "19801",
        "phone": "3025550101", "email": "kayla.simmons@example.com",
        "gender": "Female", "tshirtSize": "S",
        "ecName": "Linda Simmons", "ecRelationship": "Parent", "ecPhone": "3025550100",
        "filingStatus": "single", "multipleJobs": "no", "childCredits": "0",
        "otherDependents": "0", "additionalWithholding": "0", "exempt": "no",
        "deFilingStatus": "single", "deAllowances": "0",
        "citizenship": "citizen", "docType": "listA",
        "listAType": "U.S. Passport", "listANumber": "B11223344", "listAExpiration": "2030-01-01",
        "accountType": "checking", "bankName": "PNC Bank", "routingNumber": "031000053", "accountNumber": "445566778",
        "drugPolicyAck": "on", "conductAck": "on", "handbookAck": "on", "foodHandlerAck": "on",
        "typedName": "Kayla Simmons",
        "_post": {"status": "onboarding"},
    },
    {
        "mode": "demo", "firstName": "Tyler", "lastName": "Nguyen",
        "dob": "2010-01-08", "ssn": "222-33-4455",
        "address1": "815 Oak Street", "city": "Newark", "state": "DE", "zip": "19702",
        "phone": "3025550202", "email": "tyler.nguyen@example.com",
        "gender": "Male", "tshirtSize": "M",
        "ecName": "David Nguyen", "ecRelationship": "Parent", "ecPhone": "3025550201",
        "filingStatus": "single", "multipleJobs": "no", "childCredits": "0",
        "otherDependents": "0", "additionalWithholding": "0", "exempt": "no",
        "deFilingStatus": "single", "deAllowances": "0",
        "citizenship": "citizen", "docType": "listA",
        "listAType": "U.S. Passport", "listANumber": "C22334455", "listAExpiration": "2030-06-01",
        "accountType": "checking", "bankName": "TD Bank", "routingNumber": "011103093", "accountNumber": "556677889",
        "drugPolicyAck": "on", "conductAck": "on", "handbookAck": "on", "foodHandlerAck": "on",
        "typedName": "Tyler Nguyen",
        "_post": {
            "employment": {"payRate": "$12.00/hr", "position": "Team Member", "location": "Christiana Mall", "deptCode": "TM01", "hireDate": "2026-04-10"},
            "status": "active", "workingPapers": "given",
        },
    },
    {
        "mode": "demo", "firstName": "Destiny", "lastName": "Thompson",
        "dob": "2003-07-19", "ssn": "333-44-5566",
        "address1": "220 Maple Ave", "city": "Wilmington", "state": "DE", "zip": "19805",
        "phone": "3025550303", "email": "destiny.thompson@example.com",
        "gender": "Female", "tshirtSize": "S",
        "ecName": "Robert Thompson", "ecRelationship": "Parent", "ecPhone": "3025550300",
        "filingStatus": "single", "multipleJobs": "no", "childCredits": "0",
        "otherDependents": "0", "additionalWithholding": "0", "exempt": "no",
        "deFilingStatus": "single", "deAllowances": "1",
        "citizenship": "citizen", "docType": "listA",
        "listAType": "U.S. Passport", "listANumber": "D33445566", "listAExpiration": "2028-09-01",
        "accountType": "checking", "bankName": "Bank of America", "routingNumber": "026009593", "accountNumber": "667788990",
        "drugPolicyAck": "on", "conductAck": "on", "handbookAck": "on", "foodHandlerAck": "on",
        "typedName": "Destiny Thompson",
        "_post": {
            "employment": {"payRate": "$13.50/hr", "position": "Shift Lead", "location": "Christiana Mall", "deptCode": "SL01", "hireDate": "2026-04-01"},
            "status": "active",
        },
    },
    {
        "mode": "demo", "firstName": "Marcus", "lastName": "Reid",
        "dob": "2006-11-30", "ssn": "444-55-6677",
        "address1": "1102 Pine Street", "city": "Middletown", "state": "DE", "zip": "19709",
        "phone": "3025550404", "email": "marcus.reid@example.com",
        "gender": "Male", "tshirtSize": "XL",
        "ecName": "Patricia Reid", "ecRelationship": "Parent", "ecPhone": "3025550400",
        "filingStatus": "single", "multipleJobs": "no", "childCredits": "0",
        "otherDependents": "0", "additionalWithholding": "0", "exempt": "no",
        "deFilingStatus": "single", "deAllowances": "1",
        "citizenship": "citizen", "docType": "listA",
        "listAType": "U.S. Passport", "listANumber": "E44556677", "listAExpiration": "2029-03-01",
        "accountType": "checking", "bankName": "Chase", "routingNumber": "021000021", "accountNumber": "778899001",
        "drugPolicyAck": "on", "conductAck": "on", "handbookAck": "on", "foodHandlerAck": "on",
        "typedName": "Marcus Reid",
        "_post": {
            "employment": {"payRate": "$12.50/hr", "position": "Team Member", "location": "Christiana Mall", "deptCode": "TM02", "hireDate": "2026-04-20"},
            "status": "onboarding",
        },
    },
    {
        "mode": "demo", "firstName": "Ashley", "lastName": "Kim",
        "dob": "2001-04-03", "ssn": "555-66-7788",
        "address1": "333 Birch Lane", "city": "Newark", "state": "DE", "zip": "19711",
        "phone": "3025550505", "email": "ashley.kim@example.com",
        "gender": "Female", "tshirtSize": "M",
        "ecName": "James Kim", "ecRelationship": "Spouse", "ecPhone": "3025550500",
        "filingStatus": "married", "multipleJobs": "no", "childCredits": "0",
        "otherDependents": "0", "additionalWithholding": "0", "exempt": "no",
        "deFilingStatus": "married", "deAllowances": "2",
        "citizenship": "citizen", "docType": "listA",
        "listAType": "U.S. Passport", "listANumber": "F55667788", "listAExpiration": "2031-12-01",
        "accountType": "checking", "bankName": "Wells Fargo", "routingNumber": "121000248", "accountNumber": "889900112",
        "drugPolicyAck": "on", "conductAck": "on", "handbookAck": "on", "foodHandlerAck": "on",
        "typedName": "Ashley Kim",
        "_post": {
            "employment": {"payRate": "$15.00/hr", "position": "Assistant Manager", "location": "Christiana Mall", "deptCode": "AM01", "hireDate": "2026-03-15"},
            "status": "active",
        },
    },
]


def api_call(method, path, body=None, auth=False):
    url = API + path
    data = json.dumps(body).encode() if body else None
    headers = {"Content-Type": "application/json"}
    if auth:
        headers["X-API-Key"] = DEMO_KEY
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
            try:
                return resp.status, json.loads(raw)
            except Exception:
                return resp.status, raw
    except urllib.error.HTTPError as e:
        return e.code, {}
    except Exception as ex:
        return 0, str(ex)


print("\n── Step 1: Submitting employees ────────────────────────────────────")
for emp in EMPLOYEES:
    payload = {k: v for k, v in emp.items() if not k.startswith("_")}
    print(f"  {emp['firstName']} {emp['lastName']} ({emp['dob']})...", end=" ", flush=True)
    status, _ = api_call("POST", "/fill", payload)
    print("✓" if status == 200 else f"✗ HTTP {status}")
    time.sleep(2)

print("\n── Step 2: Waiting 90s for GAS background writes ──────────────────")
for i in range(9):
    print(f"  {(i+1)*10}s...", end=" ", flush=True)
    time.sleep(10)
print()

print("\n── Step 3: Fetching row IDs ────────────────────────────────────────")
status, subs = api_call("GET", "/submissions", auth=True)
if status != 200 or not isinstance(subs, list):
    print(f"  ✗ Could not fetch submissions (HTTP {status}). Try again in 60s.")
    sys.exit(1)
print(f"  Found {len(subs)} submissions in demo sheet")

name_map = {f"{e['firstName']} {e['lastName']}": e for e in EMPLOYEES}
matched = {}
for s in subs:
    full = f"{s.get('firstName','')} {s.get('lastName','')}".strip()
    if full in name_map:
        matched[full] = s
        print(f"  ✓ {full} → row {s['id']}")

if not matched:
    print("  ⚠ No matches — GAS may still be writing. Re-run in 60s.")
    sys.exit(1)

print("\n── Step 4: Employment / status / working-papers ───────────────────")
for name, emp_def in name_map.items():
    post = emp_def.get("_post", {})
    if name not in matched:
        print(f"  ⚠ {name}: not in sheet yet, skipping")
        continue
    row_id = matched[name]["id"]

    if post.get("employment"):
        s, _ = api_call("PATCH", f"/submissions/{row_id}/employment", post["employment"], auth=True)
        print(f"  {name} employment: {'✓' if s == 200 else f'✗ {s}'}")
        time.sleep(1)

    if post.get("status") and post["status"] != "onboarding":
        s, _ = api_call("PATCH", f"/submissions/{row_id}/status", {"status": post["status"]}, auth=True)
        print(f"  {name} → {post['status']}: {'✓' if s == 200 else f'✗ {s}'}")
        time.sleep(1)

    if post.get("workingPapers") == "given":
        s, _ = api_call("PATCH", f"/submissions/{row_id}/working-papers", {"field": "given", "value": True}, auth=True)
        print(f"  {name} working papers given: {'✓' if s == 200 else f'✗ {s}'}")
        time.sleep(1)

print("\n── Done! ───────────────────────────────────────────────────────────")
print("  https://previews.gomiddleman.com/onboarding/admin  (password: twist)\n")
