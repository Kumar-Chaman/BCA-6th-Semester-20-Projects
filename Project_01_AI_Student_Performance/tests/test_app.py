import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from app import app
from config import get_db_connection

print("=== PROJECT 01: AUTOMATED TEST SUITE EXECUTION ===")
client = app.test_client()
passed = 0
total = 6

# Test 1: Database Connectivity
try:
    conn = get_db_connection()
    assert conn.is_connected()
    conn.close()
    print("[TEST 1/6] MySQL Database Connectivity         -> PASS")
    passed += 1
except Exception as e:
    print(f"[TEST 1/6] MySQL Database Connectivity         -> FAIL ({e})")

# Test 2: Unauthorized Access Protection
res = client.get("/dashboard")
if res.status_code == 302:
    print("[TEST 2/6] Unauthorized Dashboard Redirect     -> PASS (HTTP 302)")
    passed += 1
else:
    print(f"[TEST 2/6] Unauthorized Dashboard Redirect     -> FAIL ({res.status_code})")

# Test 3: Admin Login Authentication
res = client.post("/login", data={"email": "admin@college.edu", "password": "admin123"}, follow_redirects=True)
if b"EduPredict AI" in res.data and b"HOD Admin" in res.data:
    print("[TEST 3/6] Admin Login Authentication          -> PASS")
    passed += 1
else:
    print("[TEST 3/6] Admin Login Authentication          -> FAIL")

# Test 4: AI Prediction & MySQL Insert (3 Sample Students)
sample_students = [
    {"roll_no": "BCA-2301", "student_name": "Aarav Sharma", "attendance": "92.0", "study_hours": "6.0", "previous_score": "88.0", "assignment_score": "90.0", "sleep_hours": "7.5"},
    {"roll_no": "BCA-2302", "student_name": "Priya Kumari", "attendance": "71.0", "study_hours": "3.0", "previous_score": "62.0", "assignment_score": "65.0", "sleep_hours": "6.0"},
    {"roll_no": "BCA-2303", "student_name": "Rohan Das", "attendance": "54.0", "study_hours": "1.5", "previous_score": "42.0", "assignment_score": "48.0", "sleep_hours": "5.0"}
]

all_inserted = True
for st in sample_students:
    r = client.post("/predict", data=st, follow_redirects=True)
    if b"AI Prediction Complete" not in r.data:
        all_inserted = False

if all_inserted:
    print("[TEST 4/6] AI Model Prediction & MySQL Save    -> PASS (3 Sample Records Saved)")
    passed += 1
else:
    print("[TEST 4/6] AI Model Prediction & MySQL Save    -> FAIL")

# Test 5: Input Validation (Invalid Attendance > 100)
res = client.post("/predict", data={
    "roll_no": "BCA-9999", "student_name": "Invalid User", "attendance": "150",
    "study_hours": "5", "previous_score": "80", "assignment_score": "80", "sleep_hours": "7"
}, follow_redirects=True)
if b"0 se 100 ke beech" in res.data:
    print("[TEST 5/6] Input Validation Guard (Att > 100)  -> PASS")
    passed += 1
else:
    print("[TEST 5/6] Input Validation Guard (Att > 100)  -> FAIL")

# Test 6: CSV Report Export
res = client.get("/export_csv")
if res.status_code == 200 and b"Aarav Sharma" in res.data:
    print("[TEST 6/6] CSV Report Generation & Export      -> PASS")
    passed += 1
else:
    print("[TEST 6/6] CSV Report Generation & Export      -> FAIL")

print(f"\nSUMMARY: {passed}/{total} Test Cases Passed Successfully!")
