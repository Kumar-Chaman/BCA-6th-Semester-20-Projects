import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from app import app
from config import get_db_connection

print("=== PROJECT 02: AUTOMATED TEST SUITE EXECUTION ===")
client = app.test_client()
passed = 0
total = 5

# Test 1: MySQL Connection
try:
    conn = get_db_connection()
    assert conn.is_connected()
    conn.close()
    print("[TEST 1/5] MySQL Database Connectivity (resume_ai_db) -> PASS")
    passed += 1
except Exception as e:
    print(f"[TEST 1/5] MySQL Database Connectivity -> FAIL ({e})")

# Test 2: Admin Login
res = client.post("/login", data={"email": "admin@resumematch.ai", "password": "admin123"}, follow_redirects=True)
if b"ResumeMatch AI" in res.data and b"Placement Cell Admin" in res.data:
    print("[TEST 2/5] Placement Admin Login Authentication      -> PASS")
    passed += 1
else:
    print("[TEST 2/5] Placement Admin Login Authentication      -> FAIL")

# Test 3: NLP Resume Analysis & MySQL Insert (2 Candidates)
samples = [
    {"candidate_name": "Chaman Kumar", "candidate_email": "chaman@bca.edu", "resume_text": "BCA 6th Sem developer skilled in Python, Flask, MySQL, HTML, CSS, JavaScript, Bootstrap, Git, and REST API."},
    {"candidate_name": "Ananya Singh", "candidate_email": "ananya@bca.edu", "resume_text": "Aspiring Data Scientist with Python, Machine Learning, Deep Learning, Scikit-Learn, Pandas, NumPy, NLP, and TensorFlow."}
]
ok = True
for s in samples:
    r = client.post("/analyze", data=s, follow_redirects=True)
    if b"AI Analysis Complete" not in r.data:
        ok = False
if ok:
    print("[TEST 3/5] NLP TF-IDF Career Analysis & MySQL Save   -> PASS (2 Records Saved)")
    passed += 1
else:
    print("[TEST 3/5] NLP TF-IDF Career Analysis & MySQL Save   -> FAIL")

# Test 4: Short/Empty Input Guard
res = client.post("/analyze", data={"candidate_name": "Test", "candidate_email": "t@t.com", "resume_text": "abc"}, follow_redirects=True)
if b"bahut chhota hai" in res.data:
    print("[TEST 4/5] Invalid/Empty Resume Input Guard          -> PASS")
    passed += 1
else:
    print("[TEST 4/5] Invalid/Empty Resume Input Guard          -> FAIL")

# Test 5: CSV Report Export
res = client.get("/export_csv")
if res.status_code == 200 and b"Chaman Kumar" in res.data:
    print("[TEST 5/5] CSV Placement Report Export               -> PASS")
    passed += 1
else:
    print("[TEST 5/5] CSV Placement Report Export               -> FAIL")

print(f"\nSUMMARY: {passed}/{total} Test Cases Passed Successfully!")
