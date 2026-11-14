import os
import csv
import io
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from werkzeug.utils import secure_filename
from config import get_db_connection
from ml_engine.career_engine import extract_text_from_pdf, analyze_resume

app = Flask(__name__)
app.secret_key = "bca_project_02_resume_secret_key"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            flash("Kripya pehle login karein!", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.route("/")
def home():
    return redirect(url_for("dashboard" if "user_id" in session else "login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "").strip()
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM users WHERE email = %s AND password = %s", (email, password))
        user = cursor.fetchone()
        cursor.close()
        conn.close()
        if user:
            session["user_id"] = user["id"]
            session["user_name"] = user["full_name"]
            session["user_role"] = user["role"]
            flash(f"Swagat hai, {user['full_name']}!", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid Email ya Password!", "danger")
    return render_template("login.html")

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "Candidate")
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO users (full_name, email, password, role) VALUES (%s, %s, %s, %s)",
                           (full_name, email, password, role))
            conn.commit()
            cursor.close()
            conn.close()
            flash("Registration successful! Ab login karein.", "success")
            return redirect(url_for("login"))
        except Exception:
            flash("Yeh Email pehle se registered hai!", "danger")
    return render_template("register.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    search_q = request.args.get("search", "").strip()
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    if search_q:
        cursor.execute("SELECT * FROM resume_analyses WHERE candidate_name LIKE %s OR recommended_role LIKE %s ORDER BY id DESC",
                       (f"%{search_q}%", f"%{search_q}%"))
    else:
        cursor.execute("SELECT * FROM resume_analyses ORDER BY id DESC")
    records = cursor.fetchall()

    cursor.execute("SELECT * FROM resume_analyses")
    all_rows = cursor.fetchall()
    cursor.close()
    conn.close()

    total = len(all_rows)
    avg_ats = round(sum(r["ats_score"] for r in all_rows) / total, 2) if total else 0
    role_counts = {}
    for r in all_rows:
        role_counts[r["recommended_role"]] = role_counts.get(r["recommended_role"], 0) + 1

    top_role = max(role_counts, key=role_counts.get) if role_counts else "N/A"
    latest_result = session.pop("latest_result", None)
    return render_template("dashboard.html", records=records, total=total, avg_ats=avg_ats,
                           top_role=top_role, role_counts=role_counts, search_q=search_q,
                           latest_result=latest_result)

@app.route("/analyze", methods=["POST"])
@login_required
def analyze():
    candidate_name = request.form.get("candidate_name", "").strip()
    candidate_email = request.form.get("candidate_email", "").strip()
    resume_text = request.form.get("resume_text", "").strip()
    pdf_file = request.files.get("resume_pdf")

    source_type = "Text Paste"
    if pdf_file and pdf_file.filename:
        if not pdf_file.filename.lower().endswith(".pdf"):
            flash("Sirf .PDF file hi upload karein!", "danger")
            return redirect(url_for("dashboard"))
        filename = secure_filename(pdf_file.filename)
        saved_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        pdf_file.save(saved_path)
        resume_text = extract_text_from_pdf(saved_path)
        source_type = f"PDF ({filename})"

    if len(resume_text) < 20:
        flash("Resume content bahut chhota hai! Kripya valid PDF upload karein ya skills text paste karein.", "danger")
        return redirect(url_for("dashboard"))

    ranked_roles = analyze_resume(resume_text)
    best = ranked_roles[0]

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO resume_analyses (candidate_name, candidate_email, source_type, recommended_role, ats_score, matched_skills, missing_skills)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
    ''', (
        candidate_name, candidate_email, source_type, best["role"], best["ats_score"],
        ", ".join(best["matched_skills"]) if best["matched_skills"] else "None",
        ", ".join(best["missing_skills"]) if best["missing_skills"] else "None"
    ))
    conn.commit()
    cursor.close()
    conn.close()

    session["latest_result"] = {
        "candidate_name": candidate_name,
        "top_3": ranked_roles[:3]
    }
    flash(f"AI Analysis Complete! Best Match: {best['role']} ({best['ats_score']}% ATS Score)", "success")
    return redirect(url_for("dashboard"))

@app.route("/delete/<int:rec_id>", methods=["POST"])
@login_required
def delete_record(rec_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM resume_analyses WHERE id = %s", (rec_id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Record delete kar diya gaya.", "info")
    return redirect(url_for("dashboard"))

@app.route("/export_csv")
@login_required
def export_csv():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM resume_analyses ORDER BY id DESC")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(["ID", "Candidate Name", "Email", "Source", "Recommended Role", "ATS Score (%)", "Matched Skills", "Missing Skills", "Date"])
    for r in rows:
        w.writerow([r["id"], r["candidate_name"], r["candidate_email"], r["source_type"], r["recommended_role"], r["ats_score"], r["matched_skills"], r["missing_skills"], r["created_at"]])
    return Response(out.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment;filename=ResumeMatch_AI_Report.csv"})

if __name__ == "__main__":
    app.run(debug=True, port=5002)
