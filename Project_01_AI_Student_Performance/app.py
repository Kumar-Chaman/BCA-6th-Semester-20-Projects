import os
import csv
import io
import joblib
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from functools import wraps
from config import get_db_connection

app = Flask(__name__)
app.secret_key = "bca_project_01_super_secret_key"

# Trained AI Models load karna
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
score_model = joblib.load(os.path.join(BASE_DIR, "ml_engine", "score_model.pkl"))
risk_model = joblib.load(os.path.join(BASE_DIR, "ml_engine", "risk_model.pkl"))

# Login Required Security Decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Kripya pehle login karein!", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/")
def home():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

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
        else:
            flash("Galat Email ya Password! Dobara koshish karein.", "danger")
    return render_template("login.html")

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "Faculty")
        
        if not full_name or not email or not password:
            flash("Saare fields bharna zaroori hai!", "danger")
            return redirect(url_for("register"))
            
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO users (full_name, email, password, role) VALUES (%s, %s, %s, %s)",
                (full_name, email, password, role)
            )
            conn.commit()
            cursor.close()
            conn.close()
            flash("Account ban gaya! Ab login karein.", "success")
            return redirect(url_for("login"))
        except Exception as e:
            flash("Yeh Email pehle se registered hai!", "danger")
    return render_template("register.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Aap safely logout ho gaye hain.", "info")
    return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    search_q = request.args.get("search", "").strip()
    risk_filter = request.args.get("risk", "").strip()
    
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # Dynamic Search & Filter Query (Parameterized for SQL Injection Safety)
    query = "SELECT * FROM student_predictions WHERE 1=1"
    params = []
    if search_q:
        query += " AND (student_name LIKE %s OR roll_no LIKE %s)"
        params.extend([f"%{search_q}%", f"%{search_q}%"])
    if risk_filter:
        query += " AND risk_status = %s"
        params.append(risk_filter)
    query += " ORDER BY id DESC"
    
    cursor.execute(query, tuple(params))
    students = cursor.fetchall()
    
    # Analytics Summary Cards & Chart Data
    cursor.execute("SELECT * FROM student_predictions")
    all_records = cursor.fetchall()
    cursor.close()
    conn.close()
    
    total_students = len(all_records)
    avg_score = round(sum(r["predicted_final_score"] for r in all_records) / total_students, 2) if total_students else 0
    safe_count = sum(1 for r in all_records if r["risk_status"] == "Safe")
    at_risk_count = sum(1 for r in all_records if r["risk_status"] == "At-Risk")
    critical_count = sum(1 for r in all_records if r["risk_status"] == "Critical")
    
    stats = {
        "total": total_students,
        "avg_score": avg_score,
        "safe": safe_count,
        "at_risk": at_risk_count,
        "critical": critical_count
    }
    return render_template("dashboard.html", students=students, stats=stats, search_q=search_q, risk_filter=risk_filter)

@app.route("/predict", methods=["POST"])
@login_required
def predict():
    try:
        roll_no = request.form.get("roll_no", "").strip()
        student_name = request.form.get("student_name", "").strip()
        attendance = float(request.form.get("attendance"))
        study_hours = float(request.form.get("study_hours"))
        previous_score = float(request.form.get("previous_score"))
        assignment_score = float(request.form.get("assignment_score"))
        sleep_hours = float(request.form.get("sleep_hours"))
        
        # Input Validation
        if not (0 <= attendance <= 100 and 0 <= previous_score <= 100 and 0 <= assignment_score <= 100):
            flash("Percentage values 0 se 100 ke beech honi chahiye!", "danger")
            return redirect(url_for("dashboard"))
            
        # AI Prediction
        features = pd.DataFrame(
            [[attendance, study_hours, previous_score, assignment_score, sleep_hours]],
            columns=["attendance", "study_hours", "previous_score", "assignment_score", "sleep_hours"]
        )
        pred_score = round(float(score_model.predict(features)[0]), 2)
        pred_risk = str(risk_model.predict(features)[0])
        
        # Save to MySQL Database
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO student_predictions 
            (roll_no, student_name, attendance, study_hours, previous_score, assignment_score, sleep_hours, predicted_final_score, risk_status)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', (roll_no, student_name, attendance, study_hours, previous_score, assignment_score, sleep_hours, pred_score, pred_risk))
        conn.commit()
        cursor.close()
        conn.close()
        
        flash(f"AI Prediction Complete: {student_name} -> {pred_score}% ({pred_risk})", "success")
    except ValueError:
        flash("Kripya sahi numeric values enter karein!", "danger")
    except Exception as e:
        flash(f"Error: {str(e)}", "danger")
    return redirect(url_for("dashboard"))

@app.route("/delete/<int:record_id>", methods=["POST"])
@login_required
def delete_record(record_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM student_predictions WHERE id = %s", (record_id,))
    conn.commit()
    cursor.close()
    conn.close()
    flash("Student prediction record delete kar diya gaya.", "info")
    return redirect(url_for("dashboard"))

@app.route("/export_csv")
@login_required
def export_csv():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT roll_no, student_name, attendance, study_hours, previous_score, assignment_score, sleep_hours, predicted_final_score, risk_status, created_at FROM student_predictions ORDER BY id DESC")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Roll No", "Student Name", "Attendance (%)", "Study Hours", "Prev Score (%)", "Assignment (%)", "Sleep Hours", "AI Predicted Score (%)", "Risk Status", "Timestamp"])
    for r in rows:
        writer.writerow([r["roll_no"], r["student_name"], r["attendance"], r["study_hours"], r["previous_score"], r["assignment_score"], r["sleep_hours"], r["predicted_final_score"], r["risk_status"], r["created_at"]])
        
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=AI_Student_Performance_Report.csv"}
    )

if __name__ == "__main__":
    app.run(debug=True, port=5001)
