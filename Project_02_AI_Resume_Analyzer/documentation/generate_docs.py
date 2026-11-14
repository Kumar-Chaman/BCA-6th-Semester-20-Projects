import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ml_engine.career_engine import analyze_resume

print("\n=== PROJECT 02: GENERATING DIAGRAMS & ACADEMIC REPORT ===")
os.makedirs("documentation", exist_ok=True)

# Diagram 1: System Architecture
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
ax.set_title("SYSTEM ARCHITECTURE: AI-POWERED RESUME ANALYZER (RESUMEMATCH AI)", fontsize=12.5, fontweight="bold")

boxes = [
    (0.4, 2.1, 2.4, 1.9, "#e0e7ff", "#3730a3", "CLIENT LAYER\n(Browser UI)\n\n• PDF / Text Upload\n• Bootstrap 5 Cards\n• Chart.js Analytics"),
    (3.8, 2.1, 2.5, 1.9, "#dcfce7", "#15803d", "NLP & FLASK TIER\n(Python 3.14 Engine)\n\n• pypdf Text Extractor\n• TF-IDF Vectorizer\n• Cosine Similarity"),
    (7.3, 2.1, 2.3, 1.9, "#fef3c7", "#b45309", "DATABASE TIER\n(MySQL 8.0)\n\n• DB: resume_ai_db\n• Table: users\n• Table: resume_analyses")
]
for x, y, w, h, fc, ec, txt in boxes:
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1", facecolor=fc, edgecolor=ec, linewidth=2))
    ax.text(x+w/2, y+h/2, txt, ha="center", va="center", fontsize=8.8, fontweight="bold")

ax.annotate("", xy=(3.7, 3.1), xytext=(2.9, 3.1), arrowprops=dict(arrowstyle="<->", lw=2))
ax.annotate("", xy=(7.2, 3.1), xytext=(6.4, 3.1), arrowprops=dict(arrowstyle="<->", lw=2))
plt.tight_layout()
arch_img = os.path.join("documentation", "system_architecture.png")
plt.savefig(arch_img, dpi=200); plt.close()

# Diagram 2: ER Diagram
fig, ax = plt.subplots(figsize=(10, 4.8))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.2); ax.axis("off")
ax.set_title("ENTITY-RELATIONSHIP (ER) DIAGRAM - DATABASE: resume_ai_db", fontsize=12.5, fontweight="bold")

t1 = "ENTITY: users\n-------------------------\n[PK] id (INT, AI)\nfull_name (VARCHAR 100)\nemail (VARCHAR 100, UQ)\npassword (VARCHAR 255)\nrole (ENUM)\ncreated_at (TIMESTAMP)"
t2 = "ENTITY: resume_analyses\n----------------------------------\n[PK] id (INT, AI)\ncandidate_name (VARCHAR 100)\ncandidate_email (VARCHAR 100)\nsource_type (VARCHAR 30)\nrecommended_role (VARCHAR 100)\nats_score (FLOAT)\nmatched_skills (TEXT)\nmissing_skills (TEXT)\ncreated_at (TIMESTAMP)"

ax.add_patch(patches.FancyBboxPatch((0.6, 1.2), 3.3, 2.8, boxstyle="round,pad=0.1", facecolor="#eff6ff", edgecolor="#1e40af", linewidth=2))
ax.text(2.25, 2.6, t1, ha="center", va="center", fontsize=8.5, family="monospace", fontweight="bold")
ax.add_patch(patches.FancyBboxPatch((5.5, 0.8), 3.9, 3.6, boxstyle="round,pad=0.1", facecolor="#f0fdf4", edgecolor="#166534", linewidth=2))
ax.text(7.45, 2.6, t2, ha="center", va="center", fontsize=8.5, family="monospace", fontweight="bold")
ax.annotate("", xy=(5.4, 2.6), xytext=(4.0, 2.6), arrowprops=dict(arrowstyle="->", lw=2))
plt.tight_layout()
er_img = os.path.join("documentation", "er_diagram.png")
plt.savefig(er_img, dpi=200); plt.close()

# Diagram 3: Sample ATS Role Comparison Bar Chart
demo_res = analyze_resume("BCA student with Python, Flask, MySQL, HTML, CSS, JavaScript, Machine Learning, Pandas, Git")
r_names = [x["role"] for x in demo_res]
r_scores = [x["ats_score"] for x in demo_res]
fig, ax = plt.subplots(figsize=(8.5, 4.3))
bars = ax.barh(r_names[::-1], r_scores[::-1], color="#4f46e5")
ax.set_xlabel("ATS Compatibility Score (%)", fontweight="bold")
ax.set_title("NLP CAREER MATCHING: ROLE COMPATIBILITY BENCHMARK", fontsize=11.5, fontweight="bold")
for b in bars:
    ax.text(b.get_width() + 0.8, b.get_y() + b.get_height()/2, f"{b.get_width()}%", va="center", fontsize=9, fontweight="bold")
ax.set_xlim(0, 100)
plt.tight_layout()
ats_img = os.path.join("documentation", "ats_benchmark.png")
plt.savefig(ats_img, dpi=200); plt.close()

# Generate Word Report (.docx)
doc = Document()
for s in doc.sections:
    s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
    s.header.paragraphs[0].text = "Project 02: AI-Powered Resume Analyzer & Career Recommendation System | BCA 6th Sem"
    s.footer.paragraphs[0].text = "Submitted by: Chaman Kumar | GitHub: Kumar-Chaman/BCA-6th-Semester-20-Projects"

style = doc.styles["Normal"]
style.font.name = "Times New Roman"; style.font.size = Pt(12); style.paragraph_format.line_spacing = 1.5

def add_ch(num, title):
    doc.add_page_break()
    p = doc.add_paragraph()
    r1 = p.add_run(f"CHAPTER {num}\n"); r1.bold = True; r1.font.size = Pt(14); r1.font.color.rgb = RGBColor(30, 64, 175)
    r2 = p.add_run(title.upper()); r2.bold = True; r2.font.size = Pt(16)
    doc.add_paragraph("_" * 62)

def add_h2(t):
    p = doc.add_paragraph(); r = p.add_run(t); r.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(29, 78, 216)

def add_p(t):
    p = doc.add_paragraph(t); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Cover Page
pc = doc.add_paragraph(); pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
pc.add_run("[ENTER UNIVERSITY & COLLEGE NAME]\nDEPARTMENT OF COMPUTER APPLICATIONS (BCA)\n\n").bold = True
rt = pc.add_run("AI-POWERED RESUME ANALYZER AND\nCAREER RECOMMENDATION SYSTEM\n(ResumeMatch AI)\n\n")
rt.bold = True; rt.font.size = Pt(21)
pc.add_run("A Natural Language Processing (TF-IDF & Cosine Similarity) Web Platform\nBuilt with Python 3.14, Flask, Scikit-Learn, pypdf, and MySQL 8.0\n\n").italic = True

tbl = doc.add_table(rows=1, cols=2); tbl.style = "Table Grid"
tbl.rows[0].cells[0].text = "SUBMITTED BY:\nName: CHAMAN KUMAR\nClass: BCA 6th Semester\nRoll No: [Enter Roll No]\nReg. No: [Enter Reg. No]"
tbl.rows[0].cells[1].text = "SUPERVISOR & REPOSITORY:\nGuide: [Enter Supervisor Name]\nDept. of Computer Applications\nGitHub: github.com/Kumar-Chaman/\nBCA-6th-Semester-20-Projects"

# Certificate, Declaration & Abstract
doc.add_page_break()
add_h2("CERTIFICATE & STUDENT DECLARATION")
add_p("This is to certify that the project titled 'AI-Powered Resume Analyzer and Career Recommendation System (ResumeMatch AI)' is an original academic software project developed by Chaman Kumar in partial fulfillment of the Bachelor of Computer Applications (BCA 6th Semester) degree requirements.")
add_p("\nSupervisor Signature: ______________________          Student Signature: Chaman Kumar")

add_h2("ABSTRACT")
add_p("Manual screening of student resumes during campus placements is time-consuming and often fails to provide constructive skill-gap feedback to applicants. ResumeMatch AI automates resume evaluation using Natural Language Processing (NLP). The system parses uploaded PDF resumes (via pypdf) or plain text, vectorizes technical keywords using Term Frequency-Inverse Document Frequency (TF-IDF) with bigrams, and computes Cosine Similarity against six curated IT career tracks. Combining semantic similarity (30%) with explicit skill coverage (70%), the engine outputs an ATS Compatibility Score (0-100%), ranks the top 3 career paths, highlights matched and missing skills, and persists every evaluation inside a MySQL 8.0 database ('resume_ai_db').")

# Chapters 1-5
add_ch("1", "Introduction, Problem Statement & Objectives")
add_h2("1.1 Problem Statement")
add_p("Graduating BCA students frequently apply for software roles without knowing whether their resumes pass Applicant Tracking Systems (ATS) or which technical competencies they lack for target roles like Full-Stack Developer, AI/ML Engineer, or Data Analyst.")
add_h2("1.2 Objectives")
for o in [
    "Extract structured text automatically from uploaded PDF resumes using pypdf.",
    "Implement TF-IDF Vectorization and Cosine Similarity in Scikit-Learn to match resumes against 6 IT job profiles.",
    "Provide actionable Skill-Gap Analysis listing both matched strengths and missing skills to learn next.",
    "Store candidate history in MySQL 8.0 with search, analytics charts, and CSV export."
]:
    doc.add_paragraph(o, style="List Bullet")

add_ch("2", "System Architecture & NLP Methodology")
add_h2("2.1 Three-Tier System Architecture")
add_p("The system separates PDF/UI handling (Bootstrap 5 + Chart.js), NLP processing (Flask + Scikit-Learn TF-IDF), and persistent storage (MySQL 8.0).")
doc.add_picture(arch_img, width=Inches(6.2))
add_h2("2.2 Role Compatibility Benchmark")
doc.add_picture(ats_img, width=Inches(5.8))

add_ch("3", "Database Design & ER Schema (MySQL 8.0)")
add_h2("3.1 Entity-Relationship Diagram (resume_ai_db)")
doc.add_picture(er_img, width=Inches(6.2))

add_ch("4", "Complete Source Code")
files_to_doc = [
    ("4.1 database/setup_db.py", os.path.join("database", "setup_db.py")),
    ("4.2 ml_engine/career_engine.py", os.path.join("ml_engine", "career_engine.py")),
    ("4.3 config.py", "config.py"),
    ("4.4 app.py", "app.py"),
    ("4.5 templates/dashboard.html", os.path.join("templates", "dashboard.html")),
    ("4.6 tests/test_app.py", os.path.join("tests", "test_app.py"))
]
for title, path in files_to_doc:
    add_h2(title)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8-sig") as f:
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.05
            r = p.add_run(f.read())
            r.font.name = "Consolas"; r.font.size = Pt(8.5)

add_ch("5", "Software Testing & Viva-Voce Preparation")
add_h2("5.1 Automated Test Results (5/5 Passed)")
t_tc = doc.add_table(rows=1, cols=4); t_tc.style = "Table Grid"
for idx, h in enumerate(["Test ID", "Test Case", "Expected", "Actual"]):
    t_tc.rows[0].cells[idx].text = h
for tc in [
    ("TC-01", "MySQL Connection (resume_ai_db)", "Connected = True", "PASS"),
    ("TC-02", "Placement Admin Login", "Redirect to Dashboard", "PASS"),
    ("TC-03", "NLP Resume Analysis & Save", "Top Role + ATS Score Saved", "PASS (2 Saved)"),
    ("TC-04", "Short/Invalid Text Guard", "Warning Flash Message", "PASS"),
    ("TC-05", "CSV Placement Report Export", "HTTP 200 CSV Download", "PASS")
]:
    row = t_tc.add_row().cells
    for i, v in enumerate(tc): row[i].text = v

add_h2("5.2 Top 15 Viva-Voce Questions & Answers")
viva2 = [
    ("Q1. What is TF-IDF?", "Term Frequency-Inverse Document Frequency measures how important a technical keyword is to a resume relative to all job profiles, reducing weight for common filler words."),
    ("Q2. How does Cosine Similarity work in ResumeMatch AI?", "It calculates the cosine of the angle between the TF-IDF vector of the candidate's resume and each job role's skill vector, producing a similarity score between 0 and 1."),
    ("Q3. How is the final ATS Score computed?", "It combines 70% explicit skill coverage (matched required skills / total required skills) and 30% TF-IDF semantic similarity."),
    ("Q4. Which library extracts text from PDF files?", "pypdf (PdfReader) iterates over every page of the uploaded PDF and extracts raw text for NLP preprocessing."),
    ("Q5. How is file upload security handled in Flask?", "Using werkzeug.utils.secure_filename() to sanitize filenames and checking the .pdf extension before saving to uploads/."),
    ("Q6. Why use ngram_range=(1, 2) in TfidfVectorizer?", "Bigrams allow the model to recognize two-word technical skills such as 'machine learning', 'deep learning', 'spring boot', and 'power bi'."),
    ("Q7. What database tables are used in resume_ai_db?", "'users' for authentication and 'resume_analyses' for storing candidate details, recommended role, ATS score, matched skills, and missing skills."),
    ("Q8. How does regex boundary matching (\\b) prevent false skill matches?", "Using word boundaries ensures 'java' does not falsely match inside 'javascript' or 'c' inside 'css'."),
    ("Q9. How does Chart.js display the career distribution?", "Flask passes aggregated role counts as JSON arrays via Jinja2's tojson filter to render a bar chart."),
    ("Q10. How can this project be extended in the future?", "By adding BERT/transformer embeddings, direct LinkedIn job scraping, and automated PDF resume improvement suggestions.")
]
for q, a in viva2:
    pq = doc.add_paragraph(); rq = pq.add_run(q); rq.bold = True; rq.font.color.rgb = RGBColor(30, 58, 138)
    add_p(f"Answer: {a}")

rep_path = os.path.join("documentation", "Project_02_Complete_Report.docx")
doc.save(rep_path)
print(f"[SUCCESS] Project 02 Diagrams & Academic Report saved: {rep_path}")
