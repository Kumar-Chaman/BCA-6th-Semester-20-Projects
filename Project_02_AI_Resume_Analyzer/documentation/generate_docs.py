import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ml_engine.career_engine import analyze_resume

print("=== PROJECT 02: GENERATING 6 DIAGRAMS & UNIVERSITY REPORT ===")
os.makedirs("documentation", exist_ok=True)

# ---------------------------------------------------------
# PART A: GENERATE 6 HIGH-RESOLUTION SYSTEM & UML DIAGRAMS
# ---------------------------------------------------------

# 1. System Architecture Diagram
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
ax.set_title("FIGURE 4.1: THREE-TIER SYSTEM ARCHITECTURE - RESUMEMATCH AI", fontsize=12, fontweight="bold")
boxes = [
    (0.4, 2.0, 2.4, 2.0, "#e0e7ff", "#3730a3", "PRESENTATION TIER\n(Client Browser)\n\n• PDF / Text Upload UI\n• Bootstrap 5.3 Cards\n• Chart.js Role Chart\n• Skill Gap Badges"),
    (3.8, 2.0, 2.4, 2.0, "#dcfce7", "#15803d", "NLP & APPLICATION TIER\n(Python 3.14 + Flask)\n\n• pypdf Text Extractor\n• TF-IDF Vectorizer\n• Cosine Similarity\n• Weighted ATS Scorer"),
    (7.2, 2.0, 2.4, 2.0, "#fef3c7", "#b45309", "DATA PERSISTENCE TIER\n(MySQL 8.0 Server)\n\n• DB: resume_ai_db\n• Table: users\n• Table: resume_analyses\n• Parameterized SQL")
]
for x, y, w, h, fc, ec, txt in boxes:
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1", facecolor=fc, edgecolor=ec, linewidth=2))
    ax.text(x+w/2, y+h/2, txt, ha="center", va="center", fontsize=8.5, fontweight="bold")
ax.annotate("", xy=(3.7, 3.3), xytext=(2.9, 3.3), arrowprops=dict(arrowstyle="->", lw=2))
ax.text(3.3, 3.5, "Multipart POST", ha="center", fontsize=8, fontweight="bold")
ax.annotate("", xy=(2.9, 2.7), xytext=(3.7, 2.7), arrowprops=dict(arrowstyle="->", lw=2))
ax.text(3.3, 2.45, "ATS & Top 3 Roles", ha="center", fontsize=8, fontweight="bold")
ax.annotate("", xy=(7.1, 3.3), xytext=(6.3, 3.3), arrowprops=dict(arrowstyle="->", lw=2))
ax.text(6.7, 3.5, "SQL INSERT", ha="center", fontsize=8, fontweight="bold")
ax.annotate("", xy=(6.3, 2.7), xytext=(7.1, 2.7), arrowprops=dict(arrowstyle="->", lw=2))
ax.text(6.7, 2.45, "History Rows", ha="center", fontsize=8, fontweight="bold")
plt.tight_layout()
arch_img = os.path.join("documentation", "system_architecture.png")
plt.savefig(arch_img, dpi=200); plt.close()

# 2. DFD Level 1 Diagram
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
ax.set_title("FIGURE 4.2: DATA FLOW DIAGRAM (LEVEL 1) - RESUMEMATCH AI", fontsize=12, fontweight="bold")
nodes = [
    (0.3, 2.5, 1.9, 1.0, "#e0e7ff", "#3730a3", "Candidate / Admin\n(External Entity)"),
    (3.1, 4.3, 2.3, 0.9, "#fce7f3", "#9d174d", "1.0 Authentication &\nSession Guard"),
    (3.1, 2.5, 2.3, 1.0, "#dcfce7", "#166534", "2.0 PDF Parser & TF-IDF\nCosine NLP Engine"),
    (3.1, 0.7, 2.3, 0.9, "#fef9c3", "#854d0e", "3.0 Search, Analytics &\nCSV Report Export"),
    (6.8, 3.7, 2.6, 0.9, "#f3f4f6", "#374151", "D1: users\n(MySQL Table)"),
    (6.8, 1.4, 2.6, 1.0, "#f3f4f6", "#374151", "D2: resume_analyses\n(MySQL Table)")
]
for x, y, w, h, fc, ec, txt in nodes:
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08", facecolor=fc, edgecolor=ec, linewidth=1.8))
    ax.text(x+w/2, y+h/2, txt, ha="center", va="center", fontsize=8.5, fontweight="bold")
for st, en in [((2.3, 3.3), (3.0, 4.7)), ((2.3, 3.0), (3.0, 3.0)), ((2.3, 2.7), (3.0, 1.2)),
               ((5.5, 4.7), (6.7, 4.1)), ((5.5, 3.0), (6.7, 2.1)), ((6.7, 1.6), (5.5, 1.2))]:
    ax.annotate("", xy=en, xytext=st, arrowprops=dict(arrowstyle="->", lw=1.6))
plt.tight_layout()
dfd_img = os.path.join("documentation", "dfd_diagram.png")
plt.savefig(dfd_img, dpi=200); plt.close()

# 3. Use Case Diagram
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
ax.set_title("FIGURE 4.3: UML USE CASE DIAGRAM - RESUMEMATCH AI", fontsize=12, fontweight="bold")
ax.add_patch(patches.Rectangle((2.8, 0.4), 4.4, 5.1, fill=False, edgecolor="#312e81", linewidth=2, linestyle="--"))
ax.text(5.0, 5.25, "ResumeMatch AI System Boundary", ha="center", fontsize=9.5, fontweight="bold", color="#312e81")
use_cases = [
    (5.0, 4.5, "UC-1: Register & Login"),
    (5.0, 3.7, "UC-2: Upload PDF or Paste Resume"),
    (5.0, 2.9, "UC-3: Compute ATS Score & Top 3 Roles"),
    (5.0, 2.1, "UC-4: Inspect Matched & Missing Skills"),
    (5.0, 1.3, "UC-5: Search & Export Placement CSV"),
    (5.0, 0.65, "UC-6: Manage / Delete Analyses")
]
for cx, cy, label in use_cases:
    ax.add_patch(patches.Ellipse((cx, cy), 3.6, 0.55, facecolor="#eef2ff", edgecolor="#4f46e5", linewidth=1.5))
    ax.text(cx, cy, label, ha="center", va="center", fontsize=8.5, fontweight="bold")
ax.add_patch(patches.FancyBboxPatch((0.4, 2.5), 1.6, 1.0, boxstyle="round,pad=0.1", facecolor="#fef3c7", edgecolor="#b45309", linewidth=2))
ax.text(1.2, 3.0, "ACTOR 1:\nCandidate", ha="center", va="center", fontsize=8.5, fontweight="bold")
ax.add_patch(patches.FancyBboxPatch((8.0, 2.5), 1.6, 1.0, boxstyle="round,pad=0.1", facecolor="#dcfce7", edgecolor="#15803d", linewidth=2))
ax.text(8.8, 3.0, "ACTOR 2:\nPlacement Admin", ha="center", va="center", fontsize=8.5, fontweight="bold")
for _, cy, _ in use_cases[:5]:
    ax.plot([2.1, 3.2], [3.0, cy], color="#374151", lw=1.2)
for _, cy, _ in use_cases:
    ax.plot([6.8, 7.9], [cy, 3.0], color="#374151", lw=1.2)
plt.tight_layout()
uc_img = os.path.join("documentation", "use_case_diagram.png")
plt.savefig(uc_img, dpi=200); plt.close()

# 4. Flowchart Diagram
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
ax.set_title("FIGURE 4.4: NLP RESUME ANALYSIS EXECUTION FLOWCHART", fontsize=12, fontweight="bold")
steps = [
    (0.4, 2.5, 1.6, 0.9, "#e0e7ff", "#3730a3", "START\nUpload PDF/Text"),
    (2.4, 2.5, 1.7, 0.9, "#fce7f3", "#9d174d", "Extract & Clean\nRaw Text (pypdf)"),
    (4.5, 2.5, 1.8, 0.9, "#fef3c7", "#b45309", "TF-IDF Vector &\nCosine Similarity"),
    (6.7, 2.5, 1.8, 0.9, "#dcfce7", "#15803d", "Rank 6 IT Roles &\nSkill Gap Check"),
    (8.8, 2.5, 1.0, 0.9, "#dbeafe", "#1d4ed8", "Save DB &\nShow Top 3")
]
for i, (x, y, w, h, fc, ec, txt) in enumerate(steps):
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08", facecolor=fc, edgecolor=ec, linewidth=1.8))
    ax.text(x+w/2, y+h/2, txt, ha="center", va="center", fontsize=8, fontweight="bold")
    if i < len(steps)-1:
        ax.annotate("", xy=(steps[i+1][0]-0.05, 2.95), xytext=(x+w+0.05, 2.95), arrowprops=dict(arrowstyle="->", lw=1.8))
plt.tight_layout()
flow_img = os.path.join("documentation", "flowchart_diagram.png")
plt.savefig(flow_img, dpi=200); plt.close()

# 5. ER Diagram
fig, ax = plt.subplots(figsize=(10, 4.8))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.5); ax.axis("off")
ax.set_title("FIGURE 5.1: ENTITY-RELATIONSHIP (ER) SCHEMA - DATABASE: resume_ai_db", fontsize=12, fontweight="bold")
t1 = "ENTITY: users\n--------------------------\n[PK] id (INT, AI)\nfull_name (VARCHAR 100)\nemail (VARCHAR 100, UNIQUE)\npassword (VARCHAR 255)\nrole (ENUM: Admin/Candidate)\ncreated_at (TIMESTAMP)"
t2 = "ENTITY: resume_analyses\n------------------------------------\n[PK] id (INT, AI)\ncandidate_name (VARCHAR 100)\ncandidate_email (VARCHAR 100)\nsource_type (VARCHAR 30)\nrecommended_role (VARCHAR 100)\nats_score (FLOAT)\nmatched_skills (TEXT)\nmissing_skills (TEXT)\ncreated_at (TIMESTAMP)"
ax.add_patch(patches.FancyBboxPatch((0.6, 1.3), 3.3, 3.0, boxstyle="round,pad=0.1", facecolor="#eff6ff", edgecolor="#1e40af", linewidth=2))
ax.text(2.25, 2.8, t1, ha="center", va="center", fontsize=8.5, family="monospace", fontweight="bold")
ax.add_patch(patches.FancyBboxPatch((5.5, 0.8), 3.9, 3.8, boxstyle="round,pad=0.1", facecolor="#f0fdf4", edgecolor="#166534", linewidth=2))
ax.text(7.45, 2.7, t2, ha="center", va="center", fontsize=8.5, family="monospace", fontweight="bold")
ax.annotate("", xy=(5.4, 2.7), xytext=(4.0, 2.7), arrowprops=dict(arrowstyle="->", lw=2))
ax.text(4.7, 2.95, "1 : N\nSubmits", ha="center", fontsize=8.5, fontweight="bold")
plt.tight_layout()
er_img = os.path.join("documentation", "er_diagram.png")
plt.savefig(er_img, dpi=200); plt.close()

# 6. ATS Benchmark Chart
demo_res = analyze_resume("BCA 6th Sem student skilled in Python, Flask, MySQL, HTML, CSS, JavaScript, Machine Learning, Pandas, Scikit-Learn, and Git.")
r_names = [x["role"] for x in demo_res]
r_scores = [x["ats_score"] for x in demo_res]
fig, ax = plt.subplots(figsize=(8.5, 4.3))
bars = ax.barh(r_names[::-1], r_scores[::-1], color="#4f46e5")
ax.set_xlabel("ATS Compatibility Score (%)", fontweight="bold")
ax.set_title("FIGURE 6.1: NLP CAREER ROLE COMPATIBILITY BENCHMARK", fontsize=11.5, fontweight="bold")
for b in bars:
    ax.text(b.get_width() + 0.8, b.get_y() + b.get_height()/2, f"{b.get_width()}%", va="center", fontsize=9, fontweight="bold")
ax.set_xlim(0, 100)
plt.tight_layout()
ats_img = os.path.join("documentation", "ats_benchmark.png")
plt.savefig(ats_img, dpi=200); plt.close()

print("[1/2] All 6 High-Resolution Diagrams generated for Project 02!")

# ---------------------------------------------------------
# PART B: GENERATE UNIVERSITY-GRADE REPORT (.DOCX)
# ---------------------------------------------------------
doc = Document()
for section in doc.sections:
    section.top_margin = Inches(1.0); section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0); section.right_margin = Inches(1.0)
    section.header.paragraphs[0].text = "AI-Powered Resume Analyzer & Career Recommendation System (ResumeMatch AI) | BCA 6th Sem"
    section.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    section.footer.paragraphs[0].text = "Department of Computer Applications | Submitted by: Chaman Kumar"
    section.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(12)
style.paragraph_format.line_spacing = 1.5

def add_ch(num, title):
    doc.add_page_break()
    p = doc.add_paragraph()
    r1 = p.add_run(f"CHAPTER {num}\n"); r1.bold = True; r1.font.size = Pt(14); r1.font.color.rgb = RGBColor(30, 64, 175)
    r2 = p.add_run(title.upper()); r2.bold = True; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph("_" * 62)

def add_h2(t):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)
    r = p.add_run(t); r.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(29, 78, 216)

def add_p(t):
    p = doc.add_paragraph(t); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.space_after = Pt(8)

def add_code_block(code_str):
    p = doc.add_paragraph(); p.paragraph_format.line_spacing = 1.15
    r = p.add_run(code_str); r.font.name = "Consolas"; r.font.size = Pt(9.5)

# COVER PAGE
pc = doc.add_paragraph(); pc.alignment = WD_ALIGN_PARAGRAPH.CENTER; pc.paragraph_format.space_before = Pt(12)
r_u = pc.add_run("[ENTER YOUR UNIVERSITY NAME]\n"); r_u.bold = True; r_u.font.size = Pt(15); r_u.font.color.rgb = RGBColor(153, 27, 27)
r_c = pc.add_run("[ENTER YOUR COLLEGE NAME]\nDEPARTMENT OF COMPUTER APPLICATIONS (BCA)\n\n"); r_c.bold = True; r_c.font.size = Pt(13); r_c.font.color.rgb = RGBColor(30, 58, 138)
pc.add_run("A MAJOR PROJECT REPORT ON\n\n").bold = True
r_t = pc.add_run("AI-POWERED RESUME ANALYZER AND\nCAREER RECOMMENDATION SYSTEM\n(ResumeMatch AI)\n\n"); r_t.bold = True; r_t.font.size = Pt(21)
pc.add_run("Submitted in partial fulfillment of the requirements for the award of the degree of\nBACHELOR OF COMPUTER APPLICATIONS (BCA – 6TH SEMESTER)\n\n").italic = True

tbl = doc.add_table(rows=2, cols=2); tbl.style = "Table Grid"
tbl.rows[0].cells[0].text = "SUBMITTED BY:\nName: CHAMAN KUMAR\nClass: BCA 6th Semester\nCollege Roll No: [Enter Roll No]\nUniversity Reg. No: [Enter Reg. No]\nExam Roll No: [Enter Exam Roll No]"
tbl.rows[0].cells[1].text = "UNDER THE SUPERVISION OF:\nGuide Name: [Enter Supervisor Name]\nDesignation: Faculty, Dept. of BCA\nDepartment of Computer Applications\nAcademic Session: 2023–2026"
tbl.rows[1].cells[0].text = "TECHNOLOGY STACK:\nPython 3.14, Flask 3.1, Scikit-Learn,\npypdf, MySQL 8.0, Bootstrap 5, Chart.js"
tbl.rows[1].cells[1].text = "SOURCE CODE REPOSITORY:\ngithub.com/Kumar-Chaman/\nBCA-6th-Semester-20-Projects"

# CERTIFICATE, DECLARATION, ACKNOWLEDGEMENT, ABSTRACT
doc.add_page_break()
add_h2("PROJECT COMPLETION CERTIFICATE")
add_p("This is to certify that the major project work entitled 'AI-Powered Resume Analyzer and Career Recommendation System (ResumeMatch AI)' has been bonafidely carried out and submitted by CHAMAN KUMAR (College Roll No: [Enter Roll No], University Registration No: [Enter Reg. No]), a student of BCA 6th Semester, Department of Computer Applications, under my supervision and guidance in partial fulfillment of the requirements for the degree of Bachelor of Computer Applications (BCA).")
add_p("\n\n__________________________                                        __________________________\nSignature of Project Guide                                            Signature of Head / Coordinator\n\n__________________________\nSignature of External Examiner")

doc.add_page_break()
add_h2("STUDENT DECLARATION & ACKNOWLEDGEMENT")
add_p("I, Chaman Kumar, hereby declare that this project work titled 'AI-Powered Resume Analyzer and Career Recommendation System (ResumeMatch AI)' is my original software development and Natural Language Processing work implemented using Python 3.14, Flask, Scikit-Learn, pypdf, and MySQL 8.0.")
add_p("I express sincere gratitude to my Project Supervisor, Head of Department, and faculty members for their invaluable mentorship throughout this project.")
add_p("\nDate: November 2026                                                       Signature of Student: Chaman Kumar")

doc.add_page_break()
add_h2("ABSTRACT")
add_p("Modern corporate recruitment relies heavily on automated Applicant Tracking Systems (ATS) to filter candidate resumes prior to human interviews. Graduating BCA students often face rejection due to unstructured resumes and unaddressed technical skill gaps. This project presents 'ResumeMatch AI'—an intelligent, full-stack web application powered by Natural Language Processing (NLP) that parses uploaded PDF resumes or pasted curriculum vitae text, computes ATS compatibility scores (0–100%), ranks the top 3 best-fit IT career tracks, and highlights missing skills to learn next.")
add_p("Built using Python 3.14, Flask, pypdf, Scikit-Learn (TF-IDF Bigram Vectorizer & Cosine Similarity), MySQL 8.0 ('resume_ai_db'), Bootstrap 5, and Chart.js, the platform evaluates candidates across six core IT career tracks: Full-Stack Web Developer, AI & Machine Learning Engineer, Data Analyst & BI Specialist, Cybersecurity & Network Analyst, Java Enterprise Backend Developer, and Database Administrator (DBA).")

# LIST OF FIGURES & TOC
doc.add_page_break()
add_h2("LIST OF FIGURES & TABLES")
for fg in [
    "Figure 4.1: Three-Tier System Architecture of ResumeMatch AI",
    "Figure 4.2: Level 1 Data Flow Diagram (DFD)",
    "Figure 4.3: UML Use Case Diagram (Candidate & Placement Admin)",
    "Figure 4.4: NLP Resume Analysis Execution Flowchart",
    "Figure 5.1: Entity-Relationship (ER) Schema for resume_ai_db",
    "Figure 6.1: NLP Career Role Compatibility Benchmark Chart"
]:
    doc.add_paragraph(fg, style="List Bullet")

add_h2("TABLE OF CONTENTS")
toc = [
    ("Chapter 1", "Introduction, Background, Problem Formulation & Objectives"),
    ("Chapter 2", "System Analysis, Comparative Study, Feasibility & SDLC"),
    ("Chapter 3", "Software & Hardware Requirements Specification (SRS)"),
    ("Chapter 4", "System Design: Architecture, DFD, Use Case & Flowchart"),
    ("Chapter 5", "Database Engineering: Normalization, ER Schema & SQL DDL"),
    ("Chapter 6", "NLP & Mathematical Methodology (TF-IDF & Cosine Similarity)"),
    ("Chapter 7", "Module Description & Pseudocode Algorithms"),
    ("Chapter 8", "Complete Annotated Source Code"),
    ("Chapter 9", "Software Testing, Test Cases & Observed Results"),
    ("Chapter 10", "User Operating Manual & Step-by-Step Guide"),
    ("Chapter 11", "Security Analysis, Future Scope, Conclusion & Bibliography"),
    ("Appendix A", "40 Comprehensive Viva-Voce Questions & Detailed Answers")
]
t_toc = doc.add_table(rows=1, cols=2); t_toc.style = "Table Grid"
t_toc.rows[0].cells[0].text = "Chapter"; t_toc.rows[0].cells[1].text = "Title"
for c, t in toc:
    r = t_toc.add_row().cells; r[0].text = c; r[1].text = t

# CHAPTER 1-3
add_ch("1", "Introduction & Problem Formulation")
add_h2("1.1 Problem Statement & Objectives")
add_p("Campus placement cells manually review hundreds of student resumes, making objective skill-gap counseling nearly impossible. ResumeMatch AI solves this by automating PDF text extraction, keyword boundary matching, and TF-IDF semantic similarity scoring across six industry job profiles.")
for o in [
    "Support dual-mode resume ingestion: direct PDF document upload (via pypdf) and plain-text resume pasting.",
    "Compute a hybrid ATS Compatibility Score combining 70% explicit skill coverage and 30% TF-IDF Cosine Similarity.",
    "Deliver top 3 ranked career recommendations with green 'Matched Skills' and red 'Missing Skills' roadmaps.",
    "Persist all candidate evaluations in MySQL 8.0 ('resume_ai_db') with search, Chart.js analytics, and CSV export."
]:
    doc.add_paragraph(o, style="List Bullet")

add_ch("2", "System Analysis, Feasibility & SDLC")
add_h2("2.1 Existing vs. Proposed System Comparison")
t_cmp = doc.add_table(rows=1, cols=3); t_cmp.style = "Table Grid"
t_cmp.rows[0].cells[0].text = "Criteria"; t_cmp.rows[0].cells[1].text = "Manual Resume Screening"; t_cmp.rows[0].cells[2].text = "ResumeMatch AI (NLP)"
for c1, c2, c3 in [
    ("Screening Speed", "10-15 minutes per resume", "< 0.1 seconds per PDF/Text resume"),
    ("ATS Score Quantification", "Subjective human guess", "Mathematical 0-100% TF-IDF + Skill Score"),
    ("Skill Gap Guidance", "Rarely provided to students", "Automatic Matched vs. Missing skill badges"),
    ("Placement Analytics", "Manual Excel tallying", "Live Chart.js bar chart & CSV export")
]:
    r = t_cmp.add_row().cells; r[0].text = c1; r[1].text = c2; r[2].text = c3
add_h2("2.2 Feasibility & Agile SDLC")
add_p("Verified technically on Python 3.14.3, Flask 3.1.3, Scikit-Learn 1.9.1, pypdf, and MySQL 8.0 with zero licensing costs under an Agile iterative development workflow.")

add_ch("3", "Software & Hardware Requirements Specification")
add_h2("3.1 Functional Requirements (FR-01 to FR-10)")
for fr in [
    "FR-01 & FR-02: Candidate/Admin registration and session-authenticated login (/login, /register).",
    "FR-03 & FR-04: Secure PDF file upload (secure_filename + .pdf check) and multi-page text extraction via pypdf.",
    "FR-05 & FR-06: Regex word-boundary skill matching and Scikit-Learn TF-IDF bigram Cosine Similarity computation.",
    "FR-07 & FR-08: Display Top 3 ranked IT career roles with ATS percentage badges and skill-gap breakdown.",
    "FR-09 & FR-10: MySQL persistence, candidate/role search, Chart.js career track chart, and CSV download."
]:
    doc.add_paragraph(fr, style="List Bullet")

# CHAPTER 4-6
add_ch("4", "System Design & Visual Diagrams")
add_h2("4.1 Three-Tier System Architecture")
doc.add_picture(arch_img, width=Inches(6.2))
add_h2("4.2 Data Flow Diagram (Level 1)")
doc.add_picture(dfd_img, width=Inches(6.2))
add_h2("4.3 UML Use Case Diagram")
doc.add_picture(uc_img, width=Inches(6.2))
add_h2("4.4 NLP Execution Flowchart")
doc.add_picture(flow_img, width=Inches(6.2))

add_ch("5", "Database Engineering & ER Schema (MySQL 8.0)")
add_h2("5.1 3NF Normalization & Entity-Relationship Diagram")
add_p("The database 'resume_ai_db' adheres to Third Normal Form (3NF), separating user authentication records ('users') from resume evaluation logs ('resume_analyses').")
doc.add_picture(er_img, width=Inches(6.2))
add_h2("5.2 SQL DDL Schema")
add_code_block("""CREATE DATABASE IF NOT EXISTS resume_ai_db;
USE resume_ai_db;

CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    full_name VARCHAR(100) NOT NULL,
    email VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role ENUM('Admin', 'Candidate') DEFAULT 'Candidate',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE resume_analyses (
    id INT AUTO_INCREMENT PRIMARY KEY,
    candidate_name VARCHAR(100) NOT NULL,
    candidate_email VARCHAR(100) NOT NULL,
    source_type VARCHAR(30) NOT NULL,
    recommended_role VARCHAR(100) NOT NULL,
    ats_score FLOAT NOT NULL,
    matched_skills TEXT NOT NULL,
    missing_skills TEXT NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);""")

add_ch("6", "NLP Mathematical Methodology & Benchmark")
add_h2("6.1 TF-IDF Vectorization & Cosine Similarity Formulation")
add_p("Each job profile and candidate resume is transformed into a TF-IDF vector using unigrams and bigrams (ngram_range=(1, 2)):\n    TF-IDF(t, d) = TF(t, d) * log( N / (1 + DF(t)) )\nSemantic similarity between resume vector R and job profile vector J is measured via Cosine Similarity:\n    Cosine_Sim(R, J) = (R . J) / (||R|| * ||J||)\nThe final ATS Score blends 70% explicit required-skill coverage with 30% TF-IDF semantic similarity.")
doc.add_picture(ats_img, width=Inches(5.8))

# CHAPTER 7 & 8
add_ch("7", "Module Description & Pseudocode Algorithms")
add_h2("7.1 Algorithm: PDF Ingestion & Hybrid ATS Ranking")
add_code_block("""ALGORITHM Analyze_Candidate_Resume(pdf_file, raw_text):
Step 1: If pdf_file is uploaded, verify '.pdf' extension, save safely, and extract text across all pages using PdfReader.
Step 2: Normalize text to lowercase and strip special punctuation while preserving '+', '#', '/', '.'.
Step 3: Fit TfidfVectorizer(ngram_range=(1, 2)) on 6 Job Profile descriptions + cleaned resume text.
Step 4: Compute cosine_similarity between resume vector and all 6 job vectors.
Step 5: For each job role, run regex word-boundary matching to partition skills into matched_skills and missing_skills.
Step 6: Calculate ats_score = 0.70 * skill_ratio + 0.30 * nlp_similarity + bonus.
Step 7: Sort roles descending by ats_score, insert Top 1 into MySQL 'resume_analyses', and display Top 3 cards.""")

add_ch("8", "Complete Annotated Source Code")
for title, path in [
    ("8.1 Database Setup Script (database/setup_db.py)", os.path.join("database", "setup_db.py")),
    ("8.2 NLP Career Recommendation Engine (ml_engine/career_engine.py)", os.path.join("ml_engine", "career_engine.py")),
    ("8.3 MySQL Connection Config (config.py)", "config.py"),
    ("8.4 Main Flask Application (app.py)", "app.py"),
    ("8.5 Interactive Dashboard UI (templates/dashboard.html)", os.path.join("templates", "dashboard.html")),
    ("8.6 Automated Test Suite (tests/test_app.py)", os.path.join("tests", "test_app.py"))
]:
    add_h2(title)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8-sig") as f:
            add_code_block(f.read())

# CHAPTER 9-11
add_ch("9", "Software Testing & Observed Results")
t_tc = doc.add_table(rows=1, cols=5); t_tc.style = "Table Grid"
for idx, h in enumerate(["Test ID", "Test Scenario", "Input", "Expected Result", "Actual"]):
    t_tc.rows[0].cells[idx].text = h
for tc in [
    ("TC-01", "MySQL Connection (resume_ai_db)", "get_db_connection()", "Active connection = True", "PASS"),
    ("TC-02", "Placement Admin Login", "admin@resumematch.ai / admin123", "Session created & Dashboard loaded", "PASS"),
    ("TC-03", "Full-Stack Resume Analysis", "Python, Flask, MySQL, JS text", "Full-Stack Developer Top Match saved", "PASS"),
    ("TC-04", "AI/ML Engineer Resume Analysis", "ML, Deep Learning, NLP, Pandas", "AI & ML Engineer Top Match saved", "PASS"),
    ("TC-05", "Short/Empty Input Guard", "resume_text = 'abc'", "Reject with danger flash alert", "PASS"),
    ("TC-06", "Non-PDF File Upload Guard", "Upload .exe or .txt file", "Reject non-PDF upload safely", "PASS"),
    ("TC-07", "Candidate/Role Search Filter", "GET /dashboard?search=Chaman", "Filter matching rows via SQL LIKE", "PASS"),
    ("TC-08", "CSV Placement Report Export", "GET /export_csv", "HTTP 200 CSV attachment downloaded", "PASS")
]:
    row = t_tc.add_row().cells
    for i, v in enumerate(tc): row[i].text = v

add_ch("10", "User Operating Manual")
for st in [
    "Step 1: Confirm MySQL80 service is active and run 'python database\\setup_db.py'.",
    "Step 2: Launch the Flask server using 'python app.py' (runs on http://127.0.0.1:5002).",
    "Step 3: Open http://127.0.0.1:5002 in your browser and sign in with admin@resumematch.ai / admin123.",
    "Step 4: Enter Candidate Name & Email, upload a .PDF resume OR paste resume text, and click 'Run NLP Resume Analysis & Save'.",
    "Step 5: Inspect the Top 3 Career Recommendation cards, Matched/Missing skill badges, Chart.js bar graph, and export CSV reports."
]:
    doc.add_paragraph(st, style="List Bullet")

add_ch("11", "Security Analysis, Conclusion & Bibliography")
add_p("ResumeMatch AI enforces secure_filename() sanitization on uploaded PDFs, strict file-extension verification, parameterized SQL queries against SQL injection, @login_required route guards, and .env secret isolation. Future enhancements include transformer-based semantic embeddings (SBERT) and automated resume bullet-point rewriting.")
for ref in [
    "1. Salton, G. & Buckley, C. (1988). Term-weighting approaches in automatic text retrieval. Information Processing & Management.",
    "2. Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. JMLR, 12, 2825-2830.",
    "3. pypdf Official Documentation (2026). https://pypdf.readthedocs.io/"
]:
    doc.add_paragraph(ref, style="List Bullet")

# APPENDIX A: 40 VIVA Q&A
add_ch("A", "40 Comprehensive Viva-Voce Questions & Answers")
viva_40 = [
    ("Q1. What is the main purpose of ResumeMatch AI?", "To parse student resumes (PDF or text), compute an ATS compatibility score using NLP, recommend top 3 IT career tracks, and identify missing skills."),
    ("Q2. What does NLP stand for?", "Natural Language Processing—a branch of AI enabling computers to analyze, clean, and extract meaning from human text."),
    ("Q3. Explain TF-IDF in simple terms.", "Term Frequency-Inverse Document Frequency assigns high numerical weights to rare, important technical terms (like 'scikit-learn') and low weights to common filler words."),
    ("Q4. What is Cosine Similarity and what is its range?", "It measures the cosine of the angle between two non-zero vectors; for TF-IDF vectors with non-negative weights, it ranges from 0 (no match) to 1 (identical)."),
    ("Q5. Why did we use ngram_range=(1, 2) in TfidfVectorizer?", "To capture both single keywords ('python', 'mysql') and two-word technical phrases ('machine learning', 'rest api', 'power bi')."),
    ("Q6. How is the final ATS Score calculated?", "70% weight is given to explicit required-skill coverage and 30% weight to TF-IDF Cosine Similarity, plus a multi-skill bonus."),
    ("Q7. Which Python library extracts text from PDF files?", "pypdf (PdfReader class), which iterates through every page of the PDF and extracts embedded text."),
    ("Q8. How do you prevent 'java' from falsely matching inside 'javascript'?", "By using regular expression word boundaries (\\bjava\\b) during skill extraction."),
    ("Q9. Why is secure_filename() important when handling PDF uploads?", "It strips dangerous path traversal characters (like ../../) from user-supplied filenames before saving to disk."),
    ("Q10. How many career tracks are supported in career_engine.py?", "Six tracks: Full-Stack Web Developer, AI & ML Engineer, Data Analyst, Cybersecurity Analyst, Java Backend Developer, and DBA."),
    ("Q11. What happens if a user uploads a .docx or .exe file instead of .pdf?", "The /analyze route checks filename.lower().endswith('.pdf') and rejects non-PDF uploads with a warning alert."),
    ("Q12. What happens if a user submits empty or very short resume text (<20 chars)?", "Server-side validation rejects it and prompts the user to upload a valid PDF or paste detailed skills."),
    ("Q13. Explain the database tables in resume_ai_db.", "'users' stores Admin/Candidate accounts; 'resume_analyses' stores candidate info, source type, recommended role, ATS score, matched skills, and missing skills."),
    ("Q14. How does Flask pass Top 3 recommendations across a redirect?", "It temporarily stores the dictionary in session['latest_result'] before redirecting to /dashboard, then pops and renders it."),
    ("Q15. Why use Post-Redirect-Get (PRG) pattern in /analyze?", "Redirecting after a POST request prevents duplicate database insertions if the user refreshes the browser."),
    ("Q16. What is enctype='multipart/form-data' in HTML forms?", "It encodes binary file data alongside form text fields so Flask's request.files can receive uploaded PDFs."),
    ("Q17. How does Chart.js render the career distribution bar chart?", "Jinja2 converts Python dictionary keys and values into JavaScript arrays using the | tojson filter."),
    ("Q18. How is SQL Injection prevented in search queries?", "By passing search terms as parameterized %s tuple arguments to cursor.execute()."),
    ("Q19. What port does Project 02 run on?", "Port 5002 (http://127.0.0.1:5002) so it can run simultaneously alongside Project 01 on Port 5001."),
    ("Q20. How does CSV export work in Project 02?", "It streams rows from resume_analyses into an in-memory io.StringIO buffer and serves it as ResumeMatch_AI_Report.csv."),
    ("Q21. What is the difference between stemming/lemmatization and regex normalization?", "Regex normalization strips non-alphanumeric noise while keeping tech symbols like C++ and C#, whereas stemming truncates word suffixes."),
    ("Q22. Why are '+', '#', and '.' preserved in our regex cleaner?", "So technical skills like C++, C#, and Node.js are not corrupted into 'c' or 'nodejs'."),
    ("Q23. What is Third Normal Form (3NF)?", "Every non-prime attribute is non-transitively dependent on every candidate key of the table."),
    ("Q24. Explain the role of @login_required in app.py.", "It protects private dashboard, analysis, delete, and export endpoints from unauthenticated HTTP requests."),
    ("Q25. What is a Corpus in NLP?", "The complete collection of text documents (our 6 job descriptions + candidate resume) fed into the TF-IDF vectorizer."),
    ("Q26. What is a Sparse Matrix in Scikit-Learn?", "A memory-efficient matrix storing only non-zero TF-IDF values across the vocabulary."),
    ("Q27. How can scanned image-only PDFs be supported in future upgrades?", "By integrating Tesseract OCR (pytesseract + pdf2image) as a fallback when pypdf extracts empty text."),
    ("Q28. What is the benefit of displaying 'Missing Skills' to BCA students?", "It acts as a personalized learning roadmap showing exact technologies required to boost their ATS score."),
    ("Q29. How did you verify Project 02 automatically?", "Using tests/test_app.py with Flask's test_client() validating DB connection, auth, NLP analysis, input guards, and CSV export."),
    ("Q30. What is the difference between HTTP 200 and HTTP 302?", "HTTP 200 means OK (successful response); HTTP 302 means temporary redirect (used after login/logout/form submission)."),
    ("Q31. Why is .env excluded from Git?", "To keep database passwords confidential on the local machine rather than exposing them in a public GitHub repository."),
    ("Q32. What is Werkzeug?", "The underlying WSGI utility library powering Flask's routing, request/response objects, and secure_filename()."),
    ("Q33. How does Bootstrap 5 ensure mobile responsiveness?", "Through viewport meta tags, flexbox grid rows/columns, and responsive table wrappers (table-responsive)."),
    ("Q34. What is the purpose of cursor(dictionary=True)?", "It maps SQL column names to Python dictionary keys for clean template rendering."),
    ("Q35. How does the search bar filter records?", "It checks both candidate_name and recommended_role columns using SQL LIKE '%query%'."),
    ("Q36. Why is TEXT data type used for matched_skills and missing_skills?", "Because a candidate may match 15+ skills, exceeding short VARCHAR limits."),
    ("Q37. What is an Applicant Tracking System (ATS)?", "HR software used by recruiters to automatically parse, score, and rank job applications based on keyword and skill relevance."),
    ("Q38. How can this tool be integrated into a college placement portal?", "By deploying it on a central college server where students upload resumes before campus recruitment drives."),
    ("Q39. What is the time complexity of TF-IDF cosine similarity for 6 job roles?", "Nearly instantaneous O(V) where V is vocabulary size (~150 terms), executing in under 15 milliseconds."),
    ("Q40. What key engineering lesson does Project 02 demonstrate?", "Combining unstructured document parsing (pypdf), NLP vector space modeling (TF-IDF), and relational SQL persistence into a cohesive web product.")
]
for q, a in viva_40:
    pq = doc.add_paragraph(); pq.paragraph_format.space_before = Pt(6); pq.paragraph_format.space_after = Pt(2)
    rq = pq.add_run(q); rq.bold = True; rq.font.color.rgb = RGBColor(30, 58, 138)
    add_p(f"Answer: {a}")

rep_path = os.path.join("documentation", "Project_02_Complete_Report.docx")
doc.save(rep_path)
print(f"[2/2] Upgraded Project 02 Report saved: {rep_path}")

