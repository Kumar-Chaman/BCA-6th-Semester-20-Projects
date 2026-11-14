import re
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

JOB_PROFILES = {
    "Full-Stack Web Developer": {
        "skills": ["html", "css", "javascript", "python", "flask", "django", "react", "node", "mysql", "bootstrap", "git", "rest api"],
        "description": "html css javascript python flask django react node mysql bootstrap git rest api web frontend backend database"
    },
    "AI & Machine Learning Engineer": {
        "skills": ["python", "machine learning", "deep learning", "scikit-learn", "pandas", "numpy", "tensorflow", "nlp", "data science", "statistics", "opencv"],
        "description": "python machine learning deep learning scikit-learn pandas numpy tensorflow nlp data science statistics opencv artificial intelligence"
    },
    "Data Analyst & BI Specialist": {
        "skills": ["sql", "mysql", "excel", "python", "pandas", "matplotlib", "power bi", "tableau", "data visualization", "statistics", "etl"],
        "description": "sql mysql excel python pandas matplotlib power bi tableau data visualization statistics etl reporting analytics dashboard"
    },
    "Cybersecurity & Network Analyst": {
        "skills": ["networking", "linux", "kali linux", "wireshark", "nmap", "cryptography", "firewall", "penetration testing", "python", "tcp/ip", "security"],
        "description": "networking linux kali linux wireshark nmap cryptography firewall penetration testing python tcp/ip security vulnerability"
    },
    "Java Enterprise Backend Developer": {
        "skills": ["java", "spring boot", "jdbc", "mysql", "oop", "multithreading", "hibernate", "maven", "git", "rest api", "sql"],
        "description": "java spring boot jdbc mysql oop multithreading hibernate maven git rest api sql enterprise backend"
    },
    "Database Administrator (DBA)": {
        "skills": ["mysql", "sql", "normalization", "indexing", "stored procedures", "backup", "recovery", "database design", "er diagram", "linux", "performance tuning"],
        "description": "mysql sql normalization indexing stored procedures backup recovery database design er diagram linux performance tuning"
    }
}

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return text.strip()

def analyze_resume(resume_text):
    clean_text = re.sub(r"[^a-zA-Z0-9+#/.\s]", " ", resume_text.lower())
    
    roles = list(JOB_PROFILES.keys())
    corpus = [JOB_PROFILES[r]["description"] for r in roles] + [clean_text]
    
    vectorizer = TfidfVectorizer(ngram_range=(1, 2))
    tfidf_matrix = vectorizer.fit_transform(corpus)
    
    resume_vec = tfidf_matrix[-1]
    role_vecs = tfidf_matrix[:-1]
    similarities = cosine_similarity(resume_vec, role_vecs)[0]
    
    ranked_roles = []
    for idx, role in enumerate(roles):
        req_skills = JOB_PROFILES[role]["skills"]
        matched = [s for s in req_skills if re.search(r"\b" + re.escape(s) + r"\b", clean_text)]
        missing = [s for s in req_skills if s not in matched]
        
        skill_ratio = (len(matched) / len(req_skills)) * 100
        nlp_sim = similarities[idx] * 100
        # Combined weighted ATS Score (70% Skill Coverage + 30% TF-IDF Semantic Similarity)
        ats_score = round(min(98.5, (0.70 * skill_ratio) + (0.30 * nlp_sim) + (10 if len(matched) >= 3 else 0)), 2)
        
        ranked_roles.append({
            "role": role,
            "ats_score": ats_score,
            "matched_skills": matched,
            "missing_skills": missing
        })
        
    ranked_roles.sort(key=lambda x: x["ats_score"], reverse=True)
    return ranked_roles

if __name__ == "__main__":
    sample_cv = "Chaman Kumar. BCA 6th Semester student skilled in Python, Flask, MySQL, HTML, CSS, JavaScript, Machine Learning, Pandas, Scikit-Learn, and Git."
    res = analyze_resume(sample_cv)
    print(f"[2/2] NLP Engine Ready! Top Match: {res[0]['role']} (ATS Score: {res[0]['ats_score']}%)")
