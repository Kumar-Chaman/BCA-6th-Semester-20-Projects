import mysql.connector
import os

def load_env():
    env_vars = {"DB_HOST": "localhost", "DB_USER": "root", "DB_PASSWORD": ""}
    env_path = r"F:\BCA_6th_Sem_Projects\.env"
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8-sig") as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    env_vars[k] = v
    return env_vars

env = load_env()
conn = mysql.connector.connect(host=env["DB_HOST"], user=env["DB_USER"], password=env["DB_PASSWORD"])
cursor = conn.cursor()

cursor.execute("CREATE DATABASE IF NOT EXISTS resume_ai_db")
cursor.execute("USE resume_ai_db")

cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INT AUTO_INCREMENT PRIMARY KEY,
        full_name VARCHAR(100) NOT NULL,
        email VARCHAR(100) UNIQUE NOT NULL,
        password VARCHAR(255) NOT NULL,
        role ENUM('Admin', 'Candidate') DEFAULT 'Candidate',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS resume_analyses (
        id INT AUTO_INCREMENT PRIMARY KEY,
        candidate_name VARCHAR(100) NOT NULL,
        candidate_email VARCHAR(100) NOT NULL,
        source_type VARCHAR(30) NOT NULL,
        recommended_role VARCHAR(100) NOT NULL,
        ats_score FLOAT NOT NULL,
        matched_skills TEXT NOT NULL,
        missing_skills TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
''')

cursor.execute("SELECT COUNT(*) FROM users WHERE email = 'admin@resumematch.ai'")
if cursor.fetchone()[0] == 0:
    cursor.execute(
        "INSERT INTO users (full_name, email, password, role) VALUES (%s, %s, %s, %s)",
        ("Placement Cell Admin", "admin@resumematch.ai", "admin123", "Admin")
    )

conn.commit()
cursor.close()
conn.close()
print("[1/2] MySQL Database 'resume_ai_db' & tables created successfully!")
