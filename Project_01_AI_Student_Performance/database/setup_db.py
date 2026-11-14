import mysql.connector
import os

def load_env():
    env_vars = {}
    with open(r"F:\BCA_6th_Sem_Projects\.env", "r", encoding="utf-8-sig") as f:
        for line in f:
            if "=" in line:
                k, v = line.strip().split("=", 1)
                env_vars[k] = v
    return env_vars

env = load_env()

try:
    conn = mysql.connector.connect(
        host=env["DB_HOST"],
        user=env["DB_USER"],
        password=env["DB_PASSWORD"]
    )
    cursor = conn.cursor()
    
    cursor.execute("CREATE DATABASE IF NOT EXISTS student_ai_db")
    cursor.execute("USE student_ai_db")
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            full_name VARCHAR(100) NOT NULL,
            email VARCHAR(100) UNIQUE NOT NULL,
            password VARCHAR(255) NOT NULL,
            role ENUM('Admin', 'Faculty') DEFAULT 'Faculty',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_predictions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            roll_no VARCHAR(30) NOT NULL,
            student_name VARCHAR(100) NOT NULL,
            attendance FLOAT NOT NULL,
            study_hours FLOAT NOT NULL,
            previous_score FLOAT NOT NULL,
            assignment_score FLOAT NOT NULL,
            sleep_hours FLOAT NOT NULL,
            predicted_final_score FLOAT NOT NULL,
            risk_status VARCHAR(30) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM users WHERE email = 'admin@college.edu'")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO users (full_name, email, password, role) VALUES (%s, %s, %s, %s)",
            ("HOD Admin", "admin@college.edu", "admin123", "Admin")
        )
    
    conn.commit()
    cursor.close()
    conn.close()
    print("[SUCCESS] MySQL Connected! 'student_ai_db' database aur tables ban gaye hain!")

except mysql.connector.Error as err:
    print(f"[ERROR] MySQL Connection Failed: {err}")
