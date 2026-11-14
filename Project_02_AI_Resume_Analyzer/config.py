import os
import mysql.connector

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

def get_db_connection():
    env = load_env()
    return mysql.connector.connect(
        host=env["DB_HOST"],
        user=env["DB_USER"],
        password=env["DB_PASSWORD"],
        database="resume_ai_db"
    )
