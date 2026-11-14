import os
import numpy as np
import pandas as pd
import joblib
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.metrics import r2_score, accuracy_score, mean_absolute_error

print("=== PROJECT 01: AI / ML MODEL TRAINING ENGINE ===")

# 1. Reproducible synthetic academic dataset generate karna (500 Students)
np.random.seed(42)
n_samples = 500

attendance = np.round(np.random.uniform(45, 100, n_samples), 1)
study_hours = np.round(np.random.uniform(0.5, 8.0, n_samples), 1)
previous_score = np.round(np.random.uniform(35, 98, n_samples), 1)
assignment_score = np.round(np.random.uniform(40, 100, n_samples), 1)
sleep_hours = np.round(np.random.uniform(4.0, 9.5, n_samples), 1)

# Realistic academic formula + slight natural noise
sleep_bonus = np.where((sleep_hours >= 6.5) & (sleep_hours <= 8.5), 3.0, -1.5)
noise = np.random.normal(0, 2.5, n_samples)

final_score = (
    0.25 * attendance +
    2.80 * study_hours +
    0.35 * previous_score +
    0.18 * assignment_score +
    sleep_bonus +
    noise
)
final_score = np.round(np.clip(final_score, 20, 99.5), 2)

# Risk Category Label assign karna
def categorize_risk(score, att):
    if score < 45 or att < 60:
        return "Critical"
    elif score < 65 or att < 75:
        return "At-Risk"
    else:
        return "Safe"

risk_status = [categorize_risk(s, a) for s, a in zip(final_score, attendance)]

df = pd.DataFrame({
    "attendance": attendance,
    "study_hours": study_hours,
    "previous_score": previous_score,
    "assignment_score": assignment_score,
    "sleep_hours": sleep_hours,
    "final_score": final_score,
    "risk_status": risk_status
})

os.makedirs("ml_engine", exist_ok=True)
dataset_path = os.path.join("ml_engine", "student_dataset.csv")
df.to_csv(dataset_path, index=False)
print(f"[1/4] Dataset saved: {dataset_path} ({len(df)} student records)")

# 2. Features (X) aur Labels (y) alag karna
X = df[["attendance", "study_hours", "previous_score", "assignment_score", "sleep_hours"]]
y_score = df["final_score"]
y_risk = df["risk_status"]

# 3. 80% Training aur 20% Testing Split
X_train, X_test, y_score_train, y_score_test, y_risk_train, y_risk_test = train_test_split(
    X, y_score, y_risk, test_size=0.20, random_state=42
)
print(f"[2/4] Train-Test Split: {len(X_train)} Training samples | {len(X_test)} Testing samples")

# 4. Random Forest Models Train karna
score_model = RandomForestRegressor(n_estimators=150, random_state=42)
score_model.fit(X_train, y_score_train)

risk_model = RandomForestClassifier(n_estimators=150, random_state=42)
risk_model.fit(X_train, y_risk_train)

# 5. Unseen Test Data par Evaluation (Asli Accuracy check)
score_preds = score_model.predict(X_test)
risk_preds = risk_model.predict(X_test)

r2 = r2_score(y_score_test, score_preds) * 100
mae = mean_absolute_error(y_score_test, score_preds)
clf_acc = accuracy_score(y_risk_test, risk_preds) * 100

print(f"[3/4] Score Regressor R2 Accuracy : {r2:.2f}% (Avg Error: +/- {mae:.2f} marks)")
print(f"[3/4] Risk Classifier Accuracy    : {clf_acc:.2f}%")

# 6. Trained AI Models ko disk par save karna
joblib.dump(score_model, os.path.join("ml_engine", "score_model.pkl"))
joblib.dump(risk_model, os.path.join("ml_engine", "risk_model.pkl"))
print("[4/4] AI Models saved successfully: ml_engine/score_model.pkl & ml_engine/risk_model.pkl")

# 7. Ek Sample Prediction Test
sample_student = pd.DataFrame([[88.0, 5.5, 82.0, 85.0, 7.5]], columns=X.columns)
pred_marks = score_model.predict(sample_student)[0]
pred_risk = risk_model.predict(sample_student)[0]
print(f"\n[DEMO PREDICTION] Student (88% Att, 5.5h Study, 82% Prev) -> Predicted Score: {pred_marks:.2f}% | Status: {pred_risk}")
