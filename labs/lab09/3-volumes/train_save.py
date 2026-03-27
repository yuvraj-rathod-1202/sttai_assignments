import pickle, os
import sklearn
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier

os.makedirs('/app/outputs', exist_ok=True)

# YOUR CODE HERE
# 1. Load breast_cancer, split (random_state=42)
X, y = load_breast_cancer(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 2. Train RandomForestClassifier(n_estimators=100, random_state=42)
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 3. Save model to /app/outputs/model.pkl using pickle
with open('/app/outputs/model.pkl', 'wb') as f:
    pickle.dump(model, f)   

# 4. Write a log file /app/outputs/training_log.txt containing:
from datetime import datetime
accuracy = model.score(X_test, y_test)
sklearn_version = sklearn.__version__
timestamp = datetime.now().isoformat()
with open('/app/outputs/training_log.txt', 'w') as f:
    f.write(f"accuracy: {accuracy:.4f}")
    f.write(f"sklearn version: {sklearn_version}")
    f.write(f"timestamp: {timestamp}")

# 5. Print "Model and log saved to /app/outputs/"
print("Model and log saved to /app/outputs/")
