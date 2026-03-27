
# env_train.py — reads MODEL_TYPE and N_ESTIMATORS from environment
import os
from sklearn.datasets import load_wine
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.svm import SVC
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline

MODEL_TYPE   = os.environ.get('MODEL_TYPE',   'rf')    # default: Random Forest
N_ESTIMATORS = int(os.environ.get('N_ESTIMATORS', '100'))  # default: 100

# YOUR CODE HERE
# 1. Print f"Model: {MODEL_TYPE}, N_estimators: {N_ESTIMATORS}"
print(f"Model: {MODEL_TYPE}, N_estimators: {N_ESTIMATORS}")

# 2. Load wine, split (random_state=42)
X, y = load_wine(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 3. Pick model based on MODEL_TYPE:
if MODEL_TYPE == 'rf':
    model = RandomForestClassifier(n_estimators=N_ESTIMATORS, random_state=42)
elif MODEL_TYPE == 'gb':
    model = GradientBoostingClassifier(n_estimators=N_ESTIMATORS, random_state=42)
elif MODEL_TYPE == 'svm':
    model = Pipeline([('scaler', StandardScaler()), ('svm', SVC(random_state=42))])
else:
    raise ValueError(f"Unknown model: {MODEL_TYPE}")

# 4. Fit and print accuracy
model.fit(X_train, y_train)
accuracy = model.score(X_test, y_test)
print(f"Accuracy: {accuracy:.2f}")
