
# Requirements:
#  - set_seed(42) called at the top
#  - random_state=42 in all sklearn calls
#  - Train on Digits dataset, Random Forest classifier
#  - Log the run with trackio (init, log, finish)
#     config must include: model, n_estimators, max_depth, random_state, dataset
#     logged metrics: train_accuracy, test_accuracy, n_classes, n_features
#  - Save trained model to /app/outputs/model.pkl
#  - Save a metadata JSON to /app/outputs/meta.json:
#     {"accuracy": ..., "sklearn_version": ..., "seed": 42, "timestamp": ...}
#  - Print a summary at the end

# YOUR CODE HERE
import os
import numpy as np
import sklearn
import pickle
import trackio
from datetime import datetime
from sklearn.datasets import load_digits
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
import json

def set_seed(seed=42):
    np.random.seed(seed)

set_seed(42)

X, y = load_digits(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

model = RandomForestClassifier(n_estimators=100, max_depth=10, random_state=42)

trackio.init(project='cs203-capstone', name='train-digits', config={'model': 'RandomForest', 'n_estimators': 100, 'max_depth': 10, 'random_state': 42, 'dataset': 'digits'})
model.fit(X_train, y_train)

train_acc = model.score(X_train, y_train)
test_acc = model.score(X_test, y_test)

n_classes = len(np.unique(y))
n_features = X.shape[1]
trackio.log({'train_accuracy': float(train_acc), 'test_accuracy': float(test_acc), 'n_classes': n_classes, 'n_features': n_features})

os.makedirs('/app/outputs', exist_ok=True)
with open('/app/outputs/model.pkl', 'wb') as f:
    pickle.dump(model, f)

meta = {
    'accuracy': test_acc,
    'sklearn_version': sklearn.__version__,
    'seed': 42,
    'timestamp': datetime.now().isoformat()
}

with open('/app/outputs/meta.json', 'w') as f:
    json.dump(meta, f, indent=4)

print(f"Training complete! Test accuracy: {test_acc:.4f}, model saved to /app/outputs/model.pkl, metadata saved to /app/outputs/meta.json")
trackio.finish()
