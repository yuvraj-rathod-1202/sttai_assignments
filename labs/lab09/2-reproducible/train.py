
# 1. Import sklearn, numpy
import numpy as np
import sklearn
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split

# 2. set random_state=42 everywhere
np.random.seed(42)

# 3. Load breast_cancer dataset
X, y = load_breast_cancer(return_X_y=True)

# 4. Train-test split (test_size=0.2, random_state=42)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 5. Train RandomForestClassifier(n_estimators=100, random_state=42)
model = sklearn.ensemble.RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 6. Print accuracy rounded to 4 decimal places
print(f"Accuracy: {model.score(X_test, y_test):.4f}")

# 7. Print sklearn version used
print(f"sklearn version: {sklearn.__version__}")
