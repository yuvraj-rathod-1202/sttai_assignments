
# Requirements:
#  - Load /app/outputs/model.pkl with pickle
#  - Load Digits dataset, take the first test sample
#  - Print: predicted class, true class, and whether they match
#  - Also read /app/outputs/meta.json and print the accuracy it was trained with

# YOUR CODE HERE
import os
import pickle
import json
from sklearn.datasets import load_digits
from sklearn.model_selection import train_test_split
# Load model
with open('/app/outputs/model.pkl', 'rb') as f:
    model = pickle.load(f)
# Load test sample
X, y = load_digits(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
test_sample = X_test[0].reshape(1, -1)
true_class = y_test[0]
# Predict
pred_class = model.predict(test_sample)[0]
# Load metadata
with open('/app/outputs/meta.json', 'r') as f:
    meta = json.load(f)
accuracy = meta.get('accuracy', 'N/A')
# Print results
print(f"Predicted class: {pred_class}")
print(f"True class: {true_class}")
print(f"Match: {'Yes' if pred_class == true_class else 'No'}")
print(f"Model was trained with accuracy: {accuracy:.4f}")
