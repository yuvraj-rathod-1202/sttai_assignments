
# app.py — simple Gradio ML demo
import gradio as gr
import pickle
import numpy as np
from sklearn.datasets import load_breast_cancer
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

# Train at startup (in production you'd load a saved model)
X, y = load_breast_cancer(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)
feature_names = load_breast_cancer().feature_names.tolist()

def predict(*args):
    # YOUR CODE HERE
    # args is a tuple of 30 floats (one per feature)
    # Return a string: "Malignant" or "Benign" with the probability
    input_array = np.array(args).reshape(1, -1)
    proba = model.predict_proba(input_array)[0]
    pred = model.predict(input_array)[0]
    label = "Malignant" if pred == 0 else "Benign"
    return f"{label} (proba: {proba[pred]:.2f})"

# Build Gradio interface with 5 key features (radius_mean, texture_mean,
# perimeter_mean, area_mean, smoothness_mean) as Number inputs
# Use gr.Interface or gr.Blocks
# YOUR CODE HERE
inputs = [gr.Number(label=feature) for feature in ['radius_mean', 'texture_mean', 'perimeter_mean', 'area_mean', 'smoothness_mean']]
interface = gr.Interface(fn=predict, inputs=inputs, outputs="text", title="Breast Cancer Prediction", description="Enter the values of 5 features to predict if the tumor is malignant or benign.")
interface.launch(server_name="0.0.0.0")

# IMPORTANT for Docker:
demo.launch(server_name="0.0.0.0", server_port=7860)
