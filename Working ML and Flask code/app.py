from flask import Flask, request, jsonify
import os
import pickle
import numpy as np
from gensim.models import Word2Vec
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import re
import PyPDF2
import nltk

# Download NLTK data
nltk.download('stopwords')
nltk.download('wordnet')

# Initialize Flask app
app = Flask(__name__)

# Load the trained Random Forest model and Word2Vec model
with open("random_forest_model.pkl", 'rb') as file:
    rf_model = pickle.load(file)

word2vec_model = Word2Vec.load("word2vec_model.model")  # Ensure Word2Vec model is saved

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text() if page.extract_text() else ""
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None
    return text if text.strip() else None

# Function to preprocess text
def preprocess_text(text):  # move to outside
    text = text.lower()
    text = re.sub(r'[^a-zA-Z\s]', '', text)  # Remove special characters
    tokens = text.split()  # Tokenize text
    tokens = [word for word in tokens if word not in stopwords.words('english')]  # Remove stopwords
    lemmatizer = WordNetLemmatizer()
    tokens = [lemmatizer.lemmatize(word) for word in tokens]  # Lemmatize tokens
    return tokens

# Function to generate document vector
def document_vector(tokens):
    tokens = [word for word in tokens if word in word2vec_model.wv.key_to_index]
    if not tokens:
        return np.zeros(word2vec_model.vector_size)  # Return zero vector if no tokens match
    return np.mean(word2vec_model.wv[tokens], axis=0)

# Define the /predict endpoint
@app.route('/predict', methods=['POST'])
def predict():
    try:
        # Check if a file is uploaded
        if 'file' not in request.files:
            return jsonify({'error': "No file provided"}), 400
        
        pdf_file = request.files['file']

        # Save the uploaded file temporarily
        temp_path = os.path.join("temp", pdf_file.filename)
        os.makedirs("temp", exist_ok=True)  # Create temp directory if not exists
        pdf_file.save(temp_path)

        # Extract text from the uploaded PDF
        extracted_text = extract_text_from_pdf(temp_path)
        if not extracted_text:
            return jsonify({'error': "Failed to extract text from PDF"}), 400

        # Preprocess the text
        tokens = preprocess_text(extracted_text)

        # Convert to document vector
        vector = document_vector(tokens).reshape(1, -1)  # Ensure it is a 2D array

        # Predict using the loaded model
        prediction = rf_model.predict(vector)[0]  # Get the first prediction

        # Clean up temporary file
        os.remove(temp_path)

        # Map prediction to label
        label_map = {0: "Non-Annual Report", 1: "Annual Report"}
        predicted_label = label_map.get(prediction, "Unknown")

        return jsonify({'prediction': predicted_label}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Run the Flask application
if __name__ == '__main__':
    app.run(debug=True)