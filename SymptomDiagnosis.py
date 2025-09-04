from fuzzywuzzy import process
import pandas as pd
import sys
import json
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder

# Load the dataset
training = pd.read_csv('c:/Users/mayank.c/OneDrive - Optimus Information Inc/Desktop/Chatbot_Project/Chatbot/MedicalFiles/Training.csv')

# Preprocess the data
cols = training.columns[:-1]
training['prognosis'] = training['prognosis'].str.strip()

# Encode the prognosis column (disease labels)
label_encoder = LabelEncoder()
training['prognosis'] = label_encoder.fit_transform(training['prognosis'])

# Train a RandomForest model for disease prediction
X = training[cols]
y = training['prognosis']
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Initialize and train the model
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# Function to match symptoms using fuzzy matching
def match_symptoms(user_symptoms, dataset_symptoms):
    matched_symptoms = []
    for symptom in user_symptoms:
        match = process.extractOne(symptom, dataset_symptoms)
        if match and match[1] > 70:
            matched_symptoms.append(match[0])
    return matched_symptoms

# Function to predict the disease using the trained ML model
def predict_disease(matched_symptoms, dataset):
    if not matched_symptoms:
        return {
            "status": "no_match",
            "message": "No matching symptoms found. Please consult a doctor for further evaluation.",
            "matched_symptoms": []
        }

    # Convert matched symptoms to a binary vector
    symptom_vector = [1 if symptom in matched_symptoms else 0 for symptom in cols]
    
    # Predict disease using the trained model
    predicted_disease_index = model.predict([symptom_vector])[0]
    predicted_disease = label_encoder.inverse_transform([predicted_disease_index])[0]

    return {
        "status": "success",
        "predicted_disease": predicted_disease,
        "matched_symptoms": matched_symptoms,
        "message": f"Based on your symptoms, you may have {predicted_disease}."
    }

# Main function to handle user input and run the prediction
def main():
    if len(sys.argv) < 2:
        print(json.dumps({
            "error": "No input provided. Please provide symptoms as a comma-separated list."
        }, indent=4))
        return

    user_input = " ".join(sys.argv[1:])
    user_symptoms = [symptom.strip() for symptom in user_input.split(',')]
    dataset_symptoms = list(cols)
    matched_symptoms = match_symptoms(user_symptoms, dataset_symptoms)
    result = predict_disease(matched_symptoms, training)

    print(json.dumps(result, indent=4))

# Run the main function
if __name__ == "__main__":
    main()
