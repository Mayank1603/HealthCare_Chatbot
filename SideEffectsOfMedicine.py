# medicine_chatbot.py

import os
import sys
import json
import pickle
import pandas as pd
import numpy as np
import faiss
from sentence_transformers import SentenceTransformer

# File paths
CSV_PATH = 'c:/Users/mayank.c/OneDrive - Optimus Information Inc/Desktop/Chatbot_Project/Chatbot/MedicalFiles/SideEffectsOFMedicine.csv'
EMBEDDINGS_FILE = 'medicine_embeddings.pkl'

# Load CSV
try:
    df = pd.read_csv(CSV_PATH)
except Exception as e:
    print(json.dumps({"error": f"Failed to load CSV: {str(e)}"}, indent=4))
    sys.exit(1)

# Columns to use for names
name_columns = ['name', 'substitute0', 'substitute1', 'substitute2', 'substitute3', 'substitute4']

# Step 1: Load or generate embeddings
if os.path.exists(EMBEDDINGS_FILE):
    with open(EMBEDDINGS_FILE, 'rb') as f:
        medicine_names, name_embeddings = pickle.load(f)
else:
    print("🔄 Generating embeddings...")

    names = df[name_columns].values.flatten()
    medicine_names = list(set([str(name).strip() for name in names if pd.notna(name)]))

    model = SentenceTransformer('all-MiniLM-L6-v2')
    name_embeddings = model.encode(medicine_names)

    with open(EMBEDDINGS_FILE, 'wb') as f:
        pickle.dump((medicine_names, name_embeddings), f)

    print("✅ Embeddings generated and saved to file.")

# Step 2: Prepare FAISS index
name_embeddings = np.array(name_embeddings).astype('float32')
index = faiss.IndexFlatL2(name_embeddings.shape[1])
index.add(name_embeddings)

# Step 3: Load model (if not loaded yet)
if 'model' not in locals():
    model = SentenceTransformer('all-MiniLM-L6-v2')

# Step 4: Matching function
def get_medicine_details(user_input):
    input_embedding = model.encode([user_input])[0].astype('float32')
    _, top_index = index.search(np.array([input_embedding]), 1)
    best_match = medicine_names[top_index[0][0]]

    # Find row in CSV
    medicine_row = df[
        df['name'].str.contains(best_match, case=False, na=False) |
        df[['substitute0', 'substitute1', 'substitute2', 'substitute3', 'substitute4']].apply(
            lambda x: x.str.contains(best_match, case=False, na=False).any(), axis=1
        )
    ]

    if medicine_row.empty:
        return {
            "status": "not_found",
            "message": "No matching medicine found.",
            "medicine_input": user_input
        }

    # Collect info
    substitutes = medicine_row[name_columns].values.flatten()
    substitutes = list(dict.fromkeys([sub.strip() for sub in substitutes if pd.notna(sub)]))

    side_effect_cols = [col for col in df.columns if col.startswith('sideEffect')]
    side_effects = medicine_row[side_effect_cols].values.flatten()
    side_effects = list(dict.fromkeys([s.strip() for s in side_effects if pd.notna(s)]))

    use_cols = [col for col in df.columns if col.startswith('use')]
    uses = medicine_row[use_cols].values.flatten()
    uses = list(dict.fromkeys([u.strip() for u in uses if pd.notna(u)]))

    return {
        "status": "success",
        "medicine_input": user_input,
        "matched_with": best_match,
        "substitutes": substitutes,
        "side_effects": side_effects,
        "uses": uses
    }

# Step 5: CLI interaction
def chatbot():
    if len(sys.argv) < 2:
        print(json.dumps({
            "error": "No input provided. Please provide the name of the medicine as an argument."
        }, indent=4))
        return

    user_input = " ".join(sys.argv[1:])
    result = get_medicine_details(user_input)
    print(json.dumps(result, indent=4))

if __name__ == "__main__":
    chatbot()
