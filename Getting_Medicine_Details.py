import pandas as pd
import sys
import json
from sentence_transformers import SentenceTransformer, util
import os
import pickle

# File paths
CSV_PATH = 'c:/Users/mayank.c/OneDrive - Optimus Information Inc/Desktop/Chatbot_Project/Chatbot/MedicalFiles/Medicine_Details.csv'
CACHE_PATH = 'embedding_cache.pkl'

# Load sentence transformer model
model = SentenceTransformer('all-MiniLM-L6-v2')

# Load dataset
df = pd.read_csv(CSV_PATH)
df['search_text'] = df['Medicine Name'].fillna('') + " " + df['Side_effects'].fillna('')

# Load or compute embeddings
def load_or_compute_embeddings():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, 'rb') as f:
            print("[INFO] Loading cached embeddings...")
            df['embedding'] = pickle.load(f)
    else:
        print("[INFO] Computing embeddings for the first time... This might take a moment.")
        df['embedding'] = df['search_text'].apply(lambda x: model.encode(x, convert_to_tensor=True))
        with open(CACHE_PATH, 'wb') as f:
            pickle.dump(df['embedding'].tolist(), f)
        print("[INFO] Embeddings cached successfully!")

# Perform semantic search
def semantic_search(query, top_k=5):
    query_embedding = model.encode(query, convert_to_tensor=True)
    similarities = [float(util.pytorch_cos_sim(query_embedding, emb)) for emb in df['embedding']]
    df['similarity'] = similarities
    return df.sort_values(by='similarity', ascending=False).head(top_k)[['Medicine Name', 'Side_effects']]

# Command-line chatbot interface
def chatbot():
    if len(sys.argv) < 2:
        print(json.dumps({"error": "No input provided. Please provide a medicine name or symptom as an argument."}))
        return

    user_input = " ".join(sys.argv[1:])
    results = semantic_search(user_input)

    if results.empty:
        print(json.dumps({"message": "No relevant results found."}, indent=4))
    else:
        formatted_results = results.to_dict(orient='records')
        print(json.dumps({"semantic_matches": formatted_results}, indent=4))

# Main
if __name__ == "__main__":
    try:
        load_or_compute_embeddings()
        chatbot()
    except Exception as e:
        print(json.dumps({"error": str(e)}))
