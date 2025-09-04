from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd
import sys
import json

# Load the CSV file into a pandas DataFrame
df = pd.read_csv('c:/Users/mayank.c/OneDrive - Optimus Information Inc/Desktop/Chatbot_Project/Chatbot/MedicalFiles/medquad.csv')

# Load a pre-trained sentence transformer model
model = SentenceTransformer('all-MiniLM-L6-v2')

# Embed all focus_area entries in advance
focus_embeddings = model.encode(df['focus_area'].fillna("").tolist(), show_progress_bar=False)

def get_answers_by_question(user_question):
    # Get embedding for the user question
    question_embedding = model.encode([user_question])[0]

    # Calculate cosine similarity with all focus areas
    similarities = cosine_similarity([question_embedding], focus_embeddings)[0]

    # Get top N results (e.g., top 3)
    top_indices = similarities.argsort()[-3:][::-1]

    top_results = df.iloc[top_indices].copy()
    top_results['similarity'] = similarities[top_indices]

    if not top_results.empty:
        formatted_results = top_results.to_dict(orient='records')
        return {
            "status": "success",
            "message": "Top matches based on your question:",
            "results": formatted_results
        }
    else:
        return {
            "status": "no_match",
            "message": "Sorry, I couldn't find anything related."
        }

# Chatbot interface
def chatbot():
    if len(sys.argv) < 2:
        print(json.dumps({"error": "No input provided. Please provide a question as an argument."}, indent=4))
        return

    user_input = " ".join(sys.argv[1:])
    response = get_answers_by_question(user_input)
    print(json.dumps(response, indent=4))

if __name__ == "__main__":
    chatbot()
