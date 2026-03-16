from flask import Flask, request, render_template, send_file
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import json
from sklearn.feature_extraction.text import TfidfVectorizer

app = Flask(__name__)

# Default keyword scores
default_keyword_scores = {
    "confocal": 5, "microscope": 5, "microscopy": 5, "resolution": 5,
    "imaging": 4, "expansion": 4, "localization": 4, "fluorescent": 4,
    "motility": 3, "throughput": 3, "screen": 3, "content": 3,
    "tissue": 1, "cell": 1, "electron": -5, "clinical": -4, "patient": -4
}

KEYWORDS_FILE = "keywords.json"
PROJECTS_FILE = "projects.xlsx"

# Ensure keywords.json exists
if not os.path.exists(KEYWORDS_FILE):
    with open(KEYWORDS_FILE, "w") as f:
        json.dump(default_keyword_scores, f, indent=4)


def load_keywords():
    with open(KEYWORDS_FILE, "r") as f:
        return json.load(f)


def save_keywords(keywords_dict):
    with open(KEYWORDS_FILE, "w") as f:
        json.dump(keywords_dict, f, indent=4)


def calculate_score(text, keywords_dict):
    if pd.isna(text):
        return 0
    text = str(text).lower()
    return sum(points for word, points in keywords_dict.items() if word in text)


# --- Main Upload & Scoring Route ---
@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file:
            message = "No file uploaded!"
            return render_template("index.html", message=message)
        try:
            df = pd.read_excel(file)
            # Save a copy for /keywords TF-IDF reference
            df.to_excel(PROJECTS_FILE, index=False)

            keywords_dict = load_keywords()

            # --- Keyword Scoring ---
            df['PT score'] = df['Project Title'].apply(lambda x: calculate_score(x, keywords_dict))
            df['Abstract Score'] = df['Project Abstract'].apply(lambda x: calculate_score(x, keywords_dict))

            # --- TF-IDF Scoring ---
            combined_text = df[['Project Title', 'Project Abstract']].fillna('').agg(' '.join, axis=1)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform(combined_text.tolist())
            df['TF-IDF Score'] = np.array(tfidf_matrix.sum(axis=1)).flatten()

            # --- Top TF-IDF Keywords per project ---
            feature_names = np.array(vectorizer.get_feature_names_out())
            top_n = 5
            top_keywords_list = []
            for row in tfidf_matrix:
                row_data = row.toarray().flatten()
                if row_data.sum() == 0:
                    top_keywords_list.append("")
                    continue
                top_indices = row_data.argsort()[-top_n:][::-1]
                top_words = feature_names[top_indices]
                top_keywords_list.append(", ".join(top_words))
            df['Top Keywords'] = top_keywords_list

            # --- Combined Weighted Score ---
            PT_WEIGHT = 2
            ABSTRACT_WEIGHT = 1
            TFIDF_WEIGHT = 1
            df['Combined Score'] = (
                df['PT score'] * PT_WEIGHT +
                df['Abstract Score'] * ABSTRACT_WEIGHT +
                df['TF-IDF Score'] * TFIDF_WEIGHT
            )

            # Filter & sort
            filtered_df = df[df['PT score'] > 0].sort_values(by='Combined Score', ascending=False)

            # --- Export to Excel ---
            excel_columns = [
                'Project Title','Contact PI / Project Leader', 'Top Keywords', 'PT score', 'Abstract Score',
                'TF-IDF Score', 'Combined Score', 'Organization Name',
                'Total Cost', 'Fiscal Year', 'Activity', 'Project Abstract'
            ]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                filtered_df.to_excel(writer, index=False, columns=excel_columns, sheet_name='Scored Projects')
                worksheet = writer.sheets['Scored Projects']
                worksheet.set_column('A:A', 40)
                worksheet.set_column('B:B', 40)
                worksheet.set_column('C:F', 15)
                worksheet.set_column('G:K', 25)

            output.seek(0)
            excel_filename = f"Scored_report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"

            return send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=excel_filename
            )

        except Exception as e:
            message = f"Error processing Excel file: {e}"

    return render_template("index.html", message=message)


# --- Keyword Editor
@app.route("/keywords", methods=["GET", "POST"])
def keywords():
    keywords_dict = load_keywords()

    if request.method == "POST":
        new_keywords = {}
        words = request.form.getlist("word")
        scores = request.form.getlist("score")
        for w, s in zip(words, scores):
            if w.strip():
                new_keywords[w.strip()] = int(s)
        save_keywords(new_keywords)
        keywords_dict = new_keywords

    # Compute top TF-IDF keywords if project file exists
    top_tfidf_keywords = []
    if os.path.exists(PROJECTS_FILE):
        try:
            df = pd.read_excel(PROJECTS_FILE)
            combined_text = df[['Project Title','Project Abstract']].fillna('').agg(' '.join, axis=1)
            all_text = " ".join(combined_text)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform([all_text])
            feature_names = vectorizer.get_feature_names_out()
            top_indices = tfidf_matrix.toarray()[0].argsort()[::-1][:5]
            top_tfidf_keywords = [feature_names[i] for i in top_indices]
        except:
            top_tfidf_keywords = []

    return render_template(
        "keywords.html",
        keywords=keywords_dict,
        top_tfidf_keywords=top_tfidf_keywords
    )


# --- Secondary Page ---
@app.route("/secondary")
def secondary():
    return render_template("secondary.html")


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
