from flask import Flask, request, render_template_string, send_file
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


# --- HTML Templates ---
MAIN_PAGE = '''
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Excel Scoring Tool</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-5">
  <div class="row justify-content-center">
    <div class="col-md-6 p-4 bg-white shadow rounded">
      <h2 class="mb-4 text-center">Upload Excel File to Score Projects</h2>
      <p>Please upload an Excel file (.xlsx or .xls). CSV is not supported yet.</p>
      <form method="post" enctype="multipart/form-data">
        <input class="form-control mb-3" type="file" name="excel_file" accept=".xlsx,.xls" required>
        <button class="btn btn-primary w-100 mb-2" type="submit">Upload & Score</button>
      </form>
      <a href="/secondary" class="btn btn-secondary w-100 mb-2">See how this works!</a>
      <a href="/keywords" class="btn btn-info w-100">Edit Keyword Scores</a>
      {% if message %}
      <div class="alert alert-info mt-3 text-center">{{ message }}</div>
      {% endif %}
    </div>
  </div>
</div>
</body>
</html>
'''

SECONDARY_PAGE = '''
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>How It Works</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-5">
  <div class="row justify-content-center">
    <div class="col-md-8 p-4 bg-white shadow rounded text-center">
      <h2 class="mb-4">How This Program Works</h2>
      <p>Your NIH Reporter Excel file is scored based on keyword matches in the Project Title and Abstract.</p>
      <p>Project titles carry more weight than abstracts. Keyword scores can be adjusted via the "Edit Keyword Scores" page. The projects are then ranked by combined score and exported to Excel.</p>
      <a href="/" class="btn btn-primary mt-3">Back to Main Page</a>
    </div>
  </div>
</div>
</body>
</html>
'''

KEYWORDS_PAGE = '''
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Keyword Editor</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #ddd;padding:8px;text-align:left;}
th{background:#f4f4f4;}
tr:nth-child(even){background:#fafafa;}
button{padding:6px 10px;border:none;background:#007BFF;color:white;border-radius:4px;cursor:pointer;}
button:hover{background:#0056b3;}
input{padding:5px;}
</style>
</head>
<body class="bg-light">
<div class="container py-5">
  <h2>Keyword Score Editor</h2>
  <form method="POST">
    <table>
      <tr>
        <th>Keyword</th>
        <th>Score</th>
        <th>Delete</th>
      </tr>
      {% for word, score in keywords.items() %}
      <tr>
        <td><input name="word" value="{{ word }}"></td>
        <td><input name="score" type="number" value="{{ score }}" style="width:70px"></td>
        <td><button type="button" onclick="deleteRow(this)">Remove</button></td>
      </tr>
      {% endfor %}
    </table>
    <br>
    <button type="button" onclick="addRow()">Add Keyword</button>
    <button type="submit">Save Changes</button>
  </form>
  {% if top_tfidf_keywords %}
  <div class="mt-3"><strong>Top TF-IDF keywords in uploaded projects:</strong> {{ top_tfidf_keywords|join(', ') }}</div>
  {% endif %}
  <p class="mt-3"><a href="/">⬅ Back to Upload</a></p>
</div>

<script>
function addRow(){
  let table=document.querySelector("table");
  let row=table.insertRow();
  row.innerHTML=`<td><input name="word"></td>
                 <td><input name="score" type="number" value="0"></td>
                 <td><button type="button" onclick="deleteRow(this)">Remove</button></td>`;
}
function deleteRow(btn){ btn.parentElement.parentElement.remove(); }
</script>
</body>
</html>
'''


# --- Routes ---
@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file:
            message = "No file uploaded!"
            return render_template_string(MAIN_PAGE, message=message)
        try:
            df = pd.read_excel(file)
            df.to_excel(PROJECTS_FILE, index=False)  # save for /keywords TF-IDF

            keywords_dict = load_keywords()

            df['PT score'] = df['Project Title'].apply(lambda x: calculate_score(x, keywords_dict))
            df['Abstract Score'] = df['Project Abstract'].apply(lambda x: calculate_score(x, keywords_dict))

            PT_WEIGHT = 2
            ABSTRACT_WEIGHT = 1
            TFIDF_WEIGHT = 1

            # TF-IDF
            combined_text = df[['Project Title','Project Abstract']].fillna('').agg(' '.join, axis=1)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform(combined_text.tolist())
            df['TF-IDF Score'] = np.array(tfidf_matrix.sum(axis=1)).flatten()

            # Top keywords per project
            feature_names = np.array(vectorizer.get_feature_names_out())
            top_n = 5
            top_keywords_list = []
            for row in tfidf_matrix:
                row_data = row.toarray().flatten()
                if row_data.sum() == 0:
                    top_keywords_list.append("")
                    continue
                top_indices = row_data.argsort()[-top_n:][::-1]
                top_keywords_list.append(", ".join(feature_names[top_indices]))
            df['Top Keywords'] = top_keywords_list

            df['Combined Score'] = (
                df['PT score'] * PT_WEIGHT +
                df['Abstract Score'] * ABSTRACT_WEIGHT +
                df['TF-IDF Score'] * TFIDF_WEIGHT
            )

            filtered_df = df[df['PT score'] > 0].sort_values(by='Combined Score', ascending=False)

            # Export Excel
            excel_columns = [
                'Project Title','Contact PI / Project Leader','Top Keywords','PT score','Abstract Score',
                'TF-IDF Score','Combined Score','Organization Name',
                'Total Cost','Fiscal Year','Activity','Project Abstract'
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
    return render_template_string(MAIN_PAGE, message=message)


@app.route("/keywords", methods=["GET","POST"])
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

    return render_template_string(KEYWORDS_PAGE, keywords=keywords_dict, top_tfidf_keywords=top_tfidf_keywords)


@app.route("/secondary")
def secondary():
    return render_template_string(SECONDARY_PAGE)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)
