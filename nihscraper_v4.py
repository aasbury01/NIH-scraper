from flask import Flask, request, render_template_string, send_file, url_for
import pandas as pd
import numpy as np
from datetime import datetime
import io
import json
from sklearn.feature_extraction.text import TfidfVectorizer

app = Flask(__name__)

# ---------------- Global storage ----------------
latest_df = None  # Stores last uploaded DataFrame
KEYWORDS_FILE = "keywords.json"

# Default keywords
default_keywords = {
    "confocal": 5, "microscope": 5, "microscopy": 5, "resolution":5,
    "imaging":4, "expansion":4, "localization":4, "fluorescent":4,
    "motility":3, "throughput":3, "screen":3, "content":3,
    "tissue":1, "cell":1, "electron":-5, "clinical":-4, "patient":-4
}

# Ensure keywords.json exists
try:
    with open(KEYWORDS_FILE,"x") as f:
        json.dump(default_keywords,f,indent=4)
except FileExistsError:
    pass

# ---------------- Helper functions ----------------

def load_keywords():
    with open(KEYWORDS_FILE,"r") as f:
        return json.load(f)

def save_keywords(keywords_dict):
    with open(KEYWORDS_FILE,"w") as f:
        json.dump(keywords_dict,f,indent=4)

def calculate_score(text, keywords_dict):
    if pd.isna(text):
        return 0
    text = str(text).lower()
    return sum(points for word, points in keywords_dict.items() if word in text)

# ---------------- Templates ----------------
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
<form method="post" enctype="multipart/form-data">
<input class="form-control mb-3" type="file" name="excel_file" accept=".xlsx,.xls" required>
<button class="btn btn-primary w-100 mb-2" type="submit">Upload & Score</button>
</form>
<a href="{{ url_for('keywords') }}" class="btn btn-secondary w-100 mb-2">Edit Keyword Scores</a>
<a href="{{ url_for('secondary') }}" class="btn btn-info w-100">See how this works!</a>
{% if message %}<div class="alert alert-info mt-3 text-center">{{ message }}</div>{% endif %}
</div></div></div></body></html>
'''

SECONDARY_PAGE = '''
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Secondary Page</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-5">
<h2>How this program works</h2>
<p>Projects are scored based on keyword matches and TF-IDF analysis of the Project Title and Abstract. Project Title words weigh more than Abstract words. The final combined score ranks the projects. You can edit keywords dynamically.</p>
<a href="{{ url_for('index') }}" class="btn btn-primary">Back to Main Page</a>
</div></body>
</html>
'''

KEYWORDS_PAGE = '''
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Keyword Editor</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<style>
table{width:100%; border-collapse: collapse; margin-bottom: 15px;}
th, td{border:1px solid #dee2e6; padding:8px;}
tr:nth-child(even){background:#f8f9fa;}
input{width:100%; padding:5px;}
.btn-add{background:#28a745; color:white;}
.btn-remove{background:#dc3545; color:white;}
.btn-save{background:#007bff; color:white;}
.btn-add:hover{background:#218838;}
.btn-remove:hover{background:#c82333;}
.btn-save:hover{background:#0056b3;}
</style>
</head>
<body class="p-4 bg-light">
<nav class="mb-4">
<a href="{{ url_for('index') }}">Home</a> |
<a href="{{ url_for('keywords') }}">Keyword Editor</a> |
<a href="{{ url_for('secondary') }}">Secondary Page</a>
</nav>

<h2>Keyword Score Editor</h2>
{% if top_tfidf_keywords %}
<div class="alert alert-info">
<h5>Top 5 TF-IDF Keywords from last upload</h5>
<ul>{% for word in top_tfidf_keywords %}<li>{{ word }}</li>{% endfor %}</ul>
</div>
{% endif %}

<form method="POST">
<table class="table">
<thead><tr><th>Keyword</th><th>Score</th><th>Delete</th></tr></thead>
<tbody id="keyword-table-body">
{% for word, score in keywords.items() %}
<tr>
<td><input name="word" value="{{ word }}"></td>
<td><input name="score" type="number" value="{{ score }}" style="width:80px"></td>
<td><button type="button" class="btn btn-remove" onclick="deleteRow(this)">Remove</button></td>
</tr>
{% endfor %}
</tbody>
</table>
<button type="button" class="btn btn-add me-2" onclick="addRow()">Add Keyword</button>
<button type="submit" class="btn btn-save">Save Changes</button>
</form>

<script>
function addRow(){
    let tbody=document.getElementById("keyword-table-body");
    let row=document.createElement("tr");
    row.innerHTML=`<td><input name="word"></td>
    <td><input name="score" type="number" value="0" style="width:80px"></td>
    <td><button type="button" class="btn btn-remove" onclick="deleteRow(this)">Remove</button></td>`;
    tbody.appendChild(row);
}
function deleteRow(btn){btn.closest("tr").remove();}
</script>
</body>
</html>
'''

# ---------------- Routes ----------------

@app.route("/", methods=["GET","POST"])
def index():
    global latest_df
    message = None
    if request.method=="POST":
        file=request.files.get("excel_file")
        if not file:
            return render_template_string(MAIN_PAGE,message="No file uploaded!")
        try:
            df=pd.read_excel(file)
            latest_df = df.copy()
            keywords_dict = load_keywords()
            df['PT score'] = df['Project Title'].apply(lambda x: calculate_score(x,keywords_dict))
            df['Abstract Score'] = df['Project Abstract'].apply(lambda x: calculate_score(x,keywords_dict))
            # TF-IDF
            combined_text = df[['Project Title','Project Abstract']].fillna('').agg(' '.join,axis=1)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform(combined_text.tolist())
            df['TF-IDF Score'] = np.array(tfidf_matrix.sum(axis=1)).flatten()
            feature_names = np.array(vectorizer.get_feature_names_out())
            top_n = 5
            top_keywords_list = []
            for row in tfidf_matrix:
                row_data = row.toarray().flatten()
                if row_data.sum()==0: top_keywords_list.append(""); continue
                top_indices = row_data.argsort()[-top_n:][::-1]
                top_words = feature_names[top_indices]
                top_keywords_list.append(", ".join(top_words))
            df['Top Keywords'] = top_keywords_list
            # Combined Score
            df['Combined Score'] = df['PT score']*2 + df['Abstract Score']*1 + df['TF-IDF Score']*1
            filtered_df = df[df['PT score']>0].sort_values('Combined Score',ascending=False)
            # Excel export
            excel_columns=['Project Title','Contact PI / Project Leader','Top Keywords','PT score','Abstract Score',
            'TF-IDF Score','Combined Score','Organization Name','Total Cost','Fiscal Year','Activity','Project Abstract']
            output=io.BytesIO()
            with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
                filtered_df.to_excel(writer,index=False,columns=excel_columns,sheet_name='Scored Projects')
                worksheet=writer.sheets['Scored Projects']
                worksheet.set_column('A:A',40)
                worksheet.set_column('B:B',40)
                worksheet.set_column('C:F',15)
                worksheet.set_column('G:K',25)
            output.seek(0)
            excel_filename=f"Scored_report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
            return send_file(output,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             as_attachment=True,download_name=excel_filename)
        except Exception as e:
            message=f"Error processing Excel file: {e}"
    return render_template_string(MAIN_PAGE,message=message)

@app.route("/keywords",methods=["GET","POST"])
def keywords():
    keywords_dict=load_keywords()
    if request.method=="POST":
        new_keywords={}
        words=request.form.getlist("word")
        scores=request.form.getlist("score")
        for w,s in zip(words,scores):
            if w.strip(): new_keywords[w.strip()]=int(s)
        save_keywords(new_keywords)
        keywords_dict=new_keywords
    # TF-IDF preview in-memory
    top_tfidf_keywords=[]
    if latest_df is not None:
        try:
            combined_text = latest_df[['Project Title','Project Abstract']].fillna('').agg(' '.join,axis=1)
            all_text = " ".join(combined_text)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform([all_text])
            feature_names = vectorizer.get_feature_names_out()
            top_indices = tfidf_matrix.toarray()[0].argsort()[::-1][:5]
            top_tfidf_keywords = [feature_names[i] for i in top_indices]
        except:
            top_tfidf_keywords = []
    return render_template_string(KEYWORDS_PAGE,keywords=keywords_dict,top_tfidf_keywords=top_tfidf_keywords)

@app.route("/secondary")
def secondary():
    return render_template_string(SECONDARY_PAGE)

if __name__=="__main__":
    app.run(debug=True)