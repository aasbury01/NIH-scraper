from flask import Flask, request, render_template_string, send_file, url_for
import pandas as pd
from datetime import datetime
import io

app = Flask(__name__)

# Keyword scores
keyword_scores = {
    "confocal": 5, "microscope": 5, "microscopy": 5, "resolution":5,
    "imaging":4, "expansion":4, "localization":4, "fluorescent":4,
    "motility":3, "throughput":3, "screen":3, "content":3,
    "tissue":1, "cell":1, "electron":-5, "clinical":-4, "patient":-4
}

def calculate_ptscore(pt):
    if pd.isna(pt):
        return 0
    pt = str(pt).lower()
    return sum(points for word, points in keyword_scores.items() if word in pt)

def calculate_abstractscore(abstract):
    if pd.isna(abstract):
        return 0
    abstract = str(abstract).lower()
    return sum(points for word, points in keyword_scores.items() if word in abstract)

# Main HTML page template
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
            <!-- Button to secondary page -->
            <a href="{{ url_for('secondary') }}" class="btn btn-secondary w-100">See how this works!</a>
            {% if message %}
            <div class="alert alert-info mt-3 text-center">{{ message }}</div>
            {% endif %}
        </div>
    </div>
</div>
</body>
</html>
'''

# Secondary page template
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
    <div class="row justify-content-center">
        <div class="col-md-6 p-4 bg-white shadow rounded text-center">
            <h2 class="mb-4">How this program is working:</h2>
            <p>Your download fron NIH reporter is used and scored based upon keyword matching in the Project Title and Abstract. Words found in the project title carry more weight than words found within the project abstract. Right now this list is a defined set of words, but can be expanded. The list of projects is then ranked based upon that keyword score (descending) and spit out into a .csv file that can be uploaded into excel.</p>
            <a href="{{ url_for('index') }}" class="btn btn-primary mt-3">Back to Main Page</a>
        </div>
    </div>
</div>
</body>
</html>
'''

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("excel_file")
        if not file:
            return render_template_string(MAIN_PAGE, message="No file uploaded!")
        
        try:
            df = pd.read_excel(file)
            df['PT score'] = df['Project Title'].apply(calculate_ptscore)
            df['Abstract Score'] = df['Project Abstract'].apply(calculate_abstractscore)
            # making the pt score weigh heavier that the abstract
            PT_WEIGHT = 2
            ABSTRACT_WEIGHT = 1
            df['Combined Score'] = df['PT score'] * PT_WEIGHT + df['Abstract Score'] * ABSTRACT_WEIGHT

            filtered_df = df[df['PT score'] > 0].sort_values(by='Combined Score', ascending=False)

            # Save to in-memory CSV with 4-digit year
            csv_buffer = io.StringIO()
            filtered_df.to_csv(
                csv_buffer,
                columns=['Project Title', 'Organization Name', 'Total Cost', 'Fiscal Year', 'Activity',
                         'Contact PI / Project Leader', 'PT score', 'Abstract Score', 'Combined Score'],
                index=False
            )
            csv_buffer.seek(0)
            csv_filename = f"Scored_report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
            return send_file(
                io.BytesIO(csv_buffer.getvalue().encode()),
                mimetype="text/csv",
                as_attachment=True,
                download_name=csv_filename
            )
        except Exception as e:
            return render_template_string(MAIN_PAGE, message=f"Error processing Excel file: {e}")
    
    return render_template_string(MAIN_PAGE)

@app.route("/secondary")
def secondary():
    return render_template_string(SECONDARY_PAGE)

if __name__ == "__main__":
    app.run(debug=True)