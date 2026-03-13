from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import io
import os  # for Render PORT

app = Flask(__name__)

# Keyword scores
keyword_scores = {
    "confocal": 5, "microscope": 5, "microscopy": 5, "resolution": 5,
    "imaging": 4, "expansion": 4, "localization": 4, "fluorescent": 4,
    "motility": 3, "throughput": 3, "screen": 3, "content": 3,
    "tissue": 1, "cell": 1, "electron": -5, "clinical": -4, "patient": -4
}

# Function to calculate project title score
def calculate_ptscore(pt):
    if pd.isna(pt):
        return 0
    pt = str(pt).lower()
    return sum(points for word, points in keyword_scores.items() if word in pt)

# Function to calculate abstract score
def calculate_abstractscore(abstract):
    if pd.isna(abstract):
        return 0
    abstract = str(abstract).lower()
    return sum(points for word, points in keyword_scores.items() if word in abstract)

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
            # Calculate scores
            df['PT score'] = df['Project Title'].apply(calculate_ptscore)
            df['Abstract Score'] = df['Project Abstract'].apply(calculate_abstractscore)

            PT_WEIGHT = 2
            ABSTRACT_WEIGHT = 1
            df['Combined Score'] = df['PT score'] * PT_WEIGHT + df['Abstract Score'] * ABSTRACT_WEIGHT

            # Filter projects with PT score > 0 and sort
            filtered_df = df[df['PT score'] > 0].sort_values(by='Combined Score', ascending=False)

            # Save CSV to memory
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
            message = f"Error processing Excel file: {e}"

    return render_template("index.html", message=message)

@app.route("/secondary")
def secondary():
    return render_template("secondary.html")

if __name__ == "__main__":
    # Use Render's dynamic PORT, default to 10000 for local testing
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)
