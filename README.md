# NIH-scraper
A program that takes data from NIHeReporter and ranks  based upon project  relevance

# these are the kewords and their associated scores that are used
keyword_scores = {
    "confocal": 5, "microscope": 5, "microscopy": 5, "resolution":5,
    "imaging":4, "expansion":4, "localization":4, "fluorescent":4,
    "motility":3, "throughput":3, "screen":3, "content":3,
    "tissue":1, "cell":1, "electron":-5, "clinical":-4, "patient":-4
}

# Thse two blocks calculate the score based on how many times a word appears, and adds the point value associated with that word previously defined
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

# checks that the uploaded file is in excel format and alerts the user if no file is uploaded
file = request.files.get("excel_file")
if not file:
            return render_template_string(MAIN_PAGE, message="No file uploaded!")

# reads the uploaded excel file  
  try:
            df = pd.read_excel(file)
            df['PT score'] = df['Project Title'].apply(calculate_ptscore)
            df['Abstract Score'] = df['Project Abstract'].apply(calculate_abstractscore)
            
# creates a formula to make PT score weigh heavier than Abstract score
            PT_WEIGHT = 2
            ABSTRACT_WEIGHT = 1
            df['Combined Score'] = df['PT score'] * PT_WEIGHT + df['Abstract Score'] * ABSTRACT_WEIGHT

            filtered_df = df[df['PT score'] > 0].sort_values(by='Combined Score', ascending=False)

# saves the created dataframe as a downloadable .csv with specified colums ranked based upon combined score
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