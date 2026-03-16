# NIH-scraper
A program that takes data from NIHeReporter and ranks based upon project relevance. Fill out the state category and select Active Projects from the 'Fiscal Year' dropdown. Export your search from NIH Reporter with an excel sheet and include Project Abstract. This model uses a combination of TF-IDF and keyword scoring to rank projects uploaded from NIHeReporter.

This can be run broswer based - currently hosted on render - or locally using the 'nihscraper_v4.py' file. Make sure you have all necessary pacakges installed.

# Current file size limit is 5MB

# these are the default kewords and their associated scores that are used
    keyword_scores = {
        "confocal": 5, "microscope": 5, "microscopy": 5, "resolution":5,
        "imaging":4, "expansion":4, "localization":4, "fluorescent":4,
        "motility":3, "throughput":3, "screen":3, "content":3,
        "tissue":1, "cell":1, "electron":-5, "clinical":-4, "patient":-4
    }

# Keyword scoring calculates the score based on how many times a word appears, and adds the point value associated with that word previously defined.
            # --- Keyword scoring
             df['PT score'] = df['Project Title'].apply(lambda x: calculate_score(x, keywords_dict))
            df['Abstract Score'] = df['Project Abstract'].apply(lambda x: calculate_score(x, keywords_dict))

# TF-IDF scoring uses a model that helps identify relevant and distinct words by comparing it to the other project titles and project abstracts.
            # --- TF-IDF Scoring
            combined_text = df[['Project Title', 'Project Abstract']].fillna('').agg(' '.join, axis=1)
            vectorizer = TfidfVectorizer(stop_words='english')
            tfidf_matrix = vectorizer.fit_transform(combined_text.tolist())
            df['TF-IDF Score'] = np.array(tfidf_matrix.sum(axis=1)).flatten()


# the program will spit back an excel file that can then be opened and further refined. 
