**Web Text Analyzer**



This project is a web text analysis tool that extracts textual content from web pages and performs various analyses, including sentiment analysis, readability assessment, and linguistic analysis. The results are stored in an Excel file for further processing.
__Features Include:__



-> Extracts text from web pages using BeautifulSoup

-> Sentiment analysis using TextBlob

-> Readability analysis using textstat

-> Linguistic analysis with spaCy

-> Counts positive and negative words

-> Computes metrics like average word length, sentence length, and syllables per word

-> Outputs results to an Excel file


**Input.xlsx should contain a column named URL_ID and URL.**


**The script generates output.xlsx with columns like:**

  -> URL ID, URL, Positive Score, Negative Score, Polarity, Subjectivity

  -> Average Sentence Length, Complex Word Percentage, Fog Index

  -> Total Words Count, Syllables per Word, Personal Pronouns, etc.
