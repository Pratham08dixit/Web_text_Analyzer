import pandas as pd
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
import textstat
import spacy

nlp = spacy.load("en_core_web_sm")

def loadwords(file_path):
    try:
        with open(file_path, 'r') as file:
            words = [line.strip().lower() for line in file.readlines()]
        return words
    except Exception as e:
        print(f"Error loading words from {file_path}: {e}")
        return []

def load_stopwords(files):
    stop_words = set()
    for file in files:
        stop_words.update(loadwords(file))
    return stop_words

def extract_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        paragraphs = soup.find_all('p')
        text = ' '.join([para.get_text() for para in paragraphs])
        
        return text
    except Exception as e:
        return f"Error: {e}"

def sentiment_analysis(text, stop_words):
    words = [word.lower() for word in text.split() if word.lower() not in stop_words]
    filtered_text = ' '.join(words)
    blob = TextBlob(filtered_text)
    polarity = blob.sentiment.polarity
    subjectivity = blob.sentiment.subjectivity
    return polarity, subjectivity

def readability_analysis(text):
    fog_index = textstat.gunning_fog(text)
    return fog_index

def word_sentence_analysis(text):
    doc = nlp(text)
    words = [token.text for token in doc if token.is_alpha]  
    sentences = list(doc.sents)  
    avg_sentence_length = len(words) / len(sentences) if sentences else 0
    return len(words), len(sentences), avg_sentence_length

def advanced_analysis(text):
    doc = nlp(text)
    complex_word_count = sum(1 for token in doc if token.is_alpha and len(token) > 6)  
    personal_pronouns = sum(1 for token in doc if token.pos_ == "PRON" and token.tag_ in ["PRP", "PRP$"])
    syllables_per_word = textstat.syllable_count(text) / len(doc) if len(doc) > 0 else 0
    return complex_word_count, personal_pronouns, syllables_per_word

def average_word_length(text):
    doc = nlp(text)
    words = [token.text for token in doc if token.is_alpha]
    avg_length = sum(len(word) for word in words) / len(words) if words else 0
    return avg_length

def sentiment_scores(text, positive_words, negative_words, stop_words):
    words = [token.text.lower() for token in nlp(text) if token.is_alpha and token.text.lower() not in stop_words]
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    return positive_score, negative_score

# Main function to process all URLs
def process_urls(input_file, output_file, pos_words_file, neg_words_file, stop_word_files):
    print("START")
    # Load positive and negative word lists
    positive_words = loadwords(pos_words_file)
    negative_words = loadwords(neg_words_file)

    stop_words = load_stopwords(stop_word_files)

    df = pd.read_excel(input_file)

    urls = df['URL']  
    url_ids = df['URL_ID']  
    results = []

    for url, url_id in zip(urls, url_ids):
        text = extract_text(url)
        if not text.startswith("Error"):
            polarity, subjectivity = sentiment_analysis(text, stop_words)
            fog = readability_analysis(text)
            word_count, sentence_count, avg_sentence_length = word_sentence_analysis(text)
            complex_words, pronouns, syllables_per_word = advanced_analysis(text)
            avg_word_len = average_word_length(text)
            positive_score, negative_score = sentiment_scores(text, positive_words, negative_words, stop_words)
            if word_count>0:
                complex_word_percentage=(complex_words/word_count)*100
            else:
                0
            avg_words_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
            results.append({
                "URL ID": url_id,
                "URL": url,
                "Positive Score": positive_score,
                "Negative Score": negative_score,
                "Polarity": polarity,
                "Subjectivity": subjectivity,
                "Avg Sentence Length": avg_sentence_length,
                "Complex Word Percentage": complex_word_percentage,
                "Fog Index": fog,
                "Avg Words per Sentence": avg_words_per_sentence,
                "Complex Words": complex_words,
                "Total Words Count": word_count,
                "Syllables per Word": syllables_per_word,
                "Personal Pronouns": pronouns,
                "Avg Word Length": avg_word_len,                  
            })
        
    output_df = pd.DataFrame(results)

    output_df.to_excel(output_file, index=False)
    print(f"END")

if __name__ == "__main__":
    input_file = "Input.xlsx" 
    output_file = "output2.xlsx"  
    pos_words_file = "positive-words.txt"  
    neg_words_file = "negative-words.txt"  
    stop_word_files = [  
        "StopWords_Auditor.txt",
        "StopWords_Currencies.txt",
        "StopWords_DatesandNumbers.txt",
        "StopWords_Generic.txt",
        "StopWords_GenericLong.txt",
        "StopWords_Geographic.txt",
        "StopWords_Names.txt",
    ]
    process_urls(input_file, output_file, pos_words_file, neg_words_file, stop_word_files)
