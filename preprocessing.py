import os
import spacy
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import string

# Load the pre-trained model
nlp = spacy.load('en_core_web_sm')

# English stopwords
stop_words = set(stopwords.words('english'))

def clean_text(text):
    # 1. Convert to lowercase
    text = text.lower()
    # 2. Remove punctuation and digits
    text = ''.join([char for char in text if char not in string.punctuation and not char.isdigit()])
    # 3. Tokenize the text
    words = word_tokenize(text)
    # 4. Remove stopwords
    words = [word for word in words if word not in stop_words]
    # 5. Use spaCy for part-of-speech tagging and retain only nouns
    doc = nlp(' '.join(words))
    nouns = [token.lemma_ for token in doc if token.pos_ == 'NOUN']

    return ' '.join(nouns)

# Read text files, clean them, and save the results
def wash_txt(source_dir, destination_dir):
    i = 1
    # Iterate through all files in the source directory
    for filename in os.listdir(source_dir):
        source_file_path = os.path.join(source_dir, filename)
        with open(source_file_path, 'r', encoding='utf-8') as file:
            text = file.read()

        cleaned_text = clean_text(text)

        destination_file_path = os.path.join(destination_dir, filename)
        with open(destination_file_path, 'w', encoding='utf-8') as file:
            file.write(cleaned_text)
        print(i)
        i += 1

input_path = r"D:\数据\new_finance_goaltxt\total_fixname"
output_path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"

wash_txt(input_path, output_path)
print(f"Cleaned text saved to: {output_path}")