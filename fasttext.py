from gensim.models import FastText
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import pandas as pd
import math
import os

# Initialize list to store documents
documents = []

# Load Excel data
df = pd.read_excel(r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx")
filename_column = df.columns[0]
filename_column_list = df[filename_column].tolist()
year_column = df.columns[2]
year_column_list = df[year_column].tolist()

# Define path to text files
path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"
file_dict = {filename: 0 for filename in os.listdir(path)}

# Read and process text files
file_counter = 1
for filename in os.listdir(path):
    txt_file_path = os.path.join(path, filename)
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        txt = file.read()
        documents.append(txt)
        print(f"Processing file {file_counter}")
        file_counter += 1

# Train FastText model
tokenized_documents = [doc.split() for doc in documents]
model = FastText(vector_size=50, window=2, min_count=10, workers=4, epochs=100)
model.build_vocab(tokenized_documents)
model.train(tokenized_documents, total_examples=len(tokenized_documents), epochs=model.epochs)

# Get document vectors by averaging word vectors
def document_vector(doc):
    return np.mean([model.wv[word] for word in doc.split() if word in model.wv], axis=0)

doc_vectors = [document_vector(doc) for doc in documents]

# Calculate similarity matrix
similarity_matrix = cosine_similarity(doc_vectors)

# Print similarity matrix
print("FastText Similarity Matrix:")
print(similarity_matrix)

# Calculate Shannon entropy
current_row = 0
for filename in os.listdir(path):
    for column in range(similarity_matrix.shape[1]):
        sim_value = similarity_matrix[current_row][column]
        if sim_value > 0:
            # Shannon entropy calculation with scaling factor
            file_dict[filename] += -(sim_value * math.log(sim_value)) / 9.35
        else:
            # Handle non-positive values
            file_dict[filename] += 0
    current_row += 1

print(file_dict)

# Match documents with Excel entries
fasttext_scores = []
for index in range(len(filename_column_list)):
    for filename in os.listdir(path):
        parts = filename.split('-')
        # Verify filename pattern matching
        if (parts[0] == str(filename_column_list[index]) and
            int(parts[1]) + 2000 == year_column_list[index]):
            fasttext_scores.append(file_dict[filename])

# Validate output
print(f"Total entries: {len(fasttext_scores)}")
print("Document scores:", fasttext_scores)
print("Sample value:", file_dict["713671-01-000170-0.txt"])

# Save results to Excel
fasttext_df = pd.DataFrame(fasttext_scores, columns=['fasttext'], index=filename_column_list)
output_excel_path = r"D:\excel\fasttext.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    fasttext_df.to_excel(writer)
print('Process completed successfully')