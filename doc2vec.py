from gensim.models.doc2vec import Doc2Vec, TaggedDocument
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd
import os
import math

# Load Excel data
df = pd.read_excel(r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx")
filename_column = df.columns[0]
filename_column_list = df[filename_column].tolist()
year_column = df.columns[2]
year_column_list = df[year_column].tolist()

# Initialize document storage
documents = []

# Create filename mapping dictionary
file_path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"
file_dict = {filename: 0 for filename in os.listdir(file_path)}

# Read and process text files
file_counter = 1
for filename in os.listdir(file_path):
    txt_file_path = os.path.join(file_path, filename)
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        txt = file.read()
        documents.append(txt)
        print(f"Processing file {file_counter}")
        file_counter += 1

# Convert documents to TaggedDocument format
tagged_documents = [TaggedDocument(doc.split(), [i]) for i, doc in enumerate(documents)]

# Train Doc2Vec model
model = Doc2Vec(
    tagged_documents,
    vector_size=80,
    window=2,
    min_count=10,
    workers=4,
    epochs=100
)

# Get document vectors
doc_vectors = [model.dv[i] for i in range(len(documents))]

# Calculate cosine similarity matrix
similarity_matrix = cosine_similarity(doc_vectors)

# Print similarity matrix
print("Doc2Vec Similarity Matrix:")
print(similarity_matrix)

# Calculate Shannon entropy
current_row = 0
for filename in os.listdir(file_path):
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
doc2vec_scores = []
for index in range(len(filename_column_list)):
    for filename in os.listdir(file_path):
        parts = filename.split('-')
        # Verify filename pattern matching
        if (parts[0] == str(filename_column_list[index]) and
            int(parts[1]) + 2000 == year_column_list[index]):
            doc2vec_scores.append(file_dict[filename])

# Validate output
print(f"Total entries: {len(doc2vec_scores)}")
print("Document scores:", doc2vec_scores)
print("Sample value:", file_dict["713671-01-000170-0.txt"])

# Save results to Excel
doc2vec_df = pd.DataFrame(doc2vec_scores, columns=['doc2vec'], index=filename_column_list)
output_excel_path = r"D:\excel\doc2vec_scaled_80vs.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    doc2vec_df.to_excel(writer)
print('Process completed successfully')