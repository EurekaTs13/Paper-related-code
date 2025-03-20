import os
import pandas as pd
import math

# Read the Excel file
df = pd.read_excel(r"D:\excel\Firm_SystemRisk_Finance-20230101_modified.xlsx")
filename_column = df.columns[0]
filename_column_list = df[filename_column].tolist()
year_column = df.columns[2]
year_column_list = df[year_column].tolist()

# Read the documents
path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"
documents = []
filename_list = []
a = 0
for filename in os.listdir(path):
    txt_file_path = os.path.join(path, filename)
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        txt = file.read()
        documents.append(txt)
        filename_list.append(filename)
        print(a)
        a += 1

# Calculate the word overlap rate between two documents
def calculate_word_overlap(doc1, doc2):
    common_words = doc1.intersection(doc2)
    total_words = doc1.union(doc2)
    overlap_rate = len(common_words) / len(total_words) if len(total_words) > 0 else 0
    return overlap_rate

# Create the word overlap matrix
num_docs = len(documents)
overlap_matrix = [[0 for _ in range(num_docs)] for _ in range(num_docs)]

words_doc = []
for i in range(num_docs):
    print('\n', i)
    words = documents[i].split()  # Split the document into words
    words_doc.append(set(words))  # Add the set of words to the list

for i in range(num_docs):
    print(i)
    for j in range(i, num_docs):
        overlap_rate = calculate_word_overlap(words_doc[i], words_doc[j])
        overlap_matrix[i][j] = overlap_rate
        overlap_matrix[j][i] = overlap_matrix[i][j]

# Calculate the similarity score for each document
dict_scores = {filename: 0 for filename in filename_list}

for i, filename in enumerate(filename_list):
    print(i)
    for j in range(num_docs):
        sim_value = overlap_matrix[i][j]
        if sim_value > 0:
            dict_scores[filename] += -(sim_value * math.log(sim_value)) / 9.35

# Print the similarity scores dictionary
print(dict_scores)

# Save the similarity scores in the order of the Excel file into a list
wordoverlap = []
for index in range(len(filename_column_list)):
    for filename in filename_list:
        parts = filename.split('-')
        if parts[0] == str(filename_column_list[index]) and int(parts[1]) + 2000 == year_column_list[index]:
            wordoverlap.append(dict_scores[filename])

# Print the results
print(len(wordoverlap))
print(wordoverlap)
print(dict_scores.get("713671-01-000170-0.txt"))

# Save the results to an Excel file
wordoverlap_df = pd.DataFrame(wordoverlap, columns=['wordoverlap'], index=filename_column_list)
output_excel_path = r"D:\excel\wordoverlap.xlsx"
wordoverlap_df.to_excel(output_excel_path, engine='xlsxwriter')
print('done')