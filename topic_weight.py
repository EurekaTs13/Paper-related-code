from sklearn.feature_extraction.text import CountVectorizer
import os
import pandas as pd
from sklearn.decomposition import LatentDirichletAllocation

# Function to read processed text files
def read_txt_files_to_list(directory_path):
    file_contents = []
    file_names = []
    i = 1
    for filename in os.listdir(directory_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(directory_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                print(f"Processing file {i}")  # Print progress
                i += 1
                file_contents.append(content)
                file_names.append(filename)
    return file_contents, file_names

# Directory containing processed text files
directory_path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"
documents, file_names = read_txt_files_to_list(directory_path)

# Vectorize documents using CountVectorizer
vectorizer = CountVectorizer(max_df=0.8, min_df=100)
X = vectorizer.fit_transform(documents)
print(f"Shape of document-term matrix: {X.shape}")

# Train LDA model
num_topics = 60
lda_model = LatentDirichletAllocation(
    n_components=num_topics,
    max_iter=100,
    learning_method='online',
    random_state=42
)
lda_Z = lda_model.fit_transform(X)

# Function to print top words for each topic
def print_top_words(model, feature_names, n_top_words):
    for topic_idx, topic in enumerate(model.components_):
        print(f"Topic #{topic_idx}:")
        print(" ".join([feature_names[i] for i in topic.argsort()[:-n_top_words - 1:-1]]))
    print()

# Print top 20 words for each topic
tf_feature_names = vectorizer.get_feature_names_out()
print_top_words(lda_model, tf_feature_names, n_top_words=20)

# Create document-topic matrix for visualization
doc_topic_df = pd.DataFrame(
    lda_Z,
    columns=[f'Topic {i}' for i in range(num_topics)],
    index=file_names
)

# Save document-topic matrix to Excel with formatted floating-point numbers
output_excel_path = r"D:\excel\finance\topic_60_0.8_100_iter=5.xlsx"
with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
    doc_topic_df.to_excel(writer, float_format="%.4f")  # Format to 4 decimal places

print(f"Document-topic matrix saved to {output_excel_path}")