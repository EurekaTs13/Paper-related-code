from sklearn.feature_extraction.text import CountVectorizer
import os
import matplotlib.pyplot as plt
from sklearn.decomposition import LatentDirichletAllocation

def read_txt_files_to_list(directory_path):
    file_contents = []
    i = 1
    for filename in os.listdir(directory_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(directory_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                print(i)  # Print current file count
                i += 1
                file_contents.append(content)
    return file_contents

# Read files
directory_path = r"D:\数据\new_finance_goaltxt\total_fixname_washed"
documents = read_txt_files_to_list(directory_path)

# Vectorize using CountVectorizer
vectorizer = CountVectorizer(max_df=0.8, min_df=100)
X = vectorizer.fit_transform(documents)
print(X.shape)
# print(X.toarray())  # Be cautious with toarray() - might consume large memory

# Perplexity calculation
plexs = []
for i in range(20, 160, 10):
    print(f"Processing topics: {i}")
    # Fixed parameter syntax (moved random_state outside the string)
    lda = LatentDirichletAllocation(n_components=i,
                                   max_iter=50,
                                   learning_method='online',
                                   random_state=0)
    lda.fit(X)
    plexs.append(lda.perplexity(X))

# Plot perplexity scores
x = list(range(20, 160, 10))
plt.plot(x, plexs)
plt.xlabel("Number of Topics")
plt.ylabel("Perplexity Score")
plt.title("Perplexity vs Number of Topics")
plt.show()
print("Perplexity values:", plexs)