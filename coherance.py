from sklearn.feature_extraction.text import CountVectorizer
import os
from sklearn.decomposition import LatentDirichletAllocation
from gensim.models import CoherenceModel
from gensim.corpora import Dictionary
import matplotlib.pyplot as plt

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

def main():
    directory_path = r"D:\pdf\制造业_清洗"
    documents = read_txt_files_to_list(directory_path)

    # Vectorize using CountVectorizer
    vectorizer = CountVectorizer(max_df=0.8, min_df=50)
    X = vectorizer.fit_transform(documents)
    print(X.shape)

    # Tokenize text using the same tokenizer
    tokenizer = vectorizer.build_tokenizer()
    texts = [tokenizer(doc) for doc in documents]
    dictionary = Dictionary(texts)

    coherence = []

    for i in range(20, 160, 10):
        print(i)  # Print current topic number
        num_topics = i
        lda_model = LatentDirichletAllocation(
            n_components=num_topics,
            max_iter=50,  # Increase the number of iterations
            learning_method='online',
            random_state=42
        )
        lda_model.fit(X)  # Use fit instead of fit_transform to save memory

        tf_feature_names = vectorizer.get_feature_names_out()

        # Extract top 20 keywords for each topic
        lda_topics = []
        for topic_weights in lda_model.components_:
            # lda_model.components_ returns a 2D array: topics × features
            # Each row represents the word weight distribution for the topic
            top_word_indices = topic_weights.argsort()[:-21:-1]  # Get indices of top words
            topic_words = [tf_feature_names[i] for i in top_word_indices]  # Retrieve the top words
            lda_topics.append(topic_words)

        # Calculate coherence score
        coherence_model = CoherenceModel(
            topics=lda_topics,
            texts=texts,
            dictionary=dictionary,
            coherence='c_v'
        )
        cohere_score = coherence_model.get_coherence()
        coherence.append(cohere_score)

    # Plot coherence scores
    plt.plot(range(20, 160, 10), coherence)
    plt.show()
    print(coherence)

if __name__ == '__main__':
    main()