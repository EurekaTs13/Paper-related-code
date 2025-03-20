import re
import os

# Define file paths
txt_file_path = r"D:\数据\new_finance_goaltxt\total_fixname_washed_delete_year2000"
topic_words_path = r"D:\excel\finance\topic_60_0.8_100_iter=5.txt"

def extract_topics_from_file(file_path, num_words=10):
    """
    Extract topics and their associated words from a file.

    Args:
        file_path (str): Path to the file containing topic information.
        num_words (int): Number of words to extract per topic.

    Returns:
        dict: A dictionary where keys are topic IDs and values are lists of words.
    """
    topics = {}

    # Open and read the file
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()  # Read the entire file content

        # Use regex to extract topic information
        # Pattern matches "Topic #<ID>:" followed by topic words
        pattern = r"Topic #(\d+):\s*([^\n]+)"
        matches = re.findall(pattern, content)

        for match in matches:
            topic_id = int(match[0])  # Topic ID
            words = match[1].split()  # Split words by spaces
            topics[topic_id] = words[:num_words]  # Extract the first `num_words` words

    return topics


# Extract topics from the file
topics = extract_topics_from_file(topic_words_path)
print("Extracted topics:", topics)

# Initialize a dictionary to count word occurrences per topic
word_count_dict = {}
for topic_id in range(60):
    word_count_dict[topic_id] = {}
    for word in topics[topic_id]:
        word_count_dict[topic_id][word] = 0  # Initialize word counts to 0

# Process each text file
index = 0
for filename in os.listdir(txt_file_path):
    index += 1
    print(f"Processing file {index}: {filename}")
    txtpath = os.path.join(txt_file_path, filename)
    with open(txtpath, 'r', encoding='utf-8') as file:
        txt = file.read()
        for topic_id in range(60):
            for word in topics[topic_id]:
                # Count occurrences of each word in the document
                word_count_dict[topic_id][word] += txt.split().count(word)

# Print word counts for each topic
for topic_id in word_count_dict.keys():
    print(f"Topic {topic_id}: {word_count_dict[topic_id]}")
    print()  # Add a newline for better readability