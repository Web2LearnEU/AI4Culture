import pandas as pd
from collections import Counter
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string

# Download NLTK resources if not already downloaded
nltk.download("punkt")
nltk.download("stopwords")

# Load the extracted text dataset
input_csv = "extracted_text.csv"
output_txt = "summary.txt"
df = pd.read_csv(input_csv)

# Ensure necessary columns exist
if "File Name" not in df.columns or "Sentence" not in df.columns:
    raise ValueError("Columns 'File Name' and 'Sentence' not found in CSV!")

# Open a text file to save the output
with open(output_txt, "w", encoding="utf-8") as f:

    # Basic Statistics
    total_sentences = df.shape[0]
    unique_files = df["File Name"].nunique()
    avg_sentence_length = df["Sentence"].apply(len).mean()

    summary = (
        f"\n📊 **Basic Dataset Summary:**\n"
        f"🔹 Total Sentences: {total_sentences}\n"
        f"🔹 Unique Files: {unique_files}\n"
        f"🔹 Average Sentence Length: {round(avg_sentence_length, 2)} characters\n"
    )
    print(summary)
    f.write(summary)

    # Breakdown by file
    file_counts = df["File Name"].value_counts()
    file_summary = "\n📂 **Sentences per File:**\n" + file_counts.to_string() + "\n"
    print(file_summary)
    f.write(file_summary)

    # Breakdown by section (if available)
    if "Section" in df.columns:
        section_counts = df["Section"].value_counts()
        section_summary = "\n📜 **Sentences per Section:**\n" + section_counts.head(10).to_string() + "\n"
        print(section_summary)
        f.write(section_summary)

    # Most Common Words (Basic NLP Insight)
    def clean_text(text):
        tokens = word_tokenize(text.lower())  # Convert to lowercase and tokenize
        tokens = [word for word in tokens if word.isalpha()]  # Remove punctuation/numbers
        tokens = [word for word in tokens if word not in stopwords.words("english")]  # Remove stopwords
        return tokens

    # Flatten all sentences into a single list of words
    all_words = [word for sentence in df["Sentence"] for word in clean_text(sentence)]
    word_freq = Counter(all_words)

    word_summary = "\n🔠 **Most Common Words (Excluding Stopwords):**\n"
    word_summary += "\n".join([f"{word}: {freq}" for word, freq in word_freq.most_common(20)]) + "\n"
    print(word_summary)
    f.write(word_summary)

print(f"✅ Summary saved to '{output_txt}'")
