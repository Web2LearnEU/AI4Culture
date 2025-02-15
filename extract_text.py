import os
import docx
import textract
import pdfminer.high_level
import pandas as pd
import nltk
from nltk.tokenize import sent_tokenize

# Download NLTK sentence tokenizer (only needed once)
nltk.download('punkt')
nltk.download('punkt_tab')

# Folder containing nomination forms
input_folder = "nomination_forms"
output_csv = "extracted_text.csv"
file_data = []

def extract_text_from_docx(file_path):
    """
    Extracts text from DOCX files and splits it into sentences.
    """
    try:
        doc = docx.Document(file_path)
        text = []

        # Extract normal paragraphs
        for para in doc.paragraphs:
            text.append(para.text.strip())

        # Extract text from tables (first column as section headers)
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    text.append(" | ".join(row_text))

        # Convert to sentences
        sentences = sent_tokenize("\n".join(text))
        return sentences
    except Exception as e:
        return [f"Error extracting DOCX: {e}"]

def extract_text_from_doc(file_path):
    """
    Extracts text from older DOC files and cleans unwanted characters.
    """
    try:
        text = textract.process(file_path).decode("utf-8").strip()
        
        # Remove unwanted pipe '|' characters
        text = text.replace("|", " ")  # Replaces '|' with a space

        return sent_tokenize(text)  # Splits into sentences
    except Exception as e:
        return [f"Error extracting DOC: {e}"]

def extract_text_from_pdf(file_path):
    """
    Extracts text from PDF files and splits it into sentences.
    """
    try:
        text = pdfminer.high_level.extract_text(file_path).strip()
        sentences = sent_tokenize(text)
        return sentences
    except Exception as e:
        return [f"Error extracting PDF: {e}"]

# Process each file in the folder
for file_name in os.listdir(input_folder):
    file_path = os.path.join(input_folder, file_name)

    if file_name.endswith(".docx"):
        extracted_sentences = extract_text_from_docx(file_path)
        file_type = "DOCX"
    elif file_name.endswith(".doc"):
        extracted_sentences = extract_text_from_doc(file_path)
        file_type = "DOC"
    elif file_name.endswith(".pdf"):
        extracted_sentences = extract_text_from_pdf(file_path)
        file_type = "PDF"
    else:
        print(f"Skipping unsupported file: {file_name}")
        continue

    # Store sentences as separate rows
    for sentence in extracted_sentences:
        file_data.append({
            "File Name": file_name,
            "File Type": file_type,
            "Sentence": sentence
        })

# Convert to DataFrame and save
df = pd.DataFrame(file_data)
df.to_csv(output_csv, index=False)

print(f"✅ Sentence-level text extraction complete! Results saved to '{output_csv}'")