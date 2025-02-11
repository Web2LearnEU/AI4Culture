import os
import pdfminer.high_level
import docx
import textract
import pandas as pd

# Folder containing the nomination files
input_folder = "nomination_forms"

# Output CSV file
output_csv = "file_text_vs_scanned.csv"

# Function to check if a PDF is text-based
def is_text_pdf(file_path):
    try:
        text = pdfminer.high_level.extract_text(file_path)
        return bool(text.strip())  # True if text exists, False if empty
    except Exception as e:
        print(f"Error processing PDF {file_path}: {e}")
        return False

# Function to check if a DOCX file is text-based
def is_text_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs])
        return bool(text.strip())  # True if text exists, False if empty
    except Exception as e:
        print(f"Error processing DOCX {file_path}: {e}")
        return False

# Function to check if a DOC file is text-based
def is_text_doc(file_path):
    try:
        text = textract.process(file_path).decode("utf-8")
        return bool(text.strip())  # True if text exists, False if empty
    except Exception as e:
        print(f"Error processing DOC {file_path}: {e}")
        return False

# Process all files
results = []
for file_name in os.listdir(input_folder):
    file_path = os.path.join(input_folder, file_name)

    if file_name.endswith(".pdf"):
        text_based = is_text_pdf(file_path)
    elif file_name.endswith(".docx"):
        text_based = is_text_docx(file_path)
    elif file_name.endswith(".doc"):
        text_based = is_text_doc(file_path)
    else:
        print(f"Skipping unsupported file: {file_name}")
        continue

    # Store the result
    results.append({"File Name": file_name, "Text-Based": "Yes" if text_based else "No"})

# Convert results to DataFrame
df = pd.DataFrame(results)

# Save results as CSV
df.to_csv(output_csv, index=False)

print(f"✅ File check complete! Results saved to {output_csv}")