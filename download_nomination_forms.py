import requests
from bs4 import BeautifulSoup
import os
import time
import pandas as pd
import mimetypes

# Load URLs from the CSV file
csv_file = "urls.csv"
urls = pd.read_csv(csv_file, header=None)[0].tolist()  # Read single-column CSV

# Create output directory
output_folder = "nomination_forms"
os.makedirs(output_folder, exist_ok=True)

# Base UNESCO URL (for relative links)
base_url = "https://ich.unesco.org"

# Track file numbering
file_counter = 1

def download_nomination_form(page_url, file_number):
    """ Scrapes the nomination form URL and downloads the English version. """
    global file_counter

    try:
        # Request the page
        response = requests.get(page_url)
        if response.status_code != 200:
            print(f"❌ Failed to access {page_url}")
            return

        soup = BeautifulSoup(response.text, "html.parser")

        # Find the nomination file section
        nomination_section = soup.find("div", class_="module-content nomination-file")
        if not nomination_section:
            print(f"⚠️ No nomination file section found for {page_url}")
            return

        # Find the "Nomination form:" <li> tag
        nomination_li = None
        for li in nomination_section.find_all("li"):
            if "Nomination form:" in li.get_text():
                nomination_li = li
                break

        if not nomination_li:
            print(f"⚠️ No 'Nomination form:' entry found for {page_url}")
            return

        # Find the English link inside the <li> tag
        english_link = None
        for a_tag in nomination_li.find_all("a"):
            if "English" in a_tag.get_text():
                english_link = a_tag["href"]
                break

        if not english_link:
            print(f"⚠️ No English nomination form found for {page_url}")
            return

        # Convert relative link to full URL if necessary
        if english_link.startswith("/"):
            english_link = base_url + english_link

        # Download the file
        file_response = requests.get(english_link)
        if file_response.status_code != 200:
            print(f"⚠️ Failed to download file from {english_link}")
            return

        # Guess file extension from response headers
        content_type = file_response.headers.get("Content-Type")
        file_extension = mimetypes.guess_extension(content_type) or ".docx"  # Default to .docx if unknown

        # Define the local filename (numerically ordered)
        file_name = f"{file_number}{file_extension}"
        file_path = os.path.join(output_folder, file_name)

        # Save the file
        with open(file_path, "wb") as f:
            f.write(file_response.content)
        print(f"✅ Downloaded: {file_name} from {page_url}")

    except Exception as e:
        print(f"🚨 Error processing {page_url}: {e}")

# Loop through URLs and download each nomination form
for url in urls:
    download_nomination_form(url, file_counter)
    file_counter += 1
    time.sleep(1)  # Avoid overloading the server

print("🎉 Downloading complete!")