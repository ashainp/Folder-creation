import os
import docx
import logging
from pathlib import Path

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to extract text from a Word document
def extract_text_from_docx(doc_path):
    try:
        doc = docx.Document(doc_path)
        return [paragraph.text for paragraph in doc.paragraphs]
    except Exception as e:
        logging.error(f"Error reading the document: {e}")
        return []

# Specify the path to the Word document and the base directory
doc_path = "G:/Folders.docx"
base_directory = "G:/"

# Check if the base directory exists
if not Path(base_directory).exists():
    logging.error(f"The base directory {base_directory} does not exist.")
else:
    # Extract addresses from the Word document
    addresses = extract_text_from_docx(doc_path)

    # Create folders
    for idx, address in enumerate(addresses, start=1):
        try:
            # Replace problematic characters in folder names
            address_cleaned = address.replace("/", "_").replace("\\", "_").replace(":", "_")
            folder_name = f"{idx}-{address_cleaned}"
            folder_path = os.path.join(base_directory, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            logging.info(f"Created folder: {folder_path}")
        except Exception as e:
            logging.error(f"Error creating folder for '{address}': {e}")
