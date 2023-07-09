import csv
import os
import re
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import time

# Regular expression to match email addresses
email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Directory containing your client folders
root_directory = r"C:\Users\AFuma\Desktop\buford-closed-files-LOCAL\clients\A"  # Remember to adjust this path

start_time = time.time()

# CSV filename
csv_filename = 'client_data.csv'

with open(csv_filename, 'a', newline='') as file:
    writer = csv.writer(file)
    if file.tell() == 0:
        # Write the header row if the file is empty
        writer.writerow(["FolderName", "Filename", "Email", "PageNumber", "Status"])

    # Iterate over all subfolders in the root directory
    for foldername in os.listdir(root_directory):
        print(f"Processing Folder: {foldername}")

        folder_path = os.path.join(root_directory, foldername)
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        
        if len(pdf_files) == 0:
            writer.writerow([foldername, None, None, None, "No PDF"])
            continue
        
        if len(pdf_files) > 1:
            writer.writerow([foldername, None, None, None, "Multiple PDFs"])
            continue
        
        # Only one PDF file - let's handle it
        filename = pdf_files[0]
        full_path = os.path.join(folder_path, filename)
        
        print("Starting PDF Extraction for " + filename)

        # Initialize variables to hold the text data and extracted information
        text_data = []
        emails = set()

        print("Starting Image Conversion and OCR for " + filename)
        # Convert PDF to images and perform OCR
        images = convert_from_path(full_path)
        for i in range(len(images)):
            # Perform OCR on the image
            text = pytesseract.image_to_string(images[i])
            text_data.append(text)

        print("Starting Email Extraction for " + filename)
        for i, text in enumerate(text_data):
            # Use the regex to find all email addresses in the text
            email_matches = re.findall(email_regex, text)
            if email_matches:
                for email in email_matches:
                    if email not in emails:  # Check if the email is not already recorded
                        emails.add(email)
                        writer.writerow([foldername, filename, email, i + 1, "Processed"])  # Add page number and status

        print("Completed Processing for " + foldername)

elapsed_time = time.time() - start_time
print(f"Processing Time: {elapsed_time // 3600} hours {(elapsed_time % 3600) // 60} minutes {elapsed_time % 60} seconds")
