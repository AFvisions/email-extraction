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

# Handle folders with a single PDF
with open('client_data.csv', 'a', newline='') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow(["ClientLastName", "ClientFirstName", "Filename", "Page", "Email", "Status", "Duration"])

    # Iterate over all subfolders in the root directory
    for foldername in os.listdir(root_directory):
        # Extract the client's last name and first name from the folder name
        names = foldername.split(',')
        client_last_name = names[0].strip()
        client_first_name = names[1].strip() if len(names) > 1 else ""

        # Check the number of PDF files in the folder
        folder_path = os.path.join(root_directory, foldername)
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        if len(pdf_files) == 1:
            # Only one PDF file - let's handle it
            filename = pdf_files[0]
            full_path = os.path.join(folder_path, filename)
            
            print("Starting PDF Extraction for " + filename)

            # Initialize variables to hold the text data and extracted information
            text_data = ""
            emails = set()

            print("Starting Image Conversion and OCR for " + filename)
            # Convert PDF to images and perform OCR
            start_time = time.time()
            images = convert_from_path(full_path)
            for i in range(len(images)):
                # Perform OCR on the image
                text = pytesseract.image_to_string(images[i])
                text_data += text

                # Use the regex to find all email addresses in the text
                email_matches = re.findall(email_regex, text)
                if email_matches:
                    emails.update({(email, i + 1) for email in email_matches})  # Add all found emails to the set

            duration = time.time() - start_time

            print("Starting Email Extraction for " + filename)
            # Write the data to the CSV
            for email, page in emails:
                writer.writerow([client_last_name, client_first_name, filename, page, email, "Processed", str(int(duration // 60)) + "m" + str(int(duration % 60)) + "s"])

            print("Completed Processing for " + foldername)
        elif len(pdf_files) == 0:
            writer.writerow([client_last_name, client_first_name, "", "", "", "No PDF", ""])
        else:
            writer.writerow([client_last_name, client_first_name, "", "", "", "Multiple PDFs", ""])
