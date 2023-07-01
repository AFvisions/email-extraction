import csv
import os
import re
from PIL import Image
from pdf2image import convert_from_path
import pytesseract

# Regular expression to match email addresses
email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Directory containing your client folders
root_directory = r"C:\Users\AFuma\Desktop\buford-closed-files-LOCAL\clients\A"  # update with your path

# Part 1: Get all client last names
client_last_names = []

# Iterate over all subfolders in the root directory
for foldername in os.listdir(root_directory):
    # Extract the client's last name from the folder name
    client_last_name = foldername.split(',')[0]
    client_last_names.append(client_last_name)

# Write the client last names to a CSV file
with open('client_last_names.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow(["ClientLastName"])
    # Write the client last names
    for name in client_last_names:
        writer.writerow([name])

# Part 2: Handle folders with a single PDF

with open('client_emails.csv', 'w', newline='') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow(["ClientLastName", "Filename", "Email"])

    # Iterate over all subfolders in the root directory
    for foldername in os.listdir(root_directory):
        # Extract the client's last name from the folder name
        client_last_name = foldername.split(',')[0]

        # Check the number of PDF files in the folder
        folder_path = os.path.join(root_directory, foldername)
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        if len(pdf_files) == 1:
            # Only one PDF file - let's handle it
            filename = pdf_files[0]
            full_path = os.path.join(folder_path, filename)

            # Convert the PDF into images
            images = convert_from_path(full_path)

            # Initialize variable to hold the text data
            text_data = ""

            # Iterate over the images and perform OCR on each one
            for i, image in enumerate(images):
                text_data += pytesseract.image_to_string(image)

            # Use the regex to find all email addresses in the text
            email_matches = re.findall(email_regex, text_data)
            emails = ', '.join(email_matches) if email_matches else None

            # Write the data to the CSV
            writer.writerow([client_last_name, filename, emails])
