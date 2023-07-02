import csv
import os
import re
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import time

# Record the start time
start_time = time.time()

# Regular expression to match email addresses
email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Directory containing your client folders
root_directory = r"C:\Users\AFuma\Desktop\buford-closed-files-LOCAL\clients\A"  # Remember to adjust this path

# Part 1: Get all client last names
client_last_names = []

# Create a dictionary to store folders that do not meet the criteria of containing exactly one PDF.
folders_not_one_pdf = {}

# Iterate over all subfolders in the root directory
for foldername in os.listdir(root_directory):
    print(f"Processing Folder: {foldername}")
    # Extract the client's last name from the folder name
    client_last_name = foldername.split(',')[0]
    client_last_names.append(client_last_name)

# Write the client last names to a CSV file
with open('client_last_names.csv', 'a', newline='') as file:
    writer = csv.writer(file)
    # Write the header row
    writer.writerow(["ClientLastName"])
    # Write the client last names
    for name in client_last_names:
        writer.writerow([name])

# Part 2: Handle folders with a single PDF
with open('client_emails.csv', 'a', newline='') as file:
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
            
            print("Starting PDF Extraction for " + filename)

            # Initialize variables to hold the text data and extracted information
            text_data = ""
            emails = []

            print("Starting Image Conversion and OCR for " + filename)
            # Convert PDF to images and perform OCR
            images = convert_from_path(full_path)
            for i in range(len(images)):
                # Perform OCR on the image
                text = pytesseract.image_to_string(images[i])
                text_data += text

            print("Starting Email Extraction for " + filename)
            # Use the regex to find all email addresses in the text
            email_matches = re.findall(email_regex, text_data)
            if email_matches:
                emails.extend(email_matches)  # Add all found emails to the list

            # Write the data to the CSV
            for email in emails:
                writer.writerow([client_last_name, filename, email])

            print("Completed Processing for " + foldername)
        else:
            # Not exactly one PDF file in the folder, store it in the dictionary
            folders_not_one_pdf[foldername] = len(pdf_files)

# Print the folders that do not meet the criteria of containing exactly one PDF.
print("Folders not fitting the criteria of containing exactly one PDF:")
for folder, count in folders_not_one_pdf.items():
    print(f"Folder: {folder}, Number of PDFs: {count}")

    # Record the end time
end_time = time.time()

# Calculate and print the elapsed time
elapsed_time = end_time - start_time
hours, rem = divmod(elapsed_time, 3600)
minutes, seconds = divmod(rem, 60)
print(f"The process took {int(hours)}h {int(minutes)}m {int(seconds)}s to complete.")
