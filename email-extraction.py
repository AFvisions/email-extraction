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
root_directory = r"C:\Users\AFuma\Desktop\buford-closed-files-LOCAL\clients\H"  # Remember to adjust this path

start_time = time.time()

with open('client_data.csv', 'a', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["ClientLastName", "Filename", "Email", "Status"])

    # Iterate over all subfolders in the root directory
    for foldername in os.listdir(root_directory):
        print(f"Processing Folder: {foldername}")
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
                writer.writerow([client_last_name, filename, email, "Processed"])
            print("Completed Processing for " + foldername)
        else:
            # Write the status to the CSV
            if len(pdf_files) == 0:
                writer.writerow([client_last_name, "", "", "No PDF"])
            else:
                writer.writerow([client_last_name, "", "", "Multiple PDFs"])

end_time = time.time()
print("Time taken: " + str(end_time - start_time) + " seconds")
