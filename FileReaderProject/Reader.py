import os
import glob
from docx import Document

# Set the absolute path to the Word document folder on your desktop
folder_path = "/Users/domm/Desktop/WordDocFolder"

# Initialize the total word count
total_word_count = 0

# Create a log file
log_file = open("word_file_log.txt", "w")

# Iterate through Word files in the specified folder
for filename in glob.glob(os.path.join(folder_path, "*.docx")):
    doc = Document(filename)
    word_count = sum(len(paragraph.text.split()) for paragraph in doc.paragraphs)
    file_name = os.path.basename(filename)
    creation_date = os.path.getctime(filename)  # You may need to format this date.

    # Log information for each file
    log_file.write(f"File Name: {file_name}\n")
    log_file.write(f"Creation Date: {creation_date}\n")
    log_file.write(f"Word Count: {word_count}\n")
    log_file.write("\n")

    total_word_count += word_count

# Log the total word count
log_file.write(f"Total Word Count: {total_word_count}\n")

# Close the log file
log_file.close()
