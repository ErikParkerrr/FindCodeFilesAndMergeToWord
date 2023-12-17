import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

FILE_EXTENSION = ".vhd"

def remove_header_comments(plaintext_code):
   # Find the start and end markers
    start_marker = "----------------------------------------------------------------------------------"
    end_marker = "----------------------------------------------------------------------------------"

    # Find the indices of the start and end markers
    start_index = plaintext_code.find(start_marker)
    end_index = plaintext_code.find(end_marker, start_index + len(start_marker))

    # Remove the text between the markers
    if start_index != -1 and end_index != -1:
        plaintext_code = plaintext_code[:start_index] + plaintext_code[end_index + len(end_marker):]

    return plaintext_code

def add_plaintext_to_docx(document, plaintext_file):
    # Add file name as a heading
    document.add_heading(f'{os.path.basename(plaintext_file)}', level=2).paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Read plaintext code
    with open(plaintext_file, 'r') as file:
        plaintext_code = file.read()

    # Remove header comments block
    plaintext_code = remove_header_comments(plaintext_code)

    # Create a new paragraph for plaintext code
    paragraph = document.add_paragraph(plaintext_code, style='Code')

    # Set font size
    font = paragraph.runs[0].font
    font.size = Pt(10)

def merge_plaintext_files(output_file_name="merged_plaintext.docx"):
    # Get the current directory
    current_directory = os.path.dirname(os.path.abspath(__file__))

    # Initialize a list to store the paths of all plaintext files
    plaintext_files = []

    # Recursively search for plaintext files in all subdirectories
    for root, dirs, files in os.walk(current_directory):
        plaintext_files.extend([os.path.join(root, file) for file in files if file.endswith(FILE_EXTENSION)])

    # Check if there are any plaintext files
    if not plaintext_files:
        print("No plaintext files found in the current directory or its subdirectories.")
        return

    # Create a new Word document
    document = Document()

    # Add a custom style for plaintext code
    style = document.styles.add_style('Code', WD_PARAGRAPH_ALIGNMENT.CENTER)
    font = style.font
    font.name = 'Courier New'
    font.size = Pt(10)

    # Loop through each plaintext file and add its content to the document
    for plaintext_file in plaintext_files:
        add_plaintext_to_docx(document, plaintext_file)

    # Save the document
    document.save(output_file_name)

    print(f"Merge complete. Merged content saved in {output_file_name}.")

# Call the function to merge plaintext files
merge_plaintext_files()
