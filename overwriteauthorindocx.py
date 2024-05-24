# This script overwrites the author names in all docx files in folders and subfolders.
# It is created with the assumption that it will be called from Cron job on a web server.

from docx import Document
import os


def find_docx_files(directory):
    docx_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                full_path = os.path.join(root, file)
                docx_files.append(full_path)
    return docx_files

def modify_docx_file(filename):
    # Open the Document
    doc = Document(filename)

    # Change the author and last modified by properties
    if doc.core_properties.author != 'jsmez' or  doc.core_properties.last_modified_by != 'jsmez' :
        doc.core_properties.author = 'jsmez'
        doc.core_properties.last_modified_by = 'jsmez'
        # Save the document
        doc.save(filename)


if __name__ == '__main__':
    current_dir = '/home/yourdirectory'
    docx_files = find_docx_files(current_dir)
    for filename in docx_files:
        modify_docx_file(filename)
    
    

