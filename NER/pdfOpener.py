import webbrowser

import os

# Define the base directory and file name
base_dir = r'C:\Users\Adarsha Kumar\Downloads\WhatsApp Chat Analysis\NER'
file_name = '+91 87787 87613_entities.pdf'

# Construct the full path
pdf_output_path = os.path.join(base_dir, 'NER_Results', file_name)

import subprocess

# Open the PDF file with the default PDF viewer
subprocess.Popen([pdf_output_path], shell=True)
