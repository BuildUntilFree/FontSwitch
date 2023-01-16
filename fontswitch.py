import docx
import os

# Set the font you want to use
new_font = 'PLACE_YOUR_CHOSEN_FONT_FOR_NEW_DOCUMENT_HERE'

# Specify the directory containing the .docx files
directory = r'C:\PLACE_YOUR_FILE_HERE\FontSwitch\From'

# Iterate through all files in the directory
for filename in os.listdir(directory):
    # Only open .docx files
    if filename.endswith('.docx'):
        # Open the .docx file
        doc = docx.Document(os.path.join(directory, filename))

        # Iterate through all paragraphs in the document
        for p in doc.paragraphs:
            p.style.font.name = new_font

        # Save the changes to the .docx file
        doc.save(os.path.join(r'C:\PLACE_YOUR_FILE_HERE\FontSwitch\To', filename))

print("FontSwitched!")