import pandas as pd  # Ensure pandas is installed with "pip install pandas"
import win32com.client as win32  # Ensure pywin32 is installed with "pip install pywin32"
import os
from docx import Document  # Ensure python-docx is installed with "pip install python-docx"
import base64
from lxml import etree  # Ensure lxml is installed with "pip install lxml"

# Load your Excel file
excel_file = "contacts.xlsx"  # Update with the actual name of the Excel sheet
df = pd.read_excel(excel_file)

# Paths to the Word document and PDF document
word_doc_path = "Email.docx"  # Update with the actual name of the Word doc
pdf_doc_path = "attachment.pdf"  # Update with the actual name of the PDF file

# Function to extract text, images, and hyperlinks from a Word document
def extract_text_images_links(doc_path):
    doc = Document(doc_path)
    content_parts = []
    rels = doc.part.rels
    for para in doc.paragraphs:
        paragraph_xml = para._element

        for child in paragraph_xml:
            if child.tag.endswith('}hyperlink'):
                # Handling hyperlinks
                rId = child.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if rId and rId in rels:
                    href = rels[rId].target_ref
                    link_text = ''.join([node.text for node in child.iter() if node.tag.endswith('}t')])
                    content_parts.append(f'<a href="{href}">{link_text}</a>')
            elif child.tag.endswith('}r'):
                # Handling regular text and inline images
                run = child
                run_text = ''
                bold = False
                highlight = None

                for node in run.iter():
                    if node.tag.endswith('}t'):
                        # Text node
                        run_text += node.text
                    elif node.tag.endswith('}blip'):
                        # Inline image node
                        rId = node.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if rId and rId in rels:
                            image = rels[rId].target_part.blob
                            img_format = rels[rId].target_part.content_type.split("/")[-1]
                            img_b64 = base64.b64encode(image).decode('utf-8')
                            img_html = f'<img src="data:image/{img_format};base64,{img_b64}" />'
                            content_parts.append(img_html)
                    elif node.tag.endswith('}b'):
                        # Bold text
                        bold = True
                    elif node.tag.endswith('}highlight'):
                        # Highlighted text
                        highlight = node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")

                if run_text:
                    if bold:
                        run_text = f'<b>{run_text}</b>'
                    if highlight:
                        # Applying yellow highlight color for simplicity; adjust as needed
                        run_text = f'<span style="background-color: yellow;">{run_text}</span>'
                    content_parts.append(run_text)

        # Add a line break after each paragraph
        content_parts.append('<br>')

    return ''.join(content_parts)

# Iterate over the rows in the Excel file
for index, row in df.iterrows():
    content_parts = []
    content_parts.append(f"Dear {row['Name']},<br><br>")
    content_parts.append(f"Thank you for indicating your interest in joining <b>NTU Heritage Club Recruitment Drive 2024!</b><br><br>")
    content_parts.append(f"The following is the details for your <b>allocated timeslot</b> for the interviews:<br>")
    content_parts.append(f"Date:<b>26th August 2024</b><br>")
    content_parts.append(f"Timeslot:<b>{row['Time']}</b><br>") #*** Make sure your excel sheet has 'Time' Column
    content_parts.append(f"Venue: <b> NS TR+3 </b>for registration<br><br>") #*** Can edit this to include anything you want (e.g.: Zoom Link)
    # Extract the email body content with inline images and hyperlinks
    content_parts.append(extract_text_images_links(word_doc_path))
    email_body_html = ''.join(content_parts)

    recipient_name = row['Name']  # Ensure your Excel sheet column name is "Name"
    recipient_email = row['Email']  # Ensure your Excel sheet column name is "Email"

    # Initialize Outlook and send the email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = "Test Code"  #*** Update with the actual subject of the email
    mail.To = recipient_email

    # Set the HTML body of the email
    mail.HTMLBody = email_body_html

    # Attach the PDF document
    attachment_path = os.path.abspath(pdf_doc_path)
    mail.Attachments.Add(attachment_path)

    # Send the email
    mail.Send()
    print(f"Email sent to {recipient_name} at {recipient_email}")

print("All emails have been sent.")
