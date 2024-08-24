import pandas as pd #go terminal to install pandas "pip install pandas"
import win32com.client as win32  #go terminal to install win32 "pip install pip install pywin32"
import os
from docx import Document #"pip install python-docx" 
import base64
from lxml import etree #"pip install lxml"

#REQUIREMENTS
#1. Word docx containing the contents of the email blast (e.g. Invitation Email Blast.docx)
#2. Excel Spreadsheet containing all the intended receipients (e.g. NTU Heritage Club Membership List.xlsx)
#3. Attachment pdf file (e.g. Constitution 2024.pdf)
#4. All the required modules as stated in the import area
#5. Outlook app must be open!!!
#6. Remember to rename the items with "***""
#7. Make sure this python script, the word doc, the pdf file, the excel sheet are in the same folder in your local
#8. For spacing between each line, please "ENTER" 2 times instead of once to show 1 line of spacing

# Load your Excel file
excel_file = "contacts.xlsx" #*** This should be named as the name of the excel sheet (e.g.: "NTU Heritage Club Membership List.xlsx")
df = pd.read_excel(excel_file)

# Paths to the Word document and PDF document
word_doc_path = "email_content.docx" #*** This should be named as the name of the word doc (e.g.: "Invitation Email Blast.docx")
pdf_doc_path = "attachment.pdf"#*** This should be named as the name of the pdf file (e.g.: "(e.g. Constitution 2024.pdf)")

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
                for node in child.iter():
                    if node.tag.endswith('}t'):
                        # Text node
                        content_parts.append(node.text)
                    elif node.tag.endswith('}blip'):
                        # Inline image node
                        rId = node.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if rId and rId in rels:
                            image = rels[rId].target_part.blob
                            img_format = rels[rId].target_part.content_type.split("/")[-1]
                            img_b64 = base64.b64encode(image).decode('utf-8')
                            img_html = f'<img src="data:image/{img_format};base64,{img_b64}" />'
                            content_parts.append(img_html)

        # Add a line break after each paragraph
        content_parts.append('<br>')

    return ''.join(content_parts)

# Iterate over the rows in the Excel file
for index, row in df.iterrows():
    content_parts = []
    content_parts.append(f"Dear {row['Name']},<br><br>")
    # Extract the email body content with inline images and hyperlinks
    content_parts.append(extract_text_images_links(word_doc_path))
    email_body_html = ''.join(content_parts)

    recipient_name = row['Name'] #***Make sure your excel sheet column name is "Name"
    recipient_email = row['Email'] #***Make sure your excel sheet column name is "Email"

    # Initialize Outlook and send the email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = "Subject of the Email" #*** What is the title of the email to be sent
    mail.To = recipient_email

    # Set the HTML body of the email
    mail.HTMLBody = email_body_html

    # Attach the PDF document
    attachment_path = os.path.abspath(pdf_doc_path)
    mail.Attachments.Add(attachment_path)

    # Send the email
    mail.Send()
    print(f"Email sent to {recipient_name} at {recipient_email}")

print("All emails have been sent.") # This is to check whether all emails have been sent
