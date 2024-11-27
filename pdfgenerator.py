import subprocess
import os
import streamlit as st
from docx import Document

# Function to edit the Word template dynamically
def edit_word_template(template_path, output_path, name, designation, contact, email, location, selected_services):
    try:
        doc = Document(template_path)

        # Replace placeholders in the general paragraphs
        for para in doc.paragraphs:
            if "<<Client Name>>" in para.text:
                para.text = para.text.replace("<<Client Name>>", name)
            if "<<Client Designation>>" in para.text:
                para.text = para.text.replace("<<Client Designation>>", designation)
            if "<<Client Contact>>" in para.text:
                para.text = para.text.replace("<<Client Contact>>", contact)
            if "<<Client Email>>" in para.text:
                para.text = para.text.replace("<<Client Email>>", email)
            if "<<Client Location>>" in para.text:
                para.text = para.text.replace("<<Client Location>>", location)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "<<Client Name>>" in cell.text:
                        cell.text = cell.text.replace("<<Client Name>>", name)
                    if "<<Client Designation>>" in cell.text:
                        cell.text = cell.text.replace("<<Client Designation>>", designation)
                    if "<<Client Contact>>" in cell.text:
                        cell.text = cell.text.replace("<<Client Contact>>", contact)
                    if "<<Client Email>>" in cell.text:
                        cell.text = cell.text.replace("<<Client Email>>", email)
                    if "<<Client Location>>" in cell.text:
                        cell.text = cell.text.replace("<<Client Location>>", location)
        # Process the table to retain only selected services
        for table in doc.tables:
            # Check if the table contains the column headers (assumes headers in the first row)
            if "Name" in table.rows[0].cells[0].text and "Description" in table.rows[0].cells[1].text:
                # Filter rows based on selected services
                rows_to_keep = [table.rows[0]]  # Keep the header row
                for row in table.rows[1:]:
                    service_name = row.cells[0].text.strip()
                    if service_name in selected_services:
                        rows_to_keep.append(row)
                
                # Remove all rows and re-add only the filtered rows
                while len(table.rows) > 0:
                    table._element.remove(table.rows[0]._element)
                for row in rows_to_keep:
                    new_row = table.add_row()
                    for i, cell in enumerate(row.cells):
                        new_row.cells[i].text = cell.text

        # Save the updated document
        doc.save(output_path)
        print(f"Word document updated and saved at: {output_path}")

    except Exception as e:
        raise Exception(f"Error editing Word template: {e}")




# Function to convert Word to PDF using LibreOffice
import os
import subprocess

def convert_to_pdf(doc_path, pdf_path):
    try:
        # Construct the LibreOffice command
        command = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            doc_path,
            "--outdir", os.path.dirname(pdf_path)
        ]

        # Run the command
        subprocess.run(command, check=True)
        print(f"Converted to PDF and saved at: {pdf_path}")
    except subprocess.CalledProcessError as e:
        raise Exception(f"Error converting Word to PDF: {e}")

# Streamlit App
st.title("Client-Specific PDF Generator")

# Input fields
name = st.text_input("Name")
designation = st.text_input("Designation")
contact = st.text_input("Contact Number")
email = st.text_input("Email ID")
location = st.selectbox("Location", ["India", "ROW"])

# List of all available services
services = [
    "Landing page website (design + development)",
    "AI Automations (6 Scenarios)",
    "WhatsApp Automation + WhatsApp Cloud Business Account Setup",
    "CRM Setup",
    "Email Marketing Setup",
    "Make/Zapier Automation Setup",
    "Firefly Meeting Automation",
    "Marketing Strategy",
    "Social Media Channels",
    "Creatives (10 Per Month)",
    "Creatives (20 Per Month)",
    "Creatives (30 Per Month)",
    "Reels (10 Reels)",
    "Meta Ad Account Setup & Pages Setup",
    "Paid Ads (Lead Generation)",
    "Monthly Maintenance & Reporting",
    "AI Chatbot",
    "PDF Generation Automations",
    "AI Generated Social Media Content & Calendar",
    "Custom AI Models & Agents"
]

# Multi-select for services
selected_services = st.multiselect("Select Services", services)

# Define paths
base_dir = os.path.abspath(os.path.dirname(__file__))
template_path = os.path.join(base_dir, "DM & Automations Services Pricing - Andrew.docx")
word_output_path = os.path.join(base_dir, "Customized_Pricing.docx")
pdf_output_path = os.path.join(base_dir, "Customized_Pricing.pdf")

if st.button("Generate PDF"):
    if not all([name, designation, contact, email, location]) or not selected_services:
        st.error("All fields and at least one service must be selected!")
    else:
        try:
            # Edit the Word template
            edit_word_template(
                template_path, word_output_path, name, designation, contact, email, location, selected_services
            )
            # Convert the edited Word document to PDF
            convert_to_pdf(word_output_path, pdf_output_path)
            st.success("PDF generated successfully!")
            with open(pdf_output_path, "rb") as file:
                st.download_button("Download PDF", file, file_name="Customized_Pricing.pdf")
        except Exception as e:
            st.error(f"An error occurred: {e}")
