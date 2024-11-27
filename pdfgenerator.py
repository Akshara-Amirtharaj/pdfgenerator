import streamlit as st
from docx import Document
import os

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

        # Update the services table
        for table in doc.tables:
            # Assuming the table for services starts with specific headers (adjust accordingly)
            if "Name" in table.rows[0].cells[0].text and "Description" in table.rows[0].cells[1].text:
                # Clear all rows except the header
                for row in table.rows[1:]:
                    row._element.getparent().remove(row._element)
                
                # Add only selected services to the table
                services_data = {
                    "Landing page website (design + development)": ["Using Next JS", "£200", "5-10 Days", "One Time Fee"],
                    "AI Automations (6 Scenarios)": ["Leads Connection with CRM & AI Voice Calling Automation", "£1000", "10-20 Days", "One Time Fee"],
                    "WhatsApp Automation + WhatsApp Cloud Business Account Setup": ["Automation Setup", "£750", "10-20 Days", "One Time Fee"],
                    "CRM Setup": ["Any CRM", "£500", "5-10 Days", "One Time Fee"],
                    "Email Marketing Setup": ["Email Templates & Marketing ID", "£500", "5-10 Days", "One Time Fee"],
                    "Make/Zapier Automation Setup": ["Automation Setup", "£750", "10-20 Days", "One Time Fee"],
                    "Firefly Meeting Automation": ["Automation Setup", "£250", "10-20 Days", "One Time Fee"],
                    "Marketing Strategy": ["Custom Marketing Plan", "£1000", "10-15 Days", "One Time Fee"],
                    "Social Media Channels": ["Setup & Optimization", "£800", "7-15 Days", "One Time Fee"],
                    "Creatives (10 Per Month)": ["10 Creative Posts", "£200", "30 Days", "Monthly"],
                    "Creatives (20 Per Month)": ["20 Creative Posts", "£400", "30 Days", "Monthly"],
                    "Creatives (30 Per Month)": ["30 Creative Posts", "£600", "30 Days", "Monthly"],
                    "Reels (10 Reels)": ["10 Video Reels", "£300", "30 Days", "Monthly"],
                    "Meta Ad Account Setup & Pages Setup": ["Setup Ad Accounts & Pages", "£500", "5-10 Days", "One Time Fee"],
                    "Paid Ads (Lead Generation)": ["Lead Gen Campaigns", "£1000", "30 Days", "Monthly"],
                    "Monthly Maintenance & Reporting": ["Reports & Optimization", "£500", "30 Days", "Monthly"],
                    "AI Chatbot": ["AI-ML Model Training", "£500", "10-20 Days", "One Time Fee"],
                    "PDF Generation Automations": ["PDF Generator & Automation", "£500", "10-20 Days", "One Time Fee"],
                    "AI Generated Social Media Content & Calendar": ["Social Content Plan", "£800", "10-15 Days", "Monthly"],
                    "Custom AI Models & Agents": ["Custom AI Solution", "£2000", "30-60 Days", "One Time Fee"]
                }

                for service in selected_services:
                    if service in services_data:
                        row = table.add_row()
                        row.cells[0].text = service
                        row.cells[1].text = services_data[service][0]
                        row.cells[2].text = services_data[service][1]
                        row.cells[3].text = services_data[service][2]
                        row.cells[4].text = services_data[service][3]

        # Save the updated document
        doc.save(output_path)
        print(f"Word document updated and saved at: {output_path}")

    except Exception as e:
        raise Exception(f"Error editing Word template: {e}")


# Updated convert_to_pdf function
import pypandoc

def convert_to_pdf(doc_path, pdf_path):
    try:
        # Convert the .docx to .pdf using pypandoc
        pypandoc.convert_file(doc_path, 'pdf', outputfile=pdf_path)
        print(f"Converted to PDF and saved at: {pdf_path}")
    except Exception as e:
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
            edit_word_template(
                template_path, word_output_path, name, designation, contact, email, location, selected_services
            )
            convert_to_pdf(word_output_path, pdf_output_path)
            st.success("PDF generated successfully!")
            with open(pdf_output_path, "rb") as file:
                st.download_button("Download PDF", file, file_name="Customized_Pricing.pdf")
        except Exception as e:
            st.error(f"An error occurred: {e}")
