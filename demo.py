import streamlit as st
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from io import BytesIO


# Function to overlay text on the PDF template
def fill_pdf_template(pdf_template_path, output_path, name, designation, contact, email, location, selected_services):
    try:
        # Read the original template PDF
        template_reader = PdfReader(pdf_template_path)
        pdf_writer = PdfWriter()

        # Create an in-memory buffer for the overlay
        buffer = BytesIO()

        # Use ReportLab to create an overlay with dynamic content
        c = canvas.Canvas(buffer)
        c.setFont("Helvetica", 12)

        # Add user details dynamically
        c.drawString(100, 750, f"Client Name: {name}")
        c.drawString(100, 730, f"Designation: {designation}")
        c.drawString(100, 710, f"Contact: {contact}")
        c.drawString(100, 690, f"Email: {email}")
        c.drawString(100, 670, f"Location: {location}")

        # Add selected services
        c.drawString(100, 650, "Selected Services:")
        y_position = 630
        for service in selected_services:
            c.drawString(120, y_position, f"- {service}")
            y_position -= 20

        c.save()

        # Merge the overlay onto the template PDF
        buffer.seek(0)
        overlay_reader = PdfReader(buffer)

        for page in template_reader.pages:
            # Add overlay content to the page
            page.merge_page(overlay_reader.pages[0])
            pdf_writer.add_page(page)

        # Write the final PDF to the output path
        with open(output_path, "wb") as output_pdf:
            pdf_writer.write(output_pdf)

    except Exception as e:
        raise Exception(f"Error filling the PDF template: {e}")


# Streamlit App
st.title("Client-Specific PDF Generator")

# Input fields
name = st.text_input("Name")
designation = st.text_input("Designation")
contact = st.text_input("Contact Number")
email = st.text_input("Email ID")
location = st.selectbox("Location", ["India", "ROW"])

# List of available services
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
    "Custom AI Models & Agents",
]

# Multi-select for services
selected_services = st.multiselect("Select Services", services)

# Define paths
pdf_template_path = "DM & Automations Services Pricing - Andrew 1.pdf"  # Pre-uploaded template PDF
output_path = "Customized_Pricing.pdf"

if st.button("Generate PDF"):
    if not all([name, designation, contact, email, location]) or not selected_services:
        st.error("All fields and at least one service must be selected!")
    else:
        try:
            # Fill the PDF template dynamically
            fill_pdf_template(
                pdf_template_path, output_path, name, designation, contact, email, location, selected_services
            )
            st.success("PDF generated successfully!")

            # Allow the user to download the final PDF
            with open(output_path, "rb") as file:
                st.download_button("Download PDF", file, file_name="Customized_Pricing.pdf")
        except Exception as e:
            st.error(f"An error occurred: {e}")
