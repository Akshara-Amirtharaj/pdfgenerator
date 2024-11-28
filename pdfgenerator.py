import streamlit as st
from docx import Document
import os
import time
import requests
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
        # Process all tables
        spoc_table_found = False  # Flag to indicate if the SPOC table is found
        for table_idx, table in enumerate(doc.tables):
            # Check for the SPOC table by searching for the text "Supporting SPOC Details"
            if not spoc_table_found:  # Look for the SPOC identifier
                for para in doc.paragraphs:
                    if "Supporting SPOC Details" in para.text:
                        spoc_table_found = True
                        break
            if spoc_table_found and table_idx == 0:  # Assuming SPOC table is the first table after the identifier
                # Update placeholders in the SPOC table
                for row in table.rows:
                    if "Project Sponsor/Client’s Detail" in row.cells[0].text:
                        row.cells[1].text = name
                        row.cells[2].text = designation
                        row.cells[3].text = contact
                        row.cells[4].text = email
                spoc_table_found = False  # Reset the flag after processing the table
            else:
                # Filter rows based on selected services for other tables
                for row in table.rows[1:]:  # Skip the header row
                    service_name = row.cells[0].text.strip()
                    if service_name not in selected_services:
                        row._element.getparent().remove(row._element)
        # Save the updated document
        doc.save(output_path)
        print(f"Word document updated and saved at: {output_path}")
    except Exception as e:
        raise Exception(f"Error editing Word template: {e}")
# Updated convert_to_pdf function


def convert_to_pdf(doc_path, pdf_path):
    api_key = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMjVkNWJkMjlhYjAzZDVmMTVmM2ZiMmM1ZjQ3YjNlNmZkODZkZmE0MjJhOGZiMTlkMjhjMzFkYmJlMDNiYmFhZjRlZTMzNGQ5MzdhMGVmMTciLCJpYXQiOjE3MzI3ODcwMjUuNDA5OTM1LCJuYmYiOjE3MzI3ODcwMjUuNDA5OTM2LCJleHAiOjQ4ODg0NjA2MjUuNDA2MzQxLCJzdWIiOiI3MDMzMDc0OSIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.jkpWOMwG7jg0SzxO58A_-Ndhg3iVzgcExwIgbYheAuJ1SLxAH77x0dCd2dIm729779WJmy0UpaOpFHWQ_CFBZKJc7u-MeXcjTZvzt9xraEVvbylb_o0F9TW6-CF5vvO5ps0I7uGznmjyqrVKhxaPYgPux5gKsXe4DuhMp6oiAOb2_i70yDzDOAb9F3DDdCeJVpdtXg0lMYauM6QGYrsUDucqDPFazGY47itApyEwP7S5JIRFQPaozRicL40xaMgiqRUpI65-ByCA4ZTDFqLQBa8T-nIyUNRvST1pU_0FNP3Q4g0PZ6m5U0aNZ9GBHwkUMzveigajDWSZ2h7g2CYZkZ5B9jvxuStyS83vjpod3CvyFEvoNa6gK7IGXJw_PTjNN9EIuGxtSmwvMaonRJKwhaMABVRE1VR0MJ_wx6Ehc91bQFr9YJxyFSRiHrQVBmNh5zoBvJ0ZXyvlUNMre13iqVw5atrTntcVyLTHeZICSiaKelGPmLK9J1ETwaYIwBauUofA0gdRMI1OmY7VCcbtMkiNVUnSkVgjUA6a1eMcOCZVqJycR7T-ijXXcKQloi1Vm0GaCoOPvBoyMdjuSx19sXOdBhyeZXz-iGKVLs_jnHP3fFdKMXEJMgTf6Hvv5KZivxPk2qPmky0gHtJzOs-Ob_t5H5zOl6LpBLvJftGMg8U"  # Replace with your actual API key
    endpoint = "https://api.cloudconvert.com/v2/jobs"


    try:
        # Step 1: Create a conversion job with tasks
        job_payload = {
            "tasks": {
                "upload-task": {
                    "operation": "import/upload"
                },
                "convert-task": {
                    "operation": "convert",
                    "input": "upload-task",
                    "input_format": "docx",
                    "output_format": "pdf"
                },
                "export-task": {
                    "operation": "export/url",
                    "input": "convert-task"
                }
            }
        }

        # Create the job
        job_response = requests.post(
            endpoint,
            headers={"Authorization": f"Bearer {api_key}"},
            json=job_payload
        )
        if job_response.status_code != 201:
            raise Exception(f"Job creation failed: {job_response.json()}")

        job_data = job_response.json()
        upload_task_result = job_data["data"]["tasks"][0]["result"]["form"]
        upload_url = upload_task_result["url"]
        upload_parameters = upload_task_result["parameters"]

        # Step 2: Upload the file to the provided URL
        with open(doc_path, "rb") as file:
            upload_response = requests.post(upload_url, files={"file": file}, data=upload_parameters)
        if upload_response.status_code not in [200, 201]:
            raise Exception(f"File upload failed: {upload_response.text}")

        # Step 3: Poll the job status until it finishes
        job_id = job_data["data"]["id"]
        while True:
            status_response = requests.get(
                f"{endpoint}/{job_id}",
                headers={"Authorization": f"Bearer {api_key}"}
            )
            status_data = status_response.json()
            if status_data["data"]["status"] == "finished":
                break
            elif status_data["data"]["status"] == "error":
                raise Exception(f"Job failed: {status_data}")
            time.sleep(3)  # Wait before polling again

        # Step 4: Locate the export task and download the converted file
        output_url = None
        for task in status_data["data"]["tasks"]:
            if task["operation"] == "export/url":
                output_url = task["result"]["files"][0]["url"]
                break

        if not output_url:
            raise Exception("Export task with file URL not found.")

        # Download the converted file
        pdf_response = requests.get(output_url)
        with open(pdf_path, "wb") as output_file:
            output_file.write(pdf_response.content)

        print(f"PDF saved at {pdf_path}")

    except Exception as e:
        raise Exception(f"Error during PDF conversion: {e}")




st.title("Client-Specific PDF Generator")
# Input fields
name = st.text_input("Name")
designation = st.text_input("Designation")
contact = st.text_input("Contact Number")
email = st.text_input("Email ID")
location = st.selectbox("Location", ["India", "ROW"])

# List of all available services (ensure this matches your template)
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

# Checkbox to select all services
select_all = st.checkbox("Select All Services")
if select_all:
    selected_services = services
else:
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