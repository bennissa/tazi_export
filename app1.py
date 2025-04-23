import streamlit as st
import tempfile
import os
import io
from script import prepare_final_excel, process_uploaded_files, clean_and_format_data, extract_data, process_df_facture, extract_data_facture
from PIL import Image
import pandas as pd


st.title("Image and Excel Processing")

# Upload multiple Excel files and multiple PDF files
uploaded_excel_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)
uploaded_pdf_files = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)

if uploaded_excel_files and uploaded_pdf_files:
    # Create a temporary directory to save the uploaded files
    temp_dir = tempfile.mkdtemp()

    # Initialize variables to store combined data and tables
    data_combined = {}
    extracted_tables = []

    # Loop through each uploaded PDF file
    for uploaded_pdf in uploaded_pdf_files:
        # Save the uploaded PDF to the temporary directory
        temp_pdf_path = os.path.join(temp_dir, uploaded_pdf.name)
        with open(temp_pdf_path, "wb") as temp_file:
            temp_file.write(uploaded_pdf.getbuffer())

        # Process the PDF files    
        tables = process_uploaded_files(temp_pdf_path)
        extracted_tables.extend(tables)

        # Cleanup: Remove the temporary PDF file after processing
        os.remove(temp_pdf_path)

    for uploaded_excel in uploaded_excel_files:
        data = extract_data(uploaded_excel)
        data_combined.update(data)

    # Clean and format the extracted data
    df = clean_and_format_data(extracted_tables, data_combined)
    df_final = prepare_final_excel(df)
    
    # Save the final Excel file to a BytesIO buffer
    output = io.BytesIO()
    df_final.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)  # Reset the pointer to the start of the BytesIO buffer

    # Provide download link for the Excel file
    st.download_button(
        label="Download Processed Excel",
        data=output,
        file_name="ventillation_template_populated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Cleanup: Remove the temporary directory after processing all files
    os.rmdir(temp_dir)
