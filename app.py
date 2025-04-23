import streamlit as st
import tempfile
import os
import io
import pandas as pd
from script import prepare_final_excel, extract_data, extract_data_facture, process_df_facture

st.title("Facture Excel Processing")

# Upload multiple Excel files (for metadata)
uploaded_excel_metadata = st.file_uploader("Importer les decharges excel", type=["xlsx"], accept_multiple_files=True)

# Upload multiple Excel invoice files
uploaded_excel_factures = st.file_uploader("Importer les factures excel", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_excel_metadata and uploaded_excel_factures:
    # Create a temporary directory to save the uploaded files
    temp_dir = tempfile.mkdtemp()

    # 1. Combine metadata from extract_data()
    data_combined = {}
    for excel_file in uploaded_excel_metadata:
        data = extract_data(excel_file)
        data_combined.update(data)

    # 2. Extract and combine all product data from facture Excel files
    df_all_products = pd.DataFrame()
    for facture_file in uploaded_excel_factures:
        # Save uploaded file to a temp path for compatibility
        temp_path = os.path.join(temp_dir, facture_file.name)
        with open(temp_path, "wb") as f:
            f.write(facture_file.getbuffer())

        df_facture = extract_data_facture(temp_path)
        df_all_products = pd.concat([df_all_products, df_facture], ignore_index=True)

    # 3. Enrich facture data with metadata
    df_enriched = process_df_facture(df_all_products, data_combined)

    # 4. Format into the final Excel output
    df_final = prepare_final_excel(df_enriched)

    # 5. Save the final DataFrame to a BytesIO buffer
    output = io.BytesIO()
    df_final.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    # 6. Provide download link
    st.download_button(
        label="Télécharger la ventillation",
        data=output,
        file_name="ventillation_template_populated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Cleanup
    for file in os.listdir(temp_dir):
        os.remove(os.path.join(temp_dir, file))
    os.rmdir(temp_dir)
