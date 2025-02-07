import streamlit as st
import os
import openpyxl
import tempfile
from io import BytesIO

def process_excel_files(uploaded_files):
    """Processes uploaded Excel files and allows users to download them."""
    if not uploaded_files:
        return "❌ No files uploaded. Please upload .xlsx files."

    processed_files = []

    for uploaded_file in uploaded_files:
        try:
            # Load Excel file
            wb = openpyxl.load_workbook(uploaded_file)
            temp_file = BytesIO()
            wb.save(temp_file)
            temp_file.seek(0)  # Reset pointer for reading

            # Store processed file
            processed_files.append((uploaded_file.name, temp_file))
        
        except Exception as e:
            st.write(f"❌ Error processing {uploaded_file.name}: {e}")

    return processed_files

# Streamlit UI
st.title("📂 Nielsen File Conversion")
st.write("Upload `.xlsx` files, and the script will process them.")

uploaded_files = st.file_uploader("Drop Excel files here", type=["xlsx"], accept_multiple_files=True)

if st.button("Start Processing"):
    if uploaded_files:
        processed_files = process_excel_files(uploaded_files)
        if processed_files:
            st.success(f"✅ Processed {len(processed_files)} files successfully!")

            # Provide download links for each file
            for file_name, file_data in processed_files:
                st.download_button(
                    label=f"📥 Download {file_name}",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        st.error("Please upload at least one Excel file.")
