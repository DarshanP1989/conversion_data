import streamlit as st
import os
import openpyxl
import time
import tempfile


def process_excel_files(uploaded_files):
    """Processes uploaded Excel files and updates a progress bar."""
    if not uploaded_files:
        return "❌ No files uploaded. Please upload .xlsx files."

    total_files = len(uploaded_files)
    processed_files = 0

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        for i, uploaded_file in enumerate(uploaded_files, start=1):
            try:
                # Save uploaded file to a temporary location
                temp_file_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Open and re-save using openpyxl
                wb = openpyxl.load_workbook(temp_file_path)
                wb.save(temp_file_path)

                processed_files += 1
                status_text.text(f"✅ Processed ({i}/{total_files}): {uploaded_file.name}")

            except Exception as e:
                st.write(f"❌ Error processing {uploaded_file.name}: {e}")

            progress_bar.progress(i / total_files)
            time.sleep(0.3)

    finally:
        progress_bar.progress(1.0)
        status_text.text("🎯 All files processed successfully!")

    return f"✅ Processed {processed_files}/{total_files} Excel files successfully!"

# Streamlit UI
st.set_page_config(page_title="Excel File Processor")

st.title("📂 Nielsen File Conversion")
st.write("Upload `.xlsx` files, and the script will open and re-save them.")

# File uploader (allows multiple file uploads)
uploaded_files = st.file_uploader("Drop Excel files here", type=["xlsx"], accept_multiple_files=True)

if st.button("Start Processing"):
    if uploaded_files:
        result = process_excel_files(uploaded_files)
        st.success(result)
    else:
        st.error("Please upload at least one Excel file.")
