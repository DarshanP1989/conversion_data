import streamlit as st
import os
import openpyxl
import time
import tempfile
import zipfile

def process_excel_files(uploaded_files):
    """Processes multiple uploaded Excel files and returns a downloadable ZIP."""
    
    if not uploaded_files:
        return "❌ No files uploaded. Please upload .xlsx files."
    
    total_files = len(uploaded_files)
    processed_files = 0
    temp_folder = tempfile.mkdtemp()  # Temporary folder to store processed files
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    processed_files_list = []

    try:
        for i, uploaded_file in enumerate(uploaded_files, start=1):
            try:
                # Save uploaded file to a temporary folder
                temp_file_path = os.path.join(temp_folder, uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Open and re-save using openpyxl
                wb = openpyxl.load_workbook(temp_file_path)
                wb.save(temp_file_path)  # Overwrite the file

                processed_files += 1
                processed_files_list.append(temp_file_path)

                status_text.text(f"✅ Processed ({i}/{total_files}): {uploaded_file.name}")

            except Exception as e:
                st.write(f"❌ Error processing {uploaded_file.name}: {e}")

            progress_bar.progress(i / total_files)
            time.sleep(0.3)  # Small delay for visibility

    finally:
        progress_bar.progress(1.0)
        status_text.text(f"🎯 All {processed_files}/{total_files} files processed!")

    if processed_files_list:
        # Create a ZIP file of all processed files
        zip_path = os.path.join(tempfile.gettempdir(), "processed_files.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file_path in processed_files_list:
                zipf.write(file_path, os.path.basename(file_path))

        # Provide download link for ZIP
        with open(zip_path, "rb") as zip_file:
            st.download_button(label="📥 Download Processed Files", 
                               data=zip_file, 
                               file_name="processed_files.zip", 
                               mime="application/zip")

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
