import streamlit as st
import os
import openpyxl
import time
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor

def process_excel(uploaded_file):
    """Processes a single Excel file and returns the processed file."""
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        temp_file = BytesIO()
        wb.save(temp_file)
        temp_file.seek(0)  # Reset pointer

        return uploaded_file.name, temp_file, None  # Return processed file
    except Exception as e:
        return uploaded_file.name, None, str(e)  # Return error message

def process_excel_files_parallel(uploaded_files):
    """Processes multiple Excel files in parallel and updates the progress bar."""
    if not uploaded_files:
        return None, "❌ No files uploaded. Please upload .xlsx files."

    total_files = len(uploaded_files)
    processed_files = []
    errors = []

    progress_bar = st.progress(0)
    status_text = st.empty()

    with ThreadPoolExecutor(max_workers=4) as executor:  # Adjust max_workers as needed
        results = list(executor.map(process_excel, uploaded_files))

    for i, (file_name, file_data, error) in enumerate(results, start=1):
        if error:
            errors.append(f"❌ Error processing {file_name}: {error}")
        else:
            processed_files.append((file_name, file_data))

        progress_bar.progress(i / total_files)  # Update progress
        status_text.text(f"✅ Processed ({i}/{total_files}): {file_name}")

    progress_bar.progress(1.0)  # Ensure progress reaches 100%
    status_text.text("🎯 All files processed successfully!")

    # Display errors if any
    if errors:
        for error in errors:
            st.error(error)

    return processed_files, f"✅ Processed {len(processed_files)}/{total_files} Excel files successfully!"

# Streamlit UI
st.title("📂 Fast Nielsen File Conversion 🚀")
st.write("Upload `.xlsx` files, and the script will process them quickly in parallel.")

uploaded_files = st.file_uploader("Drop Excel files here", type=["xlsx"], accept_multiple_files=True)

if st.button("Start Processing"):
    if uploaded_files:
        processed_files, result_message = process_excel_files_parallel(uploaded_files)
        st.success(result_message)

        # Provide download links for each processed file
        for file_name, file_data in processed_files:
            st.download_button(
                label=f"📥 Download {file_name}",
                data=file_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.error("Please upload at least one Excel file.")
