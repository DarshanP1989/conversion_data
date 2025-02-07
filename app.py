import streamlit as st
import os
import win32com.client
import time
import pythoncom  # Required for COM initialization

def process_excel_files(folder_path):
    """Processes all Excel files in the specified folder and updates a progress bar."""

    if not os.path.exists(folder_path):
        return "❌ Folder does not exist. Please enter a valid path."

    total_files = 0
    processed_files = 0

    excel_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                total_files += 1
                file_path = os.path.join(root, file)
                excel_files.append(file_path)

    if not excel_files:
        return "🎉 No Excel files found. Exiting."

    # ✅ Initialize COM
    pythoncom.CoInitialize()

    # ✅ Run Excel in the background
    excel = win32com.client.DispatchEx("Excel.Application")  
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.UserControl = False  

    # 🔵 Create a progress bar
    progress_bar = st.progress(0)  # Start progress at 0%
    status_text = st.empty()  # Placeholder for status updates

    try:
        for i, file_path in enumerate(excel_files, start=1):
            try:
                workbook = excel.Workbooks.Open(file_path)
                workbook.Save()
                workbook.Close(False)
                processed_files += 1
                status_text.text(f"✅ Processed ({i}/{total_files}): {file_path}")
            except Exception as e:
                st.write(f"❌ Error processing {file_path}: {e}")

            # 🔄 Update progress bar
            progress_bar.progress(i / total_files)  # Normalize progress between 0 and 1

            time.sleep(0.3)  # Small delay for visibility

    finally:
        excel.Quit()  # Ensure Excel quits properly
        pythoncom.CoUninitialize()  # ✅ Uninitialize COM after use

    progress_bar.progress(1.0)  # Ensure progress reaches 100% at the end
    status_text.text("🎯 All files processed successfully!")

    return f"✅ Processed {processed_files}/{total_files} Excel files successfully!"

# Streamlit UI
st.set_page_config(page_title="Excel File Processor")

st.title("📂 Excel File Processor with Progress Bar")
st.write("Enter the folder path where your `.xlsx` files are stored. The script will open and re-save all files.")

# User input for folder path
folder_path = st.text_input("Enter folder path:")

if st.button("Start Processing"):
    if folder_path:
        result = process_excel_files(folder_path)
        st.success(result)
    else:
        st.error("Please enter a valid folder path.")
