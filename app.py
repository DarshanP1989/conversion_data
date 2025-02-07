import streamlit as st
import os
import time
import pandas as pd

# ✅ Ensure `openpyxl` is installed
try:
    import openpyxl
except ImportError:
    st.error("Missing dependency: `openpyxl`. Please install it using `pip install openpyxl`.")
    st.stop()

# ✅ Function to process Excel files
def process_excel_files(folder_path):
    if not os.path.exists(folder_path):
        return "❌ Folder does not exist. Please enter a valid path."

    total_files = 0
    processed_files = 0

    excel_files = [os.path.join(root, file) for root, _, files in os.walk(folder_path) for file in files if file.endswith(".xlsx")]
    total_files = len(excel_files)

    if total_files == 0:
        return "🎉 No Excel files found. Exiting."

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, file_path in enumerate(excel_files, start=1):
        try:
            # ✅ Open and re-save using openpyxl
            df = pd.read_excel(file_path, engine="openpyxl")  # Read with openpyxl
            df.to_excel(file_path, engine="openpyxl", index=False)  # Save as new file

            processed_files += 1
            status_text.text(f"✅ Processed ({i}/{total_files}): {file_path}")
        except Exception as e:
            st.write(f"❌ Error processing {file_path}: {e}")

        progress_bar.progress(i / total_files)
        time.sleep(0.3)

    progress_bar.progress(1.0)
    status_text.text("🎯 All files processed successfully!")
    return f"✅ Processed {processed_files}/{total_files} Excel files successfully!"

# ✅ Streamlit UI
st.set_page_config(page_title="Excel File Processor")

st.title("📂 Excel File Processor with Progress Bar")
st.write("Select a folder where your `.xlsx` files are stored. The script will open and re-save all files.")

# ✅ Folder selection
folder_path = st.text_input("Enter folder path (or use mounted folder in Streamlit Cloud):")

if st.button("Start Processing"):
    if folder_path:
        result = process_excel_files(folder_path)
        st.success(result)
    else:
        st.error("Please enter a valid folder path.")
