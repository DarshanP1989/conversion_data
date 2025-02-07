import streamlit as st
import os
import pandas as pd
import time

def process_uploaded_files(uploaded_files):
    """Processes uploaded Excel files and updates a progress bar."""
    
    total_files = len(uploaded_files)
    processed_files = 0

    if total_files == 0:
        return "❌ No files uploaded. Please upload Excel files."

    # 🔵 Create a progress bar
    progress_bar = st.progress(0)  
    status_text = st.empty()  

    try:
        for i, uploaded_file in enumerate(uploaded_files, start=1):
            try:
                # Read and save the file using pandas
                df = pd.read_excel(uploaded_file, engine="openpyxl")
                
                # Save processed file (optional: in a selected folder)
                save_path = os.path.join("processed_files", uploaded_file.name)
                os.makedirs("processed_files", exist_ok=True)  # Ensure output folder exists
                df.to_excel(save_path, index=False, engine="openpyxl")  

                processed_files += 1
                status_text.text(f"✅ Processed ({i}/{total_files}): {uploaded_file.name}")

            except Exception as e:
                st.write(f"❌ Error processing {uploaded_file.name}: {e}")

            # 🔄 Update progress bar
            progress_bar.progress(i / total_files)  

            time.sleep(0.3)  

    except Exception as e:
        return f"❌ Error during processing: {e}"

    progress_bar.progress(1.0)  
    status_text.text("🎯 All files processed successfully!")

    return f"✅ Processed {processed_files}/{total_files} Excel files successfully!"

# Streamlit UI
st.set_page_config(page_title="Excel File Processor")

st.title("📂 Nielsen File Converter")
st.write("Upload multiple `.xls/.xlsx` files, and the script will process them.")

# 📂 File Uploader (Multiple Files)
uploaded_files = st.file_uploader("Upload Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)

if st.button("Start Processing"):
    if uploaded_files:
        result = process_uploaded_files(uploaded_files)
        st.success(result)
    else:
        st.error("Please upload at least one Excel file.")
