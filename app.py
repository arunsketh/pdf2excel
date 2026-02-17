import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- Page Configuration ---
st.set_page_config(page_title="Op Snap PDF Converter", layout="wide")

st.title("📄 Op Snap PDF to Excel Converter")
st.markdown("""
This tool converts **Op Snap Monthly Publication** PDFs (containing CSV text) into clean Excel files.
""")

# --- Logic Function ---
def parse_pdf_to_df(uploaded_file):
    """
    Extracts text from the PDF and parses the embedded CSV data.
    """
    full_text = ""
    
    # 1. Extract text using pdfplumber
    with pdfplumber.open(uploaded_file) as pdf:
        # Update progress bar
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                full_text += text + "\n"
            progress_bar.progress((i + 1) / total_pages)
            
    # 2. Locate the start of the CSV data
    # We look for the header "Reportar" or the first quote character
    # to avoid capturing metadata at the top of the page.
    start_marker = '"Reportar'
    start_index = full_text.find(start_marker)

    if start_index != -1:
        csv_content = full_text[start_index:]
    else:
        st.warning("Specific header not found. Attempting to parse full text.")
        csv_content = full_text

    # 3. Parse into Pandas
    try:
        # skipinitialspace=True is crucial for the format seen in your snippets
        df = pd.read_csv(io.StringIO(csv_content), quotechar='"', skipinitialspace=True)
        
        # 4. Clean Data
        # The PDF contains newlines (\n) inside cells; we replace them with spaces
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        df = df.replace(r'\n', ' ', regex=True)
        
        return df
        
    except Exception as e:
        st.error(f"Error parsing CSV structure: {e}")
        return None

# --- Main GUI ---
uploaded_file = st.file_uploader("Upload your PDF file", type="pdf")

if uploaded_file is not None:
    st.info("File uploaded successfully. Processing...")
    
    # Run the extraction
    df = parse_pdf_to_df(uploaded_file)
    
    if df is not None:
        st.success("Conversion Successful!")
        
        # Show Preview
        st.subheader("Data Preview")
        st.dataframe(df.head())
        
        # Convert DF to Excel in memory
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # Download Button
        st.download_button(
            label="📥 Download Excel File",
            data=buffer.getvalue(),
            file_name=f"{uploaded_file.name.replace('.pdf', '')}_Converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
