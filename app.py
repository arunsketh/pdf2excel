import streamlit as st
import pdfplumber
import pandas as pd
import io

# --- Page Configuration ---
st.set_page_config(page_title="Universal PDF Converter", layout="wide", page_icon="📑")

st.title("📑 Universal PDF to Excel Converter")
st.markdown("""
This tool recovers data from PDFs and converts it into a clean Excel spreadsheet. 

**Choose your Extraction Mode in the sidebar:**
1. **Visual Grid:** Best for PDFs where data is inside clear table lines/borders.
2. **Raw Text / CSV:** Best for PDFs that contain raw data dumps (comma-separated text).
""")

# --- Sidebar Configuration ---
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # 1. Mode Selection
    mode = st.radio("Extraction Mode", ["Visual Grid (Standard)", "Raw Text / CSV (Data Dump)"])
    
    st.divider()
    
    # 2. Dynamic Settings based on Mode
    if mode == "Visual Grid (Standard)":
        st.subheader("Grid Settings")
        clean_newlines = st.checkbox("Remove newlines from cells", value=True, help="Merges multi-line text into a single line.")
        header_row = st.checkbox("First row is header", value=True)
        
    else: # Raw Text Mode
        st.subheader("Parsing Settings")
        separator = st.selectbox("Separator", options=[", (Comma)", "; (Semicolon)", "| (Pipe)", "\\t (Tab)"], index=0)
        quote_char = st.text_input("Quote Character", value='"', max_chars=1)
        start_keyword = st.text_input("Start Keyword (Optional)", help="Start extraction only after this word is found (e.g., 'Date' or 'ID').")

# --- Logic Functions ---

def process_visual_grid(pdf_file):
    """Extracts tables based on visual lines."""
    all_rows = []
    
    with pdfplumber.open(pdf_file) as pdf:
        progress_bar = st.progress(0)
        for i, page in enumerate(pdf.pages):
            # Extract table returns a list of lists
            tables = page.extract_table()
            if tables:
                all_rows.extend(tables)
            progress_bar.progress((i + 1) / len(pdf.pages))
            
    if not all_rows:
        return None
        
    df = pd.DataFrame(all_rows)
    
    # formatting
    if header_row:
        df.columns = df.iloc[0]
        df = df[1:]
    
    if clean_newlines:
        df = df.replace(r'\n', ' ', regex=True)
        
    return df

def process_raw_text(pdf_file, sep_char, quote, start_key):
    """Extracts raw text and parses it as a CSV."""
    full_text = ""
    
    # 1. Extract Text
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    
    # 2. Filter by Keyword (if provided)
    if start_key:
        start_idx = full_text.find(start_key)
        if start_idx != -1:
            full_text = full_text[start_idx:]
        else:
            st.warning(f"Keyword '{start_key}' not found. Using full text.")

    # 3. Parse
    try:
        # Resolve separator selection to actual character
        sep_map = {", (Comma)": ",", "; (Semicolon)": ";", "| (Pipe)": "|", "\\t (Tab)": "\t"}
        actual_sep = sep_map.get(sep_char, ",")
        
        df = pd.read_csv(
            io.StringIO(full_text),
            sep=actual_sep,
            quotechar=quote,
            skipinitialspace=True,
            on_bad_lines='skip' 
        )
        
        # Clean headers and values
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        df = df.replace(r'\n', ' ', regex=True)
        
        return df
    except Exception as e:
        st.error(f"Parsing Error: {e}")
        return None

# --- Main App Execution ---

uploaded_file = st.file_uploader("Upload your PDF", type="pdf")

if uploaded_file:
    st.info("Processing...")
    
    df_result = None
    
    # Route to correct function
    if mode == "Visual Grid (Standard)":
        df_result = process_visual_grid(uploaded_file)
    else:
        df_result = process_raw_text(uploaded_file, separator, quote_char, start_keyword)
        
    # Output
    if df_result is not None and not df_result.empty:
        st.success("Conversion Successful!")
        
        st.subheader("Data Preview")
        st.dataframe(df_result.head())
        
        # Prepare Excel File
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 Download Excel File",
            data=buffer.getvalue(),
            file_name=f"{uploaded_file.name.split('.')[0]}_converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data could be extracted. Try switching the 'Extraction Mode' in the sidebar.")
