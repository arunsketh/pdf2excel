import streamlit as st
import pdfplumber
import openpyxl
import csv
import io

# --- App Configuration ---
st.set_page_config(page_title="PDF Data Extractor", layout="wide", page_icon="📄")

st.title("📄 PDF to Excel Converter")
st.markdown("Strictly designed for extracting tabular data from PDF files.")

# --- Sidebar ---
with st.sidebar:
    st.header("Extraction Settings")
    mode = st.radio("PDF Structure Type", ["Visual Tables (Grid)", "Embedded Text (Raw Data)"])
    
    st.divider()
    
    if mode == "Visual Tables (Grid)":
        st.caption("For PDFs with clear lines and borders.")
    else:
        st.caption("For PDFs containing raw CSV text.")
        separator_map = {"Comma (,)": ",", "Semicolon (;)": ";", "Tab (\\t)": "\t", "Pipe (|)": "|"}
        sep_choice = st.selectbox("Column Separator", options=list(separator_map.keys()))
        delimiter = separator_map[sep_choice]
        start_marker = st.text_input("Start Keyphrase (Optional)", help="Ignore text before this phrase.")

# --- Helper Functions ---
def clean_rows(rows):
    """Cleans newline characters from extracted data."""
    cleaned = []
    for row in rows:
        # Remove None values and newlines
        cleaned.append([str(cell).replace('\n', ' ').strip() if cell else "" for cell in row])
    return cleaned

def generate_excel(rows):
    """Generates an Excel file in memory without Pandas."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(output)
    return output.getvalue()

def parse_grid_pdf(file):
    """Extracts data using visual table lines."""
    all_data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_table()
            if tables:
                all_data.extend(tables)
    return clean_rows(all_data)

def parse_text_pdf(file, delimiter, marker):
    """Extracts and parses raw text data."""
    full_text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    
    # Slice text if marker is found
    if marker:
        idx = full_text.find(marker)
        if idx != -1:
            full_text = full_text[idx:]
            
    # Parse text as CSV
    f = io.StringIO(full_text)
    reader = csv.reader(f, delimiter=delimiter, skipinitialspace=True)
    rows = [row for row in reader if row]
    return clean_rows(rows)

# --- Main Execution ---
uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_file:
    st.divider()
    with st.spinner("Analyzing PDF structure..."):
        try:
            if mode == "Visual Tables (Grid)":
                data = parse_grid_pdf(uploaded_file)
            else:
                data = parse_text_pdf(uploaded_file, delimiter, start_marker)
            
            if data:
                st.success(f"Successfully extracted {len(data)} rows.")
                
                # Preview (First 5 rows)
                st.subheader("Preview")
                # Create flexible columns for preview
                if len(data) > 0:
                    cols = st.columns(len(data[0]))
                    for i, col in enumerate(cols):
                        if i < len(data[0]):
                            col.write(f"**{data[0][i]}**") # Header
                    
                    for row in data[1:6]:
                        cols = st.columns(len(data[0]))
                        for i, col in enumerate(cols):
                            if i < len(row):
                                col.write(row[i])

                # Download Button
                excel_file = generate_excel(data)
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_file,
                    file_name=f"{uploaded_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No data found. Try changing the 'PDF Structure Type' in the sidebar.")
                
        except Exception as e:
            st.error(f"Error processing file: {e}")
