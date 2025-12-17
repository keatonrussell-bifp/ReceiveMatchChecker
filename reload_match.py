import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# --------------------------------------------------
# Extract LPNs from PDFs
# --------------------------------------------------
def extract_lpns_from_pdfs(pdf_files):
    lpns = set()
    lpn_pattern = re.compile(r"\b\d{8,12}\b")

    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            for match in lpn_pattern.findall(text):
                lpns.add(match)

    return lpns


# --------------------------------------------------
# Streamlit App
# --------------------------------------------------
st.set_page_config(page_title="Receive Match Checker", layout="wide")

st.title("📦 Receive Match Checker")
st.write(
    "Upload a container Excel file and one or more warehouse receipt PDFs. "
    "The app will match **PACKAGEID** against **LPNs** in the PDFs."
)

# Upload inputs
excel_file = st.file_uploader(
    "Upload Excel file (PACKAGEID on row 2)",
    type=["xlsx"]
)

pdf_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if excel_file and pdf_files:
    try:
        # Read Excel (header on row 2)
        df = pd.read_excel(excel_file, header=1, dtype=str).fillna("")
        df.columns = df.columns.astype(str).str.strip().str.upper()

        if "PACKAGEID" not in df.columns:
            st.error("Excel file must contain a 'PACKAGEID' column on row 2.")
            st.stop()

        # Extract LPNs
        lpns_in_pdfs = extract_lpns_from_pdfs(pdf_files)

        # Match logic
        pdf_lpn_col = []
        receive_match_col = []

        for pkg in df["PACKAGEID"]:
            pkg = str(pkg).strip()
            if pkg in lpns_in_pdfs:
                pdf_lpn_col.append(pkg)
                receive_match_col.append("YES")
            else:
                pdf_lpn_col.append("")
                receive_match_col.append("NO")

        # Insert new columns
        df["PDF LPN"] = pdf_lpn_col
        df["RECEIVE MATCH"] = receive_match_col

        st.success("Matching completed successfully.")
        st.dataframe(df, use_container_width=True)

        # Prepare download
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="⬇️ Download Excel with Receive Match",
            data=output,
            file_name=excel_file.name.replace(".xlsx", "_RECEIVE_MATCH.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload both an Excel file and at least one PDF.")
