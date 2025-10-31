import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

st.set_page_config(page_title="Dynamic Excel vs PDF Comparator", layout="wide")
st.title("Excel vs PDF/Text Comparator")
st.markdown("""
Upload any **Excel file** (master data) and **PDF/Text file** (to verify).  
This tool automatically converts the PDF to text and compares the content dynamically.
""")

# ---------- File Upload ----------
excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
pdf_file = st.file_uploader("Upload PDF or Text File", type=["pdf", "txt"])

if excel_file and pdf_file:
    try:
        # ---------- STEP 1: Load Excel ----------
        df_excel = pd.read_excel(excel_file)
        df_excel.columns = [str(c).strip().upper().replace("\n", " ").replace("  ", " ") for c in df_excel.columns]
        df_excel = df_excel.astype(str).apply(lambda x: x.str.strip())
        st.success("Excel file loaded successfully.")
        st.dataframe(df_excel.head())

        # ---------- STEP 2: Extract PDF as plain text ----------
        if pdf_file.name.lower().endswith(".pdf"):
            pdf_text = ""
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    pdf_text += (page.extract_text() or "") + "\n"
            if not pdf_text.strip():
                st.error("No text extracted from PDF. Try uploading a text version of the file.")
                st.stop()
        else:
            pdf_text = pdf_file.read().decode("utf-8")

        # Clean PDF text
        pdf_text = re.sub(r'\s+', ' ', pdf_text.strip()).upper()

        # ---------- STEP 3: Compare dynamically ----------
        st.info("Comparing Excel data with PDF text content dynamically...")

        mismatches = []
        total_checked = 0

        for idx, row in df_excel.iterrows():
            row_dict = row.to_dict()
            missing_fields = []
            found_all = True

            for col_name, value in row_dict.items():
                value_str = str(value).strip().upper()
                # Skip empty cells
                if not value_str or value_str in ["NAN", "NONE"]:
                    continue

                total_checked += 1
                # Check if this cell value is found in PDF text
                if value_str not in pdf_text:
                    missing_fields.append(col_name)
                    found_all = False

            if not found_all:
                mismatch_info = {"ROW_INDEX": idx + 1}
                for col in missing_fields:
                    mismatch_info[col] = row_dict[col]
                mismatches.append(mismatch_info)

        # ---------- STEP 4: Display Results ----------
        if not mismatches:
            st.success(f"All {total_checked} data values were found in the PDF/Text. No mismatches detected.")
        else:
            mismatch_df = pd.DataFrame(mismatches)
            st.warning(f"Found {len(mismatch_df)} rows where some Excel data were missing in the PDF/Text.")
            st.dataframe(mismatch_df)

            # Allow download
            csv_buffer = io.StringIO()
            mismatch_df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="📥 Download Mismatch Report (CSV)",
                data=csv_buffer.getvalue(),
                file_name="mismatched_data.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.info("Please upload both Excel and PDF/Text files to start the comparison.")
