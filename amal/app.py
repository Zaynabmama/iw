import io
from datetime import datetime

import streamlit as st

from processor import build_output_workbook, process_uploaded_pdfs


st.set_page_config(page_title="PDF to Excel", layout="wide")

st.title("PDF to Excel")
st.write(
    "Upload the SOB PDF and the IBM PO / commercial invoice PDF to generate "
    "an Excel workbook with `comm-inv` and `pack_list` sheets."
)

sob_file = st.file_uploader(
    "Upload SOB PDF",
    type=["pdf"],
    key="sob_pdf",
)

ibm_file = st.file_uploader(
    "Upload IBM PO / Commercial Invoice PDF",
    type=["pdf"],
    key="ibm_pdf",
)

if sob_file and ibm_file:
    with st.spinner("Preparing workbook..."):
        result = process_uploaded_pdfs(sob_file, ibm_file)
        workbook_bytes = build_output_workbook(result)

    st.success("Workbook prepared successfully.")
    st.warning("Debug panel is enabled in this build.")

    if result.messages:
        st.info("Current notes:")
        for message in result.messages:
            st.write(f"- {message}")

    

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    st.download_button(
        label="Download Excel Workbook",
        data=workbook_bytes.getvalue(),
        file_name=f"output_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


else:
    st.caption("Both PDF files are required before workbook generation.")
