import io
from datetime import datetime

import streamlit as st

from processor import build_output_workbook, process_uploaded_pairs


st.set_page_config(page_title="PDF to Excel", layout="wide")

st.title("Commercial Invoice and Packing List generation")

if "pair_count" not in st.session_state:
    st.session_state.pair_count = 1

pair_inputs = []
incomplete_pairs = []

for index in range(st.session_state.pair_count):
    pair_number = index + 1
    st.subheader(f"Shipment Pair {pair_number}")
    left_col, right_col = st.columns(2)
    with left_col:
        sob_file = st.file_uploader(
            "Upload SOB PDF",
            type=["pdf"],
            key=f"sob_pdf_{index}",
        )
    with right_col:
        ibm_file = st.file_uploader(
            "Upload IBM PO / Commercial Invoice PDF",
            type=["pdf"],
            key=f"ibm_pdf_{index}",
        )

    if sob_file or ibm_file:
        if sob_file and ibm_file:
            pair_inputs.append((sob_file, ibm_file))
        else:
            incomplete_pairs.append(pair_number)

if st.button("Add another pair"):
    st.session_state.pair_count += 1
    st.rerun()

if incomplete_pairs:
    st.error(
        "Each pair must include both files before workbook generation. "
        f"Incomplete pair(s): {', '.join(str(value) for value in incomplete_pairs)}"
    )

if pair_inputs and not incomplete_pairs:
    with st.spinner("Preparing workbook..."):
        result = process_uploaded_pairs(pair_inputs)
        workbook_bytes = build_output_workbook(result)

    st.success("Workbook prepared successfully.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    st.download_button(
        label="Download Excel Workbook",
        data=workbook_bytes.getvalue(),
        file_name=f"output_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
