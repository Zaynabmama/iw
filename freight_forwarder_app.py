from __future__ import annotations

from datetime import datetime

import streamlit as st

from freight_forwarder_processor import (
    JVConfig,
    create_excel_file,
    process_freight_forwarder_pdfs,
)


st.set_page_config(page_title="Freight Forwarder JV Tool", layout="wide")

st.markdown(
    """
    <style>
    [data-testid="stToolbar"],
    [data-testid="stHeaderActionElements"],
    .stAppDeployButton,
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    [data-testid="stFloatingActionButton"] {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Freight Forwarder JV Tool")

uploaded_files = st.file_uploader(
    "Upload PDF invoices",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 50:
        st.error("Please upload no more than 50 PDF files at a time.")
        st.stop()

    config = JVConfig()

    with st.spinner("Reading PDFs and building JV rows..."):
        output_df, parsed_invoices, errors = process_freight_forwarder_pdfs(uploaded_files, config)

    if not output_df.empty:
        output_buffer = create_excel_file(output_df)
        export_name = f"JV_UPLOAD_TEMPLATE_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

        st.success(f"Processed {len(output_df) // 2} invoice(s).")
        st.download_button(
            label="Download JV Upload File",
            data=output_buffer.getvalue(),
            file_name=export_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if errors:
        st.error("\n".join(errors))
