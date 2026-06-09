
import base64
import csv
import logging
import os
from pathlib import Path
import zipfile
import streamlit as st
import pandas as pd
import io
import traceback
import hashlib
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
from amal.processor import build_output_workbook as build_comm_generator_workbook
from amal.processor import process_uploaded_pairs as process_comm_generator_pairs
from lenovo_processor import build_lenovo_workbook

TOOL_OPTIONS = [
        "-- Select a tool --",
        "💻 Dell Invoice Extractor",
        "📦 Barcode PDF Generator grouped",
        "CI and Packing list - IBM",
        "Lenovo Huawei Starter"
    ]
tool = st.selectbox(
    "Select a tool:",
    TOOL_OPTIONS,
    key="tool_selector"
)

if tool == "Lenovo Huawei Starter":
    st.title("Lenovo Huawei Starter")
    st.caption("Start with one SOB PDF + one or more Huawei-style Excel files. The workbook uses the same layout style as the current Amal flow.")

    sob_file = st.file_uploader("Upload SOB PDF", type=["pdf"], key="lenovo_sob_pdf")
    excel_files = st.file_uploader(
        "Upload Huawei Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="lenovo_excel_files",
    )

    if st.button("Generate Lenovo workbook", key="lenovo_generate"):
        if sob_file and excel_files:
            with st.spinner("Preparing Lenovo workbook..."):
                workbook_bytes = build_lenovo_workbook(sob_file, list(excel_files))
            st.success("Lenovo workbook prepared successfully.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="Download Excel Workbook",
                data=workbook_bytes.getvalue(),
                file_name=f"lenovo_output_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="lenovo_download",
            )
        else:
            st.warning("Please upload one SOB PDF and at least one Huawei-style Excel file.")

if tool == "CI and Packing list - IBM":
    st.title("CI and Packing list - IBM")

    if "comm_generator_pair_count" not in st.session_state:
        st.session_state.comm_generator_pair_count = 1

    pair_inputs = []
    incomplete_pairs = []

    for index in range(st.session_state.comm_generator_pair_count):
        pair_number = index + 1
        st.subheader(f"Shipment Pair {pair_number}")
        left_col, right_col = st.columns(2)
        with left_col:
            sob_file = st.file_uploader(
                "Upload SOB PDF",
                type=["pdf"],
                key=f"comm_generator_sob_pdf_{index}",
            )
        with right_col:
            ibm_file = st.file_uploader(
                "Upload IBM PO / Commercial Invoice PDF",
                type=["pdf"],
                key=f"comm_generator_ibm_pdf_{index}",
            )

        if sob_file or ibm_file:
            if sob_file and ibm_file:
                pair_inputs.append((sob_file, ibm_file))
            else:
                incomplete_pairs.append(pair_number)

    if st.button("Add another pair", key="comm_generator_add_pair"):
        st.session_state.comm_generator_pair_count += 1
        st.rerun()

    if incomplete_pairs:
        st.error(
            "Each pair must include both files before workbook generation. "
            f"Incomplete pair(s): {', '.join(str(value) for value in incomplete_pairs)}"
        )

    if pair_inputs and not incomplete_pairs:
        with st.spinner("Preparing workbook..."):
            result = process_comm_generator_pairs(pair_inputs)
            workbook_bytes = build_comm_generator_workbook(result)

        st.success("Workbook prepared successfully.")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="Download Excel Workbook",
            data=workbook_bytes.getvalue(),
            file_name=f"output_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="comm_generator_download",
           
        )
    else:
        st.caption("Upload at least one complete SOB + IBM pair to generate the workbook.")