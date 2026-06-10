
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
from lenovo_processor import build_lenovo_ci_workbook, build_pack_list_zip

TOOL_OPTIONS = [
        "-- Select a tool --",
        "💻 Dell Invoice Extractor",
        "📦 Barcode PDF Generator grouped",
        "CI and Packing list - IBM",
        "Commercial Invoice and Packing List- Huawei"
    ]
tool = st.selectbox(
    "Select a tool:",
    TOOL_OPTIONS,
    key="tool_selector"
)

if tool == "Commercial Invoice and Packing List- Huawei":
    st.title("Commercial Invoice and Packing List- Huawei")
    st.caption("Use the three upload sections below: SOB first, then Huawei CI files (1–2 files), then PL files.")

    st.subheader("1) Upload SOB")
    sob_file = st.file_uploader("Upload SOB PDF", type=["pdf"], key="lenovo_sob_pdf")

    st.subheader("2) Upload Huawei CI files (DG + Non DG)")
    ci_files = st.file_uploader(
        "Upload Huawei CI files (min 1, max 2)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="lenovo_ci_files",
    )

    st.subheader("3) Upload PL files")
    pl_files = st.file_uploader(
        "Upload Packing List files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="lenovo_pl_files",
    )

    if "lenovo_generated" not in st.session_state:
        st.session_state.lenovo_generated = False
    if "lenovo_ci_outputs" not in st.session_state:
        st.session_state.lenovo_ci_outputs = []
    if "lenovo_pl_output" not in st.session_state:
        st.session_state.lenovo_pl_output = None
    if "lenovo_ci_names" not in st.session_state:
        st.session_state.lenovo_ci_names = []

    if st.button("Generate outputs", key="lenovo_generate"):
        if not sob_file:
            st.warning("Please upload the SOB PDF first.")
        elif not (1 <= len(ci_files) <= 2):
            st.warning("Please upload 1 to 2 Huawei CI files.")
        elif not pl_files:
            st.warning("Please upload at least one packing list file.")
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            with st.spinner("Preparing CI workbooks and PL zip..."):
                ci_workbooks = [
                    build_lenovo_ci_workbook(sob_file, ci_file, all_ci_files=list(ci_files))
                    for ci_file in ci_files
                ]
                pl_zip = build_pack_list_zip(list(pl_files), sob_file=sob_file)

            st.session_state.lenovo_generated = True
            st.session_state.lenovo_ci_outputs = [wb.getvalue() for wb in ci_workbooks]
            st.session_state.lenovo_pl_output = pl_zip.getvalue()
            st.session_state.lenovo_ci_names = [Path(file.name).stem for file in ci_files]

            st.success("CI workbooks and PL zip are ready.")

    if st.session_state.lenovo_generated and st.session_state.lenovo_ci_outputs:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        ci_names = st.session_state.get("lenovo_ci_names", [])

        for index, workbook_bytes in enumerate(st.session_state.lenovo_ci_outputs, start=1):
                ci_name = ci_names[index - 1] if index - 1 < len(ci_names) else f"ci_{index}"
                st.download_button(
                    label=f"Download CI Excel {index} - {ci_name}",
                    data=workbook_bytes,
                    file_name=f"lenovo_ci_{index}_{ci_name}_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"lenovo_ci_download_{index}",
                )

        st.download_button(
            label="Download PL ZIP",
            data=st.session_state.lenovo_pl_output,
            file_name=f"lenovo_pl_{timestamp}.zip",
            mime="application/zip",
            key="lenovo_pl_zip_download",
        )

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