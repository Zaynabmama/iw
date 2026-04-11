"""
MS Invoice Tool - Streamlit Application
Converts Excel files to MS Invoice format
"""

import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from msinvoice_processor import process_ms_invoice_file, validate_input_file, OUTPUT_HEADER
from msinvoice_srcl import create_ms_srcl_file

st.set_page_config(page_title="MS Invoice Tool", layout="wide")

st.title("MS Invoice Tool")

# Instructions
st.markdown(
    """
    <div style="
        padding: 18px 20px;
        background: linear-gradient(90deg, #fff3cd, #ffeeba);
        border: 2px solid #ffcc00;
        border-radius: 10px;
        font-weight: 700;
        color: #7a5a00;
        font-size: 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 12px;
    ">
        <span style="font-size: 18px;">📋 Instructions:</span>
        <span style="margin-left: 8px;">
        Upload your MS invoice Excel file and the tool will transform it to the required output format.
        </span>
    </div>
    """,
    unsafe_allow_html=True,
)

# File upload
uploaded_file = st.file_uploader(
    "Upload your MS Invoice Excel file", 
    type=["xlsx", "xls"],
    key="ms_invoice_upload"
)

if uploaded_file:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ File loaded: {uploaded_file.name} ({len(df)} rows)")
        
        # Validate input
        is_valid, validation_errors = validate_input_file(df)
        
        if not is_valid:
            st.error("❌ Validation failed:")
            for error in validation_errors:
                st.error(f"  • {error}")
            st.stop()
        
        # Process the file
        with st.spinner("Processing file..."):
            output_df, processing_errors = process_ms_invoice_file(df)
        
        # Display results
        if processing_errors:
            st.warning("⚠️ Some rows had processing issues:")
            for error in processing_errors[:10]:  # Show first 10 errors
                st.warning(f"  • {error}")
            if len(processing_errors) > 10:
                st.warning(f"  ... and {len(processing_errors) - 10} more")
        
        # Metrics
        col1, col2, col3 = st.columns(3)
        col1.metric("📊 Total Rows", len(output_df))
        col2.metric("✅ Successfully Processed", len(output_df) - len(processing_errors))
        col3.metric("❌ Failed Rows", len(processing_errors))

        gross_values = pd.to_numeric(output_df["Gross Value"], errors="coerce")
        positive_df = output_df[gross_values >= 0].copy()
        negative_df = output_df[gross_values < 0].copy()
        
        # Display preview
        st.subheader("Preview of Output")
        st.dataframe(positive_df.head(10), use_container_width=True)
        
        # Create Excel file for download
        output_buffer = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "MS INVOICE"
        
        # Write header
        ws.append(OUTPUT_HEADER)
        
        # Write data
        for _, row in positive_df.iterrows():
            row_list = [row.get(col, "") for col in OUTPUT_HEADER]
            ws.append(row_list)
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        # Download button
        st.download_button(
            label="⬇️ Download MS Invoice (Excel)",
            data=output_buffer.getvalue(),
            file_name="ms_invoice_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if not negative_df.empty:
            srcl_buffer = create_ms_srcl_file(negative_df)
            st.download_button(
                label="⬇️ Download SRCL File",
                data=srcl_buffer.getvalue(),
                file_name="ms_srcl_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        
        # Show full data option
        if st.checkbox("Show full dataset"):
            st.subheader("Complete Output Data")
            st.dataframe(positive_df, use_container_width=True)
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.error("Please ensure your file is in the correct Excel format.")
