import io
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook


def _get_mpc_billdate(document_location: str) -> str:
    doc_loc = str(document_location or "").strip().upper()
    if doc_loc in ["TC000", "UJ000"]:
        return "UAE - 28"
    if doc_loc == "QA000":
        return "QAR - 28"
    if doc_loc == "WT000":
        return "KWT - 28"
    return ""


def _clean_item_name(item_name: str) -> str:
    cleaned = re.sub(r"[\r\n]+", " ", str(item_name or "").strip())
    cleaned = cleaned.replace("'", "").replace('"', "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:240]


def create_ms_srcl_file(df: pd.DataFrame) -> io.BytesIO:
    """Create SRCL workbook from negative MS invoice rows only."""
    headers_head = [
        "S.No",
        "Date - (dd/MM/yyyy)",
        "Cust_Code",
        "Curr_Code",
        "FORM_CODE",
        "Doc_Src_Locn",
        "Location_Code",
        "Delivery_Location",
        "SalesmanID",
    ]

    headers_item = [
        "S.No",
        "Ref. Key",
        "Item_Code",
        "Item_Name",
        "Grade1",
        "Grade2",
        "UOM",
        "Qty",
        "Qty_Ls",
        "Rate",
        "CI Number CL",
        "End User CL",
        "Subs ID CL",
        "MPC Billdate CL",
        "Unit Cost CL",
        "Total",
    ]

    wb = Workbook()
    ws_head = wb.active
    ws_head.title = "SALES_RET_HEAD"
    ws_head.append(headers_head)

    today_str = datetime.today().strftime("%d/%m/%Y")
    header_sno_map = {}
    header_counter = 1

    for _, row in df.iterrows():
        invoice_no = str(row.get("Invoice No.", "")).strip()
        if invoice_no and invoice_no not in header_sno_map:
            header_sno_map[invoice_no] = header_counter
            ws_head.append([
                header_counter,
                today_str,
                row.get("Customer Code", ""),
                row.get("Currency Code", ""),
                "0",
                row.get("Document Location", ""),
                row.get("Document Location", ""),
                row.get("Delivery Location Code", ""),
                "ED068",
            ])
            header_counter += 1

    ws_item = wb.create_sheet(title="SALES_RET_ITEM")
    ws_item.append(headers_item)

    item_counter = 1
    for _, row in df.iterrows():
        invoice_no = str(row.get("Invoice No.", "")).strip()
        ref_key = header_sno_map.get(invoice_no, "")
        qty = abs(float(row.get("Quantity", 0) or 0))
        qty_ls = abs(float(row.get("Qty Loose", 0) or 0))
        rate = abs(float(row.get("Rate Per Qty", 0) or 0))
        unit_cost = abs(float(row.get("Cost", 0) or 0))
        total = abs(round(qty * rate, 2))

        ws_item.append([
            item_counter,
            ref_key,
            row.get("ITEM Code", ""),
            _clean_item_name(row.get("ITEM Name", "")),
            row.get("Grade code-1", ""),
            row.get("Grade code-2", ""),
            row.get("UOM", ""),
            qty,
            qty_ls,
            rate,
            invoice_no,
            row.get("End User", ""),
            row.get("Subscription Id", ""),
            _get_mpc_billdate(row.get("Document Location", "")),
            unit_cost,
            total,
        ])
        item_counter += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
