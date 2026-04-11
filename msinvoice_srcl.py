import io
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook

SRCL_EXCHANGE_RATE_MAP = {
    "UJ000": 0.272294078,
    "TC000": 0.272294078,
    "QA000": 0.274725274725,
    "OM000": 2.60078023407,
    "KA000": 0.2666666666,
}


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


def _round_2(value: float) -> float:
    return round(float(value), 2)


def _normalize_date_key(value) -> str:
    if value is None or str(value).strip() == "":
        return ""
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return str(value).strip()


def build_kuwait_exchange_lookup(raw_df: pd.DataFrame) -> Tuple[Dict[str, float], List[str], List[str]]:
    """Build exact-date Kuwait exchange-rate lookup from positive rows in the uploaded file."""
    if raw_df.empty:
        return {}, [], []

    working_df = raw_df.copy()
    invoice_series = working_df.get("Invoice No.", pd.Series("", index=working_df.index)).fillna("").astype(str).str.strip()
    gross_series = pd.to_numeric(working_df.get("Gross Value", pd.Series(0, index=working_df.index)), errors="coerce")
    rate_series = pd.to_numeric(working_df.get("Exchange Rate", pd.Series(index=working_df.index)), errors="coerce")
    date_series = working_df.get("Invoice Date", pd.Series("", index=working_df.index)).apply(_normalize_date_key)

    is_kuwait = invoice_series.str.startswith("DNKW")
    positive_kwt = working_df[is_kuwait & (gross_series >= 0)].copy()
    positive_kwt["_date_key"] = date_series[positive_kwt.index]
    positive_kwt["_rate_value"] = rate_series[positive_kwt.index].apply(
        lambda rate: (1 / float(rate)) if pd.notna(rate) and float(rate) > 0 else None
    )

    lookup = {}
    ambiguous_dates = []

    for date_key, group in positive_kwt.groupby("_date_key"):
        valid_rates = sorted({round(float(rate), 10) for rate in group["_rate_value"].dropna() if float(rate) > 0})
        if len(valid_rates) == 1:
            lookup[date_key] = valid_rates[0]
        elif len(valid_rates) > 1:
            ambiguous_dates.append(date_key)

    negative_dates = sorted({
        date_series[idx]
        for idx in working_df[is_kuwait & (gross_series < 0)].index
        if date_series[idx]
    })

    return lookup, negative_dates, sorted(ambiguous_dates)


def _get_srcl_exchange_rate(
    document_location: str,
    source_invoice_date,
    kuwait_rate_lookup: Optional[Dict[str, float]] = None,
    kuwait_manual_rate: Optional[float] = None,
) -> float:
    doc_loc = str(document_location or "").strip().upper()
    if doc_loc == "WT000":
        if kuwait_manual_rate and kuwait_manual_rate > 0:
            return float(kuwait_manual_rate)
        date_key = _normalize_date_key(source_invoice_date)
        if kuwait_rate_lookup and date_key in kuwait_rate_lookup:
            return float(kuwait_rate_lookup[date_key])
        raise ValueError(
            "Kuwait SRCL exchange rate is missing. Provide a manual Kuwait rate or upload a file "
            "with a same-date Kuwait positive row that has a valid exchange rate."
        )
    return float(SRCL_EXCHANGE_RATE_MAP[doc_loc])


def _convert_usd_to_local(
    amount,
    document_location: str,
    source_invoice_date,
    kuwait_rate_lookup: Optional[Dict[str, float]] = None,
    kuwait_manual_rate: Optional[float] = None,
) -> float:
    rate = _get_srcl_exchange_rate(document_location, source_invoice_date, kuwait_rate_lookup, kuwait_manual_rate)
    amount_value = float(amount or 0)
    if rate == 0:
        return 0.0
    return amount_value / rate


def create_ms_srcl_file(
    df: pd.DataFrame,
    kuwait_rate_lookup: Optional[Dict[str, float]] = None,
    kuwait_manual_rate: Optional[float] = None,
) -> io.BytesIO:
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
        doc_loc = str(row.get("Document Location", "")).strip().upper()
        source_invoice_date = row.get("_Source Invoice Date", "")
        qty = abs(float(row.get("Quantity", 0) or 0))
        qty_ls = abs(float(row.get("Qty Loose", 0) or 0))
        local_gross = _round_2(
            _convert_usd_to_local(
                row.get("Gross Value", 0),
                doc_loc,
                source_invoice_date,
                kuwait_rate_lookup,
                kuwait_manual_rate,
            )
        )
        rate = abs(_round_2(local_gross / qty)) if qty else 0.0
        unit_cost = abs(_round_2(
            _convert_usd_to_local(
                row.get("Cost", 0),
                doc_loc,
                source_invoice_date,
                kuwait_rate_lookup,
                kuwait_manual_rate,
            )
        ))
        total = abs(_round_2(qty * rate))

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
            _get_mpc_billdate(doc_loc),
            unit_cost,
            total,
        ])
        item_counter += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
