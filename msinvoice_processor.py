"""
MS Invoice Processor - Core business logic for MS invoice transformation
Handles data transformation, validations, and calculations
"""
 
import pandas as pd
import numpy as np
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, Tuple, Optional
import logging

logger = logging.getLogger(__name__)

# === Configuration Maps ===

INVOICE_PREFIX_MAP = {
    "DNKW": "WT000",
    "DNFZ": "UJ000",
    "DNQA": "QA000",
    "DNOM": "OM000",
    "DNSA": "KA000",
    "DNAE": "TC000"
}

INVOICE_CURRENCY_MAP = {
    "DNKW": "KWD",
    "DNFZ": "USD",
    "DNQA": "QAR",
    "DNOM": "OMR",
    "DNSA": "SAR",
    "DNAE": "AED"
}

DELIVERY_LOCATION_MAP = {
    "KA000": "KA200",
    "UJ000": "UJ200",
    "QA000": "QA200",
    "WT000": "WT200",
    "TC000": "TC200",
    "OM000": "OM200"
}

TAX_CODE_MAP = {
    "WT000": "", 
    "QA000": "", 
    "TC000": "SLVAT5",
    "OM000": "SLVAT5", 
    "UJ000": "SEVAT0", 
    "KA000": "SLVAT15"
}

TAX_PERCENT_MAP = {
    "WT000": "", 
    "QA000": "", 
    "TC000": 5,
    "OM000": 5, 
    "UJ000": 0, 
    "KA000": 15
}

CURRENCY_MAP = {
    "WT000": "", 
    "QA000": "", 
    "TC000": "AED",
    "OM000": "OMR", 
    "UJ000": "USD", 
    "KA000": "SAR"
}

EXCHANGE_RATE_MAP = {
    "UJ000": 1,
    "TC000": 0.272294078,
    "QA000": 0.274725274725,
    "OM000": 2.60078023407,
    "KA000": 0.2666666666
}

KEYWORD_MAP = {
    ("windows server", "window server","Office LTSC Standard", "MSPER-CNS"): "MSPER-CNS",
    ("azure subscription", "MSAZ-CNS"): "MSAZ-CNS",
    ("google workspace", "GL-WSP-CNS"): "GL-WSP-CNS",
    ("m365", "microsoft 365", "office 365", "exchange online", "Microsoft Defender for Endpoint P1", "MS-CNS"): "MS-CNS",
    ("POWERPLATFORM - Power Apps Premium (New Commerce)", "powerapps premium", "power apps premium", "Power Apps Premium", "MS-CNS"): "MS-CNS",
    ("POWERPLATFORM - Power Automate per user plan (New Commerce)", "power automate per user", "Power Automate per user", "MS-CNS"): "MS-CNS",
    ("Excel LTSC 2024", "excel ltsc", "MSPER-CNS"): "MSPER-CNS",
    ("Project Professional 2024 (Commercial) (Subs ID)", "project professional 2024", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 - 1 User CAL (Commercial)", "sql server 2025 - 1 user cal", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 Enterprise core - 2 core License Pack (Commercial)", "sql server 2025 enterprise core", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 Standard edition Perpetual 1 Server License (Commercial)", "sql server 2025 standard edition perpetual 1 server license", "MSPER-CNS"): "MSPER-CNS",
    ("Visual Studio Professional 2026 (Commercial)", "visual studio professional 2026", "MSPER-CNS"): "MSPER-CNS",
    ("Windows 11 Enterprise LTSC 2024 Upgrade (Commercial)", "windows 11 enterprise ltsc 2024 upgrade", "MSPER-CNS"): "MSPER-CNS",
    ("Azure Plan Reserved Instances", "azure plan reserved instances", "MSRI-CNS"): "MSRI-CNS",
    ("acronis", "AS-CNS"): "AS-CNS",
    ("windows 11 pro", "MSPER-CNS"): "MSPER-CNS",
    ("power bi", "MS-CNS"): "MS-CNS",
    ("planner", "project plan", "MS-CNS"): "MS-CNS",
    ("power automate premium", "MS-CNS"): "MS-CNS",
    ("visio", "MS-CNS"): "MS-CNS",
    ("Microsoft Entra ID Governance (Education Faculty Pricing)", "Power Apps Premium (Non-Profit Pricing)", "MS-CNS"): "MS-CNS",
    ("MSRI-CNS",): "MSRI-CNS",
    ("dynamics 365", "MS-CNS"): "MS-CNS",
    ("AWS Account", "AWS"): "AWS-UTILITIES-CNS",
    ("minecraft education per user", "MS-CNS"): "MS-CNS",
}

OUTPUT_HEADER = [
    "Invoice No.", "Customer Code", "Customer Name", "Invoice Date", "Document Location",
    "Sale Location", "Delivery Location Code", "Delivery Date", "Annotation", "Currency Code",
    "Exchange Rate", "Shipment Mode", "Payment Term", "Mode Of Payment", "Status",
    "Credit Card Transaction No.", "HEADER Discount Code", "HEADER Discount %", "HEADER Currency", "HEADER Basis",
    "HEADER Disc Value", "HEADER Expense Code", "HEADER Expense %", "HEADER Expense Currency", "HEADER Expense Basis",
    "HEADER Expense Value", "Subscription Id", "Billing Cycle Start Date", "Billing Cycle End Date",
    "ITEM Code", "ITEM Name", "UOM", "Grade code-1", "Grade code-2", "Quantity", "Qty Loose",
    "Rate Per Qty", "Gross Value", "ITEM Discount Code", "ITEM Discount %", "ITEM Discount Currency", "ITEM Discount Basis",
    "ITEM Disc Value", "ITEM Expense Code", "ITEM Expense %", "ITEM Expense Currency", "ITEM Expense Basis",
    "ITEM Expense Value", "ITEM Tax Code", "ITEM Tax %", "ITEM Tax Currency", "ITEM Tax Basis", "ITEM Tax Value",
    "LPO Number", "End User", "Cost"
]


# === Helper Functions ===

def normalize_input_column_name(column_name) -> str:
    """Normalize uploaded headers so minor Excel formatting differences do not break matching."""
    if pd.isna(column_name):
        return ""
    normalized = str(column_name).replace("\ufeff", " ")
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized


def standardize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known input columns to their canonical names using normalized header matching."""
    canonical_columns = [
        "Invoice No.",
        "Customer Code",
        "Customer Name",
        "Currency Code",
        "Invoice Type",
        "Payment Method",
        "MS Subscription ID",
        "Billing Cycle Start Date",
        "Billing Cycle End Date",
        "Charge Description",
        "Quantity",
        "LPO Number",
        "End User",
        "End Customer Country",
        "Invoice Date",
        "Exchange Rate",
        "Gross Value",
        "Unit Cost",
        "ITEM Code",
    ]

    normalized_to_canonical = {
        normalize_input_column_name(column): column for column in canonical_columns
    }
    alias_columns = {
        normalize_input_column_name("Azure Consumption Description"): "Charge Description",
        normalize_input_column_name("End User Country"): "End Customer Country",
    }
    normalized_to_canonical.update(alias_columns)
    rename_map = {}

    for column in df.columns:
        canonical_name = normalized_to_canonical.get(normalize_input_column_name(column))
        if canonical_name and column != canonical_name:
            rename_map[column] = canonical_name

    if not rename_map:
        df = df.copy()
    else:
        df = df.rename(columns=rename_map)

    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]  # Keep first occurrence of duplicate columns

    return df


def get_scalar_value(value, default=""):
    """Return a single scalar from pandas Series/ndarray if needed."""
    if isinstance(value, pd.Series):
        if value.empty:
            return default
        return value.iloc[0]
    if isinstance(value, np.ndarray):
        if value.size == 0:
            return default
        return value.flat[0]
    return value


def drop_last_input_row(df: pd.DataFrame) -> pd.DataFrame:
    """Drop the last row only when its Invoice No. is empty."""
    if df.empty:
        return df
    if "Invoice No." not in df.columns:
        return df

    last_invoice_no = clean_text_value(get_scalar_value(df.iloc[-1].get("Invoice No.", "")))
    if last_invoice_no == "":
        return df.iloc[:-1].copy()
    return df

def get_document_location(invoice_no: str) -> str:
    """Extract Document Location from Invoice No. prefix"""
    if pd.isna(invoice_no):
        return ""
    
    invoice_str = str(invoice_no).strip()
    for prefix, location in INVOICE_PREFIX_MAP.items():
        if invoice_str.startswith(prefix):
            return location
    
    return ""


def get_currency_from_invoice_no(invoice_no: str) -> str:
    """Extract Currency Code from Invoice No. prefix"""
    if pd.isna(invoice_no):
        return ""

    invoice_str = str(invoice_no).strip()
    for prefix, currency in INVOICE_CURRENCY_MAP.items():
        if invoice_str.startswith(prefix):
            return currency

    return ""


def get_delivery_location_code(document_location: str) -> str:
    """Map Document Location to Delivery Location Code"""
    if pd.isna(document_location):
        return ""

    return DELIVERY_LOCATION_MAP.get(str(document_location).strip(), "")


def extract_payment_term(payment_method: str) -> str:
    """Extract payment term number from Payment Method column"""
    if pd.isna(payment_method):
        return ""
    
    method_str = str(payment_method).strip()
    
    # Look for Net 30, Net 60, Net 90 patterns (case-sensitive)
    patterns = {
        "Net 30": "30",
        "Net 60": "60",
        "Net 90": "90"
    }
    
    for pattern, value in patterns.items():
        if pattern in method_str:
            return value
    
    return method_str


def get_exchange_rate(document_location: str, uploaded_exchange_rate: float = None) -> float:
    """Calculate exchange rate based on Document Location"""
    if document_location == "WT000":
        # Special case for Kuwait: 1 / uploaded Exchange Rate
        if uploaded_exchange_rate and uploaded_exchange_rate > 0:
            return 1 / uploaded_exchange_rate
        return ""
    
    return EXCHANGE_RATE_MAP.get(document_location, "")


def find_column_with_prefix(df: pd.DataFrame, prefix: str) -> Optional[str]:
    """Find column name that starts with given prefix"""
    for col in df.columns:
        if str(col).startswith(prefix):
            return col
    return None


def get_item_code(mapping_source: str) -> str:
    """Map Charge Description text to output ITEM Code using substring matching."""
    if pd.isna(mapping_source):
        return ""
    
    item_code_value = str(mapping_source).strip().upper()
    
    for keywords, code in KEYWORD_MAP.items():
        for keyword in keywords:
            if str(keyword).strip().upper() in item_code_value:
                return code
    
    return ""


def round_to_2_decimals(value):
    """Round value to 2 decimals using Excel-style half-up rounding."""
    try:
        if pd.isna(value):
            return ""
        decimal_value = Decimal(str(value))
        return float(decimal_value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except (ValueError, TypeError, InvalidOperation):
        return ""


def calculate_rate_per_qty(gross_value_value, quantity):
    """Calculate Rate Per Qty as Gross Value / Quantity, keeping empty if Quantity is 0."""
    try:
        if pd.isna(quantity) or quantity == 0:
            return ""

        gross_value = float(gross_value_value)
        qty = float(quantity)

        if qty == 0:
            return ""

        return gross_value / qty
    except (ValueError, TypeError, ZeroDivisionError):
        return ""


def calculate_cost_per_qty(cost_value, quantity):
    """Calculate output Cost as cost divided by output Quantity, rounded half-up to 2 decimals."""
    try:
        if pd.isna(cost_value) or pd.isna(quantity) or quantity == "" or quantity == 0:
            return ""

        cost = Decimal(str(cost_value))
        qty = Decimal(str(quantity))

        if qty == 0:
            return ""

        return float((cost / qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except (ValueError, TypeError, InvalidOperation, ZeroDivisionError):
        return ""


def calculate_gross_value(rate_per_qty, exchange_rate, quantity):
    """Calculate Gross Value = ROUND(ROUND(rate_per_qty * exchange_rate, 2) * quantity, 2) using half-up rounding."""
    try:
        if pd.isna(rate_per_qty) or pd.isna(exchange_rate) or pd.isna(quantity):
            return ""

        r = Decimal(str(rate_per_qty))
        e = Decimal(str(exchange_rate))
        q = Decimal(str(quantity))

        inner = (r * e).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        gross = (inner * q).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return float(gross)
    except (ValueError, TypeError, InvalidOperation):
        return ""


def is_azure_consumption_description(value) -> bool:
    """Detect Azure consumption description text used to identify Azure invoice groups."""
    if pd.isna(value):
        return False
    text_value = str(value).strip().lower()
    if text_value == "":
        return False
    return any(keyword in text_value for keyword in [
        "azure plan",
        "azure consumption",
        "azure usage",
        "azure subscription",
    ])


def build_azure_group_keys(df: pd.DataFrame) -> set:
    """Find invoice+subscription groups that should be consolidated as Azure."""
    group_keys = set()
    if "Invoice No." not in df.columns or "MS Subscription ID" not in df.columns or "Charge Description" not in df.columns:
        logger.debug(
            "Azure detection skipped because required columns are missing. Columns present: %s",
            list(df.columns),
        )
        return group_keys

    for _, row in df.iterrows():
        invoice_no = clean_text_value(get_scalar_value(row.get("Invoice No.", "")))
        ms_sub_id = clean_text_value(get_scalar_value(row.get("MS Subscription ID", "")))
        charge_description = get_scalar_value(row.get("Charge Description", ""))
        is_azure_row = (
            invoice_no and
            ms_sub_id and
            is_azure_consumption_description(charge_description)
        )
        logger.debug(
            "Azure detection row: invoice=%s subscription=%s charge_description=%s matched=%s",
            invoice_no,
            ms_sub_id,
            charge_description,
            bool(is_azure_row),
        )
        if is_azure_row:
            group_keys.add((invoice_no, ms_sub_id))

    logger.debug("Azure group keys built: %s", sorted(group_keys))
    return group_keys


def sum_group_gross_values(group_df: pd.DataFrame) -> Decimal:
    """Sum input Gross Value from a group using Decimal precision."""
    total = Decimal("0")
    for row_index, value in group_df.get("Gross Value", []).items():
        if pd.isna(value) or str(value).strip() == "":
            logger.debug("Azure gross sum skipped empty value at source row index=%s", row_index)
            continue
        try:
            total += Decimal(str(value))
            logger.debug(
                "Azure gross sum added source row index=%s gross_value=%s running_total=%s",
                row_index,
                value,
                total,
            )
        except (ValueError, TypeError, InvalidOperation):
            logger.debug(
                "Azure gross sum skipped invalid value at source row index=%s gross_value=%s",
                row_index,
                value,
            )
            continue
    return total


def get_row_cost_value(row: pd.Series, cost_col: Optional[str]):
    """Return the source cost value for one row from Total Cost Transaction columns."""
    if cost_col:
        cost_value = round_to_2_decimals(get_scalar_value(row.get(cost_col, "")))
        logger.debug(
            "Azure cost source row index=%s cost_column=%s raw_value=%s rounded_value=%s",
            getattr(row, "name", None),
            cost_col,
            get_scalar_value(row.get(cost_col, "")),
            cost_value,
        )
        return cost_value
    logger.debug(
        "Azure cost source missing Total Cost Transaction column for source row index=%s",
        getattr(row, "name", None),
    )
    return ""


def sum_group_cost_values(group_df: pd.DataFrame, cost_col: Optional[str]) -> Decimal:
    """Sum source cost values from a group using Decimal precision."""
    total = Decimal("0")
    for _, group_row in group_df.iterrows():
        value = get_row_cost_value(group_row, cost_col)
        if pd.isna(value) or str(value).strip() == "":
            logger.debug(
                "Azure cost sum skipped empty value at source row index=%s",
                getattr(group_row, "name", None),
            )
            continue
        try:
            total += Decimal(str(value))
            logger.debug(
                "Azure cost sum added source row index=%s cost_value=%s running_total=%s",
                getattr(group_row, "name", None),
                value,
                total,
            )
        except (ValueError, TypeError, InvalidOperation):
            logger.debug(
                "Azure cost sum skipped invalid value at source row index=%s cost_value=%s",
                getattr(group_row, "name", None),
                value,
            )
            continue
    return total


def format_date_only(value) -> str:
    """Return a date object for date-like values, without timestamps."""
    value = get_scalar_value(value)
    if pd.isna(value):
        return ""
    text_value = str(value).strip()
    if text_value.lower() in ["", "nan", "none", "nat"]:
        return ""
    try:
        return pd.to_datetime(value).date()
    except (ValueError, TypeError):
        return ""


def clean_text_value(value) -> str:
    """Return empty string for blank/NaN-like text values."""
    value = get_scalar_value(value)
    if pd.isna(value):
        return ""
    text_value = str(value).strip()
    if text_value.lower() in ["", "nan", "none", "nat"]:
        return ""
    return text_value


def find_blank_rows(df: pd.DataFrame, column_name: str) -> list:
    """Return 1-based Excel row numbers for blank/NaN-like values in a required text column."""
    if column_name not in df.columns:
        return []

    blank_rows = []
    for idx, value in df[column_name].items():
        if clean_text_value(value) == "":
            blank_rows.append(idx + 2)
    return blank_rows


def build_end_user_value(end_user, end_customer_country) -> str:
    """Combine End User and End Customer Country without emitting NaN-like text."""
    end_user_value = clean_text_value(end_user)
    country_value = clean_text_value(end_customer_country)

    if end_user_value and country_value:
        return f"{end_user_value} ; {country_value}"
    if end_user_value:
        return end_user_value
    if country_value:
        return country_value
    return ""


def calculate_tax_value(gross_value: float, tax_percent: float) -> str:
    """Calculate Tax Value = Gross Value * Tax %"""
    try:
        if pd.isna(gross_value) or pd.isna(tax_percent) or tax_percent == "":
            return ""
        
        gv = float(gross_value)
        tp = float(tax_percent) if isinstance(tax_percent, str) else float(tax_percent)
        
        return round(gv * tp / 100, 2)
    except (ValueError, TypeError):
        return ""


def is_negative_credit_note(row: pd.Series) -> bool:
    """Use Invoice Type only to detect credit-note rows."""
    invoice_type = clean_text_value(row.get("Invoice Type", "")).lower()
    if invoice_type == "credit invoice":
        return True
    if invoice_type == "debit invoice":
        return False
    raise ValueError(
        "Invoice Type must be either 'Credit Invoice' or 'Debit Invoice'."
    )


def apply_invoice_number_versioning(output_df: pd.DataFrame) -> pd.DataFrame:
    """Overwrite Invoice No. with versioned values based on Invoice/LPO/End User groups"""
    if output_df.empty:
        return output_df

    version_df = output_df.copy()
    version_df["_original_invoice_no"] = version_df["Invoice No."].fillna("").astype(str).str.strip()
    version_df["_lpo_number"] = version_df["LPO Number"].fillna("").astype(str).str.strip()
    version_df["_end_user"] = version_df["End User"].fillna("").astype(str).str.strip()
    version_df["_group_key"] = (
        version_df["_original_invoice_no"] +
        version_df["_lpo_number"] +
        version_df["_end_user"]
    )

    unique_groups = version_df[["_original_invoice_no", "_group_key"]].drop_duplicates().reset_index(drop=True)
    unique_groups["_version_no"] = unique_groups.groupby("_original_invoice_no").cumcount() + 1
    version_map = dict(zip(unique_groups["_group_key"], unique_groups["_version_no"]))

    version_df["Invoice No."] = version_df.apply(
        lambda row: (
            f'{row["_original_invoice_no"]}-{version_map.get(row["_group_key"], 1)}'
            if row["_original_invoice_no"] else ""
        ),
        axis=1
    )

    return version_df.drop(columns=["_original_invoice_no", "_lpo_number", "_end_user", "_group_key"])


# === Main Processing Function ===

def process_ms_invoice_file(df: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    """
    Transform input Excel file to MS Invoice output format
    
    Args:
        df: Input DataFrame from Excel file
    
    Returns:
        Tuple of (output_df, errors_list)
    """
    df = standardize_input_columns(df)
    df = drop_last_input_row(df)
    azure_group_keys = build_azure_group_keys(df)
    processed_azure_groups = set()
    errors = []
    output_rows = []
    today = datetime.today().date()
    logger.debug(
        "MS invoice processing started: input_rows=%s azure_group_count=%s azure_groups=%s",
        len(df),
        len(azure_group_keys),
        sorted(azure_group_keys),
    )
    
    for idx, row in df.iterrows():
        try:
            out_row = {}
            
            # Get Invoice No. (as-is from input)
            invoice_no = str(get_scalar_value(row.get("Invoice No.", ""))).strip()
            out_row["Invoice No."] = invoice_no
            
            # Document Location from Invoice No. prefix
            doc_location = get_document_location(invoice_no)
            out_row["Document Location"] = doc_location
            out_row["Sale Location"] = doc_location
            out_row["Delivery Location Code"] = get_delivery_location_code(doc_location)
            
            # Customer info (as-is from input)
            out_row["Customer Code"] = clean_text_value(get_scalar_value(row.get("Customer Code", "")))
            out_row["Customer Name"] = clean_text_value(get_scalar_value(row.get("Customer Name", "")))
            out_row["_Invoice Type"] = clean_text_value(get_scalar_value(row.get("Invoice Type", "")))
            
            # Dates
            out_row["_Source Invoice Date"] = get_scalar_value(row.get("Invoice Date", ""))
            out_row["Invoice Date"] = today
            out_row["Delivery Date"] = today
            out_row["Annotation"] = ""
            
            # Currency from Invoice No. prefix
            out_row["Currency Code"] = get_currency_from_invoice_no(invoice_no)
            raw_exchange_rate_input = get_scalar_value(row.get("Exchange Rate", ""))
            uploaded_exchange_rate = None
            try:
                uploaded_exchange_rate = float(raw_exchange_rate_input)
            except (ValueError, TypeError):
                uploaded_exchange_rate = None
            
            exchange_rate = get_exchange_rate(doc_location, uploaded_exchange_rate)
            output_exchange_rate = EXCHANGE_RATE_MAP.get(doc_location, "")
            out_row["Exchange Rate"] = output_exchange_rate if output_exchange_rate != "" else ""
            
            # Fixed fields
            out_row["Shipment Mode"] = "EML"
            out_row["Payment Term"] = extract_payment_term(str(get_scalar_value(row.get("Payment Method", ""))))
            out_row["Mode Of Payment"] = "OC"
            out_row["Status"] = "Unpaid"
            out_row["Credit Card Transaction No."] = ""
            
            # Header discount/expense fields (blank)
            out_row["HEADER Discount Code"] = ""
            out_row["HEADER Discount %"] = ""
            out_row["HEADER Currency"] = ""
            out_row["HEADER Basis"] = ""
            out_row["HEADER Disc Value"] = ""
            out_row["HEADER Expense Code"] = ""
            out_row["HEADER Expense %"] = ""
            out_row["HEADER Expense Currency"] = ""
            out_row["HEADER Expense Basis"] = ""
            out_row["HEADER Expense Value"] = ""
            
            # Subscription
            ms_sub_id = clean_text_value(get_scalar_value(row.get("MS Subscription ID", "")))
            subscription_id_value = ms_sub_id if ms_sub_id else "Subs ID"
            out_row["Subscription Id"] = subscription_id_value
            out_row["Billing Cycle Start Date"] = format_date_only(get_scalar_value(row.get("Billing Cycle Start Date", "")))
            out_row["Billing Cycle End Date"] = format_date_only(get_scalar_value(row.get("Billing Cycle End Date", "")))
            
            cost_col = find_column_with_prefix(df, "Total Cost Transaction Currency")

            # ITEM Code mapped from Charge Description
            charge_desc = clean_text_value(get_scalar_value(row.get("Charge Description", "")))
            invoice_no_key = str(get_scalar_value(row.get("Invoice No.", ""))).strip()
            group_key = (invoice_no_key, subscription_id_value)
            logger.debug(
                "Processing source row index=%s invoice=%s subscription=%s charge_description=%s azure_group_match=%s",
                idx,
                invoice_no_key,
                subscription_id_value,
                charge_desc,
                group_key in azure_group_keys,
            )

            if group_key in azure_group_keys:
                # Consolidate all same invoice + subscription Azure rows into one output row
                if group_key in processed_azure_groups:
                    logger.debug(
                        "Azure group already processed, skipping source row index=%s group_key=%s",
                        idx,
                        group_key,
                    )
                    continue

                processed_azure_groups.add(group_key)
                group_rows = df[
                    df["Invoice No."].astype(str).str.strip().eq(invoice_no_key) &
                    df["MS Subscription ID"].astype(str).str.strip().eq(subscription_id_value)
                ]
                logger.debug(
                    "Azure group start: group_key=%s source_row_indexes=%s row_count=%s",
                    group_key,
                    list(group_rows.index),
                    len(group_rows),
                )

                azure_rows = group_rows[group_rows["Charge Description"].apply(is_azure_consumption_description)]
                charge_desc = clean_text_value(
                    azure_rows.iloc[0]["Charge Description"]
                    if not azure_rows.empty
                    else group_rows.iloc[0].get("Charge Description", "")
                )
                charge_desc = charge_desc or "Azure plan"
                logger.debug(
                    "Azure group description resolved: group_key=%s azure_row_indexes=%s final_charge_description=%s",
                    group_key,
                    list(azure_rows.index),
                    charge_desc,
                )
                item_code = "MSAZ-CNS"
                out_row["ITEM Code"] = item_code
                out_row["ITEM Name"] = charge_desc + (f" ({subscription_id_value})" if subscription_id_value else "")

                out_row["UOM"] = "NOS"
                out_row["Grade code-1"] = "NA"
                out_row["Grade code-2"] = "NA"

                out_row["Quantity"] = 1
                out_row["Qty Loose"] = 0

                sum_gross = sum_group_gross_values(group_rows)
                actual_exchange_rate = raw_exchange_rate_input if raw_exchange_rate_input != "" else exchange_rate
                logger.debug(
                    "Azure gross preparation: group_key=%s summed_input_gross=%s raw_exchange_rate_input=%s fallback_exchange_rate=%s actual_exchange_rate=%s",
                    group_key,
                    sum_gross,
                    raw_exchange_rate_input,
                    exchange_rate,
                    actual_exchange_rate,
                )
                if sum_gross != Decimal("0") and actual_exchange_rate != "":
                    gross_value = calculate_gross_value(sum_gross, actual_exchange_rate, 1)
                    out_row["Gross Value"] = gross_value
                    out_row["Rate Per Qty"] = gross_value
                    logger.debug(
                        "Azure gross final: group_key=%s quantity=1 gross_value=%s rate_per_qty=%s",
                        group_key,
                        gross_value,
                        gross_value,
                    )
                else:
                    out_row["Gross Value"] = ""
                    out_row["Rate Per Qty"] = ""
                    logger.debug(
                        "Azure gross final: group_key=%s gross calculation skipped because summed_input_gross=%s actual_exchange_rate=%s",
                        group_key,
                        sum_gross,
                        actual_exchange_rate,
                    )

                sum_cost = sum_group_cost_values(group_rows, cost_col)
                rounded_group_cost = float(sum_cost.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)) if sum_cost != Decimal("0") else ""
                out_row["Cost"] = calculate_cost_per_qty(rounded_group_cost, out_row["Quantity"])
                logger.debug(
                    "Azure cost final: group_key=%s cost_column=%s summed_cost=%s output_cost=%s",
                    group_key,
                    cost_col,
                    sum_cost,
                    out_row["Cost"],
                )
                logger.debug(
                    "Azure output row ready: group_key=%s item_code=%s item_name=%s quantity=%s gross_value=%s rate_per_qty=%s cost=%s",
                    group_key,
                    out_row.get("ITEM Code", ""),
                    out_row.get("ITEM Name", ""),
                    out_row.get("Quantity", ""),
                    out_row.get("Gross Value", ""),
                    out_row.get("Rate Per Qty", ""),
                    out_row.get("Cost", ""),
                )
            else:
                item_code = get_item_code(charge_desc)
                out_row["ITEM Code"] = item_code
                out_row["ITEM Name"] = charge_desc + (f" ({subscription_id_value})" if subscription_id_value else "")

                out_row["UOM"] = "NOS"
                out_row["Grade code-1"] = "NA"
                out_row["Grade code-2"] = "NA"

                quantity = get_scalar_value(row.get("Quantity", 0))
                out_row["Quantity"] = quantity
                out_row["Qty Loose"] = 0

                rate_per_qty_input = get_scalar_value(row.get("Rate Per Qty", ""))
                exchange_rate_input = get_scalar_value(row.get("Exchange Rate", ""))
                if rate_per_qty_input != "" and exchange_rate_input != "" and quantity != "" and quantity != 0:
                    gross_value = calculate_gross_value(rate_per_qty_input, exchange_rate_input, quantity)
                    out_row["Gross Value"] = gross_value
                    rate_per_qty = calculate_rate_per_qty(gross_value, quantity)
                    out_row["Rate Per Qty"] = rate_per_qty
                else:
                    out_row["Gross Value"] = ""
                    out_row["Rate Per Qty"] = ""
            
            # ITEM Discount fields (blank)
            out_row["ITEM Discount Code"] = ""
            out_row["ITEM Discount %"] = ""
            out_row["ITEM Discount Currency"] = ""
            out_row["ITEM Discount Basis"] = ""
            out_row["ITEM Disc Value"] = ""
            
            # ITEM Expense fields (blank)
            out_row["ITEM Expense Code"] = ""
            out_row["ITEM Expense %"] = ""
            out_row["ITEM Expense Currency"] = ""
            out_row["ITEM Expense Basis"] = ""
            out_row["ITEM Expense Value"] = ""
            
            # ITEM Tax fields (from Document Location mapping)
            out_row["ITEM Tax Code"] = TAX_CODE_MAP.get(doc_location, "")
            tax_percent = TAX_PERCENT_MAP.get(doc_location, "")
            out_row["ITEM Tax %"] = tax_percent
            out_row["ITEM Tax Currency"] = CURRENCY_MAP.get(doc_location, "")
            out_row["ITEM Tax Basis"] = "R"
            
            # ITEM Tax Value = Gross Value * Tax %
            gross_value_str = out_row.get("Gross Value", "")
            tax_value = calculate_tax_value(gross_value_str, tax_percent)
            out_row["ITEM Tax Value"] = tax_value
            
            # LPO and End User (as-is from input)
            out_row["LPO Number"] = clean_text_value(get_scalar_value(row.get("LPO Number", "")))
            out_row["End User"] = build_end_user_value(
                get_scalar_value(row.get("End User", "")),
                get_scalar_value(row.get("End Customer Country", "")),
            )
            
            # Cost source comes from Total Cost Transaction columns
            if group_key not in azure_group_keys:
                out_row["Cost"] = calculate_cost_per_qty(get_row_cost_value(row, cost_col), out_row["Quantity"])
            
            output_rows.append(out_row)
            
        except Exception as e:
            errors.append(f"Row {idx + 2}: {str(e)}")
            logger.error(f"Error processing row {idx + 2}: {str(e)}")
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_rows)
    output_df = apply_invoice_number_versioning(output_df)
    
    # Ensure all columns exist and reorder
    for col in OUTPUT_HEADER:
        if col not in output_df.columns:
            output_df[col] = ""
    
    helper_columns = [col for col in output_df.columns if col.startswith("_")]
    output_df = output_df[OUTPUT_HEADER + helper_columns]
    
    return output_df, errors


def validate_input_file(df: pd.DataFrame) -> Tuple[bool, list]:
    """
    Validate that input file has required columns
    
    Args:
        df: Input DataFrame
    
    Returns:
        Tuple of (is_valid, error_messages)
    """
    df = standardize_input_columns(df)
    df = drop_last_input_row(df)

    required_cols = [
        "Invoice No.", "Customer Code", "Customer Name", "Currency Code",
        "Invoice Type", "Payment Method", "MS Subscription ID", "Billing Cycle Start Date",
        "Billing Cycle End Date", "Charge Description", "Quantity", "LPO Number", "End User"
    ]
    
    errors = []
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        errors.append(f"Missing required columns: {', '.join(missing_cols)}")
        return False, errors

    blank_customer_name_rows = find_blank_rows(df, "Customer Name")
    if blank_customer_name_rows:
        errors.append(
            "Customer Name is mandatory and cannot be blank. "
            f"Blank value found on row(s): {', '.join(map(str, blank_customer_name_rows))}"
        )

    blank_end_user_rows = find_blank_rows(df, "End User")
    if blank_end_user_rows:
        errors.append(
            "End User is mandatory and cannot be blank. "
            f"Blank value found on row(s): {', '.join(map(str, blank_end_user_rows))}"
        )

    if errors:
        return False, errors
    
    return True, []
