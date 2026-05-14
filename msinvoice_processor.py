"""
MS Invoice Processor - Core business logic for MS invoice transformation
Handles data transformation, validations, and calculations
"""
 
import pandas as pd
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
    "DNKW": "KWT",
    "DNFZ": "AED",
    "DNQA": "QAT",
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
    ("POWERPLATFORM - Power Automate per user plan (New Commerce)", "power automate per user", "Power Automate per user", "MSPER-CNS"): "MSPER-CNS",
    ("Excel LTSC 2024", "excel ltsc", "MSPER-CNS"): "MSPER-CNS",
    ("Project Professional 2024 (Commercial) (Subs ID)", "project professional 2024", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 - 1 User CAL (Commercial)", "sql server 2025 - 1 user cal", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 Enterprise core - 2 core License Pack (Commercial)", "sql server 2025 enterprise core", "MSPER-CNS"): "MSPER-CNS",
    ("SQL Server 2025 Standard edition Perpetual 1 Server License (Commercial)", "sql server 2025 standard edition perpetual 1 server license", "MSPER-CNS"): "MSPER-CNS",
    ("Visual Studio Professional 2026 (Commercial)", "visual studio professional 2026", "MSPER-CNS"): "MSPER-CNS",
    ("Windows 11 Enterprise LTSC 2024 Upgrade (Commercial)", "windows 11 enterprise ltsc 2024 upgrade", "MSRI-CNS"): "MSRI-CNS",
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
    rename_map = {}

    for column in df.columns:
        canonical_name = normalized_to_canonical.get(normalize_input_column_name(column))
        if canonical_name and column != canonical_name:
            rename_map[column] = canonical_name

    if not rename_map:
        return df

    return df.rename(columns=rename_map)


def drop_last_input_row(df: pd.DataFrame) -> pd.DataFrame:
    """Ignore the last row of the uploaded file, if any rows exist."""
    if df.empty:
        return df
    return df.iloc[:-1].copy()

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


def format_date_only(value) -> str:
    """Return a date object for date-like values, without timestamps."""
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
    errors = []
    output_rows = []
    today = datetime.today().date()
    
    for idx, row in df.iterrows():
        try:
            out_row = {}
            
            # Get Invoice No. (as-is from input)
            invoice_no = str(row.get("Invoice No.", "")).strip()
            out_row["Invoice No."] = invoice_no
            
            # Document Location from Invoice No. prefix
            doc_location = get_document_location(invoice_no)
            out_row["Document Location"] = doc_location
            out_row["Sale Location"] = doc_location
            out_row["Delivery Location Code"] = get_delivery_location_code(doc_location)
            
            # Customer info (as-is from input)
            out_row["Customer Code"] = clean_text_value(row.get("Customer Code", ""))
            out_row["Customer Name"] = clean_text_value(row.get("Customer Name", ""))
            out_row["_Invoice Type"] = clean_text_value(row.get("Invoice Type", ""))
            
            # Dates
            out_row["_Source Invoice Date"] = row.get("Invoice Date", "")
            out_row["Invoice Date"] = today
            out_row["Delivery Date"] = today
            out_row["Annotation"] = ""
            
            # Currency from Invoice No. prefix
            out_row["Currency Code"] = get_currency_from_invoice_no(invoice_no)
            uploaded_exchange_rate = None
            try:
                uploaded_exchange_rate = float(row.get("Exchange Rate", 0))
            except (ValueError, TypeError):
                uploaded_exchange_rate = None
            
            exchange_rate = get_exchange_rate(doc_location, uploaded_exchange_rate)
            out_row["Exchange Rate"] = exchange_rate if exchange_rate != "" else ""
            
            # Fixed fields
            out_row["Shipment Mode"] = "EML"
            out_row["Payment Term"] = extract_payment_term(str(row.get("Payment Method", "")))
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
            ms_sub_id = clean_text_value(row.get("MS Subscription ID", ""))
            subscription_id_value = ms_sub_id if ms_sub_id else "Subs ID"
            out_row["Subscription Id"] = subscription_id_value
            out_row["Billing Cycle Start Date"] = format_date_only(row.get("Billing Cycle Start Date", ""))
            out_row["Billing Cycle End Date"] = format_date_only(row.get("Billing Cycle End Date", ""))
            
            # ITEM Code mapped from Charge Description
            charge_desc = clean_text_value(row.get("Charge Description", ""))
            item_code = get_item_code(charge_desc)
            out_row["ITEM Code"] = item_code
            
            # ITEM Name = Charge Description + MS Subscription ID
            out_row["ITEM Name"] = charge_desc + (f" ({subscription_id_value})" if subscription_id_value else "")
            
            # Fixed ITEM fields
            out_row["UOM"] = "NOS"
            out_row["Grade code-1"] = "NA"
            out_row["Grade code-2"] = "NA"
            
            # Quantity and Price
            quantity = row.get("Quantity", 0)
            out_row["Quantity"] = quantity
            out_row["Qty Loose"] = 0
            
            # Calculate Gross Value using the formula from input Rate Per Qty and Exchange Rate
            rate_per_qty_input = row.get("Rate Per Qty", "")
            exchange_rate_input = row.get("Exchange Rate", "")
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
            out_row["ITEM Tax Basis"] = ""
            
            # ITEM Tax Value = Gross Value * Tax %
            gross_value_str = out_row.get("Gross Value", "")
            tax_value = calculate_tax_value(gross_value_str, tax_percent)
            out_row["ITEM Tax Value"] = tax_value
            
            # LPO and End User (as-is from input)
            out_row["LPO Number"] = clean_text_value(row.get("LPO Number", ""))
            out_row["End User"] = build_end_user_value(
                row.get("End User", ""),
                row.get("End Customer Country", ""),
            )
            
            # Cost source follows the same positive/credit-note split as Gross Value
            cost_col = find_column_with_prefix(df, "Unit Cost Transaction Currency")
            is_credit_note = is_negative_credit_note(row)
            if is_credit_note:
                cost_value = round_to_2_decimals(row.get("Unit Cost", ""))
                out_row["Cost"] = cost_value
            elif cost_col:
                cost_value = round_to_2_decimals(row.get(cost_col, ""))
                out_row["Cost"] = cost_value
            else:
                out_row["Cost"] = ""
            
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
