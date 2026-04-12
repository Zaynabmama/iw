"""
MS Invoice Processor - Core business logic for MS invoice transformation
Handles data transformation, validations, and calculations
"""
 
import pandas as pd
import re
from datetime import datetime
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
    "UJ000": 0.272294078,
    "TC000": 0.272294078,
    "QA000": 0.274725274725,
    "OM000": 2.60078023407,
    "KA000": 0.2666666666
}

KEYWORD_MAP = {
    ("windows server", "window server", "MSPER-CNS"): "MSPER-CNS",
    ("azure subscription", "MSAZ-CNS"): "MSAZ-CNS",
    ("google workspace", "GL-WSP-CNS"): "GL-WSP-CNS",
    ("m365", "microsoft 365", "office 365", "exchange online", "Microsoft Defender for Endpoint P1", "MS-CNS"): "MS-CNS",
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
        # Special case: round(1/Exchange Rate / 2)
        if uploaded_exchange_rate and uploaded_exchange_rate > 0:
            return round(1 / uploaded_exchange_rate / 2, 10)
        return ""
    
    return EXCHANGE_RATE_MAP.get(document_location, "")


def find_column_with_prefix(df: pd.DataFrame, prefix: str) -> Optional[str]:
    """Find column name that starts with given prefix"""
    for col in df.columns:
        if str(col).startswith(prefix):
            return col
    return None


def get_item_code(input_item_code: str) -> str:
    """Map uploaded ITEM Code to output ITEM Code using substring matching."""
    if pd.isna(input_item_code):
        return ""
    
    item_code_value = str(input_item_code).strip().upper()
    
    for keywords, code in KEYWORD_MAP.items():
        for keyword in keywords:
            if str(keyword).strip().upper() in item_code_value:
                return code
    
    return ""


def round_to_2_decimals(value) -> str:
    """Round value to 2 decimals, return empty string if NaN or 0"""
    try:
        if pd.isna(value):
            return ""
        num = float(value)
        rounded = round(num, 2)
        return str(rounded)
    except (ValueError, TypeError):
        return ""


def calculate_rate_per_qty(gross_value_str, quantity) -> str:
    """Calculate Rate Per Qty as Gross Value / Quantity, keeping empty if Quantity is 0"""
    try:
        if pd.isna(quantity) or quantity == 0:
            return ""
        
        gross_value = float(gross_value_str) if isinstance(gross_value_str, str) else float(gross_value_str)
        qty = float(quantity)
        
        if qty == 0:
            return ""
        
        rate = round(gross_value / qty, 2)
        return str(rate)
    except (ValueError, TypeError, ZeroDivisionError):
        return ""


def calculate_tax_value(gross_value: float, tax_percent: float) -> str:
    """Calculate Tax Value = Gross Value * Tax %"""
    try:
        if pd.isna(gross_value) or pd.isna(tax_percent) or tax_percent == "":
            return ""
        
        gv = float(gross_value)
        tp = float(tax_percent) if isinstance(tax_percent, str) else float(tax_percent)
        
        tax_val = round(gv * tp / 100, 2)
        return str(tax_val)
    except (ValueError, TypeError):
        return ""


def is_negative_credit_note(row: pd.Series) -> bool:
    """Use input Gross Value sign to detect credit-note rows."""
    try:
        return float(row.get("Gross Value", 0) or 0) < 0
    except (ValueError, TypeError):
        return False


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
    errors = []
    output_rows = []
    today = datetime.today().strftime("%d/%m/%Y")
    
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
            out_row["Customer Code"] = str(row.get("Customer Code", "")).strip()
            out_row["Customer Name"] = str(row.get("Customer Name", "")).strip()
            
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
            out_row["Subscription Id"] = str(row.get("MS Subscription ID", "")).strip()
            out_row["Billing Cycle Start Date"] = str(row.get("Billing Cycle Start Date", "")).strip()
            out_row["Billing Cycle End Date"] = str(row.get("Billing Cycle End Date", "")).strip()
            
            # ITEM Code mapped from uploaded ITEM Code
            input_item_code = str(row.get("ITEM Code", "")).strip()
            item_code = get_item_code(input_item_code)
            out_row["ITEM Code"] = item_code
            
            # ITEM Name = Charge Description + MS Subscription ID
            charge_desc = str(row.get("Charge Description", "")).strip()
            ms_sub_id = str(row.get("MS Subscription ID", "")).strip()
            out_row["ITEM Name"] = charge_desc + (f" ({ms_sub_id})" if ms_sub_id else "")
            
            # Fixed ITEM fields
            out_row["UOM"] = "NOS"
            out_row["Grade code-1"] = "NA"
            out_row["Grade code-2"] = "NA"
            
            # Quantity and Price
            quantity = row.get("Quantity", 0)
            out_row["Quantity"] = quantity
            out_row["Qty Loose"] = 0
            
            # Positive rows keep existing local-currency logic; credit notes use input Gross Value/Unit Cost
            gross_value_col = find_column_with_prefix(df, "Gross Value Transaction Currency")
            is_credit_note = is_negative_credit_note(row)

            if is_credit_note:
                gross_value_raw = row.get("Gross Value", 0)
            elif gross_value_col:
                gross_value_raw = row.get(gross_value_col, 0)
            else:
                gross_value_raw = ""

            if gross_value_raw != "":
                gross_value_rounded = round_to_2_decimals(gross_value_raw)
                out_row["Gross Value"] = gross_value_rounded
                
                # Rate Per Qty = Gross Value / Quantity
                rate_per_qty = calculate_rate_per_qty(gross_value_rounded, quantity)
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
            out_row["LPO Number"] = str(row.get("LPO Number", "")).strip()
            out_row["End User"] = str(row.get("End User", "")).strip()
            
            # Cost source follows the same positive/credit-note split as Gross Value
            cost_col = find_column_with_prefix(df, "Unit Cost Transaction Currency")
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
    required_cols = [
        "Invoice No.", "Customer Code", "Customer Name", "Currency Code",
        "Payment Method", "MS Subscription ID", "Billing Cycle Start Date",
        "Billing Cycle End Date", "Charge Description", "Quantity", "LPO Number", "End User"
    ]
    
    errors = []
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        errors.append(f"Missing required columns: {', '.join(missing_cols)}")
        return False, errors
    
    return True, []
