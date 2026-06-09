import io
from pathlib import Path

import pandas as pd

try:
    from amal.pdf_utils import extract_text_from_pdf
    from amal.sob_parser import extract_comm_inv_fields_from_sob
    from amal.workbook_builder import create_workbook_bytes
except ImportError:
    from amal.pdf_utils import extract_text_from_pdf
    from amal.sob_parser import extract_comm_inv_fields_from_sob
    from amal.workbook_builder import create_workbook_bytes


def normalize_decimal(value) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    text = str(value).replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def _find_column(columns, *candidates) -> str:
    lowered = {str(col).strip().lower(): col for col in columns}
    for candidate in candidates:
        if candidate.lower() in lowered:
            return lowered[candidate.lower()]
    return ""


def read_huawei_excel_rows(uploaded_file) -> list[dict]:
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    uploaded_file.seek(0)

    item_code_col = _find_column(df.columns, "Item Code", "ItemCode", "Code")
    desc_col = _find_column(df.columns, "Desc", "Description", "Item Description")
    origin_col = _find_column(df.columns, "Origin")
    hs_col = _find_column(df.columns, "HS Code", "HSCode", "HSCODE")
    qty_col = _find_column(df.columns, "Qty", "Quantity", "Qty.")

    rows = []
    for _, row in df.iterrows():
        item_code = str(row.get(item_code_col, "")).strip() if item_code_col else ""
        desc = str(row.get(desc_col, "")).strip() if desc_col else ""
        origin = str(row.get(origin_col, "")).strip() if origin_col else ""
        hs_code = str(row.get(hs_col, "")).strip() if hs_col else ""
        qty = normalize_decimal(row.get(qty_col, 0)) if qty_col else 0.0

        if not item_code and not desc and not qty:
            continue

        rows.append({
            "item_code": item_code,
            "desc": desc,
            "case_no": "",
            "origin": origin,
            "hs_code": hs_code,
            "qty": qty,
        })

    return rows


def build_lenovo_workbook(sob_file, excel_files) -> io.BytesIO:
    sob_text = extract_text_from_pdf(sob_file)
    sob_fields = extract_comm_inv_fields_from_sob(sob_text)
    sob_fields["date"] = ""
    sob_fields["commercial_invoice_no"] = sob_fields.get("commercial_invoice_no", "") or Path(sob_file.name).stem

    all_items = []
    for excel_file in excel_files:
        all_items.extend(read_huawei_excel_rows(excel_file))

    total_amount = normalize_decimal(sob_fields.get("sob_total", 0))
    total_qty = sum(item.get("qty", 0) for item in all_items)
    unit_price = round(total_amount / total_qty, 2) if total_qty else 0.0

    comm_inv_items = []
    for item in all_items:
        qty = float(item.get("qty", 0) or 0)
        amount = round(qty * unit_price, 2)
        comm_inv_items.append({
            "item_code": item.get("item_code", ""),
            "desc": item.get("desc", ""),
            "case_no": item.get("case_no", ""),
            "origin": item.get("origin", ""),
            "hs_code": item.get("hs_code", ""),
            "qty": qty,
            "unit_price": unit_price,
            "amount": amount,
        })

    sob_fields["total_amount"] = round(total_amount, 2)
    sob_fields["freight_charges"] = sob_fields.get("freight_charges", "")

    # Keep the workbook output consistent with the Amal style.
    return create_workbook_bytes(
        comm_inv_fields=sob_fields,
        comm_inv_items=comm_inv_items,
        comm_inv_unmatched_items=[],
        pack_list_fields={
            "commercial_invoice_no": sob_fields.get("commercial_invoice_no", ""),
            "date": sob_fields.get("date", ""),
            "bill_to": sob_fields.get("bill_to", ""),
            "ship_to": sob_fields.get("ship_to", ""),
            "total_packages": 0,
            "total_gross_weight": 0,
            "total_qty": round(total_qty, 2),
        },
        pack_list_items=[
            {
                "item_code": item.get("item_code", ""),
                "desc": item.get("desc", ""),
                "case_no": item.get("case_no", "") or f"CASE-{idx + 1}",
                "origin": item.get("origin", ""),
                "hs_code": item.get("hs_code", ""),
                "qty": item.get("qty", 0),
                "gross_weight": 0,
                "package": 1,
                "dimensions_cm": "",
            }
            for idx, item in enumerate(comm_inv_items)
        ],
        comm_inv_df=pd.DataFrame([{"source_file": sob_file.name, "document_type": "lenovo_invoice", "status": "ready"}]),
        pack_list_df=pd.DataFrame([{"source_file": sob_file.name, "document_type": "lenovo_pack_list", "status": "ready"}]),
    )
