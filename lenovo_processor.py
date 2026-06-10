import io
import re
import zipfile
from pathlib import Path

import pandas as pd

try:
    from amal.pdf_utils import extract_text_from_pdf
    from amal.sob_parser import extract_comm_inv_fields_from_sob, extract_sob_line_items, normalize_item_code
    from lenovo.workbook_builder import create_workbook_bytes
except ImportError:
    from amal.pdf_utils import extract_text_from_pdf
    from amal.sob_parser import extract_comm_inv_fields_from_sob, extract_sob_line_items, normalize_item_code
from amal.workbook_builder import build_pack_list_sheet, compute_address_rows
from lenovo.workbook_builder import fill_pack_list_sheet_static


def normalize_decimal(value) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    text = str(value).replace(",", "").strip()
    if text.lower() in {"nan", "none", ""}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _find_column(columns, *candidates) -> str:
    lowered = {re.sub(r"[^A-Za-z0-9]+", " ", str(col)).strip().lower(): col for col in columns}
    for candidate in candidates:
        normalized_candidate = re.sub(r"[^A-Za-z0-9]+", " ", candidate).strip().lower()
        if normalized_candidate in lowered:
            return lowered[normalized_candidate]
    return ""


def _excel_engine_for_file(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix.lower()
    return "xlrd" if suffix == ".xls" else "openpyxl"


def _read_excel_sheet_rows(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    engine = _excel_engine_for_file(uploaded_file)
    df = pd.read_excel(uploaded_file, engine=engine, header=None)
    uploaded_file.seek(0)
    return df


def _extract_pack_list_table_rows(uploaded_file) -> list[list]:
    df = _read_excel_sheet_rows(uploaded_file)
    rows = df.fillna("").astype(str).values.tolist()

    start_index = 0
    for idx, row in enumerate(rows):
        text = " ".join(str(value).strip() for value in row if str(value).strip()).lower()
        if any(token in text for token in ("case no", "material", "part number", "qty", "gw(kg)", "description")):
            start_index = idx + 1
            break

    filtered_rows = []
    for row in rows[start_index:]:
        text = " ".join(str(value).strip() for value in row if str(value).strip())
        if not text or text.lower().startswith("total:"):
            continue
        filtered_rows.append(row)

    return filtered_rows


def _looks_like_invoice_table(uploaded_file) -> bool:
    df = _read_excel_sheet_rows(uploaded_file)
    text_blocks = []
    for _, row in df.iterrows():
        text_blocks.append(" ".join(str(value).strip() for value in row.tolist() if str(value).strip()))

    combined = re.sub(r"[^A-Za-z0-9]+", " ", " ".join(text_blocks).lower())
    invoice_score = sum(1 for token in ("unit price", "total amount", "hscode") if token in combined)
    has_part_desc = any(token in combined for token in ("part no", "description"))
    return invoice_score >= 1 and has_part_desc


def _extract_pack_list_table_rows(uploaded_file) -> list[list]:
    df = _read_excel_sheet_rows(uploaded_file)
    rows = df.fillna("").astype(str).values.tolist()

    start_index = 0
    for idx, row in enumerate(rows):
        text = " ".join(str(value).strip() for value in row if str(value).strip()).lower()
        if any(token in text for token in ("case no", "material", "part number", "qty", "gw(kg)", "description")):
            start_index = idx + 1
            break

    table_rows = [row for row in rows[start_index:] if any(str(value).strip() for value in row)]
    return table_rows


def _detect_huawei_table(df: pd.DataFrame) -> pd.DataFrame:
    for index, row in df.iterrows():
        row_values = [str(value).strip() if pd.notna(value) else "" for value in row.tolist()]
        text = re.sub(r"[^A-Za-z0-9]+", " ", " ".join(row_values).lower())
        if any(token in text for token in ("part no", "description", "hscode", "qty", "total amount", "unit price")):
            data_df = df.iloc[index + 1:].copy()
            data_df = data_df.dropna(how="all")
            data_df.columns = row_values
            return data_df
    return df


def _extract_lenovo_address_block(sob_text: str) -> tuple[str, str]:
    start = sob_text.find("Bill To Ship To")
    end = sob_text.find("Forwarder")
    if start == -1 or end == -1 or end <= start:
        return "", ""

    raw_lines = [line.strip() for line in sob_text[start:end].splitlines() if line.strip()]
    if raw_lines and raw_lines[0].lower() == "bill to ship to":
        raw_lines = raw_lines[1:]

    cleaned_lines: list[str] = []
    seen_lines: set[str] = set()

    def add_unique_line(value: str) -> None:
        normalized = value.replace("\xa0", " ").strip()
        if not normalized:
            return

        normalized_key = normalized.lower()
        is_duplicate_fragment = (
            normalized_key in seen_lines
            and (
                re.search(r"\d[\d\s]{5,}\d", normalized)
                or "@" in normalized
                or normalized.startswith("TRN:")
            )
        )
        if is_duplicate_fragment:
            return

        seen_lines.add(normalized_key)
        cleaned_lines.append(normalized)

    for line in raw_lines:
        normalized = line.replace("\xa0", " ")

        phone_company_match = re.match(r"^(\d[\d\s]{5,}\d)([A-Z].*)$", normalized)
        if phone_company_match:
            add_unique_line(phone_company_match.group(1))
            add_unique_line(phone_company_match.group(2))
            continue

        trn_match = re.search(r"(TRN:\d+)(TRN:NA)", normalized)
        if trn_match:
            add_unique_line(trn_match.group(1))
            add_unique_line(trn_match.group(2))
            continue

        add_unique_line(normalized)

    split_index = next((idx for idx, line in enumerate(cleaned_lines) if "Uruk Engineering Services" in line), None)
    if split_index is None:
        split_index = next((idx for idx, line in enumerate(cleaned_lines) if "PO BOX:" in line), None)

    if split_index is None or split_index <= 1:
        return "", ""

    bill_to = "\n".join(cleaned_lines[:split_index]).strip()
    ship_to = "\n".join(cleaned_lines[split_index:]).strip()
    return bill_to, ship_to


def read_huawei_excel_rows(uploaded_file) -> list[dict]:
    df = _read_excel_sheet_rows(uploaded_file)
    data_df = _detect_huawei_table(df)

    no_col = _find_column(data_df.columns, "no.", "no")
    item_code_col = _find_column(data_df.columns, "part no", "item code", "itemcode", "code")
    desc_col = _find_column(data_df.columns, "description", "desc", "item description")
    origin_col = _find_column(data_df.columns, "c/o", "origin")
    hs_col = _find_column(data_df.columns, "hscode", "hs code")
    qty_col = _find_column(data_df.columns, "qty", "quantity", "qty.")
    unit_price_col = _find_column(data_df.columns, "unit price", "unitprice")
    total_amount_col = _find_column(data_df.columns, "total amount", "total")

    rows = []
    for _, row in data_df.iterrows():
        if no_col:
            try:
                numeric_no = float(str(row.get(no_col, "")).strip())
            except ValueError:
                numeric_no = None
            if numeric_no is None:
                break
        try:
            item_code = str(row.get(item_code_col, "")).strip() if item_code_col else ""
            desc = str(row.get(desc_col, "")).strip() if desc_col else ""
            origin = str(row.get(origin_col, "")).strip() if origin_col else ""
            hs_code = str(row.get(hs_col, "")).strip() if hs_col else ""
            qty = normalize_decimal(row.get(qty_col, 0)) if qty_col else 0.0
            unit_price = normalize_decimal(row.get(unit_price_col, 0)) if unit_price_col else 0.0
            total_amount = normalize_decimal(row.get(total_amount_col, 0)) if total_amount_col else 0.0

            if not item_code and not desc and not qty and not total_amount:
                continue

            rows.append({
                "item_code": item_code,
                "desc": desc,
                "case_no": "",
                "origin": origin,
                "hs_code": hs_code,
                "qty": qty,
                "unit_price": unit_price,
                "total_amount": total_amount,
            })
        except Exception:
            continue

    return rows


def build_lenovo_ci_workbook(sob_file, ci_file, all_ci_files=None) -> io.BytesIO:
    current_files = [ci_file] if not isinstance(ci_file, (list, tuple)) else list(ci_file)
    reference_files = list(all_ci_files) if all_ci_files else current_files
    return build_lenovo_workbook(
        sob_file,
        current_files,
        include_pack_list_sheet=False,
        all_ci_files=reference_files,
    )


def _first_sheet_to_bytes(uploaded_file, sob_fields: dict | None = None) -> bytes:
    df = _read_excel_sheet_rows(uploaded_file)
    rows = df.fillna("").astype(str).values.tolist()

    start_index = 0
    for idx, row in enumerate(rows):
        text = " ".join(str(value).strip() for value in row if str(value).strip()).lower()
        if any(token in text for token in ("case no", "material", "part number", "qty", "gw(kg)", "description")):
            start_index = idx
            break

    output_buffer = io.BytesIO()
    wb = __import__("openpyxl").Workbook()
    ws = wb.active
    ws.title = "pack_list"

    build_pack_list_sheet(ws)
    fill_pack_list_sheet_static(
        ws,
        sob_fields or {"bill_to": "", "ship_to": "", "commercial_invoice_no": "", "date": ""},
        compute_address_rows(sob_fields.get("bill_to", "") if sob_fields else "", sob_fields.get("ship_to", "") if sob_fields else ""),
    )

    hdr_row = 8 + compute_address_rows(sob_fields.get("bill_to", "") if sob_fields else "", sob_fields.get("ship_to", "") if sob_fields else "")
    header_values = rows[start_index]
    for col_idx, value in enumerate(header_values, start=1):
        ws.cell(row=hdr_row, column=col_idx, value=value)

    for row_idx, row_values in enumerate(rows[start_index + 1:], start=hdr_row + 1):
        for col_idx, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()


def build_pack_list_zip(pl_files, sob_file=None) -> io.BytesIO:
    sob_fields = None
    if sob_file:
        sob_text = extract_text_from_pdf(sob_file)
        sob_fields = extract_comm_inv_fields_from_sob(sob_text)
        bill_to, ship_to = _extract_lenovo_address_block(sob_text)
        if bill_to or ship_to:
            sob_fields["bill_to"] = bill_to
            sob_fields["ship_to"] = ship_to
        sob_fields["commercial_invoice_no"] = sob_fields.get("commercial_invoice_no", "") or Path(sob_file.name).stem
        sob_fields["date"] = sob_fields.get("date", "")

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for uploaded_file in pl_files:
            archive_name = f"{Path(uploaded_file.name).stem}.xlsx"
            archive.writestr(archive_name, _first_sheet_to_bytes(uploaded_file, sob_fields=sob_fields))
    output.seek(0)
    return output


def _collect_missing_sob_items(sob_items: list[dict], ci_files: list) -> list[dict]:
    covered_codes = set()

    for excel_file in ci_files:
        for item in read_huawei_excel_rows(excel_file):
            base_code = normalize_item_code(str(item.get("item_code", "")).split("/")[0].strip())
            if not base_code:
                continue

            covered_codes.update(
                sob_item.get("normalized_item_code", "")
                for sob_item in sob_items
                if sob_item.get("normalized_item_code", "").startswith(base_code)
            )

    missing_items = []
    seen_codes = set()
    for sob_item in sob_items:
        normalized_code = sob_item.get("normalized_item_code", "")
        if normalized_code in covered_codes or normalized_code in seen_codes:
            continue

        seen_codes.add(normalized_code)
        missing_items.append({
            "item_code": sob_item.get("item_code", ""),
            "desc": sob_item.get("description", ""),
            "case_no": "",
            "origin": "",
            "hs_code": "",
            "qty": sob_item.get("qty", 0),
            "lenovo_qty": sob_item.get("qty", 0),
            "unit_price": 0.0,
            "amount": 0.0,
            "append_only": True,
        })

    return missing_items


def build_lenovo_workbook(
    sob_file,
    excel_files,
    include_pack_list_sheet: bool = True,
    all_ci_files=None,
) -> io.BytesIO:
    sob_text = extract_text_from_pdf(sob_file)
    sob_fields = extract_comm_inv_fields_from_sob(sob_text)
    sob_items = extract_sob_line_items(sob_text)
    bill_to, ship_to = _extract_lenovo_address_block(sob_text)
    if bill_to or ship_to:
        sob_fields["bill_to"] = bill_to
        sob_fields["ship_to"] = ship_to
    sob_fields["date"] = ""
    sob_fields["commercial_invoice_no"] = sob_fields.get("commercial_invoice_no", "") or Path(sob_file.name).stem

    all_items = []
    pack_list_files = []

    invoice_candidates = [excel_file for excel_file in excel_files if _looks_like_invoice_table(excel_file)]
    if invoice_candidates:
        candidate_files = invoice_candidates
    else:
        candidate_files = list(excel_files)

    for excel_file in candidate_files:
        all_items.extend(read_huawei_excel_rows(excel_file))

    for excel_file in excel_files:
        if excel_file not in candidate_files:
            pack_list_files.append(excel_file)

    sob_total = normalize_decimal(sob_fields.get("sob_total", 0))
    total_qty = sum(item.get("qty", 0) for item in all_items)

    comm_inv_items = []
    for item in all_items:
        qty = float(item.get("qty", 0) or 0)
        base_code = str(item.get("item_code", "")).split("/")[0].strip()
        normalized_code = normalize_item_code(base_code)
        prefix_matches = [sob_item for sob_item in sob_items if sob_item.get("normalized_item_code", "").startswith(normalized_code)]

        sob_qty = round(sum(sob_item.get("qty", 0.0) for sob_item in prefix_matches), 2)
        row_total = round(sum(sob_item.get("total", 0.0) for sob_item in prefix_matches), 2)
        qty_for_pricing = sob_qty or qty
        unit_price = round(row_total / qty_for_pricing, 2) if qty_for_pricing and row_total else 0.0
        amount = row_total
        comm_inv_items.append({
            "item_code": item.get("item_code", ""),
            "desc": item.get("desc", ""),
            "case_no": item.get("case_no", ""),
            "origin": item.get("origin", ""),
            "hs_code": item.get("hs_code", ""),
            "qty": qty,
            "lenovo_qty": sob_qty or qty,
            "unit_price": unit_price,
            "amount": amount,
        })

    reference_files = list(all_ci_files) if all_ci_files else candidate_files
    comm_inv_items.extend(_collect_missing_sob_items(sob_items, reference_files))

    sob_fields["total_amount"] = round(sob_total, 2)
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
        extra_pack_list_files=pack_list_files,
        include_pack_list_sheet=include_pack_list_sheet,
    )
