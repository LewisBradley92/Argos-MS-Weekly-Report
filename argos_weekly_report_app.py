import io
import zipfile
import re
import unicodedata
from difflib import SequenceMatcher

import pandas as pd
import openpyxl
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator

st.set_page_config(page_title="Argos MS Weekly Report", page_icon="📊", layout="wide")
st.title("Argos MS Weekly Report Builder")
st.caption("Add the latest audit export to the weekly-report template and download the raw data and finished report.")

def norm(value):
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip().lower()
    text = text.replace("£", "pounds").replace("€", "euros")
    text = text.replace("“", '"').replace("”", '"').replace("’", "'").replace("‘", "'")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9]+", "", text)
    return text

def find_header_mapping(template_headers, source_headers):
    """Map source columns to template columns by normalized header, then cautious fuzzy match."""
    source_by_norm = {}
    for source in source_headers:
        key = norm(source)
        if key and key not in source_by_norm:
            source_by_norm[key] = source
    mapping = {}
    for target in template_headers:
        key = norm(target)
        if not key:
            continue
        if key in source_by_norm:
            mapping[target] = source_by_norm[key]
            continue
        # Only accept high-confidence fuzzy matches, to avoid putting data under the wrong question.
        candidates = [(SequenceMatcher(None, key, norm(s)).ratio(), s) for s in source_headers if norm(s)]
        if candidates:
            score, candidate = max(candidates, key=lambda x: x[0])
            if score >= 0.96:
                mapping[target] = candidate
    return mapping

def copy_cell_style(source, target):
    # StyleProxy objects are not hashable; copy the underlying style components.
    from copy import copy
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.protection:
        target.protection = copy(source.protection)

def build_outputs(template_bytes, csv_bytes, store_db_bytes):
    """Update only Raw Data in the LIVE workbook, append unseen audits, and extend Summary detail formulas."""
    try:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="cp1252")

    # Load store database and build lookups by postcode and store name.
    store_wb = load_workbook(io.BytesIO(store_db_bytes), data_only=True, read_only=True)
    store_sheet_name = next((name for name in store_wb.sheetnames if "store listing" in name.lower()), store_wb.sheetnames[0])
    store_ws = store_wb[store_sheet_name]
    store_rows = store_ws.iter_rows(values_only=True)
    db_headers = next(store_rows, ())
    db_index = {norm(header): idx for idx, header in enumerate(db_headers) if header is not None}
    required = ("postcode", "storename", "region", "district")
    if not all(key in db_index for key in required):
        store_wb.close()
        raise ValueError("The store database must contain Store Name, Post Code, Region and District columns.")

    region_by_postcode = {}
    region_by_name = {}
    for db_row in store_rows:
        postcode = norm(db_row[db_index["postcode"]]) if db_row[db_index["postcode"]] is not None else ""
        store_name = norm(db_row[db_index["storename"]]) if db_row[db_index["storename"]] is not None else ""
        region = db_row[db_index["region"]]
        district = db_row[db_index["district"]]
        if postcode:
            region_by_postcode[postcode] = (region, district)
        if store_name:
            region_by_name[store_name] = (region, district)
    store_wb.close()

    # Match audit rows to the store database by site_internal_id (preferred key).
    # Keep postcode/name as fallbacks only when the ID field is unavailable in either file.
    db_id_col = next((key for key in db_index if key in ("siteinternalid", "internalid")), None)
    region_by_site_id = {}
    if db_id_col:
        id_wb = load_workbook(io.BytesIO(store_db_bytes), data_only=True, read_only=True)
        id_ws = id_wb[store_sheet_name]
        for db_row in id_ws.iter_rows(values_only=True, min_row=2):
            site_id = db_row[db_index[db_id_col]]
            if site_id is not None and str(site_id).strip():
                region_by_site_id[str(site_id).strip().casefold()] = (
                    db_row[db_index["region"]], db_row[db_index["district"]]
                )
        id_wb.close()

    audit_id_col = next((col for col in audit_df.columns if norm(col) == "siteinternalid"), None)
    post_col = next((col for col in audit_df.columns if norm(col) in ("sitepostcode", "postcode")), None)
    name_col = next((col for col in audit_df.columns if norm(col) in ("sitename", "storename")), None)
    # Assign object dtype so database values (which may be numeric IDs/labels) can be added safely.
    if "Region" not in audit_df.columns:
        audit_df["Region"] = pd.Series([""] * len(audit_df), index=audit_df.index, dtype=object)
    else:
        audit_df["Region"] = audit_df["Region"].astype(object)
    if "District" not in audit_df.columns:
        audit_df["District"] = pd.Series([""] * len(audit_df), index=audit_df.index, dtype=object)
    else:
        audit_df["District"] = audit_df["District"].astype(object)
    for ridx, record in audit_df.iterrows():
        match = None
        if audit_id_col and db_id_col:
            site_id = str(record.get(audit_id_col, "")).strip().casefold()
            if site_id:
                match = region_by_site_id.get(site_id)
        if not match and post_col:
            match = region_by_postcode.get(norm(record.get(post_col, "")))
        if not match and name_col:
            match = region_by_name.get(norm(record.get(name_col, "")))
        if match:
            audit_df.at[ridx, "Region"] = match[0] if match[0] is not None else ""
            audit_df.at[ridx, "District"] = match[1] if match[1] is not None else ""

    wb = load_workbook(io.BytesIO(template_bytes))
    if "Raw Data" not in wb.sheetnames:
        raise ValueError("The LIVE workbook must contain a 'Raw Data' sheet.")
    if "Summary" not in wb.sheetnames:
        raise ValueError("The LIVE workbook must contain a 'Summary' sheet.")

    raw_ws = wb["Raw Data"]
    summary_ws = wb["Summary"]
    template_headers = [raw_ws.cell(1, c).value for c in range(1, raw_ws.max_column + 1)]
    source_headers = list(audit_df.columns)
    mapping = find_header_mapping(template_headers, source_headers)

    # Use audit ID (internal_id) to prevent duplicates. Fall back to a composite key if absent.
    def source_key(record):
        for h in source_headers:
            if norm(h) == "internalid":
                value = str(record.get(h, "")).strip()
                if value:
                    return ("audit_id", value.casefold())
        return tuple(str(record.get(h, "")).strip().casefold() for h in source_headers[:min(8, len(source_headers))])

    audit_id_col = next((h for h in source_headers if norm(h) == "internalid"), None)
    if audit_id_col is None:
        raise ValueError("The audit export needs an 'internal_id' column so existing audits can be deduplicated safely.")

    existing_keys = set()
    for row in range(2, raw_ws.max_row + 1):
        existing_id = raw_ws.cell(row, 4).value  # internal_id is column D in the LIVE template
        if existing_id is not None and str(existing_id).strip():
            existing_keys.add(str(existing_id).strip().casefold())

    # Preserve existing rows; append only audits not already in Raw Data.
    seen_in_export = set()
    new_records = []
    for _, record in audit_df.iterrows():
        audit_id = str(record.get(audit_id_col, "")).strip()
        if not audit_id:
            continue
        key = audit_id.casefold()
        if key in existing_keys or key in seen_in_export:
            continue
        seen_in_export.add(key)
        new_records.append(record)

    append_start = raw_ws.max_row + 1
    # If the workbook contains formatted but empty rows, find the first genuinely empty row after data.
    while append_start > 2 and all(raw_ws.cell(append_start - 1, c).value is None for c in range(1, raw_ws.max_column + 1)):
        append_start -= 1

    # Preserve the existing column order. Map by heading; unmatched template columns remain blank.
    for offset, record in enumerate(new_records):
        row_idx = append_start + offset
        for col_idx, target_header in enumerate(template_headers, start=1):
            if target_header is None or target_header not in mapping:
                continue
            source_header = mapping[target_header]
            value = record.get(source_header, "")
            if value == "":
                value = None
            if value is not None and any(token in norm(target_header) for token in ("date", "deadline")):
                parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
                if not pd.isna(parsed):
                    value = parsed.to_pydatetime()
            cell = raw_ws.cell(row_idx, col_idx, value=value)
            # Every appended audit row inherits the canonical first data-row formatting.
            if row_idx > 2:
                style_row = 6 if raw_ws.max_row >= 6 else 2
                copy_cell_style(raw_ws.cell(style_row, col_idx), cell)
        # Match row height exactly to the template's first data row, rather than inheriting
        # a potentially irregular height from a preceding record.
        # The LIVE workbook's actual audit rows start at row 6 (rows 2-5 are layout/blank rows).
        # Row 6 has the intended display height; copy it explicitly to every new audit row.
        template_data_row = 6 if raw_ws.max_row >= 6 else 2
        template_height = raw_ws.row_dimensions[template_data_row].height
        if template_height is not None:
            raw_ws.row_dimensions[row_idx].height = template_height
        raw_ws.row_dimensions[row_idx].hidden = raw_ws.row_dimensions[template_data_row].hidden

    raw_ws.freeze_panes = "A2"

    # Summary detail starts at row 20, with one row per Raw Data audit (Raw Data row 2 -> Summary row 20).
    # Extend formulas and formatting down to cover every raw audit, without leaving stale detail rows.
    raw_last_row = raw_ws.max_row
    audit_count = max(0, raw_last_row - 1)
    summary_first_row = 20
    summary_last_row = summary_first_row + audit_count - 1

    # The workbook's existing row 20 is the formula/style template for the audit detail table.
    if audit_count:
        for summary_row in range(summary_first_row, summary_last_row + 1):
            source_row = summary_first_row if summary_row == summary_first_row else summary_first_row
            for col_idx in range(1, summary_ws.max_column + 1):
                source_cell = summary_ws.cell(source_row, col_idx)
                target_cell = summary_ws.cell(summary_row, col_idx)
                if summary_row != source_row:
                    copy_cell_style(source_cell, target_cell)
                if isinstance(source_cell.value, str) and source_cell.value.startswith("="):
                    try:
                        target_cell.value = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                    except Exception:
                        target_cell.value = source_cell.value
                elif summary_row != source_row:
                    target_cell.value = source_cell.value

        # Link Summary Region and District columns to the actual matching headers in Raw Data.
        raw_header_cols = {norm(raw_ws.cell(1, c).value): c for c in range(1, raw_ws.max_column + 1) if raw_ws.cell(1, c).value}
        summary_header_cols = {norm(summary_ws.cell(19, c).value): c for c in range(1, summary_ws.max_column + 1) if summary_ws.cell(19, c).value}
        for label in ("region", "district"):
            raw_col = raw_header_cols.get(label)
            summary_col = summary_header_cols.get(label)
            if raw_col and summary_col:
                raw_letter = openpyxl.utils.get_column_letter(raw_col)
                for summary_row in range(summary_first_row, summary_last_row + 1):
                    raw_row = summary_row - summary_first_row + 2
                    summary_ws.cell(summary_row, summary_col).value = f'=IF(\'Raw Data\'!{raw_letter}{raw_row}="","-",\'Raw Data\'!{raw_letter}{raw_row})'

        # Clear obsolete audit detail rows if the new raw data contains fewer audits than before.
        existing_detail_end = max(summary_ws.max_row, summary_first_row)
        for summary_row in range(summary_last_row + 1, existing_detail_end + 1):
            for col_idx in range(1, summary_ws.max_column + 1):
                summary_ws.cell(summary_row, col_idx).value = None
    else:
        for summary_row in range(summary_first_row, max(summary_ws.max_row, summary_first_row) + 1):
            for col_idx in range(1, summary_ws.max_column + 1):
                summary_ws.cell(summary_row, col_idx).value = None

    # Recalculate formulas when the workbook is opened in Excel.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(new_records), len(existing_keys), len(mapping), len(template_headers), raw_last_row

with st.sidebar:
    st.header("Inputs")
    template_file = st.file_uploader("Weekly Report LIVE workbook (.xlsx)", type=["xlsx"])
    audit_file = st.file_uploader("audit_export (.csv)", type=["csv"])
    store_database_file = st.file_uploader("Argos Store Database (.xlsx)", type=["xlsx"])
    st.caption("The output workbook keeps the LIVE workbook's existing worksheets and formatting.")

st.subheader("What this app does")
st.markdown("""
1. Uses the LIVE workbook as the only output workbook and preserves all tabs.
2. Adds only audit records not already present, deduplicating by `internal_id`.
3. Keeps the Raw Data columns in their existing order.
4. Extends the Summary detail table from row 20 so each raw audit has a matching summary row.
""")

if template_file and audit_file and store_database_file:
    if st.button("Generate raw workbook", type="primary", use_container_width=True):
        try:
            raw_bytes, added_count, existing_count, mapped_count, template_count, raw_last_row = build_outputs(
                template_file.getvalue(), audit_file.getvalue(), store_database_file.getvalue()
            )
            output_name = "Argos MS Weekly Report Raw Data.xlsx"
            st.session_state["argos_raw_workbook"] = raw_bytes
            st.session_state["argos_raw_filename"] = output_name
            st.session_state["argos_raw_stats"] = (added_count, existing_count, mapped_count, template_count, raw_last_row)
            st.success("Raw workbook generated. All original tabs have been retained.")
        except Exception as exc:
            st.error(f"Could not generate the file: {exc}")

if st.session_state.get("argos_raw_workbook"):
    added_count, existing_count, mapped_count, template_count, raw_last_row = st.session_state["argos_raw_stats"]
    st.info(
        f"Added {added_count} new audit(s); skipped duplicates. "
        f"Existing unique audit IDs before this run: {existing_count}. "
        f"Matched {mapped_count} of {template_count} template columns. "
        f"Summary detail rows now cover Summary 20 onward for {max(0, raw_last_row - 1)} audit(s)."
    )
    st.download_button(
        "Download updated raw workbook",
        data=st.session_state["argos_raw_workbook"],
        file_name=st.session_state["argos_raw_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        on_click="ignore",
    )
else:
    st.info("Upload the LIVE workbook, audit_export CSV, and Argos Store Database workbook to begin.")
