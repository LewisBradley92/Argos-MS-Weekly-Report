import io
import zipfile
import re
import unicodedata
from datetime import date, datetime
from difflib import SequenceMatcher

import pandas as pd
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

def build_outputs(template_bytes, csv_bytes, summary_reference_bytes, report_date):
    """Update only Raw Data in the LIVE workbook, append unseen audits, and extend Summary detail formulas."""
    try:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="cp1252")

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
            # Match formatting to the previous data row, without changing existing cells.
            if row_idx > 2:
                copy_cell_style(raw_ws.cell(row_idx - 1, col_idx), cell)
        if raw_ws.row_dimensions[max(2, row_idx - 1)].height:
            raw_ws.row_dimensions[row_idx].height = raw_ws.row_dimensions[max(2, row_idx - 1)].height

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
    reference_file = st.file_uploader("Previous completed report / Summary reference (.xlsx)", type=["xlsx"])
    report_day = st.date_input("Report date", value=date.today())
    st.caption("The output filename uses DD.MM, e.g. Argos MS Weekly Report 23.10.")

st.subheader("What this app does")
st.markdown("""
1. Uses the LIVE workbook as the only output workbook and preserves all tabs.
2. Adds only audit records not already present, deduplicating by `internal_id`.
3. Keeps the Raw Data columns in their existing order.
4. Extends the Summary detail table from row 20 so each raw audit has a matching summary row.
""")

if template_file and audit_file:
    if st.button("Generate raw workbook", type="primary", use_container_width=True):
        try:
            raw_bytes, added_count, existing_count, mapped_count, template_count, raw_last_row = build_outputs(
                template_file.getvalue(), audit_file.getvalue(),
                reference_file.getvalue() if reference_file else template_file.getvalue(),
                report_day
            )
            output_name = f"Argos MS Weekly Report Raw Data {report_day.strftime('%d.%m')}.xlsx"
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
    st.info("Upload the LIVE workbook and audit_export CSV in the sidebar to begin. The previous completed report is optional.")
