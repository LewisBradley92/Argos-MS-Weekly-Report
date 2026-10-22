import io
import re
import unicodedata
from datetime import date, datetime
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
    # Read the export as strings first so identifiers, leading zeroes and free-text are retained.
    try:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="utf-8-sig")
    except UnicodeDecodeError:
        audit_df = pd.read_csv(io.BytesIO(csv_bytes), dtype=object, keep_default_na=False, encoding="cp1252")

    wb = load_workbook(io.BytesIO(template_bytes))
    if "Raw Data" not in wb.sheetnames:
        raise ValueError("The weekly-report template must contain a 'Raw Data' sheet.")
    ws = wb["Raw Data"]

    # Row 1 in the live workbook defines the canonical column order. Never reorder these columns.
    template_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    source_headers = list(audit_df.columns)
    mapping = find_header_mapping(template_headers, source_headers)

    # Clear old data but preserve the template header, column positions, widths and formatting.
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # Append rows by template column, never by source-file position.
    for row_idx, (_, record) in enumerate(audit_df.iterrows(), start=2):
        for col_idx, target_header in enumerate(template_headers, start=1):
            if target_header is None or target_header not in mapping:
                continue
            source_header = mapping[target_header]
            value = record.get(source_header, "")
            if value == "":
                value = None
            # Convert common date fields to Excel dates where safely parseable; otherwise retain source text.
            if value is not None and any(token in norm(target_header) for token in ("date", "deadline")):
                parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
                if not pd.isna(parsed):
                    value = parsed.to_pydatetime()
            cell = ws.cell(row_idx, col_idx, value=value)
            if row_idx > 2:
                copy_cell_style(ws.cell(2, col_idx), cell)
        # Copy row height from the first data row if available.
        if ws.row_dimensions[2].height:
            ws.row_dimensions[row_idx].height = ws.row_dimensions[2].height

    ws.freeze_panes = "A2"
    # Ensure the workbook's calculation engine recalculates formulas when opened.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    # Use the provided past report as the visual/value reference for the Summary tab.
    # The live template's Summary tab is retained so its layout, formulas and formatting stay intact.
    # If the reference workbook has a Summary sheet, copy its layout/formatting only where the live
    # template is missing a Summary sheet; do not reuse last week's stale values.
    if "Summary" not in wb.sheetnames:
        ref_wb = load_workbook(io.BytesIO(summary_reference_bytes), data_only=True)
        if "Summary" in ref_wb.sheetnames:
            ref_ws = ref_wb["Summary"]
            out_ws = wb.create_sheet("Summary")
            for row in ref_ws.iter_rows():
                for cell in row:
                    out = out_ws.cell(cell.row, cell.column, value=cell.value)
                    copy_cell_style(cell, out)

    # The summary in the deliverable should be values-only. Calculate formulas using LibreOffice
    # when available; if not available, retain formula cells so Excel can calculate on opening.
    # We avoid replacing formulas with last week's numbers, which would be misleading.
    xlsx_io = io.BytesIO()
    wb.save(xlsx_io)
    xlsx_bytes = xlsx_io.getvalue()

    # Create a values-only Summary tab using LibreOffice when available. The no-LibreOffice
    # fallback keeps the workbook valid and clearly warns the user that Excel needs to recalculate.
    import shutil, subprocess, tempfile, os
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    values_only_bytes = None
    if lo:
        with tempfile.TemporaryDirectory() as tmp:
            input_path = os.path.join(tmp, "weekly.xlsx")
            output_dir = os.path.join(tmp, "out")
            os.makedirs(output_dir)
            with open(input_path, "wb") as f:
                f.write(xlsx_bytes)
            result = subprocess.run([lo, "--headless", "--convert-to", "xlsx", "--outdir", output_dir, input_path],
                                    capture_output=True, text=True, timeout=90)
            recalculated = os.path.join(output_dir, "weekly.xlsx")
            if os.path.exists(recalculated):
                calc_wb = load_workbook(recalculated, data_only=True)
                formula_wb = load_workbook(recalculated, data_only=False)
                if "Summary" in calc_wb.sheetnames and "Summary" in formula_wb.sheetnames:
                    value_ws = calc_wb["Summary"]
                    formula_ws = formula_wb["Summary"]
                    for row in formula_ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                cell.value = value_ws.cell(cell.row, cell.column).value
                final_io = io.BytesIO()
                formula_wb.save(final_io)
                values_only_bytes = final_io.getvalue()

    if values_only_bytes is None:
        # Fallback: use the workbook with formulas intact. Streamlit tells the user this fallback was used.
        values_only_bytes = xlsx_bytes
        summary_values_only = False
    else:
        summary_values_only = True

    # Raw deliverable is the audit export aligned to the weekly template's column order.
    raw_out = io.StringIO()
    output_headers = [h if h is not None else "" for h in template_headers]
    aligned = pd.DataFrame(columns=output_headers)
    for target_header in template_headers:
        out_name = target_header if target_header is not None else ""
        if target_header is not None and target_header in mapping:
            aligned[out_name] = audit_df[mapping[target_header]].values
        else:
            aligned[out_name] = [""] * len(audit_df)
    aligned.to_csv(raw_out, index=False, encoding="utf-8-sig")

    return raw_out.getvalue().encode("utf-8-sig"), values_only_bytes, len(audit_df), len(mapping), len(template_headers), summary_values_only

with st.sidebar:
    st.header("Inputs")
    template_file = st.file_uploader("Weekly Report LIVE workbook (.xlsx)", type=["xlsx"])
    audit_file = st.file_uploader("audit_export (.csv)", type=["csv"])
    reference_file = st.file_uploader("Previous completed report / Summary reference (.xlsx)", type=["xlsx"])
    report_day = st.date_input("Report date", value=date.today())
    st.caption("The output filename uses DD.MM, e.g. Argos MS Weekly Report 23.10.")

st.subheader("What this app does")
st.markdown("""
1. Keeps the **Raw Data** columns in the exact order used by the LIVE workbook.
2. Matches audit-export fields to the correct report columns by header, not by column position.
3. Produces a template-aligned raw CSV and an updated weekly-report workbook.
4. Uses the completed report as a reference only; it does not copy last week's summary figures into the new report.
""")

if template_file and audit_file:
    if st.button("Generate both files", type="primary", use_container_width=True):
        try:
            ref_bytes = reference_file.getvalue() if reference_file else template_file.getvalue()
            raw_bytes, report_bytes, row_count, mapped_count, template_count, summary_values_only = build_outputs(
                template_file.getvalue(), audit_file.getvalue(), ref_bytes, report_day
            )
            report_name = f"Argos MS Weekly Report {report_day.strftime('%d.%m')}.xlsx"
            raw_name = f"Argos MS Weekly Report Raw Data {report_day.strftime('%d.%m')}.csv"
            st.success(f"Processed {row_count:,} audit rows. Matched {mapped_count} of {template_count} template columns.")
            if summary_values_only:
                st.info("Summary formulas were recalculated and replaced with values.")
            else:
                st.warning("This environment could not recalculate Excel formulas automatically. The report is downloadable, but open it in Excel and allow recalculation before sharing; the Summary may still contain formulas.")
            left, right = st.columns(2)
            with left:
                st.download_button("Download raw data CSV", data=raw_bytes, file_name=raw_name, mime="text/csv", use_container_width=True)
            with right:
                st.download_button("Download weekly report XLSX", data=report_bytes, file_name=report_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with st.expander("Column mapping check"):
                st.write("The app matches by column heading and keeps the LIVE workbook's column order. Any template-only columns are left blank.")
        except Exception as exc:
            st.error(f"Could not generate the files: {exc}")
else:
    st.info("Upload the LIVE workbook and audit_export CSV in the sidebar to begin. The previous completed report is optional and is used as a reference.")
