    expected_sheet_names = list(wb.sheetnames)
    raw_xlsx_io = io.BytesIO()
    wb.save(raw_xlsx_io)
    raw_workbook_bytes = raw_xlsx_io.getvalue()
    # Verify that every original worksheet is still present in the raw workbook.
    check_wb = load_workbook(io.BytesIO(raw_workbook_bytes), read_only=True, data_only=False)
    missing_sheets = [name for name in expected_sheet_names if name not in check_wb.sheetnames]
    check_wb.close()
    if missing_sheets:
        raise ValueError("The raw workbook is missing worksheet(s): " + ", ".join(missing_sheets))

    # The final weekly report is processed separately so the raw workbook stays unchanged.
    xlsx_bytes = raw_workbook_bytes

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

    return raw_workbook_bytes, values_only_bytes, len(audit_df), len(mapping), len(template_headers), summary_values_only

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
3. Produces two Excel workbooks that retain all the original tabs: a refreshed raw workbook and the final weekly report.
4. Packages both workbooks into one ZIP download. The completed report is a reference only; last week's summary figures are not copied into the new report.
""")

if template_file and audit_file:
    if st.button("Generate both files", type="primary", use_container_width=True):
        try:
            ref_bytes = reference_file.getvalue() if reference_file else template_file.getvalue()
            raw_bytes, report_bytes, row_count, mapped_count, template_count, summary_values_only = build_outputs(
                template_file.getvalue(), audit_file.getvalue(), ref_bytes, report_day
            )
            report_name = f"Argos MS Weekly Report {report_day.strftime('%d.%m')}.xlsx"
            raw_name = f"Argos MS Weekly Report Raw Data {report_day.strftime('%d.%m')}.xlsx"
            zip_name = f"Argos MS Weekly Report Files {report_day.strftime('%d.%m')}.zip"
            bundle = io.BytesIO()
            with zipfile.ZipFile(bundle, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr(raw_name, raw_bytes)
                archive.writestr(report_name, report_bytes)

            st.session_state["argos_report_bundle"] = bundle.getvalue()
            st.session_state["argos_report_zip_name"] = zip_name
            st.session_state["argos_report_stats"] = (row_count, mapped_count, template_count, summary_values_only)
            st.success("Both workbooks have been generated.")
        except Exception as exc:
            st.error(f"Could not generate the files: {exc}")

# Keep the download control available after Streamlit reruns, rather than nesting it inside
# the Generate button event. A single ZIP download avoids browsers blocking multiple downloads.
if st.session_state.get("argos_report_bundle"):
    row_count, mapped_count, template_count, summary_values_only = st.session_state["argos_report_stats"]
    st.info(f"Processed {row_count:,} audit rows. Matched {mapped_count} of {template_count} template columns.")
    if summary_values_only:
        st.info("Summary formulas were recalculated and replaced with values.")
    else:
        st.warning("This environment could not recalculate Excel formulas automatically. The weekly report may still contain Summary formulas; open it in Excel and allow recalculation before sharing.")
    st.download_button(
        "Download both files (ZIP)",
        data=st.session_state["argos_report_bundle"],
        file_name=st.session_state["argos_report_zip_name"],
        mime="application/zip",
        type="primary",
        use_container_width=True,
        on_click="ignore",
    )
    with st.expander("Column mapping check"):
        st.write("The raw workbook is based on the LIVE workbook and retains its Raw Data, Summary, Pack Data, and any other existing tabs. Only the Raw Data sheet's contents are refreshed. The ZIP contains both the raw workbook and the final weekly report.")
else:
    st.info("Upload the LIVE workbook and audit_export CSV in the sidebar to begin. The previous completed report is optional and is used as a reference.")
