from __future__ import annotations

import hashlib

import streamlit as st

from link_checker import output_filename, process_workbook, workbook_sheet_names


st.set_page_config(page_title="Internal Link Implementation Checker", page_icon="🔗", layout="wide")

st.title("Internal Link Implementation Checker")
st.caption(
    "Upload the Excel sheet to verify whether each suggested anchor is linked to the suggested URL "
    "inside the specified paragraph."
)

with st.expander("Required sheet format"):
    st.markdown(
        "- **Column C:** source page URL on a page header row; linking paragraph on the rows below\n"
        "- **Column D:** suggested anchor text\n"
        "- **Column E:** suggested hyperlink\n\n"
        "A row is marked **Implemented** only when the exact anchor is linked to the suggested URL "
        "within the specified paragraph."
    )

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded:
    workbook_bytes = uploaded.getvalue()
    upload_signature = hashlib.sha256(workbook_bytes).hexdigest()
    if st.session_state.get("active_upload") != upload_signature:
        for key in ("checker_output", "checker_results", "checker_filename"):
            st.session_state.pop(key, None)
        st.session_state["active_upload"] = upload_signature
    try:
        sheets = workbook_sheet_names(workbook_bytes)
    except Exception as exc:
        st.error(f"The workbook could not be opened: {exc}")
        st.stop()

    selected_sheet = st.selectbox("Sheet to check", sheets)
    workers = st.slider("Pages checked simultaneously", min_value=1, max_value=10, value=5)

    if st.button("Check internal links", type="primary", use_container_width=True):
        progress_bar = st.progress(0, text="Preparing checks…")

        def update_progress(done: int, total: int) -> None:
            progress_bar.progress(done / total, text=f"Checked {done} of {total} source pages")

        try:
            output_bytes, results = process_workbook(
                workbook_bytes,
                selected_sheet,
                max_workers=workers,
                progress=update_progress,
            )
        except Exception as exc:
            progress_bar.empty()
            st.error(str(exc))
        else:
            progress_bar.progress(1.0, text="Check complete")
            st.session_state["checker_output"] = output_bytes
            st.session_state["checker_results"] = results
            st.session_state["checker_filename"] = output_filename(uploaded.name)

if st.session_state.get("checker_results"):
    results = st.session_state["checker_results"]
    implemented = sum(row["Implementation Status"] == "Implemented" for row in results)
    not_implemented = len(results) - implemented

    col1, col2, col3 = st.columns(3)
    col1.metric("Suggestions checked", len(results))
    col2.metric("Implemented", implemented)
    col3.metric("Not Implemented", not_implemented)

    st.download_button(
        "Download checked Excel file",
        data=st.session_state["checker_output"],
        file_name=st.session_state["checker_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.subheader("Detailed results")
    visible_columns = [
        "Excel Row",
        "Source Page URL",
        "Suggested Anchor Text",
        "Suggested Hyperlink",
        "Implementation Status",
        "Actual Hyperlink Found",
        "Check Details",
        "Page HTTP Status",
    ]
    st.dataframe(
        [{key: row.get(key, "") for key in visible_columns} for row in results],
        use_container_width=True,
        hide_index=True,
    )
